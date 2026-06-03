import streamlit as st
from streamlit_gsheets import GSheetsConnection
import pandas as pd
import io
from datetime import datetime, date

# ==========================================
# Configuration & Constants
# ==========================================
st.set_page_config(page_title="Vet Sales CRM", page_icon="🐾", layout="wide")

MASTER_COLUMNS = [
    "Doctor Name", "Area", "Clinic Location", "Medicine Name",
    "Quantity", "Total Invoice", "Amount Paid", "Remaining Balance",
    "Payment Status", "Last Visit Date", "Next Follow-up Date", "Notes"
]

HISTORY_COLUMNS = [
    "Timestamp", "Doctor Name", "Area", "Action Type", 
    "Invoice Added", "Amount Paid", "Current Balance", "Notes Log"
]

# إنشاء اتصال ذكي ومباشر بـ Google Sheets
conn = st.connection("gsheets", type=GSheetsConnection)

# ==========================================
# Helper Functions
# ==========================================
def load_sheet_data(worksheet_name, columns):
    """Loads data from a specific Google Sheet worksheet."""
    try:
        # ttl=0 تضمن عدم كاش البيانات وقراءتها حية في كل مرة
        df = conn.read(worksheet=worksheet_name, ttl=0)
        if df.empty or len(df.columns) == 0:
            return pd.DataFrame(columns=columns)
        # التأكد من نظافة البيانات من أي سطور فارغة بالكامل
        df = df.dropna(how='all')
        return df
    except Exception:
        return pd.DataFrame(columns=columns)

def save_sheet_data(df, worksheet_name):
    """Saves/Updates the DataFrame back to the Google Sheet worksheet."""
    conn.update(worksheet=worksheet_name, data=df)

def calculate_status(balance, amount_paid):
    if balance == 0:
        return "Fully Paid"
    elif balance < 0:
        return "Credit (مقدم)"
    elif amount_paid > 0:
        return "Partially Paid"
    else:
        return "Pending"

def log_to_history(doc_name, area, action_type, invoice, paid, current_balance, notes):
    """Appends a new immutable log entry to the history sheet."""
    history_df = load_sheet_data("History", HISTORY_COLUMNS)
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    new_log = {
        "Timestamp": now_str, "Doctor Name": doc_name, "Area": area, "Action Type": action_type,
        "Invoice Added": invoice, "Amount Paid": paid, "Current Balance": current_balance, "Notes Log": notes
    }
    history_df = pd.concat([history_df, pd.DataFrame([new_log])], ignore_index=True)
    save_sheet_data(history_df, "History")

def to_excel_bytes(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    return output.getvalue()

# ==========================================
# Main App Logic
# ==========================================
def main():
    st.title("🐾 Vet Sales CRM & Cloud Dashboard")
    
    # تحميل البيانات حية من جوجل شيت
    master_df = load_sheet_data("Master", MASTER_COLUMNS)
    history_df = load_sheet_data("History", HISTORY_COLUMNS)

    # ضبط أنواع البيانات الحسابية لمنع أخطاء الجمع
    for col in ["Quantity", "Total Invoice", "Amount Paid", "Remaining Balance"]:
        if not master_df.empty and col in master_df.columns:
            master_df[col] = pd.to_numeric(master_df[col], errors='coerce').fillna(0.0)

    # --- 1. KPI DASHBOARD ---
    total_sales = master_df["Total Invoice"].sum() if not master_df.empty else 0.0
    total_collected = master_df["Amount Paid"].sum() if not master_df.empty else 0.0
    total_debt = master_df["Remaining Balance"].sum() if not master_df.empty else 0.0

    st.markdown("### 📊 Financial Overview (Live from Cloud)")
    kpi1, kpi2, kpi3 = st.columns(3)
    kpi1.metric("Total Sales (EGP)", f"{total_sales:,.2f}")
    kpi2.metric("Total Collected (EGP)", f"{total_collected:,.2f}")
    kpi3.metric("Net Outstanding Balance (EGP)", f"{total_debt:,.2f}")
    
    st.divider()

    # --- 2. SIDEBAR ACTIONS ---
    st.sidebar.header("Actions & Data Entry")

    # Action A: Add or Update Record
    with st.sidebar.expander("➕ Add / Update Doctor Record"):
        with st.form("add_form", clear_on_submit=True):
            doc_name = st.text_input("Doctor's Name").strip()
            area = st.text_input("Area / Region").strip()
            location = st.text_input("Clinic Location (Google Maps URL)")
            medicine = st.text_input("Medicine Name / Product")
            qty = st.number_input("Quantity Sold (Boxes)", min_value=0, step=1)
            invoice = st.number_input("Invoice Amount", min_value=0.0, step=10.0)
            paid = st.number_input("Amount Paid", min_value=0.0, step=10.0)
            last_visit = st.date_input("Last Visit Date", value=date.today())
            next_visit = st.date_input("Next Follow-up Date")
            notes = st.text_area("Notes / Remarks")

            submit_new = st.form_submit_button("Save / Update Record")
            
            if submit_new:
                if not doc_name or not area:
                    st.error("يرجى إدخال اسم الدكتور والمنطقة كشرط أساسي!")
                else:
                    if not master_df.empty:
                        match_mask = (master_df["Doctor Name"].astype(str).str.strip().str.lower() == doc_name.lower()) & \
                                     (master_df["Area"].astype(str).str.strip().str.lower() == area.lower())
                    else:
                        match_mask = pd.Series([False])

                    if not master_df.empty and match_mask.any():
                        idx = master_df[match_mask].index[0]
                        master_df.loc[idx, "Total Invoice"] += invoice
                        master_df.loc[idx, "Amount Paid"] += paid
                        master_df.loc[idx, "Remaining Balance"] = master_df.loc[idx, "Total Invoice"] - master_df.loc[idx, "Amount Paid"]
                        master_df.loc[idx, "Payment Status"] = calculate_status(master_df.loc[idx, "Remaining Balance"], master_df.loc[idx, "Amount Paid"])
                        master_df.loc[idx, "Quantity"] += qty
                        
                        if medicine:
                            existing_med = str(master_df.loc[idx, "Medicine Name"])
                            master_df.loc[idx, "Medicine Name"] = existing_med + f", {medicine}" if existing_med and existing_med != "nan" else medicine
                        
                        master_df.loc[idx, "Last Visit Date"] = str(last_visit)
                        master_df.loc[idx, "Next Follow-up Date"] = str(next_visit)
                        if location: master_df.loc[idx, "Clinic Location"] = location
                        if notes:
                            existing_notes = str(master_df.loc[idx, "Notes"])
                            master_df.loc[idx, "Notes"] = existing_notes + f" | [{last_visit}]: {notes}" if existing_notes and existing_notes != "nan" else f"[{last_visit}]: {notes}"
                        
                        current_bal = master_df.loc[idx, "Remaining Balance"]
                        save_sheet_data(master_df, "Master")
                        st.success(f"تم تحديث الحساب على جوجل شيت للدكتور {doc_name}!")
                        log_to_history(doc_name, area, "Order / Visit Update", invoice, paid, current_bal, notes)
                    else:
                        balance = invoice - paid
                        status = calculate_status(balance, paid)
                        new_row = {
                            "Doctor Name": doc_name, "Area": area, "Clinic Location": location,
                            "Medicine Name": medicine, "Quantity": qty, "Total Invoice": invoice,
                            "Amount Paid": paid, "Remaining Balance": balance, "Payment Status": status,
                            "Last Visit Date": str(last_visit), "Next Follow-up Date": str(next_visit), 
                            "Notes": f"[{last_visit}]: {notes}" if notes else ""
                        }
                        master_df = pd.concat([master_df, pd.DataFrame([new_row])], ignore_index=True)
                        save_sheet_data(master_df, "Master")
                        st.success(f"تم تسجيل الدكتور {doc_name} في السحابة بنجاح!")
                        log_to_history(doc_name, area, "New Client Registration", invoice, paid, balance, notes)
                    
                    st.rerun()

    # Action B: Quick Payment Update
    with st.sidebar.expander("💵 Quick Payment (تحصيل سريع)"):
        if not master_df.empty:
            options = master_df.index.tolist()
            format_func = lambda i: f"{master_df.loc[i, 'Doctor Name']} ({master_df.loc[i, 'Area']}) | Balance: {master_df.loc[i, 'Remaining Balance']} EGP"
            selected_idx = st.selectbox("Select Doctor", options, format_func=format_func)
            add_payment = st.number_input("Amount to Pay", min_value=0.0, step=10.0)
            
            if st.button("Apply Quick Payment"):
                if add_payment > 0:
                    doc_name = master_df.loc[selected_idx, "Doctor Name"]
                    area = master_df.loc[selected_idx, "Area"]
                    master_df.loc[selected_idx, "Amount Paid"] += add_payment
                    master_df.loc[selected_idx, "Remaining Balance"] = master_df.loc[selected_idx, "Total Invoice"] - master_df.loc[selected_idx, "Amount Paid"]
                    master_df.loc[selected_idx, "Payment Status"] = calculate_status(master_df.loc[selected_idx, "Remaining Balance"], master_df.loc[selected_idx, "Amount Paid"])
                    master_df.loc[selected_idx, "Last Visit Date"] = str(date.today())
                    
                    existing_notes = str(master_df.loc[selected_idx, "Notes"])
                    pay_note = f"Quick Payment of {add_payment} EGP received."
                    master_df.loc[selected_idx, "Notes"] = existing_notes + f" | [{date.today()}]: {pay_note}" if existing_notes and existing_notes != "nan" else f"[{date.today()}]: {pay_note}"
                    
                    current_bal = master_df.loc[selected_idx, "Remaining Balance"]
                    save_sheet_data(master_df, "Master")
                    st.success("تم تحديث الدفعة على جوجل شيت الحقيقي!")
                    log_to_history(doc_name, area, "Quick Payment Collection", 0.0, add_payment, current_bal, pay_note)
                    st.rerun()
                else:
                    st.warning("يرجى إدخال مبلغ أكبر من 0.")
        else:
            st.info("لا توجد بيانات حالياً.")

    # Action C: Export Downloads
    st.sidebar.divider()
    st.sidebar.markdown("### 📥 تحميل البيانات الاحتياطية (Excel)")
    if not master_df.empty:
        st.sidebar.download_button(label="📊 تحميل الحسابات الموحدة (Excel)", data=to_excel_bytes(master_df), file_name='vet_sales_master.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', use_container_width=True)
    if not history_df.empty:
        st.sidebar.download_button(label="📜 تحميل سجل العمليات الكامل (Excel)", data=to_excel_bytes(history_df), file_name='vet_sales_history.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', use_container_width=True)

    # --- 3. TABS PRESENTATION ---
    tab1, tab2, tab3 = st.tabs([
        "📋 قاعدة البيانات الحالية (Master)", 
        "📜 سجل الحركات التاريخي (History Log)", 
        "📊 داشبورد والتحليلات (Dashboard)"
    ])

    with tab1:
        st.markdown("### 🔍 البحث والتصفية")
        col_f1, col_f2 = st.columns(2)
        search_doc = col_f1.text_input("البحث باسم الدكتور")
        search_area = col_f2.text_input("البحث بالمنطقة")
        
        filtered_master = master_df.copy()
        if search_doc and not filtered_master.empty:
            filtered_master = filtered_master[filtered_master["Doctor Name"].astype(str).str.contains(search_doc, case=False, na=False)]
        if search_area and not filtered_master.empty:
            filtered_master = filtered_master[filtered_master["Area"].astype(str).str.contains(search_area, case=False, na=False)]
            
        if not filtered_master.empty:
            st.dataframe(
                filtered_master,
                column_config={
                    "Clinic Location": st.column_config.LinkColumn("Clinic Location", display_text="Open Map"),
                    "Total Invoice": st.column_config.NumberColumn("Total Invoice", format="%.2f EGP"),
                    "Amount Paid": st.column_config.NumberColumn("Amount Paid", format="%.2f EGP"),
                    "Remaining Balance": st.column_config.NumberColumn("Remaining Balance", format="%.2f EGP"),
                },
                use_container_width=True, hide_index=True
            )
        else:
            st.info("لا توجد بيانات عملاء حالياً.")

    with tab2:
        st.markdown("### ⏱️ جميع المعاملات بترتيب حدوثها")
        if not history_df.empty:
            st.dataframe(
                history_df.iloc[::-1], 
                column_config={
                    "Invoice Added": st.column_config.NumberColumn("Invoice Added", format="%.2f EGP"),
                    "Amount Paid": st.column_config.NumberColumn("Amount Paid", format="%.2f EGP"),
                    "Current Balance": st.column_config.NumberColumn("Current Balance", format="%.2f EGP"),
                },
                use_container_width=True, hide_index=True
            )
        else:
            st.info("لم يتم تسجيل أي حركات تاريخية بعد.")

    with tab3:
        st.markdown("### 📈 تحليلات إحصائية حية")
        if not master_df.empty:
            col_d1, col_d2 = st.columns(2)
            with col_d1:
                st.subheader("💰 إجمالي المبيعات حسب المنطقة")
                area_sales = master_df.groupby("Area")["Total Invoice"].sum()
                st.bar_chart(area_sales)
            with col_d2:
                st.subheader("👨‍⚕️ أعلى 5 دكاترة مبيعاً وسحباً")
                top_docs = master_df.groupby("Doctor Name")["Total Invoice"].sum().sort_values(ascending=False).head(5)
                st.bar_chart(top_docs)
            st.divider()
            st.subheader("📊 توزيع العملاء بناءً على حالة الدفع المالي")
            status_counts = master_df["Payment Status"].value_counts()
            st.bar_chart(status_counts)
        else:
            st.info("أدخل بعض البيانات أولاً لتظهر لك التحليلات والرسوم البيانية هنا تلقائياً!")

if __name__ == "__main__":
    main()
