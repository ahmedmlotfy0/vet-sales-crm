import streamlit as st
from streamlit_gsheets import GSheetsConnection
import pandas as pd
import io
from datetime import datetime, date

# ==========================================
# Configuration & Constants
# ==========================================
st.set_page_config(page_title="Vet Sales CRM", page_icon="🐾", layout="wide")

# تم إضافة عمود "Governorate" في الترتيب الثاني لقاعدة البيانات
MASTER_COLUMNS = [
    "Doctor Name", "Governorate", "Area", "Clinic Location", "Medicine Name",
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
def normalize_arabic(text):
    """تنظيف وتوحيد النصوص العربية لمنع التكرار بسبب مسافات أو حروف متشابهة"""
    if pd.isna(text) or not isinstance(text, str):
        return ""
    text = text.strip().lower()
    text = text.replace("ى", "ي")
    text = text.replace("ة", "ه")
    return text

def load_sheet_data(worksheet_name, columns):
    """Loads data from a specific Google Sheet worksheet."""
    try:
        df = conn.read(worksheet=worksheet_name, ttl=0)
        if df.empty or len(df.columns) == 0:
            return pd.DataFrame(columns=columns)
        df = df.dropna(how='all')
        
        # إجبار الحقول الحسابية على التحول لأرقام نقية لتجنب أخطاء السالب والموجب
        for col in ["Quantity", "Total Invoice", "Amount Paid", "Remaining Balance"]:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0.0).astype(float)
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
        "Invoice Added": float(invoice), "Amount Paid": float(paid), "Current Balance": float(current_balance), "Notes Log": notes
    }
    history_df = pd.concat([history_df, pd.DataFrame([new_log])], ignore_index=True)
    save_sheet_data(history_df, "History")

def to_excel_bytes(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    return output.getvalue()

def style_financial_rows(df):
    """Applies conditional coloring based on the Remaining Balance column."""
    def get_row_colors(row):
        try:
            bal = float(row["Remaining Balance"])
        except (ValueError, TypeError):
            bal = 0.0
            
        if bal > 0:
            return ['background-color: #f8d7da; color: #721c24; font-weight: 500;'] * len(row)
        elif bal < 0:
            return ['background-color: #d4edda; color: #155724; font-weight: 500;'] * len(row)
        else:
            return [''] * len(row)
            
    return df.style.apply(get_row_colors, axis=1)

# ==========================================
# Main App Logic
# ==========================================
def main():
    st.title("🐾 Vet Sales CRM & Cloud Dashboard")
    
    # تحميل البيانات حية من جوجل شيت
    master_df = load_sheet_data("Master", MASTER_COLUMNS)
    history_df = load_sheet_data("History", HISTORY_COLUMNS)

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
            governorate = st.selectbox("المحافظة", ["دقهلية", "دمياط"])
            area = st.text_input("Area / Region").strip()
            location = st.text_input("Clinic Location (Google Maps URL)").strip()
            medicine = st.text_input("Medicine Name / Product").strip()
            qty = st.number_input("Quantity Sold (Boxes)", min_value=0, step=1)
            invoice = st.number_input("Invoice Amount", min_value=0.0, step=10.0)
            paid = st.number_input("Amount Paid", min_value=0.0, step=10.0)
            last_visit = st.date_input("Last Visit Date", value=date.today())
            next_visit = st.date_input("Next Follow-up Date")
            notes = st.text_area("Notes / Remarks").strip()

            submit_new = st.form_submit_button("Save / Update Record")
            
            if submit_new:
                if not doc_name or not area:
                    st.error("يرجى إدخال اسم الدكتور والمنطقة كشرط أساسي!")
                else:
                    # تطبيق الفلترة والتوحيد الذكي للحروف لمطابقة السجلات بدقة وثقة
                    norm_input_doc = normalize_arabic(doc_name)
                    norm_input_area = normalize_arabic(area)
                    norm_input_gov = normalize_arabic(governorate)

                    if not master_df.empty:
                        # جلب عمود المحافظة القديم إن لم يكن موجوداً لتجنب الأخطاء
                        if "Governorate" not in master_df.columns:
                            master_df["Governorate"] = governorate
                            
                        match_mask = (master_df["Doctor Name"].astype(str).apply(normalize_arabic) == norm_input_doc) & \
                                     (master_df["Area"].astype(str).apply(normalize_arabic) == norm_input_area) & \
                                     (master_df["Governorate"].astype(str).apply(normalize_arabic) == norm_input_gov)
                    else:
                        match_mask = pd.Series([False])

                    if not master_df.empty and match_mask.any():
                        idx = master_df[match_mask].index[0]
                        master_df.loc[idx, "Total Invoice"] = float(master_df.loc[idx, "Total Invoice"]) + float(invoice)
                        master_df.loc[idx, "Amount Paid"] = float(master_df.loc[idx, "Amount Paid"]) + float(paid)
                        master_df.loc[idx, "Remaining Balance"] = float(master_df.loc[idx, "Total Invoice"]) - float(master_df.loc[idx, "Amount Paid"])
                        master_df.loc[idx, "Payment Status"] = calculate_status(master_df.loc[idx, "Remaining Balance"], master_df.loc[idx, "Amount Paid"])
                        master_df.loc[idx, "Quantity"] = int(master_df.loc[idx, "Quantity"]) + int(qty)
                        
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
                        st.success(f"تم تحديث الحساب الموحد للدكتور {doc_name}!")
                        log_to_history(doc_name, area, "Order / Visit Update", invoice, paid, current_bal, notes)
                    else:
                        balance = float(invoice) - float(paid)
                        status = calculate_status(balance, paid)
                        new_row = {
                            "Doctor Name": doc_name, "Governorate": governorate, "Area": area, "Clinic Location": location,
                            "Medicine Name": medicine, "Quantity": int(qty), "Total Invoice": float(invoice),
                            "Amount Paid": float(paid), "Remaining Balance": float(balance), "Payment Status": status,
                            "Last Visit Date": str(last_visit), "Next Follow-up Date": str(next_visit), 
                            "Notes": f"[{last_visit}]: {notes}" if notes else ""
                        }
                        master_df = pd.concat([master_df, pd.DataFrame([new_row])], ignore_index=True)
                        save_sheet_data(master_df, "Master")
                        st.success(f"تم تسجيل الدكتور {doc_name} في السحابة بنجاح!")
                        log_to_history(doc_name, area, "New Client Registration", invoice, paid, balance, notes)
                    
                    st.rerun()

    # Action B: Quick Payment Update (تم الفلترة لإظهار المديونين فقط واحتواء المبالغ الزائدة)
    with st.sidebar.expander("💵 Quick Payment (تحصيل سريع)"):
        if not master_df.empty:
            # تصفية الجدول لإظهار من رصيده المتبقي أكبر من صفر فقط (عليه ديون)
            pending_debtors = master_df[master_df["Remaining Balance"] > 0]
            
            if not pending_debtors.empty:
                options = pending_debtors.index.tolist()
                format_func = lambda i: f"{master_df.loc[i, 'Doctor Name']} ({master_df.loc[i, 'Area']}) | Owes: {master_df.loc[i, 'Remaining Balance']} EGP"
                selected_idx = st.selectbox("Select Doctor", options, format_func=format_func)
                add_payment = st.number_input("Amount to Pay", min_value=0.0, step=10.0)
                
                if st.button("Apply Quick Payment"):
                    if add_payment > 0:
                        doc_name = master_df.loc[selected_idx, "Doctor Name"]
                        area = master_df.loc[selected_idx, "Area"]
                        
                        # حسابات دقيقة تقبل التحول لرصيد مقدم (سالب) بدون إخراج أي أخطاء
                        master_df.loc[selected_idx, "Amount Paid"] = float(master_df.loc[selected_idx, "Amount Paid"]) + float(add_payment)
                        master_df.loc[selected_idx, "Remaining Balance"] = float(master_df.loc[selected_idx, "Total Invoice"]) - float(master_df.loc[selected_idx, "Amount Paid"])
                        master_df.loc[selected_idx, "Payment Status"] = calculate_status(master_df.loc[selected_idx, "Remaining Balance"], master_df.loc[selected_idx, "Amount Paid"])
                        master_df.loc[selected_idx, "Last Visit Date"] = str(date.today())
                        
                        existing_notes = str(master_df.loc[selected_idx, "Notes"])
                        pay_note = f"Quick Payment of {add_payment} EGP received."
                        master_df.loc[selected_idx, "Notes"] = existing_notes + f" | [{date.today()}]: {pay_note}" if existing_notes and existing_notes != "nan" else f"[{date.today()}]: {pay_note}"
                        
                        current_bal = master_df.loc[selected_idx, "Remaining Balance"]
                        save_sheet_data(master_df, "Master")
                        st.success("تم تحديث الدفعة المالية بنجاح على جوجل شيت!")
                        log_to_history(doc_name, area, "Quick Payment Collection", 0.0, add_payment, current_bal, pay_note)
                        st.rerun()
                    else:
                        st.warning("يرجى إدخال مبلغ أكبر من 0.")
            else:
                st.info("🎉 لا يوجد أي دكاترة عليهم مديونيات حالياً! مبيعاتك كلها محصلة.")
        else:
            st.info("لا توجد بيانات حالياً.")

    # Action C: Delete Doctor Record
    with st.sidebar.expander("❌ Delete Doctor Record (حذف عميل بالكامل)"):
        if not master_df.empty:
            options_del = master_df.index.tolist()
            format_func_del = lambda i: f"{master_df.loc[i, 'Doctor Name']} ({master_df.loc[i, 'Area']})"
            selected_del_idx = st.selectbox("Select Doctor to Delete", options_del, format_func=format_func_del, key="del_doc_select")
            
            confirm_del = st.checkbox("أنا متأكد من رغبتي في حذف هذا الدكتور وحسابه تماماً", key="confirm_del_check")
            
            if st.button("Delete Permanently (حذف نهائي)", type="primary"):
                if confirm_del:
                    doc_name_del = master_df.loc[selected_del_idx, "Doctor Name"]
                    area_del = master_df.loc[selected_del_idx, "Area"]
                    
                    master_df = master_df.drop(selected_del_idx).reset_index(drop=True)
                    save_sheet_data(master_df, "Master")
                    
                    log_to_history(doc_name_del, area_del, "Client Record Deleted", 0.0, 0.0, 0.0, "تم حذف سجل الدكتور بالكامل من جدول الحسابات الرئيسي")
                    st.success(f"تم حذف الدكتور {doc_name_del} بنجاح!")
                    st.rerun()
                else:
                    st.warning("برجاء تفعيل مربع التأكيد أولاً لتفعيل زر الحذف نهائياً.")
        else:
            st.info("قاعدة البيانات فارغة، لا يوجد دكاترة لحذفهم.")

    # Action D: Export Downloads
    st.sidebar.divider()
    st.sidebar.markdown("### 📥 تحميل البيانات الاحتياطية (Excel)")
    if not master_df.empty:
        st.sidebar.download_button(label="📊 تحميل الحسابات الموحدة (Excel)", data=to_excel_bytes(master_df), file_name='vet_sales_master.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', use_container_width=True)
    if not history_df.empty:
        st.sidebar.download_button(label="📜 تحميل سجل العمليات الكامل (Excel)", data=to_excel_bytes(history_df), file_name='vet_sales_history.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', use_container_width=True)

    # --- 3. TABS PRESENTATION ---
    tab1, tab2, tab3, tab4 = st.tabs([
        "📋 قاعدة البيانات الحالية (Master)", 
        "🔔 متابعات اليوم (Today's Follow-ups)",
        "📜 سجل الحركات التاريخي (History Log)", 
        "📊 داشبورد والتحليلات (Dashboard)"
    ])

    with tab1:
        st.markdown("### 🔍 البحث والتصفية")
        col_f1, col_f2 = st.columns(2)
        search_doc = col_f1.text_input("البحث باسم الدكتور")
        search_area = col_f2.text_input("البحث بالمنطقة أو المحافظة")
        
        filtered_master = master_df.copy()
        if search_doc and not filtered_master.empty:
            filtered_master = filtered_master[filtered_master["Doctor Name"].astype(str).str.contains(search_doc, case=False, na=False)]
        if search_area and not filtered_master.empty:
            # تصفية البحث لتشمل المنطقة أو المحافظة معاً لراحة يدك في السوق
            area_mask = filtered_master["Area"].astype(str).str.contains(search_area, case=False, na=False)
            gov_mask = filtered_master["Governorate"].astype(str).str.contains(search_area, case=False, na=False) if "Governorate" in filtered_master.columns else False
            filtered_master = filtered_master[area_mask | gov_mask]
            
        if not filtered_master.empty:
            styled_df = style_financial_rows(filtered_master)
            st.dataframe(
                styled_df,
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
        st.markdown(f"### 📅 دكاترة مطلوب متابعتهم اليوم ({date.today()})")
        today_str = str(date.today())
        
        if not master_df.empty:
            follow_ups_today = master_df[
                (master_df["Next Follow-up Date"].astype(str) == today_str) & 
                (master_df["Next Follow-up Date"].astype(str) != master_df["Last Visit Date"].astype(str))
            ]
        else:
            follow_ups_today = pd.DataFrame()
            
        if not follow_ups_today.empty:
            st.warning(f"⚠️ لديك {len(follow_ups_today)} متابعة مطلوبة اليوم! يرجى التواصل معهم:")
            styled_follow_up = style_financial_rows(follow_ups_today)
            st.dataframe(
                styled_follow_up,
                column_config={
                    "Clinic Location": st.column_config.LinkColumn("Clinic Location", display_text="Open Map"),
                    "Total Invoice": st.column_config.NumberColumn("Total Invoice", format="%.2f EGP"),
                    "Amount Paid": st.column_config.NumberColumn("Amount Paid", format="%.2f EGP"),
                    "Remaining Balance": st.column_config.NumberColumn("Remaining Balance", format="%.2f EGP"),
                },
                use_container_width=True, hide_index=True
            )
        else:
            st.success("🎉 لا توجد أي متابعات مجدولة لهذا اليوم! كل دكاترتك مئة بالمئة.")

    with tab3:
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

    with tab4:
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
