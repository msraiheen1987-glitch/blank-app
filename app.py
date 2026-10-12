import streamlit as st
import pandas as pd
import datetime

# إعداد الصفحة وتكوين العرض
st.set_page_config(
    page_title="نظام متابعة عمليات المشتريات",
    page_icon="📊",
    layout="wide"
)

# دعم اللغة العربية والاتجاه من اليمين لليسار
st.markdown("""
    <style>
    body {
        direction: rtl;
        text-align: right;
    }
    .stDataFrame {
        direction: rtl;
    }
    h1, h2, h3, h4, h5, h6, p, label, span {
        direction: rtl;
        text-align: right;
    }
    </style>
""", unsafe_allow_html=True)

EXCEL_FILE = 'Copy of جدول سير عمليات المشتريات.xlsx'

@st.cache_data(ttl=60)
def load_data():
    try:
        xls = pd.ExcelFile(EXCEL_FILE)
        sheets = xls.sheet_names
        data = {}
        for sheet in sheets:
            data[sheet] = pd.read_excel(EXCEL_FILE, sheet_name=sheet)
        return data, sheets
    except Exception as e:
        st.error(f"خطأ أثناء قراءة ملف الأكسيل: {e}")
        return None, None

data_dict, sheets = load_data()

if data_dict:
    st.title("📋 نظام إدارة ومتابعة عمليات المشتريات")
    st.markdown("نظام مركزي لتتبع وعرض وتحديث عمليات المشتريات ولجان الشراء ومشاركتها مع فريق العمل.")

    # الشريط الجانبي للتنقل
    st.sidebar.header("📁 لوحة التحكم")
    selected_sheet = st.sidebar.selectbox("اختر القسم / الشيت:", [s for s in sheets if 'تحليل' not in s])
    
    df = data_dict[selected_sheet].copy()

    # شريط البحث والتصفية
    st.sidebar.markdown("---")
    st.sidebar.subheader("🔍 البحث والتصفية")
    search_query = st.sidebar.text_input("بحث برقم المعاملة أو اسم الشركة أو الموضوع:")
    
    if 'حالة المعاملة' in df.columns:
        statuses = ['الكل'] + list(df['حالة المعاملة'].dropna().unique())
        selected_status = st.sidebar.selectbox("تصفية حسب حالة المعاملة:", statuses)
        if selected_status != 'الكل':
            df = df[df['حالة المعاملة'] == selected_status]

    if search_query:
        # البحث في كافة الأعمدة النصية
        mask = df.astype(str).apply(lambda x: x.str.contains(search_query, case=False, na=False)).any(axis=1)
        df = df[mask]

    # عرض المؤشرات الإحصائية السريعة
    st.markdown("---")
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("إجمالي المعاملات المعروضة", len(df))
    with col2:
        if 'التكلفة التقديرية' in df.columns:
            total_est = df['التكلفة التقديرية'].sum()
            st.metric("إجمالي التكاليف التقديرية", f"{total_est:,.2f} دينار")
    with col3:
        if 'قيمة الإحالة' in df.columns:
            total_award = df['قيمة الإحالة'].sum()
            st.metric("إجمالي قيم الإحالة", f"{total_award:,.2f} دينار")

    st.markdown("---")
    
    # تبويبات النظام: عرض البيانات، التعديل، إضافة معاملة جديدة
    tab1, tab2, tab3 = st.tabs(["📂 استعراض وتحديث البيانات", "➕ إضافة معاملة جديدة", "📊 التحليلات والتقارير"])

    with tab1:
        st.subheader(f"بيانات: {selected_sheet}")
        st.info("💡 يمكنك تعديل الخلايا مباشرة في الجدول أدناه، ثم الضغط على زر حفظ التغييرات.")
        
        # تحرير البيانات باستخدام DataEditor
        edited_df = st.data_editor(df, num_rows="dynamic", use_container_width=True, key=f"editor_{selected_sheet}")

        if st.button("💾 حفظ التغييرات وتحديث الملف"):
            try:
                # تحديث الشيت المحدد في القاموس وحفظه لملف الأكسيل
                data_dict[selected_sheet] = edited_df
                with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl') as writer:
                    for s_name, s_df in data_dict.items():
                        s_df.to_excel(writer, sheet_name=s_name, index=False)
                st.success("✨ تم حفظ التغييرات وتحديث ملف الأكسيل بنجاح!")
            except Exception as e:
                st.error(f"حدث خطأ أثناء الحفظ: {e}")

        # زر تحميل الملف المحدث
        with open(EXCEL_FILE, "rb") as file:
            st.download_button(
                label="📥 تحميل ملف الأكسيل المحدث",
                data=file,
                file_name="جدول_سير_عمليات_المشتريات_المحدث.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    with tab2:
        st.subheader("➕ إضافة معاملة جديدة")
        with st.form(key=f"add_form_{selected_sheet}"):
            st.markdown("أدخل تفاصيل المعاملة الجديدة:")
            new_row = {}
            cols = st.columns(2)
            for idx, col_name in enumerate(df.columns):
                with cols[idx % 2]:
                    if 'تاريخ' in col_name:
                        new_row[col_name] = st.date_input(col_name, value=datetime.date.today())
                    elif 'التكلفة' in col_name or 'قيمة' in col_name or col_name == 'الرقم':
                        new_row[col_name] = st.number_input(col_name, value=0.0)
                    else:
                        new_row[col_name] = st.text_input(col_name, value="")
            
            submit_button = st.form_submit_button(label="إضافة المعاملة")
            if submit_button:
                new_df_row = pd.DataFrame([new_row])
                data_dict[selected_sheet] = pd.concat([data_dict[selected_sheet], new_df_row], ignore_index=True)
                with pd.ExcelWriter(EXCEL_FILE, engine='openpyxl') as writer:
                    for s_name, s_df in data_dict.items():
                        s_df.to_excel(writer, sheet_name=s_name, index=False)
                st.success("✅ تمت إضافة المعاملة بنجاح وتحديث الجدول!")
                st.rerun()

    with tab3:
        st.subheader("📊 لوحة المؤشرات التحليلية")
        if 'حالة المعاملة' in df.columns:
            status_counts = df['حالة المعاملة'].value_counts()
            st.markdown("##### توزيع المعاملات حسب الحالة:")
            st.bar_chart(status_counts)
        
        if 'الجهة' in df.columns:
            st.markdown("##### التوزيع حسب الجهة:")
            st.bar_chart(df['الجهة'].value_counts())
            
        # عرض شيتات التحليل الأصلية إن وجدت
        for s in sheets:
            if 'تحليل' in s:
                st.markdown(f"---")
                st.markdown(f"##### تقرير: {s}")
                st.dataframe(data_dict[s], use_container_width=True)

else:
    st.warning("الرجاء التأكد من وجود ملف الأكسيل 'Copy of جدول سير عمليات المشتريات.xlsx' في نفس مجلد التطبيق.")