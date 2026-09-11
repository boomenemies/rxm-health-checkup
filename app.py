import streamlit as st
import pandas as pd
import io
import plotly.express as px
import re

st.set_page_config(page_title="ระบบสรุปผลการตรวจสุขภาพ 2025", layout="wide")
st.title("🏥 ระบบจัดการข้อมูลตรวจสุขภาพพนักงาน 2025")

# --- ข้อมูลแม่พิมพ์รายการตรวจและราคา (อ้างอิงตามใบเสนอราคา BPK 8) ---
TEST_MAPPING = {
    'W/H,BP, BMI': ['Pro.1', 'Pro.2', 'Pro.3'], 
    'PE': ['Pro.1', 'Pro.2', 'Pro.3'],
    'X-RAY Digital': ['Pro.1', 'Pro.2', 'Pro.3'], 
    'CBC': ['Pro.1', 'Pro.2', 'Pro.3'],
    'UA': ['Pro.1', 'Pro.2', 'Pro.3'], 
    'FBS': ['Pro.1', 'Pro.2', 'Pro.3'],
    'BUN/Creatinine': ['Pro.1', 'Pro.2', 'Pro.3'], 
    'URIC ACID': ['Pro.1', 'Pro.2', 'Pro.3'],
    'CHOLESTEROL': ['Pro.1', 'Pro.2', 'Pro.3'], 
    'TRIGLYCERIDE': ['Pro.1', 'Pro.2', 'Pro.3'],
    'HDL': ['Pro.1', 'Pro.2', 'Pro.3'], 
    'LDL': ['Pro.1', 'Pro.2', 'Pro.3'],
    'SGOT/SGPT': ['Pro.1', 'Pro.2', 'Pro.3'], 
    'Hbs Ag Elisa': ['Pro.1', 'Pro.2', 'Pro.3'],
    'EKG': ['Pro.2', 'Pro.3'], 
    'AFP': ['Pro.3'],
    'ALK.PHOSE': ['Pro.3'], 
    'VISION TEST': ['Pro.1', 'Pro.2', 'Pro.3']
}

PRICE_MAPPING = {
    'W/H,BP, BMI': [0, 0, 0], 'PE': [40, 40, 40], 'X-RAY Digital': [80, 80, 80],
    'CBC': [30, 30, 0], 'UA': [20, 20, 0], 'FBS': [20, 20, 0],
    'BUN/Creatinine': [40, 40, 0], 'URIC ACID': [30, 30, 0],
    'CHOLESTEROL': [20, 20, 0], 'TRIGLYCERIDE': [20, 20, 0],
    'HDL': [40, 40, 0], 'LDL': [40, 40, 0], 'SGOT/SGPT': [40, 40, 0],
    'Hbs Ag Elisa': [80, 80, 80], 'EKG': [0, 150, 150],
    'AFP': [0, 0, 150], 'ALK.PHOSE': [0, 0, 0], 'VISION TEST': [0, 0, 0]
}

def export_full_excel(df_main, df_count, df_price, filename):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df_main.to_excel(writer, index=False, sheet_name='Checklist รายบุคคล')
        df_count.to_excel(writer, index=False, sheet_name='สรุปจำนวนรายการตรวจ')
        df_price.to_excel(writer, index=False, sheet_name='สรุปราคาแยกโปรแกรม')
    return output.getvalue()

uploaded_file = st.file_uploader("อัปโหลดไฟล์รายชื่อพนักงาน (Excel)", type=["xlsx"])

if uploaded_file is not None:
    df_raw = pd.read_excel(uploaded_file)
    df_raw.columns = df_raw.columns.str.strip()
    
    df_raw['หมายเหตุ'] = df_raw.get('หมายเหตุ', '').fillna('')
    fail_probation_count = df_raw['หมายเหตุ'].str.contains('ไม่ผ่าน|ทดลองงาน').sum()
    df_filtered = df_raw[~df_raw['หมายเหตุ'].str.contains('ไม่ผ่าน|ทดลองงาน')].copy()
    
    def assign_program(row):
        level = str(row.get('ระดับนักงาน', '')).strip()
        age_match = re.search(r'\d+', str(row.get('อายุ', '0')))
        age = int(age_match.group()) if age_match else 0
        is_pro3 = any(kw in level for kw in ['ผู้จัดการ', 'ผู้บริหาร', 'หัวหน้างาน', 'เภสัชกร', 'วิศวกร'])
        
        if is_pro3: return pd.Series(['Pro.3', 500])
        elif age >= 35: return pd.Series(['Pro.2', 650])
        else: return pd.Series(['Pro.1', 500])

    df_filtered[['โปรแกรม', 'ราคาพื้นฐาน']] = df_filtered.apply(assign_program, axis=1)
    if 'ใบรับรองแพทย์ 5 โรค' not in df_filtered.columns:
        df_filtered['ใบรับรองแพทย์ 5 โรค'] = False

    # สร้างเครื่องหมาย ✓ สำหรับรายการที่ต้องตรวจ
    for test, programs in TEST_MAPPING.items():
        df_filtered[test] = df_filtered['โปรแกรม'].apply(lambda p: '✓' if p in programs else '-')

    base_cols = ['รหัสพนักงาน', 'ชื่อ - นามสกุล', 'ตำแหน่ง', 'หน่วยงาน', 'อายุ', 'โปรแกรม', 'ราคาพื้นฐาน', 'ใบรับรองแพทย์ 5 โรค']
    test_cols = list(TEST_MAPPING.keys())
    display_cols = [c for c in base_cols + test_cols if c in df_filtered.columns]
    df_display = df_filtered[display_cols].copy()

    tab1, tab2, tab3 = st.tabs(["📋 ตารางแจกแจงรายการตรวจ", "🏢 แยกตามหน่วยงาน", "📊 Dashboard & สรุปราคา"])
    
    with tab1:
        st.subheader("ตารางแจกแจงรายการตรวจรายบุคคล")
        st.markdown("*(ติ๊กเลือกช่อง 'ใบรับรองแพทย์ 5 โรค' เพื่อคำนวณราคาเพิ่มได้)*")
        
        edited_df = st.data_editor(
            df_display, use_container_width=True, hide_index=True,
            disabled=[c for c in display_cols if c != 'ใบรับรองแพทย์ 5 โรค']
        )
        edited_df['ราคารวม (บาท)'] = edited_df['ราคาพื้นฐาน'] + (edited_df['ใบรับรองแพทย์ 5 โรค'] * 100)

    count_p1 = (edited_df['โปรแกรม'] == 'Pro.1').sum()
    count_p2 = (edited_df['โปรแกรม'] == 'Pro.2').sum()
    count_p3 = (edited_df['โปรแกรม'] == 'Pro.3').sum()
    count_med = edited_df['ใบรับรองแพทย์ 5 โรค'].sum()
    total_emp = len(edited_df)

    count_data = []
    for test, programs in TEST_MAPPING.items():
        total_test = 0
        if 'Pro.1' in programs: total_test += count_p1
        if 'Pro.2' in programs: total_test += count_p2
        if 'Pro.3' in programs: total_test += count_p3
        count_data.append({'รายการตรวจสุขภาพ': test, 'จำนวนคน': total_test})
    count_data.append({'รายการตรวจสุขภาพ': 'ใบรับรองแพทย์ 5 โรค', 'จำนวนคน': count_med})
    df_summary_count = pd.DataFrame(count_data)

    price_data = []
    for test, prices in PRICE_MAPPING.items():
        price_data.append({
            'รายการตรวจสุขภาพ': test, 
            'Pro.1 (<35)': prices[0] if prices[0] > 0 else 'ฟรี',
            'Pro.2 (>=35)': prices[1] if prices[1] > 0 else 'ฟรี',
            'Pro.3 (บริหาร)': prices[2] if prices[2] > 0 else 'ฟรี'
        })
    price_data.append({'รายการตรวจสุขภาพ': 'ใบรับรองแพทย์ 5 โรค', 'Pro.1 (<35)': 100, 'Pro.2 (>=35)': 100, 'Pro.3 (บริหาร)': 100})
    price_data.append({'รายการตรวจสุขภาพ': '--- ราคาเหมาจ่าย/คน ---', 'Pro.1 (<35)': 500, 'Pro.2 (>=35)': 650, 'Pro.3 (บริหาร)': 500})
    price_data.append({'รายการตรวจสุขภาพ': '--- จำนวนพนักงาน ---', 'Pro.1 (<35)': count_p1, 'Pro.2 (>=35)': count_p2, 'Pro.3 (บริหาร)': count_p3})
    price_data.append({'รายการตรวจสุขภาพ': '--- รวมราคา (บาท) ---', 'Pro.1 (<35)': count_p1*500, 'Pro.2 (>=35)': count_p2*650, 'Pro.3 (บริหาร)': count_p3*500})
    df_summary_price = pd.DataFrame(price_data)

    with tab2:
        st.subheader("ข้อมูลแยกตามหน่วยงาน")
        if 'หน่วยงาน' in edited_df.columns:
            depts = edited_df['หน่วยงาน'].dropna().unique()
            selected_dept = st.selectbox("เลือกหน่วยงาน:", depts)
            dept_df = edited_df[edited_df['หน่วยงาน'] == selected_dept]
            st.dataframe(dept_df, use_container_width=True, hide_index=True)
            
            excel_data = export_full_excel(dept_df, df_summary_count, df_summary_price, "Dept_Export")
            st.download_button(f"📥 Export ข้อมูลหน่วยงาน {selected_dept}", data=excel_data, file_name=f"Checkup_{selected_dept}.xlsx")

    with tab3:
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown(f"**จำนวนพนักงานทั้งหมด:** {total_emp + fail_probation_count} คน | **ไม่ผ่านทดลองงาน:** {fail_probation_count} คน")
            st.markdown(f"### **พนักงานเข้ารับการตรวจจริง: {total_emp} คน**")
            st.write("📌 **สรุปจำนวนรายการตรวจ**")
            st.dataframe(df_summary_count, use_container_width=True, hide_index=True)

        with col2:
            st.write("📌 **สรุปข้อมูลราคาและโปรแกรม**")
            st.dataframe(df_summary_price, use_container_width=True, hide_index=True)
            
            st.success(f"### 💰 ยอดค่าใช้จ่ายรวมทั้งสิ้น: {edited_df['ราคารวม (บาท)'].sum():,.2f} บาท")
            
            chart_df = pd.DataFrame({
                'โปรแกรม': ['Pro.1', 'Pro.2', 'Pro.3'],
                'จำนวนพนักงาน': [count_p1, count_p2, count_p3]
            })
            fig = px.bar(chart_df, x='โปรแกรม', y='จำนวนพนักงาน', text='จำนวนพนักงาน', color='โปรแกรม', title='สัดส่วนพนักงานในแต่ละโปรแกรม')
            fig.update_traces(textposition='outside')
            st.plotly_chart(fig, use_container_width=True)

        st.download_button("📥 Export รายงานฉบับสมบูรณ์ (รวมทุก Sheet)", data=export_full_excel(edited_df, df_summary_count, df_summary_price, "Full"), file_name="Full_Summary_Checkup_2025.xlsx")
else:
    st.info("กรุณาอัปโหลดไฟล์ Excel เพื่อเริ่มต้นการทำงานครับ")
