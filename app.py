import streamlit as st
import pandas as pd
import io
import plotly.express as px
import re

# ==========================================
# 1. การตั้งค่าหน้าเว็บ
# ==========================================
st.set_page_config(page_title="ระบบสรุปผลการตรวจสุขภาพ 2025", layout="wide")
st.title("🏥 ระบบจัดการข้อมูลตรวจสุขภาพพนักงาน 2025")

# ฟังก์ชันสำหรับสร้างไฟล์ Excel เพื่อให้ดาวน์โหลด
def to_excel(df):
    output = io.BytesIO()
    # ใช้ openpyxl เป็น engine ในการเขียนไฟล์
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Data')
    processed_data = output.getvalue()
    return processed_data

# ==========================================
# 2. ส่วนรับข้อมูล (อัปโหลดไฟล์)
# ==========================================
uploaded_file = st.file_uploader("อัปโหลดไฟล์รายชื่อพนักงาน (Excel)", type=["xlsx"])

if uploaded_file is not None:
    # อ่านข้อมูลจากไฟล์ Excel
    df_raw = pd.read_excel(uploaded_file)
    
    # ==========================================
    # 3. การประมวลผลข้อมูล (Data Processing)
    # ==========================================
    
    # 3.1 ตัดพนักงานที่ "ไม่ผ่านทดลองงาน" หรือถูก remark ไว้
    df_raw['หมายเหตุ'] = df_raw.get('หมายเหตุ', '').fillna('')
    df_filtered = df_raw[~df_raw['หมายเหตุ'].str.contains('ไม่ผ่าน|ทดลองงาน')].copy()
    
    # 3.2 ฟังก์ชันสำหรับตรวจสอบเงื่อนไขโปรแกรม 1, 2, 3
    def assign_program(row):
        # ดึงระดับพนักงาน (รองรับกรณีชื่อคอลัมน์พิมพ์ตกเป็น 'ระดับนักงาน')
        level = str(row.get('ระดับนักงาน', '')).strip()
        if level == 'nan': level = ''
            
        # ดึงตัวเลขจากข้อความอายุ (เช่น "44 ปี" -> 44)
        age_str = str(row.get('อายุ', '0'))
        age_match = re.search(r'\d+', age_str)
        age = int(age_match.group()) if age_match else 0
        
        # เงื่อนไขกลุ่มผู้บริหาร
        pro3_keywords = ['ผู้จัดการ', 'ผู้บริหาร', 'หัวหน้างาน', 'เภสัชกร', 'วิศวกร']
        is_pro3 = any(kw in level for kw in pro3_keywords)
        
        if is_pro3:
            return pd.Series(['Pro.3 (ผู้บริหาร)', 500])
        elif age >= 35:
            return pd.Series(['Pro.2 (>= 35 ปี)', 650])
        else:
            return pd.Series(['Pro.1 (< 35 ปี)', 500])

    # นำฟังก์ชันไปสร้างคอลัมน์ใหม่
    df_filtered[['โปรแกรม', 'ราคาพื้นฐาน']] = df_filtered.apply(assign_program, axis=1)
    
    # สร้างคอลัมน์สำหรับติ๊กเลือกใบรับรองแพทย์ (ค่าเริ่มต้นคือ False)
    df_filtered['ใบรับรองแพทย์ 5 โรค'] = False
    
    # เลือกเฉพาะคอลัมน์ที่ต้องการแสดงผล
    display_cols = ['รหัสพนักงาน ', 'ชื่อ - นามสกุล', 'ตำแหน่ง', 'หน่วยงาน', 'อายุ', 'โปรแกรม', 'ราคาพื้นฐาน', 'ใบรับรองแพทย์ 5 โรค']
    display_cols = [c for c in display_cols if c in df_filtered.columns]
    df_display = df_filtered[display_cols].copy()

    # ==========================================
    # 4. ส่วนแสดงผล (แบ่งเป็น 3 Tabs)
    # ==========================================
    tab1, tab2, tab3 = st.tabs(["📋 ข้อมูลทั้งหมด", "🏢 แยกตามหน่วยงาน", "📊 Dashboard & สรุปผล"])
    
    with tab1:
        st.subheader("ข้อมูลพนักงานทั้งหมด")
        st.markdown("💡 *คุณสามารถติ๊กเลือกช่อง **'ใบรับรองแพทย์ 5 โรค'** ในตารางด้านล่างได้เลย ระบบจะคำนวณราคารวมให้อัตโนมัติ*")
        
        # ใช้ data_editor เพื่อให้ตารางโต้ตอบได้
        edited_df = st.data_editor(
            df_display, 
            use_container_width=True, 
            hide_index=True,
            disabled=[c for c in display_cols if c != 'ใบรับรองแพทย์ 5 โรค'] 
        )
        
        # คำนวณราคารวมแบบเรียลไทม์หลังจากแก้ไข
        edited_df['ราคารวม'] = edited_df['ราคาพื้นฐาน'] + (edited_df['ใบรับรองแพทย์ 5 โรค'] * 100)
        
        # แสดงตารางผลลัพธ์แบบ Read-only เพื่อความชัดเจน
        st.write("ตารางสรุปราคารวม:")
        st.dataframe(edited_df[['รหัสพนักงาน ', 'ชื่อ - นามสกุล', 'โปรแกรม', 'ราคาพื้นฐาน', 'ใบรับรองแพทย์ 5 โรค', 'ราคารวม']], use_container_width=True, hide_index=True)
        
        # ปุ่ม Export ข้อมูลทั้งหมด
        excel_all = to_excel(edited_df)
        st.download_button(
            label="📥 Export ข้อมูลทั้งหมด (Excel)",
            data=excel_all,
            file_name="All_Employees_Checkup.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    with tab2:
        st.subheader("ข้อมูลแยกตามหน่วยงาน")
        # สร้าง Dropdown เลือกหน่วยงาน
        departments = edited_df['หน่วยงาน'].dropna().unique()
        selected_dept = st.selectbox("เลือกหน่วยงานที่ต้องการดูข้อมูล:", departments)
        
        # กรองข้อมูลตามหน่วยงานที่เลือก
        dept_df = edited_df[edited_df['หน่วยงาน'] == selected_dept]
        st.dataframe(dept_df, use_container_width=True, hide_index=True)
        
        # ปุ่ม Export แยกรายหน่วยงาน
        excel_dept = to_excel(dept_df)
        st.download_button(
            label=f"📥 Export ข้อมูลหน่วยงาน {selected_dept} (Excel)",
            data=excel_dept,
            file_name=f"Checkup_{selected_dept}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        
    with tab3:
        st.subheader("Dashboard สรุปผลการตรวจสุขภาพ")
        
        col1, col2 = st.columns([1, 1])
        
        # จัดกลุ่มนับจำนวนและราคารวมแยกตามโปรแกรม
        summary_pro = edited_df.groupby('โปรแกรม').agg(
            จำนวนพนักงาน=('รหัสพนักงาน ', 'count'),
            ราคารวม=('ราคารวม', 'sum')
        ).reset_index()
        
        # คำนวณยอดรายการพิเศษ
        med_cert_count = edited_df['ใบรับรองแพทย์ 5 โรค'].sum()
        med_cert_price = med_cert_count * 100
        total_employees = len(edited_df)
        grand_total = edited_df['ราคารวม'].sum()
        
        with col1:
            st.write("**ตารางแจกแจงจำนวนคนและราคาแยกตามโปรแกรม**")
            st.dataframe(summary_pro, use_container_width=True, hide_index=True)
            
            st.info(f"**สรุปรายการเพิ่มเติม (กลุ่มเสี่ยง):**\n"
                    f"- ขอใบรับรองแพทย์ 5 โรค: {med_cert_count} คน (รวม {med_cert_price:,} บาท)")
            
            st.success(f"**จำนวนพนักงานที่ตรวจทั้งหมด:** {total_employees} คน\n\n"
                       f"### **ยอดค่าใช้จ่ายรวมทั้งสิ้น:** {grand_total:,.2f} บาท")
            
            # ปุ่ม Export สรุป
            st.download_button("📥 Export ตารางสรุปนี้", data=to_excel(summary_pro), file_name="Summary_Dashboard.xlsx")
            
        with col2:
            st.write("**แผนภูมิเปรียบเทียบจำนวนคนในแต่ละโปรแกรม**")
            # เช็กก่อนว่ามีข้อมูลให้สร้างกราฟหรือไม่
            if not summary_pro.empty:
                # สร้างกราฟแท่งด้วย Plotly
                fig = px.bar(
                    summary_pro, 
                    x='โปรแกรม', 
                    y='จำนวนพนักงาน', 
                    color='โปรแกรม',
                    text='จำนวนพนักงาน',
                    color_discrete_sequence=px.colors.qualitative.Pastel
                )
                fig.update_traces(textposition='outside')
                fig.update_layout(showlegend=False)
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.warning("ไม่มีข้อมูลเพียงพอสำหรับสร้างแผนภูมิครับ")
else:
    st.info("กรุณาอัปโหลดไฟล์ Excel เพื่อเริ่มต้นการทำงานครับ")
          
