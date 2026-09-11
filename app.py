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

def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Data')
    return output.getvalue()

# ==========================================
# 2. ส่วนรับข้อมูล (อัปโหลดไฟล์)
# ==========================================
uploaded_file = st.file_uploader("อัปโหลดไฟล์รายชื่อพนักงาน (Excel)", type=["xlsx"])

if uploaded_file is not None:
    df_raw = pd.read_excel(uploaded_file)
    
    # [แก้ไขหลัก] ทำความสะอาดชื่อคอลัมน์โดยตัดช่องว่างหน้า-หลังออกทั้งหมด
    df_raw.columns = df_raw.columns.str.strip()
    
    # ==========================================
    # 3. การประมวลผลข้อมูล (Data Processing)
    # ==========================================
    df_raw['หมายเหตุ'] = df_raw.get('หมายเหตุ', '').fillna('')
    df_filtered = df_raw[~df_raw['หมายเหตุ'].str.contains('ไม่ผ่าน|ทดลองงาน')].copy()
    
    def assign_program(row):
        # อ้างอิงตามชื่อคอลัมน์จริงในไฟล์ (ระดับนักงาน)
        level = str(row.get('ระดับนักงาน', '')).strip()
        if level == 'nan': level = ''
            
        age_str = str(row.get('อายุ', '0'))
        age_match = re.search(r'\d+', age_str)
        age = int(age_match.group()) if age_match else 0
        
        pro3_keywords = ['ผู้จัดการ', 'ผู้บริหาร', 'หัวหน้างาน', 'เภสัชกร', 'วิศวกร']
        is_pro3 = any(kw in level for kw in pro3_keywords)
        
        if is_pro3:
            return pd.Series(['Pro.3 (ผู้บริหาร)', 500])
        elif age >= 35:
            return pd.Series(['Pro.2 (>= 35 ปี)', 650])
        else:
            return pd.Series(['Pro.1 (< 35 ปี)', 500])

    df_filtered[['โปรแกรม', 'ราคาพื้นฐาน']] = df_filtered.apply(assign_program, axis=1)
    df_filtered['ใบรับรองแพทย์ 5 โรค'] = False
    
    # อัปเดตชื่อคอลัมน์รหัสพนักงานให้ตรงกัน
    display_cols = ['รหัสพนักงาน', 'ชื่อ - นามสกุล', 'ตำแหน่ง', 'หน่วยงาน', 'อายุ', 'โปรแกรม', 'ราคาพื้นฐาน', 'ใบรับรองแพทย์ 5 โรค']
    display_cols = [c for c in display_cols if c in df_filtered.columns]
    df_display = df_filtered[display_cols].copy()

    # ==========================================
    # 4. ส่วนแสดงผล (แบ่งเป็น 3 Tabs)
    # ==========================================
    tab1, tab2, tab3 = st.tabs(["📋 ข้อมูลทั้งหมด", "🏢 แยกตามหน่วยงาน", "📊 Dashboard & สรุปผล"])
    
    with tab1:
        st.subheader("ข้อมูลพนักงานทั้งหมด")
        st.markdown("💡 *คุณสามารถติ๊กเลือกช่อง **'ใบรับรองแพทย์ 5 โรค'** ในตารางด้านล่างได้เลย ระบบจะคำนวณราคารวมให้อัตโนมัติ*")
        
        edited_df = st.data_editor(
            df_display, 
            use_container_width=True, 
            hide_index=True,
            disabled=[c for c in display_cols if c != 'ใบรับรองแพทย์ 5 โรค'] 
        )
        
        edited_df['ราคารวม'] = edited_df['ราคาพื้นฐาน'] + (edited_df['ใบรับรองแพทย์ 5 โรค'] * 100)
        
        st.write("ตารางสรุปราคารวม:")
        summary_cols = [c for c in ['รหัสพนักงาน', 'ชื่อ - นามสกุล', 'โปรแกรม', 'ราคาพื้นฐาน', 'ใบรับรองแพทย์ 5 โรค', 'ราคารวม'] if c in edited_df.columns]
        st.dataframe(edited_df[summary_cols], use_container_width=True, hide_index=True)
        
        st.download_button("📥 Export ข้อมูลทั้งหมด (Excel)", data=to_excel(edited_df), file_name="All_Employees_Checkup.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        
    with tab2:
        st.subheader("ข้อมูลแยกตามหน่วยงาน")
        if 'หน่วยงาน' in edited_df.columns:
            departments = edited_df['หน่วยงาน'].dropna().unique()
            selected_dept = st.selectbox("เลือกหน่วยงานที่ต้องการดูข้อมูล:", departments)
            dept_df = edited_df[edited_df['หน่วยงาน'] == selected_dept]
            st.dataframe(dept_df, use_container_width=True, hide_index=True)
            st.download_button(f"📥 Export ข้อมูลหน่วยงาน {selected_dept} (Excel)", data=to_excel(dept_df), file_name=f"Checkup_{selected_dept}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        
    with tab3:
        st.subheader("Dashboard สรุปผลการตรวจสุขภาพ")
        col1, col2 = st.columns([1, 1])
        
        if 'รหัสพนักงาน' in edited_df.columns and 'โปรแกรม' in edited_df.columns:
            summary_pro = edited_df.groupby('โปรแกรม').agg(
                จำนวนพนักงาน=('รหัสพนักงาน', 'count'),
                ราคารวม=('ราคารวม', 'sum')
            ).reset_index()
            
            med_cert_count = edited_df['ใบรับรองแพทย์ 5 โรค'].sum()
            med_cert_price = med_cert_count * 100
            total_employees = len(edited_df)
            grand_total = edited_df['ราคารวม'].sum()
            
            with col1:
                st.write("**ตารางแจกแจงจำนวนคนและราคาแยกตามโปรแกรม**")
                st.dataframe(summary_pro, use_container_width=True, hide_index=True)
                
                st.info(f"**สรุปรายการเพิ่มเติม (กลุ่มเสี่ยง):**\n- ขอใบรับรองแพทย์ 5 โรค: {med_cert_count} คน (รวม {med_cert_price:,} บาท)")
                st.success(f"**จำนวนพนักงานที่ตรวจทั้งหมด:** {total_employees} คน\n\n### **ยอดค่าใช้จ่ายรวมทั้งสิ้น:** {grand_total:,.2f} บาท")
                
                st.download_button("📥 Export ตารางสรุปนี้", data=to_excel(summary_pro), file_name="Summary_Dashboard.xlsx")
                
            with col2:
                st.write("**แผนภูมิเปรียบเทียบจำนวนคนในแต่ละโปรแกรม**")
                if not summary_pro.empty:
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
            st.error("ไม่พบคอลัมน์ 'รหัสพนักงาน' หรือ 'โปรแกรม' ในข้อมูล กรุณาตรวจสอบหัวตารางในไฟล์ Excel")
else:
    st.info("กรุณาอัปโหลดไฟล์ Excel เพื่อเริ่มต้นการทำงานครับ")
