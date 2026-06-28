import streamlit as st
import pandas as pd
import datetime
import plotly.express as px
import gspread
from google.oauth2.service_account import Credentials
import json

# ตั้งค่าหน้าจอเริ่มต้น
st.set_page_config(page_title="Minimal Finance Pro", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Prompt:wght@300;400;500;700&display=swap');
    h1, h2, h3, h4, h5, h6, p, label, input, select, .stButton>button { font-family: 'Prompt', sans-serif !important; }
    h1 { font-weight: 700; color: #2C3E50; text-align: center; margin-bottom: 1.5rem; }
    .stButton>button { border-radius: 8px; font-weight: 500; padding: 10px; }
    .quick-add-text { font-size: 16px; font-weight: bold; color: #2C3E50; margin-bottom: 5px; }
    .metric-card { background-color: #F8F9FA; padding: 15px; border-radius: 10px; text-align: center; box-shadow: 0 2px 5px rgba(0,0,0,0.05); }
    input[type=number]::-webkit-inner-spin-button, 
    input[type=number]::-webkit-outer-spin-button { -webkit-appearance: none; margin: 0; }
    </style>
""", unsafe_allow_html=True)

st.title("☁️ Minimal Finance Pro")

# --- ระบบเชื่อมต่อคลาวด์ (Google Sheets) ---
@st.cache_resource
def init_connection():
    creds_dict = json.loads(st.secrets["google_credentials"])
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    return gspread.authorize(creds)

client = init_connection()
spreadsheet_name = "Minimal Finance Pro"

try:
    sheet = client.open(spreadsheet_name).sheet1
except Exception as e:
    st.error(f"❌ หาไฟล์ Google Sheets ที่ชื่อ '{spreadsheet_name}' ไม่เจอครับ")
    st.stop()

# ตรวจสอบและสร้างตารางปุ่มลัด
try:
    qa_sheet = client.open(spreadsheet_name).worksheet("QuickAdds")
except gspread.exceptions.WorksheetNotFound:
    qa_sheet = client.open(spreadsheet_name).add_worksheet(title="QuickAdds", rows="50", cols="5")
    qa_sheet.append_row(["ชื่อปุ่ม", "ประเภท", "หมวดหมู่", "จำนวนเงิน"])
    qa_sheet.append_row(["🍛 ข้าวเช้า 50฿", "รายจ่าย", "อาหาร/เครื่องดื่ม: ข้าวเช้า", 50.0])
    qa_sheet.append_row(["☕ กาแฟ 60฿", "รายจ่าย", "อาหาร/เครื่องดื่ม: กาแฟ", 60.0])

# ตรวจสอบและสร้างตารางหมวดหมู่ (Dynamic Categories)
try:
    cat_sheet = client.open(spreadsheet_name).worksheet("Categories")
except gspread.exceptions.WorksheetNotFound:
    cat_sheet = client.open(spreadsheet_name).add_worksheet(title="Categories", rows="100", cols="3")
    cat_sheet.append_row(["ประเภท", "หมวดหมู่หลัก", "หมวดหมู่ย่อย"])
    default_cats = [
        ["💸 รายจ่าย", "อาหาร/เครื่องดื่ม", "ข้าวเช้า"], ["💸 รายจ่าย", "อาหาร/เครื่องดื่ม", "ข้าวเที่ยง"],
        ["💸 รายจ่าย", "อาหาร/เครื่องดื่ม", "กาแฟ/ชา"], ["💸 รายจ่าย", "ค่าเดินทาง", "เติมน้ำมัน"],
        ["💸 รายจ่าย", "ค่าเดินทาง", "ไปหา Alice"], ["💸 รายจ่าย", "หนังสือ/เตรียมสอบ", "คู่มือสอบ/GRE"],
        ["💸 รายจ่าย", "อุปกรณ์ไอที/เขียนงาน", "เครื่องเขียน/สมุดบันทึก"], ["📥 รายรับ", "เงินเดือน/ค่าจ้าง", "งานประจำ"],
        ["📥 รายรับ", "รายได้เสริม", "เขียนงาน/นิยาย"], ["🐷 เงินออม", "ออมเพื่อเรียนต่อ/อนาคต", "ทุนศึกษาต่อตปท."]
    ]
    cat_sheet.append_rows(default_cats)

# --- ฟังก์ชันโหลดข้อมูล ---
def load_data():
    records = sheet.get_all_records()
    if records:
        df = pd.DataFrame(records)
        df['วันที่'] = pd.to_datetime(df['วันที่']).dt.date
        df['จำนวนเงิน'] = pd.to_numeric(df['จำนวนเงิน'], errors='coerce').fillna(0)
        df['หมวดหมู่หลัก'] = df['หมวดหมู่'].apply(lambda x: str(x).split(":")[0].strip())
        df['หมวดหมู่ย่อย'] = df['หมวดหมู่'].apply(lambda x: str(x).split(":")[1].strip() if ":" in str(x) else "ทั่วไป")
        return df
    return pd.DataFrame(columns=["วันที่", "ประเภท", "หมวดหมู่", "จำนวนเงิน", "รายละเอียด", "หมวดหมู่หลัก", "หมวดหมู่ย่อย"])

def load_quick_adds():
    records = qa_sheet.get_all_records()
    return pd.DataFrame(records) if records else pd.DataFrame(columns=["ชื่อปุ่ม", "ประเภท", "หมวดหมู่", "จำนวนเงิน"])

def load_categories():
    records = cat_sheet.get_all_records()
    df = pd.DataFrame(records) if records else pd.DataFrame(columns=["ประเภท", "หมวดหมู่หลัก", "หมวดหมู่ย่อย"])
    
    # แปลงโครงสร้างให้เป็น Dictionary เพื่อเรียกใช้งานง่าย
    cat_dict = {"📥 รายรับ": {}, "💸 รายจ่าย": {}, "🐷 เงินออม": {}, "📈 เงินลงทุน": {}}
    for _, row in df.iterrows():
        p = str(row['ประเภท']).strip()
        m = str(row['หมวดหมู่หลัก']).strip()
        y = str(row['หมวดหมู่ย่อย']).strip()
        if p in cat_dict:
            if m not in cat_dict[p]:
                cat_dict[p][m] = []
            if y and y not in cat_dict[p][m]:
                cat_dict[p][m].append(y)
    return df, cat_dict

df = load_data()
qa_df = load_quick_adds()
cat_raw_df, SUB_CATEGORIES = load_categories()

# --- แถบเมนูด้านข้างสลับโหมด ---
st.sidebar.markdown("## ⚙️ โหมดการใช้งาน")
app_mode = st.sidebar.radio("เลือกหน้าตาแอปให้เหมาะกับอุปกรณ์:", ["📱 โหมดมือถือ (เน้นบันทึกไว)", "💻 โหมดคอมพิวเตอร์ (จัดเต็ม)"])
st.sidebar.markdown("---")

# ==========================================
# 📱 โหมดมือถือ (Mobile Mode)
# ==========================================
if app_mode == "📱 โหมดมือถือ (เน้นบันทึกไว)":
    st.title("☁️ Finance (Mobile)")
    st.markdown("<p class='quick-add-text'>⚡ บันทึกด่วน</p>", unsafe_allow_html=True)
    
    if not qa_df.empty:
        for i, row in qa_df.iterrows():
            if st.button(str(row['ชื่อปุ่ม']), use_container_width=True, key=f"mb_qa_{i}"):
                sheet.append_row([str(datetime.date.today()), str(row['ประเภท']), str(row['หมวดหมู่']), float(row['จำนวนเงิน']), "บันทึกด่วน"])
                st.toast("บันทึกด่วนสำเร็จ! ✨")
                st.cache_data.clear()
                st.rerun()
                
    st.markdown("---")
    st.markdown("<p class='quick-add-text'>📝 บันทึกใหม่</p>", unsafe_allow_html=True)
    
    type_entry = st.selectbox("ประเภท", ["💸 รายจ่าย", "📥 รายรับ", "🐷 เงินออม", "📈 เงินลงทุน"])
    
    main_options = list(SUB_CATEGORIES[type_entry].keys()) if SUB_CATEGORIES[type_entry] else ["ทั่วไป"]
    main_cat = st.selectbox("หมวดหมู่หลัก", main_options, key="mb_main")
    
    sub_options = SUB_CATEGORIES[type_entry].get(main_cat, ["ทั่วไป"]) if main_cat in SUB_CATEGORIES[type_entry] else ["ทั่วไป"]
    sub_cat = st.selectbox("รายละเอียดหมวดหมู่ (:เจาะจง)", sub_options, key="mb_sub")

    # เครื่องมืออำนวยความสะดวก: ปุ่มลัดวันที่
    date_shortcut = st.radio("เลือกวันที่", ["วันนี้", "เมื่อวาน", "ระบุเอง"], horizontal=True)
    if date_shortcut == "วันนี้":
        chosen_date = datetime.date.today()
    elif date_shortcut == "เมื่อวาน":
        chosen_date = datetime.date.today() - datetime.timedelta(days=1)
    else:
        chosen_date = st.date_input("เลือกวัน", datetime.date.today(), key="mb_date_picker")

    with st.form("mobile_form", clear_on_submit=True):
        amount = st.number_input("จำนวนเงิน (บาท)", min_value=0.0, step=50.0, format="%.2f")
        note = st.text_input("บันทึกสั้นๆ", placeholder="บันทึกกันลืม...")
        
        if st.form_submit_button("💾 ยืนยันการบันทึก", use_container_width=True) and amount > 0:
            full_category = f"{main_cat}: {sub_cat}" if sub_cat != "ทั่วไป" else main_cat
            sheet.append_row([str(chosen_date), type_entry.split(" ")[1], full_category, amount, note])
            st.cache_data.clear()
            st.rerun()

# ==========================================
# 💻 โหมดคอมพิวเตอร์ (Desktop Mode)
# ==========================================
else:
    st.title("☁️ Minimal Finance Pro")
    tab1, tab2, tab3, tab4 = st.tabs(["✨ บันทึกเงิน", "📊 วิเคราะห์ Infographic", "🎯 เป้าหมาย", "⚙️ ตั้งค่า/จัดการ"])

    # --- Tab 1: บันทึกเงิน ---
    with tab1:
        col_main, col_space = st.columns([2, 1])
        with col_main:
            st.markdown("<p class='quick-add-text'>⚡ บันทึกด่วน</p>", unsafe_allow_html=True)
            if not qa_df.empty:
                cols = st.columns(4)
                for i, row in qa_df.iterrows():
                    col = cols[i % 4]
                    if col.button(str(row['ชื่อปุ่ม']), use_container_width=True, key=f"dt_qa_{i}"):
                        sheet.append_row([str(datetime.date.today()), str(row['ประเภท']), str(row['หมวดหมู่']), float(row['จำนวนเงิน']), "บันทึกด่วน"])
                        st.toast("บันทึกด่วนสำเร็จ! ✨")
                        st.cache_data.clear()
                        st.rerun()
                        
            st.markdown("---")
            st.markdown("<p class='quick-add-text'>📝 บันทึกลงรายละเอียด</p>", unsafe_allow_html=True)
            type_entry = st.radio("ประเภท", ["📥 รายรับ", "💸 รายจ่าย", "🐷 เงินออม", "📈 เงินลงทุน"], horizontal=True, label_visibility="collapsed")
            
            c_main, c_sub = st.columns(2)
            with c_main:
                main_options = list(SUB_CATEGORIES[type_entry].keys()) if SUB_CATEGORIES[type_entry] else ["ทั่วไป"]
                main_cat = st.selectbox("หมวดหมู่หลัก", main_options, key="dt_main")
            with c_sub:
                sub_options = SUB_CATEGORIES[type_entry].get(main_cat, ["ทั่วไป"]) if main_cat in SUB_CATEGORIES[type_entry] else ["ทั่วไป"]
                sub_cat = st.selectbox("รายละเอียดหมวดหมู่ (:เจาะจง)", sub_options, key="dt_sub")

            c_date_tool, c_note_tool = st.columns([1, 2])
            with c_date_tool:
                date_shortcut_dt = st.radio("เลือกวันที่", ["วันนี้", "เมื่อวาน", "ระบุเอง"], horizontal=True, key="dt_date_shortcut")
                if date_shortcut_dt == "วันนี้":
                    chosen_date_dt = datetime.date.today()
                elif date_shortcut_dt == "เมื่อวาน":
                    chosen_date_dt = datetime.date.today() - datetime.timedelta(days=1)
                else:
                    chosen_date_dt = st.date_input("เลือกวัน", datetime.date.today(), key="dt_date_picker")

            with st.form("desktop_form", clear_on_submit=True):
                amount = st.number_input("จำนวนเงิน (บาท)", min_value=0.0, step=50.0, format="%.2f")
                note = st.text_input("รายละเอียดเพิ่มเติม", placeholder="บันทึกสั้นๆ...")
                if st.form_submit_button("บันทึกรายการ", use_container_width=True) and amount > 0:
                    full_category = f"{main_cat}: {sub_cat}" if sub_cat != "ทั่วไป" else main_cat
                    sheet.append_row([str(chosen_date_dt), type_entry.split(" ")[1], full_category, amount, note])
                    st.cache_data.clear()
                    st.rerun()

    # --- Tab 2: วิเคราะห์ Infographic (เวอร์ชันอัปเกรดความละเอียดสูง) ---
    with tab2:
        if not df.empty:
            df_chart = df.copy()
            df_chart['วันที่'] = pd.to_datetime(df_chart['วันที่'])
            df_chart['ปี'] = df_chart['วันที่'].dt.year
            df_chart['เดือน'] = df_chart['วันที่'].dt.month
            df_chart['ชื่อเดือน'] = df_chart['วันที่'].dt.strftime('%b')
            
            # --- เครื่องมือปรับแต่ง Infographic ---
            st.markdown("### 🎨 เครื่องมือปรับแต่งสถิติ & Infographic")
            ctrl_col1, ctrl_col2, ctrl_col3, ctrl_col4 = st.columns(4)
            
            year_list = ["ภาพรวมทุกปี"] + sorted(list(df_chart['ปี'].unique()), reverse=True)
            selected_year = ctrl_col1.selectbox("เลือกปี", year_list)
            
            if selected_year != "ภาพรวมทุกปี":
                df_filtered = df_chart[df_chart['ปี'] == selected_year]
                month_list = ["ภาพรวมทั้งปี"] + sorted(list(df_filtered['เดือน'].unique()))
                selected_month = ctrl_col2.selectbox("เลือกเดือน", month_list)
                if selected_month != "ภาพรวมทั้งปี":
                    df_filtered = df_filtered[df_filtered['เดือน'] == selected_month]
            else:
                df_filtered = df_chart
                selected_month = "ภาพรวมทั้งปี"
            
            # ปุ่มสลับรูปแบบกราฟและธีมสี (แก้ไขบั๊กชุดสีตรงนี้ครับ)
            chart_type = ctrl_col3.selectbox("รูปแบบแผนภูมิรายจ่าย", ["🍕 แผนภูมิโดนัท (Donut)", "🌀 วงกลมสลับชั้น (Sunburst)", "📦 แผนภูมิกล่องสัดส่วน (Treemap)"])
            color_theme = ctrl_col4.selectbox("โทนสีเครื่องแต่งกราฟ", ["Pastel (พาสเทลละมุน)", "Vivid (สีสันสดใส)", "Cool (โทนเย็นสบายตา)"])
            
            theme_map = {
                "Pastel (พาสเทลละมุน)": px.colors.qualitative.Pastel, 
                "Vivid (สีสันสดใส)": px.colors.qualitative.Vivid, 
                "Cool (โทนเย็นสบายตา)": px.colors.qualitative.Set3  # เปลี่ยนเป็น Set3 ที่ระบบรู้จัก
            }
            chosen_color = theme_map[color_theme]

            st.markdown("---")
            
            # คำนวณยอดรวม
            inc = df_filtered[df_filtered['ประเภท'] == 'รายรับ']['จำนวนเงิน'].sum()
            exp = df_filtered[df_filtered['ประเภท'] == 'รายจ่าย']['จำนวนเงิน'].sum()
            sav = df_filtered[df_filtered['ประเภท'] == 'เงินออม']['จำนวนเงิน'].sum()
            inv = df_filtered[df_filtered['ประเภท'] == 'เงินลงทุน']['จำนวนเงิน'].sum()
            net = inc - (exp + sav + inv)

            # เครื่องมือความสะดวก: ระบบแจ้งเตือนงบประมาณเกิน (Budget Alert)
            net_color = "#2ECC71" if net >= 0 else "#E74C3C"
            net_text = "เงินสุทธิคงเหลือ" if net >= 0 else "⚠️ งบประมาณติดลบ เกินรายรับ!"

            m1, m2, m3, m4, m5 = st.columns(5)
            m1.markdown(f"<div class='metric-card'><p style='margin:0;color:#7f8c8d;'>{net_text}</p><h3 style='margin:0;color:{net_color};'>฿{net:,.0f}</h3></div>", unsafe_allow_html=True)
            m2.markdown(f"<div class='metric-card'><p style='margin:0;color:#7f8c8d;'>📥 รายรับทั้งหมด</p><h3 style='margin:0;color:#3498db;'>฿{inc:,.0f}</h3></div>", unsafe_allow_html=True)
            m3.markdown(f"<div class='metric-card'><p style='margin:0;color:#7f8c8d;'>💸 รายจ่ายรวม</p><h3 style='margin:0;color:#e74c3c;'>฿{exp:,.0f}</h3></div>", unsafe_allow_html=True)
            m4.markdown(f"<div class='metric-card'><p style='margin:0;color:#7f8c8d;'>🐷 เงินออมสะสม</p><h3 style='margin:0;color:#f1c40f;'>฿{sav:,.0f}</h3></div>", unsafe_allow_html=True)
            m5.markdown(f"<div class='metric-card'><p style='margin:0;color:#7f8c8d;'>📈 พอร์ตลงทุน</p><h3 style='margin:0;color:#9b59b6;'>฿{inv:,.0f}</h3></div>", unsafe_allow_html=True)
            
            st.write("")
            expense_df = df_filtered[df_filtered['ประเภท'] == 'รายจ่าย']
            
            col_chart1, col_chart2 = st.columns(2)
            with col_chart1:
                st.markdown("#### 📊 แผนภูมิวิเคราะห์สัดส่วนเชิงลึก")
                if not expense_df.empty:
                    if "Donut" in chart_type:
                        pie_data = expense_df.groupby('หมวดหมู่หลัก')['จำนวนเงิน'].sum().reset_index()
                        fig = px.pie(pie_data, values='จำนวนเงิน', names='หมวดหมู่หลัก', hole=0.4, color_discrete_sequence=chosen_color)
                    elif "Sunburst" in chart_type:
                        fig = px.sunburst(expense_df, path=['หมวดหมู่หลัก', 'หมวดหมู่ย่อย'], values='จำนวนเงิน', color_discrete_sequence=chosen_color)
                    else:
                        fig = px.treemap(expense_df, path=['หมวดหมู่หลัก', 'หมวดหมู่ย่อย'], values='จำนวนเงิน', color_discrete_sequence=chosen_color)
                    
                    fig.update_layout(margin=dict(t=10, b=10, l=10, r=10))
                    st.plotly_chart(fig, use_container_width=True)
                else:
                    st.info("ไม่มีข้อมูลรายจ่ายในช่วงเวลาที่เลือก")
                    
            with col_chart2:
                st.markdown("#### 🔬 ตารางความถี่และรายละเอียดการวิเคราะห์ยอดเงิน")
                if not expense_df.empty:
                    analysis_table = expense_df.groupby(['หมวดหมู่หลัก', 'หมวดหมู่ย่อย'])['จำนวนเงิน'].sum().reset_index()
                    analysis_table = analysis_table.sort_values(by="จำนวนเงิน", ascending=False)
                    analysis_table['สัดส่วน (%)'] = (analysis_table['จำนวนเงิน'] / exp * 100).round(1)
                    analysis_table['จำนวนเงิน'] = analysis_table['จำนวนเงิน'].map('฿{:,.2f}'.format)
                    st.dataframe(analysis_table, use_container_width=True, hide_index=True)
                else:
                    st.info("ไม่มีข้อมูล")
        else:
            st.info("ยังไม่มีข้อมูลในระบบ")

    # --- Tab 3: หน้าติดตามเป้าหมาย ---
    with tab3:
        st.subheader("🎯 เป้าหมายการเงิน")
        total_study_savings = df[(df['ประเภท'] == 'เงินออม') & (df['หมวดหมู่หลัก'] == 'ออมเพื่อเรียนต่อ/อนาคต')]['จำนวนเงิน'].sum() if not df.empty else 0
        GOAL_STUDY = 100000 
        progress_percent = min(total_study_savings / GOAL_STUDY, 1.0)
        st.write("✈️ **กองทุนเพื่ออนาคต (เรียนต่อ/สอบ)**")
        st.progress(progress_percent)
        st.caption(f"สะสมแล้ว ฿{total_study_savings:,.2f} จากเป้าหมาย ฿{GOAL_STUDY:,.2f} ({progress_percent*100:.1f}%)")

    # --- Tab 4: ตั้งค่าและจัดการ ---
    with tab4:
        st.subheader("📁 เพิ่ม / ลด / แก้ไข หมวดหมู่ทั้งหมด")
        st.write("หมอสามารถพิมพ์แก้ไข เพิ่มแถวใหม่ หรือลบแถวหมวดหมู่หลักและย่อยได้ที่ตารางนี้เลยครับ (เมื่อแก้เสร็จแล้วกดบันทึกเพื่ออัปเดตระบบระบบกรอกข้อมูล)")
        edited_cat = st.data_editor(cat_raw_df, use_container_width=True, num_rows="dynamic", key="editor_cat_v4")
        if st.button("💾 บันทึกการตั้งค่าหมวดหมู่ลงคลาวด์", use_container_width=True):
            cat_sheet.clear()
            cat_data = [edited_cat.columns.values.tolist()] + edited_cat.values.tolist()
            cat_sheet.update(range_name="A1", values=cat_data)
            st.success("อัปเดตฐานข้อมูลหมวดหมู่เรียบร้อยแล้วครับ! ✨")
            st.rerun()

        st.markdown("---")
        st.subheader("⚡ จัดการปุ่มบันทึกด่วน")
        edited_qa = st.data_editor(qa_df, use_container_width=True, num_rows="dynamic", key="editor_qa_v4")
        if st.button("💾 บันทึกปุ่มด่วนลงคลาวด์", use_container_width=True):
            qa_sheet.clear()
            data = [edited_qa.columns.values.tolist()] + edited_qa.values.tolist()
            qa_sheet.update(range_name="A1", values=data)
            st.success("อัปเดตปุ่มสำเร็จ!")
            st.rerun()
            
        st.markdown("---")
        st.subheader("✏️ แก้ไข/ลบ ประวัติการเงินทั้งหมด")
        if not df.empty:
            clean_df_edit = df[["วันที่", "ประเภท", "หมวดหมู่", "จำนวนเงิน", "รายละเอียด"]]
            edited_df = st.data_editor(clean_df_edit, use_container_width=True, num_rows="dynamic", key="editor_finance_v4")
            if st.button("💾 บันทึกประวัติลงคลาวด์", use_container_width=True):
                sheet.clear()
                edited_df['วันที่'] = edited_df['วันที่'].astype(str)
                data = [edited_df.columns.values.tolist()] + edited_df.values.tolist()
                sheet.update(range_name="A1", values=data)
                st.success("อัปเดตประวัติการเงินสำเร็จ!")
                st.rerun()
