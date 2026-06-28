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

# --- ระบบเชื่อมต่อคลาวด์ (Google Sheets) ---
@st.cache_resource
def init_connection():
    creds_dict = json.loads(st.secrets["google_credentials"])
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    return gspread.authorize(creds)

client = init_connection()

try:
    sheet = client.open("Minimal Finance Pro").sheet1
except Exception as e:
    st.error("❌ หาไฟล์ Google Sheets ที่ชื่อ 'Minimal Finance Pro' ไม่เจอครับ")
    st.stop()

try:
    qa_sheet = client.open("Minimal Finance Pro").worksheet("QuickAdds")
except gspread.exceptions.WorksheetNotFound:
    qa_sheet = client.open("Minimal Finance Pro").add_worksheet(title="QuickAdds", rows="50", cols="5")
    qa_sheet.append_row(["ชื่อปุ่ม", "ประเภท", "หมวดหมู่", "จำนวนเงิน"])
    qa_sheet.append_row(["🍛 ข้าวเช้า 50฿", "รายจ่าย", "อาหาร/เครื่องดื่ม: ข้าวเช้า", 50.0])
    qa_sheet.append_row(["☕ กาแฟ 60฿", "รายจ่าย", "อาหาร/เครื่องดื่ม: กาแฟ", 60.0])
    qa_sheet.append_row(["📚 ซื้อคู่มือ GRE", "เงินออม", "หนังสือ/เตรียมสอบ: คู่มือสอบ", 500.0])
    qa_sheet.append_row(["🚌 ไปหา Alice", "รายจ่าย", "ค่าเดินทาง: ต่างจังหวัด", 150.0])

def load_data():
    records = sheet.get_all_records()
    if records:
        df = pd.DataFrame(records)
        df['วันที่'] = pd.to_datetime(df['วันที่']).dt.date
        df['จำนวนเงิน'] = pd.to_numeric(df['จำนวนเงิน'], errors='coerce').fillna(0)
        
        # แยกหมวดหมู่หลัก และหมวดหมู่ย่อยออกจากกันเพื่อใช้ทำกราฟ
        df['หมวดหมู่หลัก'] = df['หมวดหมู่'].apply(lambda x: str(x).split(":")[0].strip())
        df['หมวดหมู่ย่อย'] = df['หมวดหมู่'].apply(lambda x: str(x).split(":")[1].strip() if ":" in str(x) else "ทั่วไป")
        return df
    return pd.DataFrame(columns=["วันที่", "ประเภท", "หมวดหมู่", "จำนวนเงิน", "รายละเอียด", "หมวดหมู่หลัก", "หมวดหมู่ย่อย"])

def load_quick_adds():
    records = qa_sheet.get_all_records()
    if records:
        df = pd.DataFrame(records)
        df['จำนวนเงิน'] = pd.to_numeric(df['จำนวนเงิน'], errors='coerce').fillna(0)
        return df
    return pd.DataFrame(columns=["ชื่อปุ่ม", "ประเภท", "หมวดหมู่", "จำนวนเงิน"])

def add_entry(date, entry_type, main_cat, sub_cat, amount, note):
    # รวมร่างหมวดหมู่หลักและย่อยเข้าด้วยกันก่อนเซฟลง Google Sheets
    full_category = f"{main_cat}: {sub_cat}" if sub_cat else main_cat
    sheet.append_row([str(date), entry_type, full_category, amount, note])
    st.toast(f"บันทึก {amount} บาท สำเร็จ! ✨")
    st.cache_data.clear()

df = load_data()
qa_df = load_quick_adds()

# --- โครงสร้างหมวดหมู่ย่อยอัจฉริยะ ---
SUB_CATEGORIES = {
    "📥 รายรับ": {
        "เงินเดือน/ค่าจ้าง": ["งานประจำ", "พาร์ทไทม์", "ค่าสอนพิเศษ"],
        "รายได้เสริม": ["ขายของออนไลน์", "งานฟรีแลนซ์", "เขียนงาน/นิยาย"],
        "ทุนการศึกษา": ["ทุนรายเดือน", "ทุนวิจัย"],
        "อื่นๆ": ["เงินโอนเข้า", "ทั่วไป"]
    },
    "💸 รายจ่าย": {
        "อาหาร/เครื่องดื่ม": ["ข้าวเช้า", "ข้าวเที่ยง", "ข้าวเย็น", "กาแฟ/ชา", "ขนม/ของว่าง"],
        "ค่าเดินทาง": ["เติมน้ำมัน", "รถสาธารณะ", "เดินทางไกล/ไปหา Alice"],
        "หนังสือ/เตรียมสอบ": ["คู่มือสอบ/GRE", "หนังสือเรียนฟิสิกส์", "หนังสืออ่านเล่น/นิยาย"],
        "อุปกรณ์ไอที/เขียนงาน": ["อุปกรณ์คอม/อัปคอม", "เครื่องเขียน/สมุดบันทึก"],
        "ของใช้ส่วนตัว": ["เสื้อผ้า", "ของใช้ในห้อง", "สกินแคร์/ยา"],
        "อื่นๆ": ["ทั่วไป", "ค่าธรรมเนียม"]
    },
    "🐷 เงินออม": {
        "ออมเพื่อเรียนต่อ/อนาคต": ["ทุนศึกษาต่อตปท.", "ค่าสมัครสอบ"],
        "ออมสำรองฉุกเฉิน": ["เงินออมส่วนตัว", "กองทุนสำรอง"],
        "ท่องเที่ยว": ["ทริปเดินทาง", "พักผ่อน"]
    },
    "📈 เงินลงทุน": {
        "หุ้นไทย/ต่างประเทศ": ["พอร์ตหุ้นหลัก", "หุ้นปันผล"],
        "กองทุนรวม": ["กองทุนลดหย่อนภาษี", "กองทุนทั่วไป"],
        "สินทรัพย์อื่นๆ": ["ทองคำ", "คริปโต"]
    }
}

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
    
    # ดึงรายชื่อหมวดหมู่ตามประเภทรายการ
    main_options = list(SUB_CATEGORIES[type_entry].keys())
    main_cat = st.selectbox("หมวดหมู่หลัก", main_options, key="mb_main")
    sub_options = SUB_CATEGORIES[type_entry][main_cat]
    sub_cat = st.selectbox("รายละเอียดหมวดหมู่ (:เจาะจง)", sub_options, key="mb_sub")

    with st.form("mobile_form", clear_on_submit=True):
        amount = st.number_input("จำนวนเงิน (บาท)", min_value=0.0, step=50.0, format="%.2f")
        note = st.text_input("บันทึกสั้นๆ (ไม่ใส่ก็ได้)", placeholder="เช่น ร้านป้าแก้ว, อัปเกรดแรม")
        date = st.date_input("วันที่", datetime.date.today())
        
        if st.form_submit_button("💾 ยืนยันการบันทึก", use_container_width=True) and amount > 0:
            add_entry(date, type_entry.split(" ")[1], main_cat, sub_cat, amount, note)
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
                main_options = list(SUB_CATEGORIES[type_entry].keys())
                main_cat = st.selectbox("หมวดหมู่หลัก", main_options, key="dt_main")
            with c_sub:
                sub_options = SUB_CATEGORIES[type_entry][main_cat]
                sub_cat = st.selectbox("รายละเอียดหมวดหมู่ (:เจาะจง)", sub_options, key="dt_sub")

            with st.form("desktop_form", clear_on_submit=True):
                amount = st.number_input("จำนวนเงิน (บาท)", min_value=0.0, step=50.0, format="%.2f")
                note = st.text_input("รายละเอียดเพิ่มเติม", placeholder="บันทึกสั้นๆ...")
                date = st.date_input("วันที่", datetime.date.today())
                if st.form_submit_button("บันทึกรายการ", use_container_width=True) and amount > 0:
                    add_entry(date, type_entry.split(" ")[1], main_cat, sub_cat, amount, note)
                    st.rerun()

    # --- Tab 2: วิเคราะห์ Infographic ---
    with tab2:
        if not df.empty:
            df_chart = df.copy()
            df_chart['วันที่'] = pd.to_datetime(df_chart['วันที่'])
            df_chart['ปี'] = df_chart['วันที่'].dt.year
            df_chart['เดือน'] = df_chart['วันที่'].dt.month
            df_chart['ชื่อเดือน'] = df_chart['วันที่'].dt.strftime('%b')
            
            st.markdown("### 🔍 ตัวกรองข้อมูล")
            f_col1, f_col2 = st.columns(2)
            year_list = ["ภาพรวมทุกปี"] + sorted(list(df_chart['ปี'].unique()), reverse=True)
            selected_year = f_col1.selectbox("เลือกปี", year_list)
            
            if selected_year != "ภาพwatchทุกปี":
                df_filtered = df_chart[df_chart['ปี'] == selected_year]
                month_list = ["ภาพรวมทั้งปี"] + sorted(list(df_filtered['เดือน'].unique()))
                selected_month = f_col2.selectbox("เลือกเดือน", month_list)
                if selected_month != "ภาพรวมทั้งปี":
                    df_filtered = df_filtered[df_filtered['เดือน'] == selected_month]
            else:
                df_filtered = df_chart
                selected_month = "ภาพรวมทั้งปี"
                
            st.markdown("---")
            
            # คำนวณยอดรวม
            inc = df_filtered[df_filtered['ประเภท'] == 'รายรับ']['จำนวนเงิน'].sum()
            exp = df_filtered[df_filtered['ประเภท'] == 'รายจ่าย']['จำนวนเงิน'].sum()
            sav = df_filtered[df_filtered['ประเภท'] == 'เงินออม']['จำนวนเงิน'].sum()
            inv = df_filtered[df_filtered['ประเภท'] == 'เงินลงทุน']['จำนวนเงิน'].sum()
            net = inc - (exp + sav + inv)

            m1, m2, m3, m4, m5 = st.columns(5)
            m1.markdown(f"<div class='metric-card'><p style='margin:0;color:#7f8c8d;'>เงินสุทธิ</p><h3 style='margin:0;color:{'#2ECC71' if net >= 0 else '#E74C3C'};'>฿{net:,.0f}</h3></div>", unsafe_allow_html=True)
            m2.markdown(f"<div class='metric-card'><p style='margin:0;color:#7f8c8d;'>📥 รายรับ</p><h3 style='margin:0;color:#3498db;'>฿{inc:,.0f}</h3></div>", unsafe_allow_html=True)
            m3.markdown(f"<div class='metric-card'><p style='margin:0;color:#7f8c8d;'>💸 รายจ่าย</p><h3 style='margin:0;color:#e74c3c;'>฿{exp:,.0f}</h3></div>", unsafe_allow_html=True)
            m4.markdown(f"<div class='metric-card'><p style='margin:0;color:#7f8c8d;'>🐷 เงินออม</p><h3 style='margin:0;color:#f1c40f;'>฿{sav:,.0f}</h3></div>", unsafe_allow_html=True)
            m5.markdown(f"<div class='metric-card'><p style='margin:0;color:#7f8c8d;'>📈 ลงทุน</p><h3 style='margin:0;color:#9b59b6;'>฿{inv:,.0f}</h3></div>", unsafe_allow_html=True)
            
            expense_df = df_filtered[df_filtered['ประเภท'] == 'รายจ่าย']
            
            col_chart1, col_chart2 = st.columns(2)
            with col_chart1:
                st.markdown("#### 🍕 สัดส่วนตามหมวดหมู่หลัก (Donut Chart)")
                if not expense_df.empty:
                    pie_data = expense_df.groupby('หมวดหมู่หลัก')['จำนวนเงิน'].sum().reset_index()
                    fig_pie = px.pie(pie_data, values='จำนวนเงิน', names='หมวดหมู่หลัก', hole=0.4, color_discrete_sequence=px.colors.qualitative.Pastel)
                    fig_pie.update_layout(margin=dict(t=0, b=0, l=0, r=0))
                    st.plotly_chart(fig_pie, use_container_width=True)
                else:
                    st.info("ไม่มีข้อมูลรายจ่าย")
                    
            with col_chart2:
                st.markdown("#### 🔍 เจาะลึกรายจ่ายรายหมวดหมู่ย่อย")
                if not expense_df.empty:
                    # แสดงกราฟแท่งที่กระจายข้อมูลเป็นหมวดหมู่ย่อยให้เห็นพฤติกรรมชัดเจน
                    sub_data = expense_df.groupby(['หมวดหมู่หลัก', 'หมวดหมู่ย่อย'])['จำนวนเงิน'].sum().reset_index()
                    fig_sub = px.bar(sub_data, x='หมวดหมู่ย่อย', y='จำนวนเงิน', color='หมวดหมู่หลัก', text='จำนวนเงิน',
                                     color_discrete_sequence=px.colors.qualitative.Pastel)
                    fig_sub.update_traces(texttemplate='฿%{text:,.0f}', textposition='outside')
                    fig_sub.update_layout(margin=dict(t=30, b=0, l=0, r=0), xaxis_title="หมวดหมู่ย่อย", yaxis_title="")
                    st.plotly_chart(fig_sub, use_container_width=True)
                    
            st.markdown("---")
            st.markdown("#### 🏆 อันดับหมวดหมู่ย่อยที่ใช้เงินเยอะที่สุด")
            if not expense_df.empty:
                top_sub = expense_df.groupby('หมวดหมู่')['จำนวนเงิน'].sum().sort_values(ascending=False).head(3)
                medals = ["🥇", "🥈", "🥉"]
                cols = st.columns(3)
                for i, (cat, amt) in enumerate(top_sub.items()):
                    cols[i].markdown(f"<div style='text-align:center; padding:10px; background-color:#fff3cd; border-radius:10px;'><h4>{medals[i]} {cat}</h4><h3 style='color:#d35400;'>฿{amt:,.0f}</h3></div>", unsafe_allow_html=True)
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
        st.subheader("⚡ จัดการปุ่มบันทึกด่วน")
        st.caption("พิมพ์ชื่อหมวดหมู่ในรูปแบบ 'หมวดหมู่หลัก: หมวดหมู่ย่อย' เพื่อให้ระบบกรองข้อมูลได้ถูกต้อง")
        edited_qa = st.data_editor(qa_df, use_container_width=True, num_rows="dynamic", key="editor_qa_v3")
        if st.button("💾 บันทึกปุ่มด่วนลงคลาวด์", use_container_width=True):
            qa_sheet.clear()
            data = [edited_qa.columns.values.tolist()] + edited_qa.values.tolist()
            qa_sheet.update(range_name="A1", values=data)
            st.success("อัปเดตปุ่มสำเร็จ!")
            st.rerun()
            
        st.markdown("---")
        st.subheader("✏️ แก้ไข/ลบ ประวัติทั้งหมด")
        if not df.empty:
            # ซ่อนคอลลัมน์คำนวณชั่วคราวเพื่อลดความสับสนในการแก้ไข
            clean_df_edit = df[["วันที่", "ประเภท", "หมวดหมู่", "จำนวนเงิน", "รายละเอียด"]]
            edited_df = st.data_editor(clean_df_edit, use_container_width=True, num_rows="dynamic", key="editor_finance_v3")
            if st.button("💾 บันทึกประวัติลงคลาวด์", use_container_width=True):
                sheet.clear()
                edited_df['วันที่'] = edited_df['วันที่'].astype(str)
                data = [edited_df.columns.values.tolist()] + edited_df.values.tolist()
                sheet.update(range_name="A1", values=data)
                st.success("อัปเดตสำเร็จ!")
                st.rerun()
