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
    
    /* ซ่อนลูกศรใน number_input บนมือถือให้กดง่ายขึ้น */
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
    qa_sheet.append_row(["🍛 ข้าว/อาหาร 50฿", "รายจ่าย", "อาหาร/เครื่องดื่ม", 50.0])
    qa_sheet.append_row(["🚌 เดินทางไปหา Alice", "รายจ่าย", "ค่าเดินทาง", 150.0])
    qa_sheet.append_row(["📚 เก็บเงินสอบ GRE Physics", "เงินออม", "ออมเพื่อเรียนต่อ/อนาคต", 500.0])
    qa_sheet.append_row(["💻 ซื้ออุปกรณ์เขียนนิยาย", "รายจ่าย", "อุปกรณ์ไอที/เขียนงาน", 200.0])

def load_data():
    records = sheet.get_all_records()
    if records:
        df = pd.DataFrame(records)
        df['วันที่'] = pd.to_datetime(df['วันที่']).dt.date
        df['จำนวนเงิน'] = pd.to_numeric(df['จำนวนเงิน'], errors='coerce').fillna(0)
        return df
    return pd.DataFrame(columns=["วันที่", "ประเภท", "หมวดหมู่", "จำนวนเงิน", "รายละเอียด"])

def load_quick_adds():
    records = qa_sheet.get_all_records()
    if records:
        df = pd.DataFrame(records)
        df['จำนวนเงิน'] = pd.to_numeric(df['จำนวนเงิน'], errors='coerce').fillna(0)
        return df
    return pd.DataFrame(columns=["ชื่อปุ่ม", "ประเภท", "หมวดหมู่", "จำนวนเงิน"])

def add_entry(date, entry_type, category, amount, note):
    sheet.append_row([str(date), entry_type, category, amount, note])
    st.toast(f"บันทึก {amount} บาท ลงคลาวด์สำเร็จ! ✨")
    st.cache_data.clear()

df = load_data()
qa_df = load_quick_adds()

# ---------------------------------------------------------
# 🎛️ สร้างเมนูด้านข้าง (Sidebar) สำหรับเลือกโหมด
# ---------------------------------------------------------
st.sidebar.markdown("## ⚙️ โหมดการใช้งาน")
app_mode = st.sidebar.radio(
    "เลือกหน้าตาแอปให้เหมาะกับอุปกรณ์:",
    ["📱 โหมดมือถือ (เน้นบันทึกไว)", "💻 โหมดคอมพิวเตอร์ (จัดเต็ม)"]
)
st.sidebar.markdown("---")
st.sidebar.info("💡 **ทริค:** เวลาใช้บนมือถือ ให้กดเครื่องหมาย > มุมซ้ายบนเพื่อสลับโหมด")

# ==========================================
# 📱 โหมดที่ 1: โหมดมือถือ (Mobile Mode)
# ==========================================
if app_mode == "📱 โหมดมือถือ (เน้นบันทึกไว)":
    st.title("☁️ Finance (Mobile)")
    st.markdown("<p class='quick-add-text'>⚡ บันทึกด่วน</p>", unsafe_allow_html=True)
    
    # วางปุ่มด่วนเรียงลงมาเป็นแนวตั้ง หรือ 2 คอลัมน์ให้กดง่ายๆ ในมือถือ
    if not qa_df.empty:
        for i, row in qa_df.iterrows():
            if st.button(str(row['ชื่อปุ่ม']), use_container_width=True):
                add_entry(datetime.date.today(), str(row['ประเภท']), str(row['หมวดหมู่']), float(row['จำนวนเงิน']), "บันทึกด่วน")
                st.rerun()
                
    st.markdown("---")
    st.markdown("<p class='quick-add-text'>📝 บันทึกใหม่</p>", unsafe_allow_html=True)
    
    type_entry = st.selectbox("ประเภท", ["💸 รายจ่าย", "📥 รายรับ", "🐷 เงินออม", "📈 เงินลงทุน"])
    
    if "รายรับ" in type_entry:
        categories, clean_type = ["เงินเดือน/ค่าจ้าง", "รายได้เสริม", "ทุนการศึกษา", "อื่นๆ"], "รายรับ"
    elif "รายจ่าย" in type_entry:
        categories, clean_type = ["อาหาร/เครื่องดื่ม", "ค่าเดินทาง", "หนังสือ/เตรียมสอบ", "อุปกรณ์ไอที/เขียนงาน", "ของใช้ส่วนตัว", "อื่นๆ"], "รายจ่าย"
    elif "เงินออม" in type_entry:
        categories, clean_type = ["ออมเพื่อเรียนต่อ/อนาคต", "ออมสำรองฉุกเฉิน", "ท่องเที่ยว"], "เงินออม"
    else:
        categories, clean_type = ["หุ้นไทย/ต่างประเทศ", "กองทุนรวม", "สินทรัพย์อื่นๆ"], "เงินลงทุน"

    with st.form("mobile_form", clear_on_submit=True):
        amount = st.number_input("จำนวนเงิน (บาท)", min_value=0.0, step=50.0, format="%.2f")
        category = st.selectbox("หมวดหมู่", categories)
        note = st.text_input("รายละเอียด", placeholder="บันทึกกันลืม...")
        date = st.date_input("วันที่", datetime.date.today())
        
        # ปุ่มบันทึกใหญ่พิเศษ
        if st.form_submit_button("💾 ยืนยันการบันทึก", use_container_width=True) and amount > 0:
            add_entry(date, clean_type, category, amount, note)
            st.rerun()

# ==========================================
# 💻 โหมดที่ 2: โหมดคอมพิวเตอร์ (Desktop Mode)
# ==========================================
else:
    st.title("☁️ Minimal Finance Pro")
    tab1, tab2, tab3, tab4 = st.tabs(["✨ บันทึกเงิน", "📊 วิเคราะห์ Infographic", "🎯 เป้าหมาย", "⚙️ ตั้งค่า/จัดการ"])

    # --- Tab 1: หน้าบันทึกเงิน ---
    with tab1:
        col_main, col_space = st.columns([2, 1])
        with col_main:
            st.markdown("<p class='quick-add-text'>⚡ บันทึกด่วน</p>", unsafe_allow_html=True)
            if not qa_df.empty:
                cols = st.columns(4)
                for i, row in qa_df.iterrows():
                    col = cols[i % 4]
                    if col.button(str(row['ชื่อปุ่ม']), use_container_width=True, key=f"desktop_btn_{i}"):
                        add_entry(datetime.date.today(), str(row['ประเภท']), str(row['หมวดหมู่']), float(row['จำนวนเงิน']), "บันทึกด่วน")
                        st.rerun()
                        
            st.markdown("---")
            st.markdown("<p class='quick-add-text'>📝 บันทึกลงรายละเอียด</p>", unsafe_allow_html=True)
            type_entry = st.radio("ประเภท", ["📥 รายรับ", "💸 รายจ่าย", "🐷 เงินออม", "📈 เงินลงทุน"], horizontal=True, label_visibility="collapsed")
            
            if "รายรับ" in type_entry:
                categories, clean_type = ["เงินเดือน/ค่าจ้าง", "รายได้เสริม", "ทุนการศึกษา", "อื่นๆ"], "รายรับ"
            elif "รายจ่าย" in type_entry:
                categories, clean_type = ["อาหาร/เครื่องดื่ม", "ค่าเดินทาง", "หนังสือ/เตรียมสอบ", "อุปกรณ์ไอที/เขียนงาน", "ของใช้ส่วนตัว", "อื่นๆ"], "รายจ่าย"
            elif "เงินออม" in type_entry:
                categories, clean_type = ["ออมเพื่อเรียนต่อ/อนาคต", "ออมสำรองฉุกเฉิน", "ท่องเที่ยว"], "เงินออม"
            else:
                categories, clean_type = ["หุ้นไทย/ต่างประเทศ", "กองทุนรวม", "สินทรัพย์อื่นๆ"], "เงินลงทุน"

            with st.form("desktop_form", clear_on_submit=True):
                amount = st.number_input("จำนวนเงิน (บาท)", min_value=0.0, step=50.0, format="%.2f")
                category = st.selectbox("หมวดหมู่", categories)
                note = st.text_input("รายละเอียด", placeholder="เช่น เติมน้ำมัน, ค่าหนังสือ")
                date = st.date_input("วันที่", datetime.date.today())
                if st.form_submit_button("บันทึกรายการ", use_container_width=True) and amount > 0:
                    add_entry(date, clean_type, category, amount, note)
                    st.rerun()

    # --- Tab 2: หน้าวิเคราะห์ Infographic ---
    with tab2:
        if not df.empty:
            df_chart = df.copy()
            df_chart['วันที่'] = pd.to_datetime(df_chart['วันที่'])
            df_chart['ปี'] = df_chart['วันที่'].dt.year
            df_chart['เดือน'] = df_chart['วันที่'].dt.month
            df_chart['ชื่อเดือน'] = df_chart['วันที่'].dt.strftime('%b')
            
            st.markdown("### 🔍 ตัวกรองข้อมูล")
            f_col1, f_col2, f_col3 = st.columns(3)
            year_list = ["ภาพรวมทุกปี"] + sorted(list(df_chart['ปี'].unique()), reverse=True)
            selected_year = f_col1.selectbox("เลือกปี", year_list)
            
            if selected_year != "ภาพรวมทุกปี":
                df_filtered = df_chart[df_chart['ปี'] == selected_year]
                month_list = ["ภาพรวมทั้งปี"] + sorted(list(df_filtered['เดือน'].unique()))
                selected_month = f_col2.selectbox("เลือกเดือน", month_list)
                if selected_month != "ภาพรวมทั้งปี":
                    df_filtered = df_filtered[df_filtered['เดือน'] == selected_month]
            else:
                df_filtered = df_chart
                selected_month = "ภาพรวมทั้งปี"
                
            st.markdown("---")
            
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
                st.markdown("#### 🍕 สัดส่วนการใช้เงิน")
                if not expense_df.empty:
                    pie_data = expense_df.groupby('หมวดหมู่')['จำนวนเงิน'].sum().reset_index()
                    fig_pie = px.pie(pie_data, values='จำนวนเงิน', names='หมวดหมู่', hole=0.4, color_discrete_sequence=px.colors.qualitative.Pastel)
                    fig_pie.update_traces(textposition='inside', textinfo='percent+label')
                    fig_pie.update_layout(margin=dict(t=0, b=0, l=0, r=0))
                    st.plotly_chart(fig_pie, use_container_width=True)
                else:
                    st.info("ไม่มีข้อมูลรายจ่ายในช่วงนี้")
                    
            with col_chart2:
                st.markdown("#### 📈 แนวโน้มการใช้เงิน")
                if not expense_df.empty:
                    if selected_month != "ภาพรวมทั้งปี" and selected_year != "ภาพรวมทุกปี":
                        trend_data = expense_df.groupby(expense_df['วันที่'].dt.day)['จำนวนเงิน'].sum().reset_index()
                        fig_line = px.line(trend_data, x='วันที่', y='จำนวนเงิน', markers=True, line_shape='spline', color_discrete_sequence=['#e74c3c'])
                    else:
                        trend_data = expense_df.groupby('ชื่อเดือน')['จำนวนเงิน'].sum().reset_index()
                        months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
                        trend_data['ชื่อเดือน'] = pd.Categorical(trend_data['ชื่อเดือน'], categories=months, ordered=True)
                        trend_data = trend_data.sort_values('ชื่อเดือน')
                        fig_line = px.bar(trend_data, x='ชื่อเดือน', y='จำนวนเงิน', text='จำนวนเงิน', color_discrete_sequence=['#e74c3c'])
                        fig_line.update_traces(texttemplate='฿%{text:,.0f}', textposition='outside')
                    
                    fig_line.update_layout(margin=dict(t=0, b=0, l=0, r=0), xaxis_title="", yaxis_title="")
                    st.plotly_chart(fig_line, use_container_width=True)
                else:
                    st.info("ไม่มีข้อมูลเพียงพอ")
        else:
            st.info("ยังไม่มีข้อมูล กรุณาบันทึกข้อมูลก่อนครับ")

    # --- Tab 3: หน้าติดตามเป้าหมาย ---
    with tab3:
        st.subheader("🎯 เป้าหมายการเงิน")
        total_study_savings = df[(df['ประเภท'] == 'เงินออม') & (df['หมวดหมู่'] == 'ออมเพื่อเรียนต่อ/อนาคต')]['จำนวนเงิน'].sum() if not df.empty else 0
        GOAL_STUDY = 100000 
        progress_percent = min(total_study_savings / GOAL_STUDY, 1.0)
        st.write("✈️ **กองทุนเพื่ออนาคต (เรียนต่อ/สอบ)**")
        st.progress(progress_percent)
        st.caption(f"สะสมแล้ว ฿{total_study_savings:,.2f} จากเป้าหมาย ฿{GOAL_STUDY:,.2f} ({progress_percent*100:.1f}%)")

    # --- Tab 4: ตั้งค่าและจัดการ ---
    with tab4:
        st.subheader("⚡ จัดการปุ่มบันทึกด่วน")
        edited_qa = st.data_editor(qa_df, use_container_width=True, num_rows="dynamic", key="editor_qa_desk")
        if st.button("💾 บันทึกปุ่มด่วนลงคลาวด์", use_container_width=True):
            qa_sheet.clear()
            data = [edited_qa.columns.values.tolist()] + edited_qa.values.tolist()
            qa_sheet.update(range_name="A1", values=data)
            st.success("อัปเดตปุ่มสำเร็จ!")
            st.rerun()
            
        st.markdown("---")
        st.subheader("✏️ แก้ไข/ลบ ประวัติทั้งหมด")
        if not df.empty:
            edited_df = st.data_editor(df, use_container_width=True, num_rows="dynamic", key="editor_finance_desk")
            if st.button("💾 บันทึกประวัติลงคลาวด์", use_container_width=True):
                sheet.clear()
                edited_df['วันที่'] = edited_df['วันที่'].astype(str)
                data = [edited_df.columns.values.tolist()] + edited_df.values.tolist()
                sheet.update(range_name="A1", values=data)
                st.success("อัปเดตสำเร็จ!")
                st.rerun()
