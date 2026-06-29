import streamlit as st
import pandas as pd
import datetime
import plotly.express as px
import plotly.graph_objects as go
import gspread
from google.oauth2.service_account import Credentials
import json

# ตั้งค่าหน้าจอเริ่มต้น
st.set_page_config(page_title="Minimal Finance Pro", layout="wide", initial_sidebar_state="expanded")

# 🎨 แก้ไข CSS ให้รองรับโหมด Dark/Light อัตโนมัติ (ใช้ CSS Variables ของ Streamlit)
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Prompt:wght@300;400;500;700&display=swap');
    html, body, [class*="css"] { font-family: 'Prompt', sans-serif !important; }
    h1 { font-weight: 700; text-align: center; margin-bottom: 1.5rem; }
    
    /* ปุ่มต่างๆ */
    .stButton>button { border-radius: 8px; font-weight: 500; padding: 10px; }
    .quick-add-text { font-size: 16px; font-weight: bold; margin-bottom: 5px; opacity: 0.8; }
    
    /* กล่อง Metric Card ปรับให้เข้ากับโหมดมืด/สว่าง */
    .metric-card { 
        background-color: var(--secondary-background-color); 
        padding: 20px; 
        border-radius: 12px; 
        text-align: center; 
        border: 1px solid var(--border-color); 
    }
    
    /* ซ่อนลูกศรในช่องกรอกตัวเลข */
    input[type=number]::-webkit-inner-spin-button, input[type=number]::-webkit-outer-spin-button { -webkit-appearance: none; margin: 0; }
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

# ตรวจสอบตาราง (สร้างอัตโนมัติถ้าไม่มี)
try:
    qa_sheet = client.open(spreadsheet_name).worksheet("QuickAdds")
except gspread.exceptions.WorksheetNotFound:
    qa_sheet = client.open(spreadsheet_name).add_worksheet(title="QuickAdds", rows="50", cols="5")
    qa_sheet.append_row(["ชื่อปุ่ม", "ประเภท", "หมวดหมู่", "จำนวนเงิน"])

try:
    cat_sheet = client.open(spreadsheet_name).worksheet("Categories")
except gspread.exceptions.WorksheetNotFound:
    cat_sheet = client.open(spreadsheet_name).add_worksheet(title="Categories", rows="100", cols="3")
    cat_sheet.append_row(["ประเภท", "หมวดหมู่หลัก", "หมวดหมู่ย่อย"])

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
    cat_dict = {"📥 รายรับ": {}, "💸 รายจ่าย": {}, "🐷 เงินออม": {}, "📈 เงินลงทุน": {}}
    for _, row in df.iterrows():
        p = str(row['ประเภท']).strip()
        m = str(row['หมวดหมู่หลัก']).strip()
        y = str(row['หมวดหมู่ย่อย']).strip()
        if p in cat_dict:
            if m not in cat_dict[p]: cat_dict[p][m] = []
            if y and y not in cat_dict[p][m]: cat_dict[p][m].append(y)
    return df, cat_dict

df = load_data()
qa_df = load_quick_adds()
cat_raw_df, SUB_CATEGORIES = load_categories()

# 🎨 กำหนดพาเลตสีสุดมินิมอล (Teal, Orange, Yellow, Purple แบบในรูป)
MINIMAL_COLORS = {
    "รายรับ": "#2dd4bf",     # Teal (ฟ้าอมเขียว)
    "รายจ่าย": "#fb923c",    # Orange (ส้ม)
    "เงินออม": "#fcd34d",    # Yellow (เหลืองพาสเทล)
    "เงินลงทุน": "#c084fc",  # Purple (ม่วง)
    "เงินสุทธิ": "#a3e635"   # Green (เขียวอ่อน)
}

# --- แถบเมนูด้านข้างสลับโหมด ---
st.sidebar.markdown("## ⚙️ โหมดการใช้งาน")
app_mode = st.sidebar.radio("เลือกหน้าตาแอป:", ["📱 โหมดมือถือ (เน้นบันทึกไว)", "💻 โหมดคอมพิวเตอร์ (จัดเต็ม)"])
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
    main_options = list(SUB_CATEGORIES[type_entry].keys()) if SUB_CATEGORIES.get(type_entry) else ["ทั่วไป"]
    main_cat = st.selectbox("หมวดหมู่หลัก", main_options, key="mb_main")
    sub_options = SUB_CATEGORIES[type_entry].get(main_cat, ["ทั่วไป"]) if main_cat in SUB_CATEGORIES.get(type_entry, {}) else ["ทั่วไป"]
    sub_cat = st.selectbox("รายละเอียดหมวดหมู่", sub_options, key="mb_sub")
    
    date_shortcut = st.radio("เลือกวันที่", ["วันนี้", "เมื่อวาน", "ระบุเอง"], horizontal=True)
    chosen_date = datetime.date.today() if date_shortcut == "วันนี้" else (datetime.date.today() - datetime.timedelta(days=1) if date_shortcut == "เมื่อวาน" else st.date_input("เลือกวัน", datetime.date.today()))

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
    tab1, tab2, tab3, tab4 = st.tabs(["✨ บันทึกเงิน", "📊 วิเคราะห์ & Infographic", "🎯 เป้าหมาย", "⚙️ ตั้งค่า/จัดการ"])

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
                main_options = list(SUB_CATEGORIES[type_entry].keys()) if SUB_CATEGORIES.get(type_entry) else ["ทั่วไป"]
                main_cat = st.selectbox("หมวดหมู่หลัก", main_options, key="dt_main")
            with c_sub:
                sub_options = SUB_CATEGORIES[type_entry].get(main_cat, ["ทั่วไป"]) if main_cat in SUB_CATEGORIES.get(type_entry, {}) else ["ทั่วไป"]
                sub_cat = st.selectbox("รายละเอียดหมวดหมู่", sub_options, key="dt_sub")

            c_date_tool, c_note_tool = st.columns([1, 2])
            with c_date_tool:
                date_shortcut_dt = st.radio("เลือกวันที่", ["วันนี้", "เมื่อวาน", "ระบุเอง"], horizontal=True, key="dt_date_shortcut")
                chosen_date_dt = datetime.date.today() if date_shortcut_dt == "วันนี้" else (datetime.date.today() - datetime.timedelta(days=1) if date_shortcut_dt == "เมื่อวาน" else st.date_input("เลือกวัน", datetime.date.today(), key="dt_date_picker"))

            with st.form("desktop_form", clear_on_submit=True):
                amount = st.number_input("จำนวนเงิน (บาท)", min_value=0.0, step=50.0, format="%.2f")
                note = st.text_input("รายละเอียดเพิ่มเติม", placeholder="บันทึกสั้นๆ...")
                if st.form_submit_button("บันทึกรายการ", use_container_width=True) and amount > 0:
                    full_category = f"{main_cat}: {sub_cat}" if sub_cat != "ทั่วไป" else main_cat
                    sheet.append_row([str(chosen_date_dt), type_entry.split(" ")[1], full_category, amount, note])
                    st.cache_data.clear()
                    st.rerun()

    # --- Tab 2: วิเคราะห์ & Infographic (ดีไซน์ตามรูป Reference) ---
    with tab2:
        if not df.empty:
            df_chart = df.copy()
            df_chart['วันที่'] = pd.to_datetime(df_chart['วันที่'])
            
            # --- กล่องสรุปตัวเลขด้านบน ---
            inc = df_chart[df_chart['ประเภท'] == 'รายรับ']['จำนวนเงิน'].sum()
            exp = df_chart[df_chart['ประเภท'] == 'รายจ่าย']['จำนวนเงิน'].sum()
            sav = df_chart[df_chart['ประเภท'] == 'เงินออม']['จำนวนเงิน'].sum()
            inv = df_chart[df_chart['ประเภท'] == 'เงินลงทุน']['จำนวนเงิน'].sum()
            net = inc - (exp + sav + inv)

            m1, m2, m3, m4, m5 = st.columns(5)
            m1.markdown(f"<div class='metric-card'><p style='margin:0;opacity:0.7;font-size:14px;'>ยอดเงินสุทธิ</p><h3 style='margin:0;color:{MINIMAL_COLORS['เงินสุทธิ'] if net >= 0 else '#ef4444'};'>฿{net:,.0f}</h3></div>", unsafe_allow_html=True)
            m2.markdown(f"<div class='metric-card'><p style='margin:0;opacity:0.7;font-size:14px;'>รายรับรวม</p><h3 style='margin:0;color:{MINIMAL_COLORS['รายรับ']};'>฿{inc:,.0f}</h3></div>", unsafe_allow_html=True)
            m3.markdown(f"<div class='metric-card'><p style='margin:0;opacity:0.7;font-size:14px;'>รายจ่ายรวม</p><h3 style='margin:0;color:{MINIMAL_COLORS['รายจ่าย']};'>฿{exp:,.0f}</h3></div>", unsafe_allow_html=True)
            m4.markdown(f"<div class='metric-card'><p style='margin:0;opacity:0.7;font-size:14px;'>เงินออม</p><h3 style='margin:0;color:{MINIMAL_COLORS['เงินออม']};'>฿{sav:,.0f}</h3></div>", unsafe_allow_html=True)
            m5.markdown(f"<div class='metric-card'><p style='margin:0;opacity:0.7;font-size:14px;'>ลงทุน</p><h3 style='margin:0;color:{MINIMAL_COLORS['เงินลงทุน']};'>฿{inv:,.0f}</h3></div>", unsafe_allow_html=True)
            
            st.markdown("<br>", unsafe_allow_html=True) # เว้นบรรทัด
            
            # --- 📈 กราฟเส้นแบบ Minimal (มีรายวัน, รายเดือน, รายปี) ---
            st.markdown("### 📈 แนวโน้มการเงิน (Trend Chart)")
            time_frame = st.radio("ความละเอียด:", ["รายวัน", "รายเดือน", "รายปี"], horizontal=True, label_visibility="collapsed")
            
            df_trend = df_chart.copy()
            if time_frame == "รายปี":
                df_trend['เวลา'] = df_trend['วันที่'].dt.strftime('%Y')
            elif time_frame == "รายเดือน":
                df_trend['เวลา'] = df_trend['วันที่'].dt.strftime('%Y-%m')
            else:
                df_trend['เวลา'] = df_trend['วันที่'].dt.strftime('%Y-%m-%d')
                
            trend_data = df_trend.groupby(['เวลา', 'ประเภท'])['จำนวนเงิน'].sum().reset_index()
            
            # สร้างกราฟเส้นที่บาง มีจุด และไม่มีเส้นตารางกวนใจ
            fig_trend = px.line(trend_data, x='เวลา', y='จำนวนเงิน', color='ประเภท', 
                                color_discrete_map=MINIMAL_COLORS, markers=True, line_shape='linear')
            
            # แต่งเส้นให้เล็ก + จุดไข่ปลา
            fig_trend.update_traces(line=dict(width=2), marker=dict(size=8, line=dict(width=1, color='white')))
            
            # ลบ Grid ให้กลายเป็นสไตล์ Minimal (ใช้คำสั่ง theme="streamlit" ด้านล่างเพื่อปรับมืด/สว่างอัตโนมัติ)
            fig_trend.update_layout(
                xaxis=dict(showgrid=False, title="", zeroline=False, showline=True, linewidth=1, linecolor='rgba(128,128,128,0.3)'),
                yaxis=dict(showgrid=True, gridcolor='rgba(128,128,128,0.1)', title="", zeroline=False),
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="center", x=0.5, title="", font=dict(size=14)),
                hovermode="x unified",
                margin=dict(t=10, b=0, l=0, r=0)
            )
            # สำคัญมาก: ใส่ theme="streamlit" เพื่อให้กราฟเปลี่ยนสีพื้นหลัง/อักษร ตามโหมด Dark/Light ของเครื่อง
            st.plotly_chart(fig_trend, use_container_width=True, theme="streamlit")
            
            st.markdown("---")
            
            # --- ⭕ Infographic วงกลมบาง (Thin Circle Chart) ---
            st.markdown("### ⭕ สัดส่วนค่าใช้จ่าย (Infographic)")
            col_chart1, col_chart2 = st.columns([1, 1])
            expense_df = df_chart[df_chart['ประเภท'] == 'รายจ่าย']
            
            with col_chart1:
                st.markdown("<p style='text-align:center; opacity:0.7;'>รายจ่ายแยกตามหมวดหมู่หลัก</p>", unsafe_allow_html=True)
                if not expense_df.empty:
                    pie_data = expense_df.groupby('หมวดหมู่หลัก')['จำนวนเงิน'].sum().reset_index()
                    
                    # hole=0.85 คือทำรูตรงกลางให้กว้างมากๆ จนกลายเป็นแค่เส้นรอบวง (เหมือนในรูป)
                    fig_pie = px.pie(pie_data, values='จำนวนเงิน', names='หมวดหมู่หลัก', hole=0.85, 
                                     color_discrete_sequence=px.colors.qualitative.Pastel)
                    
                    # เอาตัวอักษรชี้ออกไปด้านนอกวงกลม (เหมือน Infographic)
                    fig_pie.update_traces(textposition='outside', textinfo='percent+label', marker=dict(line=dict(width=0)))
                    fig_pie.update_layout(showlegend=False, margin=dict(t=20, b=20, l=20, r=20))
                    
                    st.plotly_chart(fig_pie, use_container_width=True, theme="streamlit")
                else:
                    st.info("ไม่มีข้อมูลรายจ่าย")
                    
            with col_chart2:
                st.markdown("<p style='text-align:center; opacity:0.7;'>เจาะลึกหมวดหมู่ย่อย (Sub-categories)</p>", unsafe_allow_html=True)
                if not expense_df.empty:
                    sub_data = expense_df.groupby(['หมวดหมู่หลัก', 'หมวดหมู่ย่อย'])['จำนวนเงิน'].sum().reset_index()
                    sub_data = sub_data.sort_values(by="จำนวนเงิน", ascending=False).head(10) # โชว์แค่ 10 อันดับแรกจะได้ไม่รก
                    
                    fig_bar = px.bar(sub_data, x='จำนวนเงิน', y='หมวดหมู่ย่อย', color='หมวดหมู่หลัก', orientation='h',
                                     color_discrete_sequence=px.colors.qualitative.Pastel)
                    
                    fig_bar.update_layout(
                        xaxis=dict(showgrid=False, title="", zeroline=False),
                        yaxis=dict(showgrid=False, title="", autorange="reversed"), # กลับหัวให้มากสุดอยู่บน
                        showlegend=False, margin=dict(t=20, b=20, l=0, r=20)
                    )
                    st.plotly_chart(fig_bar, use_container_width=True, theme="streamlit")
                else:
                    st.info("ไม่มีข้อมูลรายจ่าย")
                    
        else:
            st.info("ยังไม่มีข้อมูล กรุณาบันทึกข้อมูลก่อนครับ")

    with tab3:
        st.subheader("🎯 เป้าหมายการเงิน")
        total_study_savings = df[(df['ประเภท'] == 'เงินออม') & (df['หมวดหมู่หลัก'] == 'ออมเพื่อเรียนต่อ/อนาคต')]['จำนวนเงิน'].sum() if not df.empty else 0
        GOAL_STUDY = 100000 
        progress_percent = min(total_study_savings / GOAL_STUDY, 1.0)
        st.write("✈️ **กองทุนเพื่ออนาคต (เรียนต่อ/สอบ)**")
        st.progress(progress_percent)
        st.caption(f"สะสมแล้ว ฿{total_study_savings:,.2f} จากเป้าหมาย ฿{GOAL_STUDY:,.2f} ({progress_percent*100:.1f}%)")

    with tab4:
        st.subheader("📁 เพิ่ม / ลด / แก้ไข หมวดหมู่ทั้งหมด")
        edited_cat = st.data_editor(cat_raw_df, use_container_width=True, num_rows="dynamic", key="editor_cat_v6")
        if st.button("💾 บันทึกการตั้งค่าหมวดหมู่", use_container_width=True):
            cat_sheet.clear()
            cat_sheet.update(range_name="A1", values=[edited_cat.columns.values.tolist()] + edited_cat.values.tolist())
            st.success("อัปเดตฐานข้อมูลหมวดหมู่เรียบร้อยแล้วครับ! ✨")
            st.rerun()

        st.markdown("---")
        st.subheader("⚡ จัดการปุ่มบันทึกด่วน")
        edited_qa = st.data_editor(qa_df, use_container_width=True, num_rows="dynamic", key="editor_qa_v6")
        if st.button("💾 บันทึกปุ่มด่วน", use_container_width=True):
            qa_sheet.clear()
            qa_sheet.update(range_name="A1", values=[edited_qa.columns.values.tolist()] + edited_qa.values.tolist())
            st.success("อัปเดตปุ่มสำเร็จ!")
            st.rerun()
            
        st.markdown("---")
        st.subheader("✏️ แก้ไข/ลบ ประวัติการเงินทั้งหมด")
        if not df.empty:
            clean_df_edit = df[["วันที่", "ประเภท", "หมวดหมู่", "จำนวนเงิน", "รายละเอียด"]]
            edited_df = st.data_editor(clean_df_edit, use_container_width=True, num_rows="dynamic", key="editor_finance_v6")
            if st.button("💾 บันทึกประวัติลงคลาวด์", use_container_width=True):
                sheet.clear()
                edited_df['วันที่'] = edited_df['วันที่'].astype(str)
                sheet.update(range_name="A1", values=[edited_df.columns.values.tolist()] + edited_df.values.tolist())
                st.success("อัปเดตประวัติการเงินสำเร็จ!")
                st.rerun()
