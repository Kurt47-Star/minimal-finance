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

st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Prompt:wght@300;400;500;700&display=swap');
    h1, h2, h3, h4, h5, h6, p, label, input, select, .stButton>button { font-family: 'Prompt', sans-serif !important; }
    h1 { font-weight: 700; color: #2C3E50; text-align: center; margin-bottom: 1.5rem; }
    .stButton>button { border-radius: 8px; font-weight: 500; padding: 10px; }
    .quick-add-text { font-size: 16px; font-weight: bold; color: #2C3E50; margin-bottom: 5px; }
    .metric-card { background-color: #ffffff; padding: 20px; border-radius: 12px; text-align: center; border: 1px solid #f0f2f6; box-shadow: 0 4px 6px rgba(0,0,0,0.02); }
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

try:
    qa_sheet = client.open(spreadsheet_name).worksheet("QuickAdds")
except gspread.exceptions.WorksheetNotFound:
    qa_sheet = client.open(spreadsheet_name).add_worksheet(title="QuickAdds", rows="50", cols="5")
    qa_sheet.append_row(["ชื่อปุ่ม", "ประเภท", "หมวดหมู่", "จำนวนเงิน"])
    qa_sheet.append_row(["🍛 ข้าวเช้า 50฿", "รายจ่าย", "อาหาร/เครื่องดื่ม: ข้าวเช้า", 50.0])

try:
    cat_sheet = client.open(spreadsheet_name).worksheet("Categories")
except gspread.exceptions.WorksheetNotFound:
    cat_sheet = client.open(spreadsheet_name).add_worksheet(title="Categories", rows="100", cols="3")
    cat_sheet.append_row(["ประเภท", "หมวดหมู่หลัก", "หมวดหมู่ย่อย"])
    cat_sheet.append_rows([["💸 รายจ่าย", "อาหาร/เครื่องดื่ม", "ข้าวเช้า"], ["📥 รายรับ", "เงินเดือน/ค่าจ้าง", "งานประจำ"]])

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
    main_options = list(SUB_CATEGORIES[type_entry].keys()) if SUB_CATEGORIES[type_entry] else ["ทั่วไป"]
    main_cat = st.selectbox("หมวดหมู่หลัก", main_options, key="mb_main")
    sub_options = SUB_CATEGORIES[type_entry].get(main_cat, ["ทั่วไป"]) if main_cat in SUB_CATEGORIES[type_entry] else ["ทั่วไป"]
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
    tab1, tab2, tab3, tab4 = st.tabs(["✨ บันทึกเงิน", "📊 วิเคราะห์ & กราฟ (Minimal)", "🎯 เป้าหมาย", "⚙️ ตั้งค่า/จัดการ"])

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

    # --- Tab 2: วิเคราะห์ & กราฟ (สไตล์มินิมอล) ---
    with tab2:
        if not df.empty:
            df_chart = df.copy()
            df_chart['วันที่'] = pd.to_datetime(df_chart['วันที่'])
            df_chart['ปี'] = df_chart['วันที่'].dt.year
            df_chart['เดือน'] = df_chart['วันที่'].dt.month
            
            # --- กล่องสรุปตัวเลขด้านบน ---
            inc = df_chart[df_chart['ประเภท'] == 'รายรับ']['จำนวนเงิน'].sum()
            exp = df_chart[df_chart['ประเภท'] == 'รายจ่าย']['จำนวนเงิน'].sum()
            sav = df_chart[df_chart['ประเภท'] == 'เงินออม']['จำนวนเงิน'].sum()
            inv = df_chart[df_chart['ประเภท'] == 'เงินลงทุน']['จำนวนเงิน'].sum()
            net = inc - (exp + sav + inv)

            m1, m2, m3, m4, m5 = st.columns(5)
            m1.markdown(f"<div class='metric-card'><p style='margin:0;color:#95a5a6;font-size:14px;'>เงินสุทธิคงเหลือ</p><h3 style='margin:0;color:{'#2ECC71' if net >= 0 else '#E74C3C'};'>฿{net:,.0f}</h3></div>", unsafe_allow_html=True)
            m2.markdown(f"<div class='metric-card'><p style='margin:0;color:#95a5a6;font-size:14px;'>รายรับรวม</p><h3 style='margin:0;color:#3498db;'>฿{inc:,.0f}</h3></div>", unsafe_allow_html=True)
            m3.markdown(f"<div class='metric-card'><p style='margin:0;color:#95a5a6;font-size:14px;'>รายจ่ายรวม</p><h3 style='margin:0;color:#e74c3c;'>฿{exp:,.0f}</h3></div>", unsafe_allow_html=True)
            m4.markdown(f"<div class='metric-card'><p style='margin:0;color:#95a5a6;font-size:14px;'>เงินออม</p><h3 style='margin:0;color:#f1c40f;'>฿{sav:,.0f}</h3></div>", unsafe_allow_html=True)
            m5.markdown(f"<div class='metric-card'><p style='margin:0;color:#95a5a6;font-size:14px;'>ลงทุน</p><h3 style='margin:0;color:#9b59b6;'>฿{inv:,.0f}</h3></div>", unsafe_allow_html=True)
            
            st.markdown("---")
            
            # --- กราฟแนวโน้ม (Trend Graph) ---
            st.markdown("### 📈 กราฟแนวโน้มการเงิน (Trend Analysis)")
            time_frame = st.radio("เลือกความละเอียดของกราฟ:", ["รายวัน", "รายเดือน", "รายปี"], horizontal=True)
            
            # เตรียมข้อมูลสำหรับกราฟ
            df_trend = df_chart.copy()
            if time_frame == "รายปี":
                df_trend['เวลา'] = df_trend['วันที่'].dt.strftime('%Y')
            elif time_frame == "รายเดือน":
                df_trend['เวลา'] = df_trend['วันที่'].dt.strftime('%Y-%m')
            else:
                df_trend['เวลา'] = df_trend['วันที่'].dt.strftime('%Y-%m-%d')
                
            trend_data = df_trend.groupby(['เวลา', 'ประเภท'])['จำนวนเงิน'].sum().reset_index()
            
            # กำหนดสีแบบมินิมอลให้แต่ละประเภท
            color_map = {"รายรับ": "#3498db", "รายจ่าย": "#e74c3c", "เงินออม": "#f1c40f", "เงินลงทุน": "#9b59b6"}
            
            fig_trend = px.line(trend_data, x='เวลา', y='จำนวนเงิน', color='ประเภท', 
                                markers=True, color_discrete_map=color_map, line_shape='spline')
            
            # ตั้งค่าให้กราฟดูคลีนและมินิมอลที่สุด
            fig_trend.update_layout(
                plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
                xaxis=dict(showgrid=False, title="", showline=True, linewidth=1, linecolor='#ecf0f1'),
                yaxis=dict(showgrid=True, gridcolor='#f7f9f9', title="", zeroline=False),
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1, title="")
            )
            st.plotly_chart(fig_trend, use_container_width=True)
            
            st.markdown("---")
            
            # --- Infographic สัดส่วนรายจ่าย (Donut/Treemap) ---
            st.markdown("### 🍕 สัดส่วนการใช้เงิน (Infographic)")
            col_chart1, col_chart2 = st.columns([1, 2])
            expense_df = df_chart[df_chart['ประเภท'] == 'รายจ่าย']
            
            with col_chart1:
                if not expense_df.empty:
                    pie_data = expense_df.groupby('หมวดหมู่หลัก')['จำนวนเงิน'].sum().reset_index()
                    fig_pie = px.pie(pie_data, values='จำนวนเงิน', names='หมวดหมู่หลัก', hole=0.5, 
                                     color_discrete_sequence=px.colors.qualitative.Pastel)
                    fig_pie.update_traces(textposition='inside', textinfo='percent')
                    fig_pie.update_layout(plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', margin=dict(t=0, b=0, l=0, r=0), showlegend=True)
                    st.plotly_chart(fig_pie, use_container_width=True)
                else:
                    st.info("ไม่มีข้อมูลรายจ่าย")
                    
            with col_chart2:
                if not expense_df.empty:
                    sub_data = expense_df.groupby(['หมวดหมู่หลัก', 'หมวดหมู่ย่อย'])['จำนวนเงิน'].sum().reset_index()
                    fig_sub = px.bar(sub_data, x='หมวดหมู่ย่อย', y='จำนวนเงิน', color='หมวดหมู่หลัก', 
                                     color_discrete_sequence=px.colors.qualitative.Pastel)
                    fig_sub.update_layout(
                        plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
                        xaxis=dict(showgrid=False, title=""),
                        yaxis=dict(showgrid=True, gridcolor='#f7f9f9', title=""),
                        showlegend=False
                    )
                    st.plotly_chart(fig_sub, use_container_width=True)
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
        edited_cat = st.data_editor(cat_raw_df, use_container_width=True, num_rows="dynamic", key="editor_cat_v5")
        if st.button("💾 บันทึกการตั้งค่าหมวดหมู่", use_container_width=True):
            cat_sheet.clear()
            cat_sheet.update(range_name="A1", values=[edited_cat.columns.values.tolist()] + edited_cat.values.tolist())
            st.success("อัปเดตฐานข้อมูลหมวดหมู่เรียบร้อยแล้วครับ! ✨")
            st.rerun()

        st.markdown("---")
        st.subheader("⚡ จัดการปุ่มบันทึกด่วน")
        edited_qa = st.data_editor(qa_df, use_container_width=True, num_rows="dynamic", key="editor_qa_v5")
        if st.button("💾 บันทึกปุ่มด่วน", use_container_width=True):
            qa_sheet.clear()
            qa_sheet.update(range_name="A1", values=[edited_qa.columns.values.tolist()] + edited_qa.values.tolist())
            st.success("อัปเดตปุ่มสำเร็จ!")
            st.rerun()
            
        st.markdown("---")
        st.subheader("✏️ แก้ไข/ลบ ประวัติการเงินทั้งหมด")
        if not df.empty:
            clean_df_edit = df[["วันที่", "ประเภท", "หมวดหมู่", "จำนวนเงิน", "รายละเอียด"]]
            edited_df = st.data_editor(clean_df_edit, use_container_width=True, num_rows="dynamic", key="editor_finance_v5")
            if st.button("💾 บันทึกประวัติลงคลาวด์", use_container_width=True):
                sheet.clear()
                edited_df['วันที่'] = edited_df['วันที่'].astype(str)
                sheet.update(range_name="A1", values=[edited_df.columns.values.tolist()] + edited_df.values.tolist())
                st.success("อัปเดตประวัติการเงินสำเร็จ!")
                st.rerun()
