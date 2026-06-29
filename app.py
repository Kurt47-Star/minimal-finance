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

# 🎨 ธีมสี Honey Pot
COLOR_BG = "#ededed"
COLOR_CARD = "#ffffff"
COLOR_NAVY = "#102937"
COLOR_TEAL = "#124d54"
COLOR_ORANGE = "#f9744b"
COLOR_BEIGE = "#e1d9cf"

# 🔤 CSS สไตล์ Soft UI & Poppins Font
st.markdown(f"""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;500;600;700&family=Prompt:wght@300;400;500;600&display=swap');
    
    /* บังคับใช้ฟอนต์ Poppins (ตัวเลข/อังกฤษ) และ Prompt (ภาษาไทย) */
    html, body, [class*="css"] {{ 
        font-family: 'Poppins', 'Prompt', sans-serif !important; 
        color: {COLOR_NAVY};
    }}
    
    /* เปลี่ยนสีพื้นหลังแอป */
    .stApp {{ background-color: {COLOR_BG}; }}
    
    h1, h2, h3 {{ font-weight: 700; color: {COLOR_NAVY}; }}
    
    /* แต่งปุ่มสไตล์คลีน */
    .stButton>button {{ 
        border-radius: 12px; 
        font-weight: 500; 
        padding: 10px; 
        border: 1px solid {COLOR_BEIGE};
        background-color: {COLOR_CARD};
        color: {COLOR_NAVY};
        box-shadow: 0 4px 6px rgba(16, 41, 55, 0.04);
        transition: all 0.3s ease;
    }}
    .stButton>button:hover {{
        border-color: {COLOR_ORANGE};
        color: {COLOR_ORANGE};
        transform: translateY(-2px);
        box-shadow: 0 6px 12px rgba(249, 116, 75, 0.15);
    }}
    
    .quick-add-text {{ font-size: 18px; font-weight: 600; margin-bottom: 8px; color: {COLOR_NAVY}; }}
    
    /* 📌 การ์ดแสดงผลตัวเลข (Metric Card) เลียนแบบรูปตัวอย่าง */
    .metric-card {{ 
        background-color: {COLOR_CARD}; 
        padding: 24px; 
        border-radius: 20px; 
        text-align: left; 
        box-shadow: 0 10px 30px rgba(16, 41, 55, 0.05); 
        border: none;
        margin-bottom: 1rem;
    }}
    .metric-title {{ color: {COLOR_TEAL}; font-size: 15px; font-weight: 500; opacity: 0.9; margin-bottom: 5px; }}
    .metric-title-alert {{ color: {COLOR_ORANGE}; font-size: 15px; font-weight: 600; margin-bottom: 5px; }}
    .metric-value {{ color: {COLOR_NAVY}; font-size: 32px; font-weight: 700; margin: 0; line-height: 1.2; }}
    .metric-currency {{ color: {COLOR_BEIGE}; font-size: 14px; font-weight: 500; margin-top: 5px; }}
    
    /* ซ่อนลูกศรในช่องกรอกตัวเลข */
    input[type=number]::-webkit-inner-spin-button, input[type=number]::-webkit-outer-spin-button {{ -webkit-appearance: none; margin: 0; }}
    </style>
""", unsafe_allow_html=True)

st.title("Minimal Finance Pro")

# --- ระบบเชื่อมต่อคลาวด์ ---
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

# 🎨 Map สีให้ตรงกับประเภท (Honey Pot Palette)
HONEY_POT_MAP = {
    "รายรับ": COLOR_TEAL,     # เขียวอมฟ้า
    "รายจ่าย": COLOR_ORANGE,  # ส้ม
    "เงินออม": COLOR_NAVY,    # กรมท่า
    "เงินลงทุน": COLOR_BEIGE  # เบจ
}

# --- แถบเมนูด้านข้างสลับโหมด ---
st.sidebar.markdown("## ⚙️ Settings")
app_mode = st.sidebar.radio("Layout Mode:", ["📱 Mobile Mode", "💻 Desktop Mode"])
st.sidebar.markdown("---")

# ==========================================
# 📱 โหมดมือถือ (Mobile Mode)
# ==========================================
if app_mode == "📱 Mobile Mode":
    st.markdown("<p class='quick-add-text'>Quick Actions</p>", unsafe_allow_html=True)
    if not qa_df.empty:
        for i, row in qa_df.iterrows():
            if st.button(str(row['ชื่อปุ่ม']), use_container_width=True, key=f"mb_qa_{i}"):
                sheet.append_row([str(datetime.date.today()), str(row['ประเภท']), str(row['หมวดหมู่']), float(row['จำนวนเงิน']), "บันทึกด่วน"])
                st.toast("Success! ✨")
                st.cache_data.clear()
                st.rerun()
                
    st.markdown("---")
    st.markdown("<p class='quick-add-text'>New Transaction</p>", unsafe_allow_html=True)
    type_entry = st.selectbox("Type", ["💸 รายจ่าย", "📥 รายรับ", "🐷 เงินออม", "📈 เงินลงทุน"])
    main_options = list(SUB_CATEGORIES[type_entry].keys()) if SUB_CATEGORIES.get(type_entry) else ["ทั่วไป"]
    main_cat = st.selectbox("Category", main_options, key="mb_main")
    sub_options = SUB_CATEGORIES[type_entry].get(main_cat, ["ทั่วไป"]) if main_cat in SUB_CATEGORIES.get(type_entry, {}) else ["ทั่วไป"]
    sub_cat = st.selectbox("Sub-category", sub_options, key="mb_sub")
    
    date_shortcut = st.radio("Date", ["วันนี้", "เมื่อวาน", "ระบุเอง"], horizontal=True)
    chosen_date = datetime.date.today() if date_shortcut == "วันนี้" else (datetime.date.today() - datetime.timedelta(days=1) if date_shortcut == "เมื่อวาน" else st.date_input("เลือกวัน", datetime.date.today()))

    with st.form("mobile_form", clear_on_submit=True):
        amount = st.number_input("Amount (THB)", min_value=0.0, step=50.0, format="%.2f")
        note = st.text_input("Note", placeholder="Optional...")
        if st.form_submit_button("Save Transaction", use_container_width=True) and amount > 0:
            full_category = f"{main_cat}: {sub_cat}" if sub_cat != "ทั่วไป" else main_cat
            sheet.append_row([str(chosen_date), type_entry.split(" ")[1], full_category, amount, note])
            st.cache_data.clear()
            st.rerun()

# ==========================================
# 💻 โหมดคอมพิวเตอร์ (Desktop Mode)
# ==========================================
else:
    tab1, tab2, tab3, tab4 = st.tabs(["✨ Transaction", "📊 Dashboard", "🎯 Goals", "⚙️ Settings"])

    with tab1:
        col_main, col_space = st.columns([2, 1])
        with col_main:
            st.markdown("<p class='quick-add-text'>Quick Actions</p>", unsafe_allow_html=True)
            if not qa_df.empty:
                cols = st.columns(4)
                for i, row in qa_df.iterrows():
                    col = cols[i % 4]
                    if col.button(str(row['ชื่อปุ่ม']), use_container_width=True, key=f"dt_qa_{i}"):
                        sheet.append_row([str(datetime.date.today()), str(row['ประเภท']), str(row['หมวดหมู่']), float(row['จำนวนเงิน']), "บันทึกด่วน"])
                        st.toast("Success! ✨")
                        st.cache_data.clear()
                        st.rerun()
                        
            st.markdown("---")
            st.markdown("<p class='quick-add-text'>New Transaction</p>", unsafe_allow_html=True)
            type_entry = st.radio("Type", ["📥 รายรับ", "💸 รายจ่าย", "🐷 เงินออม", "📈 เงินลงทุน"], horizontal=True, label_visibility="collapsed")
            
            c_main, c_sub = st.columns(2)
            with c_main:
                main_options = list(SUB_CATEGORIES[type_entry].keys()) if SUB_CATEGORIES.get(type_entry) else ["ทั่วไป"]
                main_cat = st.selectbox("Category", main_options, key="dt_main")
            with c_sub:
                sub_options = SUB_CATEGORIES[type_entry].get(main_cat, ["ทั่วไป"]) if main_cat in SUB_CATEGORIES.get(type_entry, {}) else ["ทั่วไป"]
                sub_cat = st.selectbox("Sub-category", sub_options, key="dt_sub")

            c_date_tool, c_note_tool = st.columns([1, 2])
            with c_date_tool:
                date_shortcut_dt = st.radio("Date", ["วันนี้", "เมื่อวาน", "ระบุเอง"], horizontal=True, key="dt_date_shortcut")
                chosen_date_dt = datetime.date.today() if date_shortcut_dt == "วันนี้" else (datetime.date.today() - datetime.timedelta(days=1) if date_shortcut_dt == "เมื่อวาน" else st.date_input("เลือกวัน", datetime.date.today(), key="dt_date_picker"))

            with st.form("desktop_form", clear_on_submit=True):
                amount = st.number_input("Amount (THB)", min_value=0.0, step=50.0, format="%.2f")
                note = st.text_input("Note", placeholder="...")
                if st.form_submit_button("Save Transaction", use_container_width=True) and amount > 0:
                    full_category = f"{main_cat}: {sub_cat}" if sub_cat != "ทั่วไป" else main_cat
                    sheet.append_row([str(chosen_date_dt), type_entry.split(" ")[1], full_category, amount, note])
                    st.cache_data.clear()
                    st.rerun()

    # --- Tab 2: Dashboard (ดีไซน์ Soft UI) ---
    with tab2:
        if not df.empty:
            df_chart = df.copy()
            df_chart['วันที่'] = pd.to_datetime(df_chart['วันที่'])
            
            # --- กล่องสรุปตัวเลข (Metric Cards) ---
            inc = df_chart[df_chart['ประเภท'] == 'รายรับ']['จำนวนเงิน'].sum()
            exp = df_chart[df_chart['ประเภท'] == 'รายจ่าย']['จำนวนเงิน'].sum()
            sav = df_chart[df_chart['ประเภท'] == 'เงินออม']['จำนวนเงิน'].sum()
            inv = df_chart[df_chart['ประเภท'] == 'เงินลงทุน']['จำนวนเงิน'].sum()
            net = inc - (exp + sav + inv)

            m1, m2, m3, m4, m5 = st.columns(5)
            # เงินสุทธิ
            net_title_class = "metric-title" if net >= 0 else "metric-title-alert"
            m1.markdown(f"<div class='metric-card'><div class='{net_title_class}'>Net Balance</div><div class='metric-value'>฿{net:,.0f}</div><div class='metric-currency'>THB</div></div>", unsafe_allow_html=True)
            # รายรับ
            m2.markdown(f"<div class='metric-card'><div class='metric-title' style='color:{COLOR_TEAL};'>Income ↗</div><div class='metric-value'>฿{inc:,.0f}</div><div class='metric-currency'>THB</div></div>", unsafe_allow_html=True)
            # รายจ่าย
            m3.markdown(f"<div class='metric-card'><div class='metric-title' style='color:{COLOR_ORANGE};'>Expenses ↘</div><div class='metric-value'>฿{exp:,.0f}</div><div class='metric-currency'>THB</div></div>", unsafe_allow_html=True)
            # เงินออม
            m4.markdown(f"<div class='metric-card'><div class='metric-title' style='color:{COLOR_NAVY};'>Savings ↗</div><div class='metric-value'>฿{sav:,.0f}</div><div class='metric-currency'>THB</div></div>", unsafe_allow_html=True)
            # ลงทุน
            m5.markdown(f"<div class='metric-card'><div class='metric-title' style='color:{COLOR_NAVY};'>Investments ↗</div><div class='metric-value'>฿{inv:,.0f}</div><div class='metric-currency'>THB</div></div>", unsafe_allow_html=True)
            
            # --- 📈 กราฟเส้นแบบ Minimal (มีรายวัน, รายเดือน, รายปี) ---
            st.markdown("<p class='quick-add-text' style='margin-top: 20px;'>Trend Analysis</p>", unsafe_allow_html=True)
            time_frame = st.radio("Timeframe:", ["รายวัน", "รายเดือน", "รายปี"], horizontal=True, label_visibility="collapsed")
            
            df_trend = df_chart.copy()
            if time_frame == "รายปี":
                df_trend['เวลา'] = df_trend['วันที่'].dt.strftime('%Y')
            elif time_frame == "รายเดือน":
                df_trend['เวลา'] = df_trend['วันที่'].dt.strftime('%Y-%m')
            else:
                df_trend['เวลา'] = df_trend['วันที่'].dt.strftime('%Y-%m-%d')
                
            trend_data = df_trend.groupby(['เวลา', 'ประเภท'])['จำนวนเงิน'].sum().reset_index()
            
            fig_trend = px.line(trend_data, x='เวลา', y='จำนวนเงิน', color='ประเภท', 
                                color_discrete_map=HONEY_POT_MAP, markers=True, line_shape='spline')
            
            # ปรับแต่งกราฟให้คลีน เข้ากับ Soft UI
            fig_trend.update_traces(line=dict(width=3), marker=dict(size=8, line=dict(width=2, color=COLOR_CARD)))
            fig_trend.update_layout(
                plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
                xaxis=dict(showgrid=False, title="", showline=False, tickfont=dict(color=COLOR_NAVY, family='Poppins')),
                yaxis=dict(showgrid=True, gridcolor='rgba(16, 41, 55, 0.05)', title="", zeroline=False, tickfont=dict(color=COLOR_NAVY, family='Poppins')),
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="center", x=0.5, title="", font=dict(family='Poppins', color=COLOR_NAVY)),
                hovermode="x unified", margin=dict(t=10, b=0, l=0, r=0)
            )
            st.plotly_chart(fig_trend, use_container_width=True)
            
            st.markdown("---")
            
            # --- ⭕ Infographic วงกลมบาง (Thin Circle Chart) ---
            col_chart1, col_chart2 = st.columns([1, 1.2])
            expense_df = df_chart[df_chart['ประเภท'] == 'รายจ่าย']
            
            with col_chart1:
                st.markdown("<p class='quick-add-text'>Expense Breakdown</p>", unsafe_allow_html=True)
                if not expense_df.empty:
                    pie_data = expense_df.groupby('หมวดหมู่หลัก')['จำนวนเงิน'].sum().reset_index()
                    
                    # ใช้สี Honey Pot แบบไล่เฉด (ถ้าหมวดเยอะ ให้ Plotly ช่วยเจนเนอเรตสีในโทนใกล้เคียง)
                    fig_pie = px.pie(pie_data, values='จำนวนเงิน', names='หมวดหมู่หลัก', hole=0.75, 
                                     color_discrete_sequence=[COLOR_ORANGE, COLOR_TEAL, COLOR_NAVY, COLOR_BEIGE, "#d84f2a", "#094044"])
                    
                    fig_pie.update_traces(textposition='outside', textinfo='percent+label', marker=dict(line=dict(width=0)), textfont=dict(family='Poppins', color=COLOR_NAVY))
                    fig_pie.update_layout(plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', showlegend=False, margin=dict(t=20, b=20, l=20, r=20))
                    
                    st.plotly_chart(fig_pie, use_container_width=True)
                else:
                    st.info("No expense data.")
                    
            with col_chart2:
                st.markdown("<p class='quick-add-text'>Top Sub-categories</p>", unsafe_allow_html=True)
                if not expense_df.empty:
                    sub_data = expense_df.groupby(['หมวดหมู่หลัก', 'หมวดหมู่ย่อย'])['จำนวนเงิน'].sum().reset_index()
                    sub_data = sub_data.sort_values(by="จำนวนเงิน", ascending=False).head(8) # โชว์ 8 อันดับแรก
                    
                    fig_bar = px.bar(sub_data, x='จำนวนเงิน', y='หมวดหมู่ย่อย', orientation='h',
                                     color_discrete_sequence=[COLOR_TEAL]) # ใช้สี Teal สำหรับบาร์โค้ด
                    
                    # ทำให้กราฟแท่งขอบมนเล็กน้อย (เทคนิค Plotly) และซ่อนแกน
                    fig_bar.update_traces(marker_line_width=0, opacity=0.9)
                    fig_bar.update_layout(
                        plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
                        xaxis=dict(showgrid=False, title="", zeroline=False, showticklabels=False),
                        yaxis=dict(showgrid=False, title="", autorange="reversed", tickfont=dict(family='Poppins', color=COLOR_NAVY)),
                        showlegend=False, margin=dict(t=20, b=20, l=0, r=20)
                    )
                    # ใส่ตัวเลขกำกับด้านขวาสุดของแท่ง
                    fig_bar.update_traces(texttemplate='฿%{x:,.0f}', textposition='outside', textfont=dict(family='Poppins', color=COLOR_NAVY))
                    
                    st.plotly_chart(fig_bar, use_container_width=True)
                else:
                    st.info("No expense data.")
                    
        else:
            st.info("No data available.")

    with tab3:
        st.subheader("🎯 Goals")
        total_study_savings = df[(df['ประเภท'] == 'เงินออม') & (df['หมวดหมู่หลัก'] == 'ออมเพื่อเรียนต่อ/อนาคต')]['จำนวนเงิน'].sum() if not df.empty else 0
        GOAL_STUDY = 100000 
        progress_percent = min(total_study_savings / GOAL_STUDY, 1.0)
        st.write("✈️ **GRE / Future Studies Fund**")
        st.progress(progress_percent)
        st.caption(f"Saved ฿{total_study_savings:,.2f} of ฿{GOAL_STUDY:,.2f} ({progress_percent*100:.1f}%)")

    with tab4:
        st.subheader("📁 Categories Editor")
        edited_cat = st.data_editor(cat_raw_df, use_container_width=True, num_rows="dynamic", key="editor_cat_v7")
        if st.button("💾 Save Categories", use_container_width=True):
            cat_sheet.clear()
            cat_sheet.update(range_name="A1", values=[edited_cat.columns.values.tolist()] + edited_cat.values.tolist())
            st.success("Categories updated! ✨")
            st.rerun()

        st.markdown("---")
        st.subheader("⚡ Quick Adds Editor")
        edited_qa = st.data_editor(qa_df, use_container_width=True, num_rows="dynamic", key="editor_qa_v7")
        if st.button("💾 Save Quick Adds", use_container_width=True):
            qa_sheet.clear()
            qa_sheet.update(range_name="A1", values=[edited_qa.columns.values.tolist()] + edited_qa.values.tolist())
            st.success("Quick adds updated!")
            st.rerun()
            
        st.markdown("---")
        st.subheader("✏️ Raw Data Editor")
        if not df.empty:
            clean_df_edit = df[["วันที่", "ประเภท", "หมวดหมู่", "จำนวนเงิน", "รายละเอียด"]]
            edited_df = st.data_editor(clean_df_edit, use_container_width=True, num_rows="dynamic", key="editor_finance_v7")
            if st.button("💾 Save Data to Cloud", use_container_width=True):
                sheet.clear()
                edited_df['วันที่'] = edited_df['วันที่'].astype(str)
                sheet.update(range_name="A1", values=[edited_df.columns.values.tolist()] + edited_df.values.tolist())
                st.success("Data updated!")
                st.rerun()
