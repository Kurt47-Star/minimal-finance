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

# 🔤 CSS สไตล์ Soft UI ที่รองรับทั้ง Light & Dark Mode
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;500;600;700&family=Prompt:wght@300;400;500;600&display=swap');
    
    html, body, [class*="css"] { 
        font-family: 'Poppins', 'Prompt', sans-serif !important; 
    }
    
    h1, h2, h3 { font-weight: 700; color: var(--text-color); }
    
    .stButton>button { 
        border-radius: 12px; 
        font-weight: 500; 
        padding: 10px; 
        border: 1px solid var(--border-color);
        background-color: var(--secondary-background-color);
        color: var(--text-color);
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.05);
        transition: all 0.3s ease;
    }
    .stButton>button:hover {
        border-color: #f9744b;
        color: #f9744b;
        transform: translateY(-2px);
    }
    
    .quick-add-text { font-size: 18px; font-weight: 600; margin-bottom: 8px; color: var(--text-color); opacity: 0.9; }
    
    .metric-card { 
        background-color: var(--secondary-background-color); 
        padding: 24px; 
        border-radius: 20px; 
        text-align: left; 
        box-shadow: 0 8px 20px rgba(0, 0, 0, 0.04); 
        border: 1px solid var(--border-color);
        margin-bottom: 1rem;
    }
    .metric-title { font-size: 15px; font-weight: 500; opacity: 0.7; margin-bottom: 5px; color: var(--text-color); }
    .metric-value { color: var(--text-color); font-size: 32px; font-weight: 700; margin: 0; line-height: 1.2; }
    .metric-currency { color: var(--text-color); opacity: 0.5; font-size: 14px; font-weight: 500; margin-top: 5px; }
    
    input[type=number]::-webkit-inner-spin-button, input[type=number]::-webkit-outer-spin-button { -webkit-appearance: none; margin: 0; }
    </style>
""", unsafe_allow_html=True)

st.title("Minimal Finance Pro")

# 🌍 บังคับโซนเวลาแอป
TZ_TH = datetime.timezone(datetime.timedelta(hours=7))

# --- ระบบเชื่อมต่อคลาวด์ (เชื่อมต่อแค่ครั้งเดียว) ---
@st.cache_resource
def init_connection():
    creds_dict = json.loads(st.secrets["google_credentials"])
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    return gspread.authorize(creds)

client = init_connection()
spreadsheet_name = "Minimal Finance Pro"

# 🚀 ระบบ Smart Cache ป้องกัน API Error (Rate Limit) ดึง Worksheet เก็บไว้ในความจำ
@st.cache_resource(ttl=3600)
def get_google_sheets():
    try:
        sh = client.open(spreadsheet_name)
    except Exception:
        return None, None, None, None
        
    sheet_main = sh.sheet1
    
    try:
        sheet_qa = sh.worksheet("QuickAdds")
    except:
        sheet_qa = sh.add_worksheet(title="QuickAdds", rows="50", cols="5")
        sheet_qa.append_row(["ชื่อปุ่ม", "ประเภท", "หมวดหมู่", "จำนวนเงิน"])
        
    try:
        sheet_cat = sh.worksheet("Categories")
    except:
        sheet_cat = sh.add_worksheet(title="Categories", rows="100", cols="3")
        sheet_cat.append_row(["ประเภท", "หมวดหมู่หลัก", "หมวดหมู่ย่อย"])
        
    try:
        sheet_loan = sh.worksheet("Loans")
    except:
        sheet_loan = sh.add_worksheet(title="Loans", rows="10", cols="5")
        sheet_loan.append_row(["เงินต้น", "อัตราดอกเบี้ยปี", "ระยะเวลาเดือน", "งวดที่จ่ายแล้ว", "เดือนปีที่จ่ายล่าสุด"])
        sheet_loan.append_row([10000.0, 15.0, 12, 0, ""])
        
    return sheet_main, sheet_qa, sheet_cat, sheet_loan

sheet, qa_sheet, cat_sheet, loan_sheet = get_google_sheets()

if sheet is None:
    st.error(f"❌ หาไฟล์ Google Sheets ที่ชื่อ '{spreadsheet_name}' ไม่เจอครับ")
    st.info("💡 กรุณาตรวจสอบอีเมล Service Account ว่าได้เปิดสิทธิ์ Editor ในไฟล์ Google Sheets แล้วหรือยังครับ")
    st.stop()

# --- ฟังก์ชันโหลดข้อมูลแยก Cache ---
@st.cache_data(ttl=60)
def fetch_main_data():
    return sheet.get_all_records()

@st.cache_data(ttl=3600)
def fetch_quick_adds():
    return qa_sheet.get_all_records()

@st.cache_data(ttl=3600)
def fetch_categories():
    return cat_sheet.get_all_records()

@st.cache_data(ttl=60)
def fetch_loans():
    return loan_sheet.get_all_records()

def load_data():
    records = fetch_main_data()
    if records:
        df = pd.DataFrame(records)
        parsed_time = pd.to_datetime(df['วันที่'], format='mixed', errors='coerce')
        df['วันเวลา'] = parsed_time.apply(lambda x: x.replace(year=x.year - 543) if pd.notnull(x) and x.year > 2400 else x)
        df['วันที่_date'] = df['วันเวลา'].dt.date
        df['จำนวนเงิน'] = pd.to_numeric(df['จำนวนเงิน'], errors='coerce').fillna(0)
        df['หมวดหมู่หลัก'] = df['หมวดหมู่'].apply(lambda x: str(x).split(":")[0].strip())
        df['หมวดหมู่ย่อย'] = df['หมวดหมู่'].apply(lambda x: str(x).split(":")[1].strip() if ":" in str(x) else "ทั่วไป")
        return df
    return pd.DataFrame(columns=["วันที่", "ประเภท", "หมวดหมู่", "จำนวนเงิน", "รายละเอียด", "หมวดหมู่หลัก", "หมวดหมู่ย่อย", "วันเวลา", "วันที่_date"])

def load_categories():
    records = fetch_categories()
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
qa_records = fetch_quick_adds()
qa_df = pd.DataFrame(qa_records) if qa_records else pd.DataFrame(columns=["ชื่อปุ่ม", "ประเภท", "หมวดหมู่", "จำนวนเงิน"])
cat_raw_df, SUB_CATEGORIES = load_categories()

loan_records = fetch_loans()
if loan_records:
    loan_info = loan_records[0]
    db_principal = float(loan_info["เงินต้น"])
    db_rate = float(loan_info["อัตราดอกเบี้ยปี"])
    db_months = int(loan_info["ระยะเวลาเดือน"])
    current_month_paid = int(loan_info["งวดที่จ่ายแล้ว"])
    db_last_paid_month = str(loan_info["เดือนปีที่จ่ายล่าสุด"]).strip()
else:
    db_principal, db_rate, db_months, current_month_paid, db_last_paid_month = 10000.0, 15.0, 12, 0, ""

sav_dep = df[df['ประเภท'] == 'เงินออม']['จำนวนเงิน'].sum() if not df.empty else 0
sav_withdrawn = df[df['ประเภท'] == 'ถอนเงินออม']['จำนวนเงิน'].sum() if not df.empty else 0
sav_loan = df[df['ประเภท'] == 'กู้เงินออม']['จำนวนเงิน'].sum() if not df.empty else 0
sav_repay = df[df['ประเภท'] == 'คืนเงินกู้ออม']['จำนวนเงิน'].sum() if not df.empty else 0

total_sav_now = sav_dep + sav_repay - sav_withdrawn - sav_loan
outstanding_loan = sav_loan - sav_repay

HONEY_POT_MAP = {
    "รายรับ": "#2a9d8f",     
    "รายจ่าย": "#f9744b",    
    "เงินออม": "#457b9d",    
    "เงินลงทุน": "#e9c46a",
    "เงินสุทธิ": "#8ab17d"   
}
SUB_CAT_PALETTE = ["#124d54", "#f9744b", "#e9c46a", "#2a9d8f", "#457b9d", "#f4a261", "#8ab17d", "#e76f51"]

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
                now_str = datetime.datetime.now(TZ_TH).strftime('%Y-%m-%d %H:%M:%S')
                sheet.append_row([now_str, str(row['ประเภท']), str(row['หมวดหมู่']), float(row['จำนวนเงิน']), "บันทึกด่วน"])
                fetch_main_data.clear() # เคลียร์เฉพาะ Cache ข้อมูลหลัก
                st.toast("Success! ✨")
                st.rerun()
                
    st.markdown("---")
    st.markdown("<p class='quick-add-text'>New Transaction</p>", unsafe_allow_html=True)
    type_entry = st.selectbox("Type", ["💸 รายจ่าย", "📥 รายรับ", "🐷 เงินออม", "📈 เงินลงทุน"])
    
    if "เงินออม" in type_entry:
        sav_action = st.radio("การดำเนินการเงินออม:", ["📥 ฝากเงินเพิ่ม", "🔓 เบิกออกมาใช้", "🎯 กู้เงินคลัง (ต้องคืน)", "🔄 โอนคืนเงินกู้"], horizontal=True)
        st.markdown(f"""
            <div style='background-color: rgba(69, 123, 157, 0.1); border-left: 4px solid #457b9d; padding: 10px 15px; border-radius: 8px; margin-bottom: 10px;'>
                <p style='margin:0; font-size: 13px; opacity: 0.8;'>💰 คลังเงินออมปัจจุบัน: <b>฿{total_sav_now:,.2f}</b></p>
                {"<p style='margin:0; font-size:13px; color:#f9744b;'>⚠️ ยอดหนี้ค้างคืนคลัง: <b>฿" + f"{outstanding_loan:,.2f}</b></p>" if outstanding_loan > 0 else ""}
            </div>
        """, unsafe_allow_html=True)
        main_cat = "บริหารเงินออม"
        sub_cat = sav_action.split(" ")[1]
    else:
        main_options = list(SUB_CATEGORIES[type_entry].keys()) if SUB_CATEGORIES.get(type_entry) else ["ทั่วไป"]
        main_cat = st.selectbox("Category", main_options, key="mb_main")
        sub_options = SUB_CATEGORIES[type_entry].get(main_cat, ["ทั่วไป"]) if main_cat in SUB_CATEGORIES.get(type_entry, {}) else ["ทั่วไป"]
        sub_cat = st.selectbox("Sub-category", sub_options, key="mb_sub")
    
    date_shortcut = st.radio("Date", ["วันนี้", "เมื่อวาน", "ระบุเอง"], horizontal=True)
    chosen_date = datetime.datetime.now(TZ_TH).date() if date_shortcut == "วันนี้" else ((datetime.datetime.now(TZ_TH) - datetime.timedelta(days=1)).date() if date_shortcut == "เมื่อวาน" else st.date_input("เลือกวัน", datetime.datetime.now(TZ_TH).date()))

    with st.form("mobile_form", clear_on_submit=True):
        amount = st.number_input("Amount (THB)", min_value=0.0, step=50.0, format="%.2f", value=None, placeholder="0.00")
        note = st.text_input("Note", placeholder="Optional...")
        if st.form_submit_button("Save Transaction", use_container_width=True) and amount is not None and amount > 0:
            final_type = type_entry.split(" ")[1]
            if final_type == "เงินออม":
                if "เบิกออกมาใช้" in sav_action: final_type = "ถอนเงินออม"
                elif "กู้เงินคลัง" in sav_action: final_type = "กู้เงินออม"
                elif "โอนคืนเงินกู้" in sav_action: final_type = "คืนเงินกู้ออม"
            
            full_category = f"{main_cat}: {sub_cat}" if sub_cat != "ทั่วไป" else main_cat
            combined_datetime = datetime.datetime.combine(chosen_date, datetime.datetime.now(TZ_TH).time())
            sheet.append_row([combined_datetime.strftime('%Y-%m-%d %H:%M:%S'), final_type, full_category, amount, note])
            fetch_main_data.clear() # เคลียร์เฉพาะ Cache ข้อมูลหลัก
            st.rerun()

# ==========================================
# 💻 โหมดคอมพิวเตอร์ (Desktop Mode)
# ==========================================
else:
    tab1, tab2, tab3, tab4, tab5 = st.tabs(["✨ Transaction", "📊 Dashboard", "🎯 Goals", "⚙️ Settings", "🏦 Loan Simulator"])

    with tab1:
        col_main, col_space = st.columns([2, 1])
        with col_main:
            st.markdown("<p class='quick-add-text'>Quick Actions</p>", unsafe_allow_html=True)
            if not qa_df.empty:
                cols = st.columns(4)
                for i, row in qa_df.iterrows():
                    col = cols[i % 4]
                    if col.button(str(row['ชื่อปุ่ม']), use_container_width=True, key=f"dt_qa_{i}"):
                        now_str = datetime.datetime.now(TZ_TH).strftime('%Y-%m-%d %H:%M:%S')
                        sheet.append_row([now_str, str(row['ประเภท']), str(row['หมวดหมู่']), float(row['จำนวนเงิน']), "บันทึกด่วน"])
                        fetch_main_data.clear()
                        st.toast("Success! ✨")
                        st.rerun()
                        
            st.markdown("---")
            st.markdown("<p class='quick-add-text'>New Transaction</p>", unsafe_allow_html=True)
            type_entry = st.radio("Type", ["📥 รายรับ", "💸 รายจ่าย", "🐷 เงินออม", "📈 เงินลงทุน"], horizontal=True, label_visibility="collapsed")
            
            if "เงินออม" in type_entry:
                sav_action = st.radio("การดำเนินการเงินออม:", ["📥 ฝากเงินเพิ่ม", "🔓 เบิกออกมาใช้", "🎯 กู้เงินคลัง (ต้องคืน)", "🔄 โอนคืนเงินกู้"], horizontal=True, key="dt_sav_action")
                st.markdown(f"""
                    <div style='background-color: rgba(69, 123, 157, 0.1); border-left: 4px solid #457b9d; padding: 12px 20px; border-radius: 8px; margin: 10px 0;'>
                        <p style='margin:0; font-size: 14px; opacity: 0.8;'>💰 คลังเงินออมปัจจุบัน: <b>฿{total_sav_now:,.2f}</b></p>
                        {"<p style='margin:0; color: #f9744b; font-size: 14px; font-weight:600;'>⚠️ ยอดหนี้ค้างคืนคลัง: ฿" + f"{outstanding_loan:,.2f}</p>" if outstanding_loan > 0 else ""}
                    </div>
                """, unsafe_allow_html=True)
                main_cat = "บริหารเงินออม"
                sub_cat = sav_action.split(" ")[1]
            else:
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
                chosen_date_dt = datetime.datetime.now(TZ_TH).date() if date_shortcut_dt == "วันนี้" else ((datetime.datetime.now(TZ_TH) - datetime.timedelta(days=1)).date() if date_shortcut_dt == "เมื่อวาน" else st.date_input("เลือกวัน", datetime.datetime.now(TZ_TH).date(), key="dt_date_picker"))

            with st.form("desktop_form", clear_on_submit=True):
                amount = st.number_input("Amount (THB)", min_value=0.0, step=50.0, format="%.2f", value=None, placeholder="0.00")
                note = st.text_input("Note", placeholder="...")
                if st.form_submit_button("Save Transaction", use_container_width=True) and amount is not None and amount > 0:
                    final_type = type_entry.split(" ")[1]
                    if final_type == "เงินออม":
                        if "เบิกออกมาใช้" in sav_action: final_type = "ถอนเงินออม"
                        elif "กู้เงินคลัง" in sav_action: final_type = "กู้เงินออม"
                        elif "โอนคืนเงินกู้" in sav_action: final_type = "คืนเงินกู้ออม"
                        
                    full_category = f"{main_cat}: {sub_cat}" if sub_cat != "ทั่วไป" else main_cat
                    combined_datetime = datetime.datetime.combine(chosen_date_dt, datetime.datetime.now(TZ_TH).time())
                    sheet.append_row([combined_datetime.strftime('%Y-%m-%d %H:%M:%S'), final_type, full_category, amount, note])
                    fetch_main_data.clear()
                    st.rerun()

    with tab2:
        if not df.empty:
            df_chart = df.copy()
            df_chart['วันที่'] = pd.to_datetime(df_chart['วันเวลา'])
            
            inc = df_chart[df_chart['ประเภท'] == 'รายรับ']['จำนวนเงิน'].sum()
            exp = df_chart[df_chart['ประเภท'] == 'รายจ่าย']['จำนวนเงิน'].sum()
            inv = df_chart[df_chart['ประเภท'] == 'เงินลงทุน']['จำนวนเงิน'].sum()
            net = inc + sav_withdrawn + sav_loan - exp - sav_dep - sav_repay

            m1, m2, m3, m4, m5 = st.columns(5)
            net_title_class = "metric-title" if net >= 0 else "metric-title-alert"
            m1.markdown(f"<div class='metric-card'><div class='{net_title_class}'>Net Balance</div><div class='metric-value'>฿{net:,.0f}</div><div class='metric-currency'>THB</div></div>", unsafe_allow_html=True)
            m2.markdown(f"<div class='metric-card'><div class='metric-title'>Income <span style='color:#2a9d8f;'>↗</span></div><div class='metric-value'>฿{inc:,.0f}</div><div class='metric-currency'>THB</div></div>", unsafe_allow_html=True)
            m3.markdown(f"<div class='metric-card'><div class='metric-title'>Expenses <span style='color:#f9744b;'>↘</span></div><div class='metric-value'>฿{exp:,.0f}</div><div class='metric-currency'>THB</div></div>", unsafe_allow_html=True)
            loan_badge = f"<div style='font-size:11px;color:#f9744b;font-weight:600;margin-top:2px;'>⚠️ ค้างคืนคลัง: ฿{outstanding_loan:,.0f}</div>" if outstanding_loan > 0 else ""
            m4.markdown(f"<div class='metric-card'><div class='metric-title'>Savings <span style='color:#457b9d;'>↗</span></div><div class='metric-value'>฿{total_sav_now:,.0f}</div>{loan_badge}<div class='metric-currency'>THB</div></div>", unsafe_allow_html=True)
            m5.markdown(f"<div class='metric-card'><div class='metric-title'>Investments <span style='color:#e9c46a;'>↗</span></div><div class='metric-value'>฿{inv:,.0f}</div><div class='metric-currency'>THB</div></div>", unsafe_allow_html=True)
            
            st.markdown("---")
            
            col_trend_title, col_trend_filter = st.columns([1.5, 2])
            with col_trend_title:
                st.markdown("<p class='quick-add-text' style='margin-top:5px;'>Trend Analysis (Stock Style)</p>", unsafe_allow_html=True)
            with col_trend_filter:
                c_tf, c_ms = st.columns([1.2, 2])
                time_frame = c_tf.selectbox("Timeframe:", ["รายวัน (1D)", "รายสัปดาห์ (1W)", "รายเดือน (1M)", "รายปี (1Y)", "ราย 5 ปี (5Y)"], label_visibility="collapsed")
                visible_metrics = c_ms.multiselect("เลือกเส้นวิเคราะห์คงเหลือ:", ["รายรับ", "รายจ่าย", "เงินออม", "เงินลงทุน", "เงินสุทธิ"], default=["รายรับ", "รายจ่าย", "เงินสุทธิ"])
            
            today = datetime.datetime.now(TZ_TH).date()
            df_trend = df_chart.copy()
            df_trend = df_trend.sort_values(by='วันเวลา')
            
            if "รายวัน" in time_frame:
                df_trend = df_trend[df_trend['วันที่_date'] == today]
                df_trend['เวลา'] = df_trend['วันเวลา'].dt.floor('h')
                x_tick_format = "%H:%M"
            elif "รายสัปดาห์" in time_frame:
                df_trend = df_trend[df_trend['วันที่_date'] >= (today - datetime.timedelta(days=7))]
                df_trend['เวลา'] = df_trend['วันเวลา'].dt.floor('D')
                x_tick_format = "%d %b"
            elif "รายเดือน" in time_frame:
                df_trend = df_trend[df_trend['วันที่_date'] >= (today - datetime.timedelta(days=30))]
                df_trend['เวลา'] = df_trend['วันเวลา'].dt.floor('D')
                x_tick_format = "%d %b"
            elif "รายปี" in time_frame:
                df_trend = df_trend[df_trend['วันที่_date'] >= (today - datetime.timedelta(days=365))]
                df_trend['เวลา'] = df_trend['วันเวลา'].dt.to_period('M').dt.to_timestamp()
                x_tick_format = "%b %Y"
            else:
                df_trend = df_trend[df_trend['วันที่_date'] >= (today - datetime.timedelta(days=365*5))]
                df_trend['เวลา'] = df_trend['วันเวลา'].dt.to_period('Y').dt.to_timestamp()
                x_tick_format = "%Y"
                
            if not df_trend.empty:
                trend_data_raw = df_trend.groupby(['เวลา', 'ประเภท'])['จำนวนเงิน'].sum().reset_index()
                if not trend_data_raw.empty:
                    pivot_trend = trend_data_raw.pivot(index='เวลา', columns='ประเภท', values='จำนวนเงิน').fillna(0)
                    
                    for col in ['รายรับ', 'รายจ่าย', 'เงินออม', 'เงินลงทุน', 'ถอนเงินออม', 'กู้เงินออม', 'คืนเงินกู้ออม']:
                        if col not in pivot_trend.columns: pivot_trend[col] = 0
                    
                    r_inc = pivot_trend['รายรับ']
                    r_exp = pivot_trend['รายจ่าย']
                    r_sav = pivot_trend['เงินออม']
                    r_inv = pivot_trend['เงินลงทุน']
                    r_withdrawn = pivot_trend['ถอนเงินออม']
                    r_loan = pivot_trend['กู้เงินออม']
                    r_repay = pivot_trend['คืนเงินกู้ออม']

                    pivot_trend['รายรับ'] = r_inc
                    pivot_trend['รายจ่าย'] = r_exp
                    pivot_trend['เงินออม'] = r_sav + r_repay - r_withdrawn - r_loan
                    pivot_trend['เงินลงทุน'] = r_inv
                    pivot_trend['เงินสุทธิ'] = r_inc + r_withdrawn + r_loan - r_exp - r_sav - r_inv - r_repay
                    
                    clean_trend_df = pivot_trend[['รายรับ', 'รายจ่าย', 'เงินออม', 'เงินลงทุน', 'เงินสุทธิ']].reset_index().melt(id_vars='เวลา', var_name='ประเภท', value_name='จำนวนเงิน')
                    filtered_trend_df = clean_trend_df[clean_trend_df['ประเภท'].isin(visible_metrics)]
                    
                    if not filtered_trend_df.empty:
                        fig_trend = px.line(filtered_trend_df, x='เวลา', y='จำนวนเงิน', color='ประเภท', color_discrete_map=HONEY_POT_MAP, markers=True, line_shape='spline')
                        fig_trend.update_traces(line=dict(width=2), marker=dict(size=5, line=dict(width=1, color="white")))
                        fig_trend.update_layout(
                            plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', 
                            xaxis=dict(showgrid=False, title="", showline=False, tickformat=x_tick_format, tickfont=dict(family='Poppins', size=11)),
                            yaxis=dict(showgrid=True, gridcolor='rgba(128, 128, 128, 0.08)', title="", zeroline=False, tickfont=dict(family='Poppins', size=11)),
                            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="center", x=0.5, title="", font=dict(family='Poppins', size=12)),
                            hovermode="x unified", margin=dict(t=10, b=0, l=0, r=0)
                        )
                        st.plotly_chart(fig_trend, use_container_width=True, theme="streamlit")
                    else:
                        st.info("กรุณาเลือกเส้นกราฟอย่างน้อย 1 เส้นเพื่อแสดงผล")
            else:
                st.info("ไม่มีข้อมูลการเงินบันทึกไว้ในช่วงไทม์เฟรมนี้")
            
            st.markdown("---")
            
            expense_df = df_chart[df_chart['ประเภท'] == 'รายจ่าย']
            col_exp_title, col_exp_filter = st.columns([2, 1.5])
            with col_exp_title:
                st.markdown("<p class='quick-add-text'>Expense Analysis</p>", unsafe_allow_html=True)
            with col_exp_filter:
                if not expense_df.empty:
                    all_main_cats = sorted(list(expense_df['หมวดหมู่หลัก'].unique()))
                    selected_main_filter = st.selectbox("🔎 เจาะลึกรายละเอียดหมวดหมู่ย่อยด้านขวา:", ["แสดงทั้งหมด"] + all_main_cats)

            col_chart1, col_chart2 = st.columns([1, 1.2])
            if not expense_df.empty:
                pie_data = expense_df.groupby('หมวดหมู่หลัก')['จำนวนเงิน'].sum().reset_index()
                unique_main_cats = pie_data['หมวดหมู่หลัก'].tolist()
                cat_color_map = {cat: SUB_CAT_PALETTE[i % len(SUB_CAT_PALETTE)] for i, cat in enumerate(unique_main_cats)}
                
                with col_chart1:
                    fig_pie = px.pie(pie_data, values='จำนวนเงิน', names='หมวดหมู่หลัก', hole=0.78, color='หมวดหมู่หลัก', color_discrete_map=cat_color_map)
                    fig_pie.update_traces(textposition='outside', textinfo='percent+label', marker=dict(line=dict(width=0)), textfont=dict(family='Poppins', size=11))
                    fig_pie.update_layout(plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', showlegend=False, margin=dict(t=30, b=30, l=30, r=30))
                    st.plotly_chart(fig_pie, use_container_width=True, theme="streamlit")
                        
                with col_chart2:
                    if selected_main_filter == "แสดงทั้งหมด":
                        sub_data = expense_df.groupby(['หมวดหมู่หลัก', 'หมวดหมู่ย่อย'])['จำนวนเงิน'].sum().reset_index()
                        sub_data = sub_data.sort_values(by="จำนวนเงิน", ascending=False).head(8)
                    else:
                        sub_data = expense_df[expense_df['หมวดหมู่หลัก'] == selected_main_filter].groupby(['หมวดหมู่หลัก', 'หมวดหมู่ย่อย'])['จำนวนเงิน'].sum().reset_index()
                        sub_data = sub_data.sort_values(by="จำนวนเงิน", ascending=False)
                    
                    fig_bar = px.bar(sub_data, x='จำนวนเงิน', y='หมวดหมู่ย่อย', color='หมวดหมู่หลัก', orientation='h', color_discrete_map=cat_color_map) 
                    fig_bar.update_traces(marker_line_width=0, opacity=0.9, texttemplate='฿%{x:,.0f}', textposition='outside', textfont=dict(family='Poppins', size=11))
                    fig_bar.update_layout(
                        plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
                        xaxis=dict(showgrid=False, title="", zeroline=False, showticklabels=False),
                        yaxis=dict(showgrid=False, title="", autorange="reversed", tickfont=dict(family='Poppins', size=12)),
                        showlegend=False, margin=dict(t=10, b=10, l=0, r=30)
                    )
                    st.plotly_chart(fig_bar, use_container_width=True, theme="streamlit")
            else:
                st.info("No expense data available.")
        else:
            st.info("No data available.")

    with tab3:
        st.subheader("🎯 Goals")
        progress_percent = min(total_sav_now / 100000, 1.0)
        st.write("✈️ **GRE / Future Studies Fund**")
        st.progress(progress_percent)
        st.caption(f"Saved ฿{total_sav_now:,.2f} of ฿100,000.00 ({progress_percent*100:.1f}%)")

    with tab4:
        st.subheader("📁 Categories Editor")
        edited_cat = st.data_editor(cat_raw_df, use_container_width=True, num_rows="dynamic", key="editor_cat_v14")
        if st.button("💾 Save Categories", use_container_width=True):
            cat_sheet.clear()
            cat_sheet.update(range_name="A1", values=[edited_cat.columns.values.tolist()] + edited_cat.values.tolist())
            fetch_categories.clear()
            st.success("Categories updated! ✨")
            st.rerun()

        st.markdown("---")
        st.subheader("⚡ Quick Adds Editor")
        edited_qa = st.data_editor(qa_df, use_container_width=True, num_rows="dynamic", key="editor_qa_v14")
        if st.button("💾 Save Quick Adds", use_container_width=True):
            qa_sheet.clear()
            qa_sheet.update(range_name="A1", values=[edited_qa.columns.values.tolist()] + edited_qa.values.tolist())
            fetch_quick_adds.clear()
            st.success("Quick adds updated!")
            st.rerun()
            
        st.markdown("---")
        st.subheader("✏️ Raw Data Editor")
        if not df.empty:
            clean_df_edit = df[["วันที่", "ประเภท", "หมวดหมู่", "จำนวนเงิน", "รายละเอียด"]]
            edited_df = st.data_editor(clean_df_edit, use_container_width=True, num_rows="dynamic", key="editor_finance_v14")
            if st.button("💾 Save Data to Cloud", use_container_width=True):
                sheet.clear()
                edited_df['วันที่'] = edited_df['วันที่'].astype(str)
                sheet.update(range_name="A1", values=[edited_df.columns.values.tolist()] + edited_df.values.tolist())
                fetch_main_data.clear()
                st.success("Data updated!")
                st.rerun()

    with tab5:
        st.markdown("<p class='quick-add-text' style='font-size: 22px;'>🏦 เครื่องจำลองสินเชื่อระบบคลาวด์ถาวร (EMI Lock)</p>", unsafe_allow_html=True)
        st.caption("💡 ข้อมูลในหน้านี้จะบันทึกเข้า Google Sheets อัตโนมัติ ไม่หายเมื่อกดรีเฟรชเว็บ")
        
        with st.expander("🛠️ เปิดสัญญา / ปรับปรุงยอดเงินกู้ใหม่"):
            with st.form("loan_setup_form"):
                inp_principal = st.number_input("วงเงินกู้ที่ต้องการ (บาท)", min_value=1000.0, value=None, placeholder=f"ปัจจุบัน: ฿{db_principal:,.0f}", step=1000.0)
                inp_rate = st.number_input("อัตราดอกเบี้ยต่อปี (%)", min_value=0.1, value=None, placeholder=f"ปัจจุบัน: {db_rate}%", step=0.1)
                inp_months = st.number_input("ระยะเวลาสัญญาผ่อน (เดือน)", min_value=1, value=None, placeholder=f"ปัจจุบัน: {db_months} เดือน", step=1)
                
                if st.form_submit_button("💾 อัปเดตสัญญาเงินกู้ลงคลาวด์"):
                    new_p = inp_principal if inp_principal is not None else db_principal
                    new_r = inp_rate if inp_rate is not None else db_rate
                    new_m = inp_months if inp_months is not None else db_months
                    
                    loan_sheet.update_cell(2, 1, new_p)
                    loan_sheet.update_cell(2, 2, new_r)
                    loan_sheet.update_cell(2, 3, new_m)
                    loan_sheet.update_cell(2, 4, 0)
                    loan_sheet.update_cell(2, 5, "")
                    fetch_loans.clear() # เคลียร์เฉพาะ Cache สัญญาเงินกู้
                    st.success("เปิดสัญญาเงินกู้ฉบับใหม่เรียบร้อยครับ!")
                    st.rerun()

        def calculate_emi_schedule(P, annual_r, n):
            r = (annual_r / 100) / 12  
            if r == 0: return P/n, pd.DataFrame()
            emi = P * (r * (1 + r)**n) / ((1 + r)**n - 1)
            schedule = []
            balance = P
            for month in range(1, int(n) + 1):
                interest_payment = balance * r
                principal_payment = emi - interest_payment
                balance -= principal_payment
                schedule.append({
                    "งวดที่": month, "ยอดชำระ (EMI)": emi, "ตัดเงินต้น": principal_payment, "จ่ายดอกเบี้ย": interest_payment, "เงินต้นคงเหลือ": max(0, balance)
                })
            return emi, pd.DataFrame(schedule)

        emi_amount, df_schedule = calculate_emi_schedule(db_principal, db_rate, db_months)
        total_interest = df_schedule['จ่ายดอกเบี้ย'].sum() if not df_schedule.empty else 0
        total_payment = db_principal + total_interest

        st.markdown("---")

        m1, m2, m3, m4 = st.columns(4)
        m1.markdown(f"<div class='metric-card'><div class='metric-title'>ยอดผ่อนต่อเดือน (EMI)</div><div class='metric-value' style='color:#f9744b;'>฿{emi_amount:,.2f}</div></div>", unsafe_allow_html=True)
        m2.markdown(f"<div class='metric-card'><div class='metric-title'>เงินต้นคงค้างระบบ</div><div class='metric-value'>฿{db_principal:,.0f}</div></div>", unsafe_allow_html=True)
        m3.markdown(f"<div class='metric-card'><div class='metric-title'>ดอกเบี้ยทั้งสัญญา</div><div class='metric-value' style='color:#e9c46a;'>฿{total_interest:,.2f}</div></div>", unsafe_allow_html=True)
        m4.markdown(f"<div class='metric-card'><div class='metric-title'>รวมชำระตลอดสัญญา</div><div class='metric-value' style='color:#f9744b;'>฿{total_payment:,.2f}</div></div>", unsafe_allow_html=True)

        st.markdown("<p class='quick-add-text'>🔄 ชำระค่างวดประจำเดือน</p>", unsafe_allow_html=True)

        current_real_month = datetime.datetime.now(TZ_TH).strftime("%Y-%m")
        is_paid_this_month = (db_last_paid_month == current_real_month)

        col_pay, col_info_lock = st.columns([1.5, 3.5])
        
        with col_pay:
            if current_month_paid >= db_months:
                st.button("🎉 ผ่อนชำระครบสัญญาแล้ว", disabled=True, use_container_width=True)
            elif is_paid_this_month:
                st.button("🔒 ล็อก! จ่ายงวดของเดือนนี้แล้ว", disabled=True, use_container_width=True)
            else:
                if st.button("💸 เช็คบิลจ่ายงวดประจำเดือนนี้", use_container_width=True):
                    next_paid_count = current_month_paid + 1
                    loan_sheet.update_cell(2, 4, next_paid_count)
                    loan_sheet.update_cell(2, 5, current_real_month)
                    fetch_loans.clear() # เคลียร์เฉพาะ Cache สัญญาเงินกู้
                    st.toast(f"ชำระงวดที่ {next_paid_count} สำเร็จ! ข้อมูลซิงค์ขึ้นคลาวด์แล้ว ✨")
                    st.rerun()
                    
        with col_info_lock:
            if is_paid_this_month and current_month_paid < db_months:
                st.info(f"ระบบตรวจพบสถานะความปลอดภัย: งวดที่ {current_month_paid} ถูกตัดบัญชีไปเมื่อเดือน {db_last_paid_month} เรียบร้อยแล้ว ปุ่มชำระเงินจะเปิดให้กดใหม่อัตโนมัติเมื่อขึ้นเดือนถัดไปครับ")
            elif current_month_paid >= db_months:
                st.success("สัญญาเงินกู้ฉบับนี้เสร็จสิ้นอย่างสมบูรณ์แบบเรียบร้อยแล้ว!")

        progress_pct = current_month_paid / db_months if db_months > 0 else 0
        st.progress(progress_pct)
        st.caption(f"ชำระไปแล้ว {current_month_paid} งวด จากทั้งหมด {db_months} งวด ({progress_pct*100:.1f}%)")

        st.markdown("---")
        st.markdown("<p class='quick-add-text'>📋 ตารางแจกแจงการผ่อนชำระคลาวด์ (Amortization Schedule)</p>", unsafe_allow_html=True)

        if not df_schedule.empty:
            df_display = df_schedule.copy()
            for col in ['ยอดชำระ (EMI)', 'ตัดเงินต้น', 'จ่ายดอกเบี้ย', 'เงินต้นคงเหลือ']:
                df_display[col] = df_display[col].apply(lambda x: f"฿ {x:,.2f}")

            def highlight_paid(row):
                if row.name < current_month_paid:
                    return ['background-color: rgba(42, 157, 143, 0.1); color: #2a9d8f; font-weight: bold'] * len(row)
                return [''] * len(row)

            st.dataframe(df_display.style.apply(highlight_paid, axis=1), use_container_width=True, hide_index=True)
