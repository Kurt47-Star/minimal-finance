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
    
    .calib-box-match {
        background-color: rgba(42, 157, 143, 0.12);
        border: 1px solid #2a9d8f;
        border-radius: 16px;
        padding: 16px 20px;
        margin-bottom: 16px;
    }
    .calib-box-diff {
        background-color: rgba(249, 116, 75, 0.12);
        border: 1px solid #f9744b;
        border-radius: 16px;
        padding: 16px 20px;
        margin-bottom: 16px;
    }
    
    input[type=number]::-webkit-inner-spin-button, input[type=number]::-webkit-outer-spin-button { -webkit-appearance: none; margin: 0; }
    </style>
""", unsafe_allow_html=True)

st.title("Minimal Finance Pro")

# 🌍 บังคับโซนเวลาแอปให้อยู่ในเขตประเทศไทย (UTC+7)
TZ_TH = datetime.timezone(datetime.timedelta(hours=7))

# --- ระบบเชื่อมต่อคลาวด์ ---
@st.cache_resource
def init_connection():
    creds_dict = json.loads(st.secrets["google_credentials"])
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
    return gspread.authorize(creds)

client = init_connection()
spreadsheet_name = "Minimal Finance Pro"

# 🚀 ระบบ Smart Cache พร้อมสร้างตาราง Goals, Receivables และรองรับ Multi-Wallet
@st.cache_resource(ttl=3600)
def get_google_sheets():
    try:
        sh = client.open(spreadsheet_name)
    except Exception:
        return None, None, None, None, None, None, None
        
    sheet_main = sh.sheet1
    try:
        sheet_main.resize(rows=500, cols=8)
        headers = sheet_main.row_values(1)
        if len(headers) < 6 or headers[5] != "กระเป๋า":
            sheet_main.update_cell(1, 6, "กระเป๋า")
    except Exception:
        pass
    
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

    try:
        sheet_cycle = sh.worksheet("Cycles")
        try:
            sheet_cycle.resize(rows=50, cols=10)
            headers = sheet_cycle.row_values(1)
            if len(headers) < 6:
                sheet_cycle.update_cell(1, 5, "ยอดยกมา")
                sheet_cycle.update_cell(1, 6, "เงินจริงกรุงไทย")
        except Exception:
            pass
    except:
        sheet_cycle = sh.add_worksheet(title="Cycles", rows="50", cols="10")
        sheet_cycle.append_row(["ชื่อรอบบัญชี", "เริ่มต้น", "สิ้นสุด", "สถานะ", "ยอดยกมา", "เงินจริงกรุงไทย"])
        sheet_cycle.append_row(["July 2026", "2026-06-25 00:00:00", "2026-08-01 22:34:23", "CLOSED", 0.0, 2501.0])
        sheet_cycle.append_row(["August 2026", "2026-08-01 22:34:24", "", "ACTIVE", 2501.0, 2501.0])

    try:
        sheet_debt = sh.worksheet("Receivables")
    except:
        sheet_debt = sh.add_worksheet(title="Receivables", rows="50", cols="8")
        sheet_debt.append_row(["ID", "ชื่อคนติดเงิน", "รายการ/รายละเอียด", "จำนวนเงิน", "กระเป๋าที่จ่าย", "วันที่สร้าง", "สถานะ", "วันที่คืน"])

    try:
        sheet_goal = sh.worksheet("Goals")
    except:
        sheet_goal = sh.add_worksheet(title="Goals", rows="30", cols="5")
        sheet_goal.append_row(["ไอคอน", "ชื่อเป้าหมาย", "เป้าหมาย (บาท)", "สะสมแล้ว (บาท)"])
        sheet_goal.append_row(["✈️", "GRE / Future Studies Fund", 100000.0, 0.0])
        
    return sheet_main, sheet_qa, sheet_cat, sheet_loan, sheet_cycle, sheet_debt, sheet_goal

sheet, qa_sheet, cat_sheet, loan_sheet, cycle_sheet, debt_sheet, goal_sheet = get_google_sheets()

if sheet is None:
    st.error(f"❌ หาไฟล์ Google Sheets ที่ชื่อ '{spreadsheet_name}' ไม่เจอครับ")
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

@st.cache_data(ttl=60)
def fetch_cycles():
    return cycle_sheet.get_all_records()

@st.cache_data(ttl=60)
def fetch_receivables():
    return debt_sheet.get_all_records()

@st.cache_data(ttl=60)
def fetch_goals():
    return goal_sheet.get_all_records()

def parse_custom_time(time_str, default_time):
    try:
        parts = str(time_str).strip().split(':')
        if len(parts) == 3:
            return datetime.time(int(parts[0]), int(parts[1]), int(parts[2]))
        elif len(parts) == 2:
            return datetime.time(int(parts[0]), int(parts[1]), 0)
    except Exception:
        pass
    return default_time

def load_data():
    records = fetch_main_data()
    if records:
        df = pd.DataFrame(records)
        parsed_time = pd.to_datetime(df['วันที่'], format='mixed', errors='coerce')
        df['วันเวลา'] = parsed_time.apply(lambda x: x.replace(year=x.year - 543) if pd.notnull(x) and x.year > 2400 else x)
        df['วันที่_date'] = df['วันเวลา'].dt.date
        df['จำนวนเงิน'] = pd.to_numeric(df['จำนวนเงิน'], errors='coerce').fillna(0)
        df['หมวดหมู่หลัก'] = df['หมวดหมู่'].apply(lambda x: str(x).split(":")[0].strip() if pd.notnull(x) else "ทั่วไป")
        df['หมวดหมู่ย่อย'] = df['หมวดหมู่'].apply(lambda x: str(x).split(":")[1].strip() if pd.notnull(x) and ":" in str(x) else "ทั่วไป")
        if 'กระเป๋า' not in df.columns:
            df['กระเป๋า'] = '🏦 กรุงไทย'
        else:
            df['กระเป๋า'] = df['กระเป๋า'].fillna('🏦 กรุงไทย').replace('', '🏦 กรุงไทย')
        return df
    return pd.DataFrame(columns=["วันที่", "ประเภท", "หมวดหมู่", "จำนวนเงิน", "รายละเอียด", "กระเป๋า", "หมวดหมู่หลัก", "หมวดหมู่ย่อย", "วันเวลา", "วันที่_date"])

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

# 📌 โหลดข้อมูลเป้าหมายออมเงิน (Goals) เตรียมไว้สำหรับ Slot Selector
goals_data = fetch_goals()
df_goals = pd.DataFrame(goals_data) if goals_data else pd.DataFrame(columns=["ไอคอน", "ชื่อเป้าหมาย", "เป้าหมาย (บาท)", "สะสมแล้ว (บาท)"])
goal_options_list = ["📦 คลังออมทั่วไป (ไม่ระบุเป้าหมาย)"] + (df_goals["ชื่อเป้าหมาย"].tolist() if not df_goals.empty else [])

# คำนวณสรุปคลังเงินออมสะสมทั้งหมดล่วงหน้า
sav_dep_global = df[df['ประเภท'] == 'เงินออม']['จำนวนเงิน'].sum() if not df.empty else 0
sav_withdrawn_global = df[df['ประเภท'] == 'ถอนเงินออม']['จำนวนเงิน'].sum() if not df.empty else 0
sav_loan_global = df[df['ประเภท'] == 'กู้เงินออม']['จำนวนเงิน'].sum() if not df.empty else 0
sav_repay_global = df[df['ประเภท'] == 'คืนเงินกู้ออม']['จำนวนเงิน'].sum() if not df.empty else 0

total_sav_now = sav_dep_global + sav_repay_global - sav_withdrawn_global - sav_loan_global
outstanding_loan = sav_loan_global - sav_repay_global

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
                sheet.append_row([now_str, str(row['ประเภท']), str(row['หมวดหมู่']), float(row['จำนวนเงิน']), "บันทึกด่วน", "🏦 กรุงไทย"])
                fetch_main_data.clear()
                st.toast("Success! ✨")
                st.rerun()
                
    st.markdown("---")
    st.markdown("<p class='quick-add-text'>New Transaction</p>", unsafe_allow_html=True)
    
    type_entry = st.selectbox("Type", ["💸 รายจ่าย", "📥 รายรับ", "🔄 โอนย้ายกระเป๋า", "🐷 เงินออม", "📈 เงินลงทุน"])
    wallet_entry = st.selectbox("กระเป๋าเงิน (Wallet)", ["🏦 กรุงไทย", "📱 TrueMoney Wallet"])
    
    if "เงินออม" in type_entry:
        sav_action = st.radio("การดำเนินการเงินออม:", ["📥 ฝากเงินเพิ่ม", "🔓 เบิกออกมาใช้", "🎯 กู้เงินคลัง (ต้องคืน)", "🔄 โอนคืนเงินกู้"], horizontal=True)
        selected_goal_mb = st.selectbox("🎯 เลือกเป้าหมายออมเงิน (Slot):", goal_options_list, key="mb_goal_slot")
        
        st.markdown(f"""
            <div style='background-color: rgba(69, 123, 157, 0.1); border-left: 4px solid #457b9d; padding: 10px 15px; border-radius: 8px; margin-bottom: 10px;'>
                <p style='margin:0; font-size: 13px; opacity: 0.8;'>💰 คลังเงินออมปัจจุบัน: <b>฿{total_sav_now:,.2f}</b></p>
                {"<p style='margin:0; font-size:13px; color:#f9744b;'>⚠️ ยอดหนี้ค้างคืนคลัง: <b>฿" + f"{outstanding_loan:,.2f}</b></p>" if outstanding_loan > 0 else ""}
            </div>
        """, unsafe_allow_html=True)
        main_cat = "บริหารเงินออม"
        action_name = sav_action.split(" ")[1]
        sub_cat = f"{action_name} [{selected_goal_mb}]" if selected_goal_mb != "📦 คลังออมทั่วไป (ไม่ระบุเป้าหมาย)" else action_name
    elif "โอนย้ายกระเป๋า" in type_entry:
        transfer_dir = st.radio("ทิศทางการโอน:", ["🏦 กรุงไทย ➡️ 📱 TrueMoney", "📱 TrueMoney ➡️ 🏦 กรุงไทย"], horizontal=True)
        main_cat = "โอนย้ายระหว่างกระเป๋า"
        sub_cat = "เข้า TrueMoney" if "TrueMoney" in transfer_dir.split("➡️")[1] else "เข้ากรุงไทย"
    else:
        main_options = sorted(list(SUB_CATEGORIES[type_entry].keys())) if SUB_CATEGORIES.get(type_entry) else ["ทั่วไป"]
        main_cat = st.selectbox("Category", main_options, key="mb_main")
        sub_options = sorted(SUB_CATEGORIES[type_entry].get(main_cat, ["ทั่วไป"])) if main_cat in SUB_CATEGORIES.get(type_entry, {}) else ["ทั่วไป"]
        sub_cat = st.selectbox("Sub-category", sub_options, key="mb_sub")
    
    c_md1, c_mt1 = st.columns(2)
    with c_md1:
        date_shortcut = st.radio("วันที่ (Date)", ["วันนี้", "เมื่อวาน", "ระบุเอง"], horizontal=True, key="mb_date_mode")
        chosen_date = datetime.datetime.now(TZ_TH).date() if date_shortcut == "วันนี้" else ((datetime.datetime.now(TZ_TH) - datetime.timedelta(days=1)).date() if date_shortcut == "เมื่อวาน" else st.date_input("เลือกวัน", datetime.datetime.now(TZ_TH).date(), key="mb_date_picker"))
    with c_mt1:
        time_shortcut = st.radio("เวลา (Time)", ["⏱️ เวลาปัจจุบัน", "⏰ พิมพ์ระบุเอง"], horizontal=True, key="mb_time_mode")
        if time_shortcut == "⏱️ เวลาปัจจุบัน":
            chosen_time_str = datetime.datetime.now(TZ_TH).strftime('%H:%M:%S')
            st.text_input("เวลา", value=chosen_time_str, disabled=True, key="mb_time_show")
        else:
            chosen_time_str = st.text_input("⏰ พิมพ์เวลา (เช่น 22:34:23)", value=datetime.datetime.now(TZ_TH).strftime('%H:%M:%S'), placeholder="HH:MM:SS", key="mb_time_type")

    with st.form("mobile_form", clear_on_submit=True):
        amount = st.number_input("Amount (THB)", min_value=0.0, step=50.0, format="%.2f", value=None, placeholder="0.00")
        note = st.text_input("Note", placeholder="Optional...")
        if st.form_submit_button("Save Transaction", use_container_width=True) and amount is not None and amount > 0:
            final_type = type_entry.split(" ")[1]
            if final_type == "เงินออม":
                if "เบิกออกมาใช้" in sav_action: final_type = "ถอนเงินออม"
                elif "กู้เงินคลัง" in sav_action: final_type = "กู้เงินออม"
                elif "โอนคืนเงินกู้" in sav_action: final_type = "คืนเงินกู้ออม"
                
                if selected_goal_mb != "📦 คลังออมทั่วไป (ไม่ระบุเป้าหมาย)" and not df_goals.empty:
                    for g_idx, g_row in df_goals.iterrows():
                        if str(g_row["ชื่อเป้าหมาย"]) == selected_goal_mb:
                            curr_saved = float(g_row["สะสมแล้ว (บาท)"]) if pd.notnull(g_row["สะสมแล้ว (บาท)"]) else 0.0
                            if final_type in ["เงินออม", "คืนเงินกู้ออม"]:
                                new_saved = curr_saved + float(amount)
                            else:
                                new_saved = max(0.0, curr_saved - float(amount))
                            goal_sheet.update_cell(int(g_idx) + 2, 4, new_saved)
                            fetch_goals.clear()
                            break
            
            full_category = f"{main_cat}: {sub_cat}" if sub_cat != "ทั่วไป" else main_cat
            final_time = datetime.datetime.now(TZ_TH).time() if time_shortcut == "⏱️ เวลาปัจจุบัน" else parse_custom_time(chosen_time_str, datetime.datetime.now(TZ_TH).time())
            combined_datetime = datetime.datetime.combine(chosen_date, final_time)
            
            sheet.append_row([combined_datetime.strftime('%Y-%m-%d %H:%M:%S'), final_type, full_category, amount, note, wallet_entry])
            fetch_main_data.clear()
            st.rerun()

# ==========================================
# 💻 โหมดคอมพิวเตอร์ (Desktop Mode - 6 Tabs)
# ==========================================
else:
    tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(["✨ Transaction", "📊 Dashboard", "🤝 ลูกหนี้ & หารบิล", "🎯 Goals", "⚙️ Settings", "🏦 Loan Simulator"])

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
                        sheet.append_row([now_str, str(row['ประเภท']), str(row['หมวดหมู่']), float(row['จำนวนเงิน']), "บันทึกด่วน", "🏦 กรุงไทย"])
                        fetch_main_data.clear()
                        st.toast("Success! ✨")
                        st.rerun()
                        
            st.markdown("---")
            st.markdown("<p class='quick-add-text'>New Transaction</p>", unsafe_allow_html=True)
            
            c_type, c_wallet = st.columns([3, 1.5])
            with c_type:
                type_entry = st.radio("Type", ["📥 รายรับ", "💸 รายจ่าย", "🔄 โอนย้ายกระเป๋า", "🐷 เงินออม", "📈 เงินลงทุน"], horizontal=True, label_visibility="collapsed")
            with c_wallet:
                wallet_entry = st.selectbox("กระเป๋าเงิน (Wallet):", ["🏦 กรุงไทย", "📱 TrueMoney Wallet"], label_visibility="collapsed")
            
            if "เงินออม" in type_entry:
                sav_action = st.radio("การดำเนินการเงินออม:", ["📥 ฝากเงินเพิ่ม", "🔓 เบิกออกมาใช้", "🎯 กู้เงินคลัง (ต้องคืน)", "🔄 โอนคืนเงินกู้"], horizontal=True, key="dt_sav_action")
                selected_goal_dt = st.selectbox("🎯 เลือกเป้าหมายออมเงิน (Slot):", goal_options_list, key="dt_goal_slot")
                
                st.markdown(f"""
                    <div style='background-color: rgba(69, 123, 157, 0.1); border-left: 4px solid #457b9d; padding: 12px 20px; border-radius: 8px; margin: 10px 0;'>
                        <p style='margin:0; font-size: 14px; opacity: 0.8;'>💰 คลังเงินออมปัจจุบัน: <b>฿{total_sav_now:,.2f}</b></p>
                        {"<p style='margin:0; color: #f9744b; font-size: 14px; font-weight:600;'>⚠️ ยอดหนี้ค้างคืนคลัง: ฿" + f"{outstanding_loan:,.2f}</p>" if outstanding_loan > 0 else ""}
                    </div>
                """, unsafe_allow_html=True)
                main_cat = "บริหารเงินออม"
                action_name = sav_action.split(" ")[1]
                sub_cat = f"{action_name} [{selected_goal_dt}]" if selected_goal_dt != "📦 คลังออมทั่วไป (ไม่ระบุเป้าหมาย)" else action_name
            elif "โอนย้ายกระเป๋า" in type_entry:
                transfer_dir = st.radio("ทิศทางการโอน:", ["🏦 กรุงไทย ➡️ 📱 TrueMoney", "📱 TrueMoney ➡️ 🏦 กรุงไทย"], horizontal=True)
                main_cat = "โอนย้ายระหว่างกระเป๋า"
                sub_cat = "เข้า TrueMoney" if "TrueMoney" in transfer_dir.split("➡️")[1] else "เข้ากรุงไทย"
            else:
                c_main, c_sub = st.columns(2)
                with c_main:
                    main_options = sorted(list(SUB_CATEGORIES[type_entry].keys())) if SUB_CATEGORIES.get(type_entry) else ["ทั่วไป"]
                    main_cat = st.selectbox("Category", main_options, key="dt_main")
                with c_sub:
                    sub_options = sorted(SUB_CATEGORIES[type_entry].get(main_cat, ["ทั่วไป"])) if main_cat in SUB_CATEGORIES.get(type_entry, {}) else ["ทั่วไป"]
                    sub_cat = st.selectbox("Sub-category", sub_options, key="dt_sub")

            c_date_tool, c_time_tool = st.columns([1, 1])
            with c_date_tool:
                date_shortcut_dt = st.radio("วันที่ (Date)", ["วันนี้", "เมื่อวาน", "ระบุเอง"], horizontal=True, key="dt_date_shortcut")
                chosen_date_dt = datetime.datetime.now(TZ_TH).date() if date_shortcut_dt == "วันนี้" else ((datetime.datetime.now(TZ_TH) - datetime.timedelta(days=1)).date() if date_shortcut_dt == "เมื่อวาน" else st.date_input("เลือกวัน", datetime.datetime.now(TZ_TH).date(), key="dt_date_picker"))
            with c_time_tool:
                time_shortcut_dt = st.radio("เวลา (Time)", ["⏱️ เวลาปัจจุบัน", "⏰ พิมพ์ระบุเอง"], horizontal=True, key="dt_time_shortcut")
                if time_shortcut_dt == "⏱️ เวลาปัจจุบัน":
                    chosen_time_dt_str = datetime.datetime.now(TZ_TH).strftime('%H:%M:%S')
                    st.text_input("เวลา", value=chosen_time_dt_str, disabled=True, key="dt_time_show")
                else:
                    chosen_time_dt_str = st.text_input("⏰ พิมพ์เวลา (เช่น 22:34:23)", value=datetime.datetime.now(TZ_TH).strftime('%H:%M:%S'), placeholder="HH:MM:SS", key="dt_time_type")

            with st.form("desktop_form", clear_on_submit=True):
                amount = st.number_input("Amount (THB)", min_value=0.0, step=50.0, format="%.2f", value=None, placeholder="0.00")
                note = st.text_input("Note", placeholder="...")
                if st.form_submit_button("Save Transaction", use_container_width=True) and amount is not None and amount > 0:
                    final_type = type_entry.split(" ")[1]
                    if final_type == "เงินออม":
                        if "เบิกออกมาใช้" in sav_action: final_type = "ถอนเงินออม"
                        elif "กู้เงินคลัง" in sav_action: final_type = "กู้เงินออม"
                        elif "โอนคืนเงินกู้" in sav_action: final_type = "คืนเงินกู้ออม"
                        
                        if selected_goal_dt != "📦 คลังออมทั่วไป (ไม่ระบุเป้าหมาย)" and not df_goals.empty:
                            for g_idx, g_row in df_goals.iterrows():
                                if str(g_row["ชื่อเป้าหมาย"]) == selected_goal_dt:
                                    curr_saved = float(g_row["สะสมแล้ว (บาท)"]) if pd.notnull(g_row["สะสมแล้ว (บาท)"]) else 0.0
                                    if final_type in ["เงินออม", "คืนเงินกู้ออม"]:
                                        new_saved = curr_saved + float(amount)
                                    else:
                                        new_saved = max(0.0, curr_saved - float(amount))
                                    goal_sheet.update_cell(int(g_idx) + 2, 4, new_saved)
                                    fetch_goals.clear()
                                    break
                        
                    full_category = f"{main_cat}: {sub_cat}" if sub_cat != "ทั่วไป" else main_cat
                    final_time_dt = datetime.datetime.now(TZ_TH).time() if time_shortcut_dt == "⏱️ เวลาปัจจุบัน" else parse_custom_time(chosen_time_dt_str, datetime.datetime.now(TZ_TH).time())
                    combined_datetime = datetime.datetime.combine(chosen_date_dt, final_time_dt)
                    
                    sheet.append_row([combined_datetime.strftime('%Y-%m-%d %H:%M:%S'), final_type, full_category, amount, note, wallet_entry])
                    fetch_main_data.clear()
                    st.rerun()

    # ==========================================
    # 📊 Tab 2: Dashboard
    # ==========================================
    with tab2:
        if not df.empty:
            df_chart = df.copy()
            df_chart['วันที่'] = pd.to_datetime(df_chart['วันเวลา'])
            
            cycles_data = fetch_cycles()
            df_cycles = pd.DataFrame(cycles_data) if cycles_data else pd.DataFrame(columns=["ชื่อรอบบัญชี", "เริ่มต้น", "สิ้นสุด", "สถานะ", "ยอดยกมา", "เงินจริงกรุงไทย"])
            
            cycle_options = []
            active_cycle_name = "รอบปัจจุบัน"
            active_row_idx = None
            active_carry_forward = 0.0
            
            for idx, row in df_cycles.iterrows():
                c_name = str(row['ชื่อรอบบัญชี']).strip()
                c_status = str(row['สถานะ']).strip()
                c_start = str(row['เริ่มต้น']).strip()
                c_end = str(row['สิ้นสุด']).strip()
                c_carry = float(row.get('ยอดยกมา', 0.0)) if pd.notnull(row.get('ยอดยกมา')) and str(row.get('ยอดยกมา')).strip() != "" else 0.0
                c_kt_real = float(row.get('เงินจริงกรุงไทย', 0.0)) if pd.notnull(row.get('เงินจริงกรุงไทย')) and str(row.get('เงินจริงกรุงไทย')).strip() != "" else 0.0
                
                if c_status == "ACTIVE":
                    active_cycle_name = c_name
                    active_row_idx = idx + 2
                    active_carry_forward = c_carry
                    cycle_options.append((f"🟢 {c_name} (เริ่ม {c_start[:10]})", c_start, None, c_carry, c_kt_real, idx + 2))
                else:
                    cycle_options.append((f"📅 {c_name} ({c_start[:10]} - {c_end[:10]})", c_start, c_end, c_carry, c_kt_real, idx + 2))
            
            cycle_options.reverse()
            cycle_labels = [opt[0] for opt in cycle_options] + ["🌟 แสดงข้อมูลทั้งหมด (All Time)"]
            
            col_dash_title, col_cycle_select = st.columns([1, 2])
            with col_dash_title:
                st.markdown("<p class='quick-add-text' style='margin-top:5px;'>📊 Overview (รอบบัญชี)</p>", unsafe_allow_html=True)
            with col_cycle_select:
                selected_cycle_label = st.selectbox("⏳ เลือก Circle ในการแสดงผล:", cycle_labels, index=0, label_visibility="collapsed")
            
            df_dash = df_chart.copy()
            selected_carry = 0.0
            selected_kt_real = 0.0
            selected_row_idx = None
            
            if selected_cycle_label != "🌟 แสดงข้อมูลทั้งหมด (All Time)":
                for label, start_str, end_str, carry_val, kt_val, r_idx in cycle_options:
                    if label == selected_cycle_label:
                        selected_carry = carry_val
                        selected_kt_real = kt_val
                        selected_row_idx = r_idx
                        start_dt = pd.to_datetime(start_str)
                        if end_str and pd.notnull(end_str) and end_str != "":
                            end_dt = pd.to_datetime(end_str)
                            df_dash = df_dash[(df_dash['วันที่'] >= start_dt) & (df_dash['วันที่'] <= end_dt)]
                        else:
                            df_dash = df_dash[df_dash['วันที่'] >= start_dt]
                        break

            inc = df_dash[df_dash['ประเภท'] == 'รายรับ']['จำนวนเงิน'].sum()
            exp = df_dash[df_dash['ประเภท'] == 'รายจ่าย']['จำนวนเงิน'].sum()
            inv = df_dash[df_dash['ประเภท'] == 'เงินลงทุน']['จำนวนเงิน'].sum()
            
            sav_dep_d = df_dash[df_dash['ประเภท'] == 'เงินออม']['จำนวนเงิน'].sum()
            sav_withdrawn_d = df_dash[df_dash['ประเภท'] == 'ถอนเงินออม']['จำนวนเงิน'].sum()
            sav_loan_d = df_dash[df_dash['ประเภท'] == 'กู้เงินออม']['จำนวนเงิน'].sum()
            sav_repay_d = df_dash[df_dash['ประเภท'] == 'คืนเงินกู้ออม']['จำนวนเงิน'].sum()

            lend_d = df_dash[df_dash['ประเภท'] == '🤝 เงินทดจ่าย']['จำนวนเงิน'].sum()
            refund_d = df_dash[df_dash['ประเภท'] == '🤝 รับคืนเงินทดจ่าย']['จำนวนเงิน'].sum()

            net_in_cycle = selected_carry + inc + sav_withdrawn_d + sav_loan_d + refund_d - exp - inv - sav_dep_d - sav_repay_d - lend_d
            sav_flow = sav_dep_d + sav_repay_d - sav_withdrawn_d - sav_loan_d

            tm_inc = df_dash[(df_dash['กระเป๋า'] == '📱 TrueMoney Wallet') & (df_dash['ประเภท'] == 'รายรับ')]['จำนวนเงิน'].sum()
            tm_exp = df_dash[(df_dash['กระเป๋า'] == '📱 TrueMoney Wallet') & (df_dash['ประเภท'] == 'รายจ่าย')]['จำนวนเงิน'].sum()
            tm_inv = df_dash[(df_dash['กระเป๋า'] == '📱 TrueMoney Wallet') & (df_dash['ประเภท'] == 'เงินลงทุน')]['จำนวนเงิน'].sum()
            tm_transfer_in = df_dash[df_dash['หมวดหมู่'].str.contains('เข้า TrueMoney', na=False)]['จำนวนเงิน'].sum()
            tm_transfer_out = df_dash[df_dash['หมวดหมู่'].str.contains('เข้ากรุงไทย', na=False)]['จำนวนเงิน'].sum()
            tm_lend = df_dash[(df_dash['กระเป๋า'] == '📱 TrueMoney Wallet') & (df_dash['ประเภท'] == '🤝 เงินทดจ่าย')]['จำนวนเงิน'].sum()
            tm_refund = df_dash[(df_dash['กระเป๋า'] == '📱 TrueMoney Wallet') & (df_dash['ประเภท'] == '🤝 รับคืนเงินทดจ่าย')]['จำนวนเงิน'].sum()
            tm_sav_dep = df_dash[(df_dash['กระเป๋า'] == '📱 TrueMoney Wallet') & (df_dash['ประเภท'] == 'เงินออม')]['จำนวนเงิน'].sum()
            tm_sav_withdrawn = df_dash[(df_dash['กระเป๋า'] == '📱 TrueMoney Wallet') & (df_dash['ประเภท'] == 'ถอนเงินออม')]['จำนวนเงิน'].sum()
            tm_sav_loan = df_dash[(df_dash['กระเป๋า'] == '📱 TrueMoney Wallet') & (df_dash['ประเภท'] == 'กู้เงินออม')]['จำนวนเงิน'].sum()
            tm_sav_repay = df_dash[(df_dash['กระเป๋า'] == '📱 TrueMoney Wallet') & (df_dash['ประเภท'] == 'คืนเงินกู้ออม')]['จำนวนเงิน'].sum()
            
            tm_balance = tm_inc + tm_transfer_in + tm_refund + tm_sav_withdrawn + tm_sav_loan - tm_exp - tm_inv - tm_transfer_out - tm_lend - tm_sav_dep - tm_sav_repay
            kt_balance = net_in_cycle - tm_balance

            st.markdown(f"""
                <div style='display: flex; gap: 15px; margin-bottom: 15px;'>
                    <div style='flex: 1; background-color: rgba(42, 157, 143, 0.1); border: 1px solid #2a9d8f; padding: 12px 18px; border-radius: 14px;'>
                        <span style='font-size: 13px; opacity: 0.8;'>🏦 กรุงไทย (Krungthai)</span>
                        <h3 style='margin: 0; color: #2a9d8f;'>฿{kt_balance:,.2f}</h3>
                    </div>
                    <div style='flex: 1; background-color: rgba(244, 162, 97, 0.15); border: 1px solid #f4a261; padding: 12px 18px; border-radius: 14px;'>
                        <span style='font-size: 13px; opacity: 0.8;'>📱 TrueMoney Wallet</span>
                        <h3 style='margin: 0; color: #f4a261;'>฿{tm_balance:,.2f}</h3>
                    </div>
                    <div style='flex: 1; background-color: var(--secondary-background-color); border: 1px solid var(--border-color); padding: 12px 18px; border-radius: 14px;'>
                        <span style='font-size: 13px; opacity: 0.8;'>💰 รวมเงินสดสุทธิ</span>
                        <h3 style='margin: 0;'>฿{net_in_cycle:,.2f}</h3>
                    </div>
                </div>
            """, unsafe_allow_html=True)

            if selected_cycle_label != "🌟 แสดงข้อมูลทั้งหมด (All Time)":
                diff = kt_balance - selected_kt_real
                is_balanced = abs(diff) < 0.01
                
                with st.expander(f"⚖️ คาลิเบรทบัญชีกรุงไทยจริง (Bank Calibration): {'✅ ตรงกัน 100%' if is_balanced else f'⚠️ ส่วนต่าง ฿{diff:,.2f}'}", expanded=not is_balanced):
                    c_cal1, c_cal2, c_cal3, c_cal4 = st.columns([1.2, 1.2, 1.2, 2.2])
                    c_cal1.markdown(f"**💰 เงินตั้งต้นรอบนี้**<br><h4>฿{selected_carry:,.2f}</h4>", unsafe_allow_html=True)
                    c_cal2.markdown(f"**📥 รายรับรอบนี้**<br><h4 style='color:#2a9d8f;'>+฿{inc:,.2f}</h4>", unsafe_allow_html=True)
                    c_cal3.markdown(f"**🏦 กรุงไทยในแอป**<br><h4>฿{kt_balance:,.2f}</h4>", unsafe_allow_html=True)
                    
                    with c_cal4:
                        with st.form("calibrate_bank_form", clear_on_submit=True):
                            inp_carry = st.number_input("💰 แก้ไขเงินตั้งต้นรอบนี้ (ยอดยกมาจริง)", value=None, placeholder=f"ปัจจุบัน: ฿{selected_carry:,.2f}", step=100.0, format="%.2f")
                            inp_kt = st.number_input("🏦 เงินจริงในบัญชีกรุงไทย (ปัจจุบัน)", value=None, placeholder=f"ปัจจุบัน: ฿{selected_kt_real:,.2f}", step=100.0, format="%.2f")
                            
                            sub1, sub2 = st.columns(2)
                            with sub1:
                                save_cal = st.form_submit_button("💾 บันทึกยอดลงคลาวด์", use_container_width=True)
                            with sub2:
                                sync_clean = st.form_submit_button("✂️ ล้างส่วนต่างเดือนก่อน", use_container_width=True)
                            
                            if save_cal:
                                if selected_row_idx:
                                    if inp_carry is not None:
                                        cycle_sheet.update_cell(int(selected_row_idx), 5, float(inp_carry))
                                    if inp_kt is not None:
                                        cycle_sheet.update_cell(int(selected_row_idx), 6, float(inp_kt))
                                    fetch_cycles.clear()
                                    st.success("อัปเดตข้อมูลคาลิเบรทเรียบร้อย!")
                                    st.rerun()
                                    
                            if sync_clean:
                                if selected_row_idx:
                                    flow_in_cycle = kt_balance - selected_carry
                                    target_kt = inp_kt if inp_kt is not None else selected_kt_real
                                    new_carry = target_kt - flow_in_cycle
                                    cycle_sheet.update_cell(int(selected_row_idx), 5, float(new_carry))
                                    if inp_kt is not None:
                                        cycle_sheet.update_cell(int(selected_row_idx), 6, float(inp_kt))
                                    fetch_cycles.clear()
                                    st.success(f"ล้างส่วนต่างอดีตเรียบร้อย! ปรับเงินตั้งต้นเป็น ฿{new_carry:,.2f}")
                                    st.rerun()
                    
                    if not is_balanced:
                        st.markdown(f"<div class='calib-box-diff'><b>💡 ข้อเสนอแนะ:</b> ยอดในแอป {'มากกว่า' if diff > 0 else 'น้อยกว่า'} เงินจริงในกรุงไทยอยู่ <b>฿{abs(diff):,.2f}</b><br>สามารถกดปุ่ม <b>'✂️ ล้างส่วนต่างเดือนก่อน'</b> เพื่อให้ตรงกับธนาคารจริงทันทีครับ</div>", unsafe_allow_html=True)
                    else:
                        st.markdown("<div class='calib-box-match'><b>🎉 ยอดเยี่ยมมากครับหมอ!</b> ยอดกรุงไทยใน Minimal Finance Pro ตรงกับเงินจริงในธนาคารเป๊ะ 100% ครับ</div>", unsafe_allow_html=True)

            m1, m2, m3, m4, m5 = st.columns(5)
            net_title_class = "metric-title" if net_in_cycle >= 0 else "metric-title-alert"
            m1.markdown(f"<div class='metric-card'><div class='{net_title_class}'>Net Balance</div><div class='metric-value'>฿{net_in_cycle:,.0f}</div><div class='metric-currency'>THB</div></div>", unsafe_allow_html=True)
            m2.markdown(f"<div class='metric-card'><div class='metric-title'>Income (รอบนี้) <span style='color:#2a9d8f;'>↗</span></div><div class='metric-value'>฿{inc:,.0f}</div><div class='metric-currency'>THB</div></div>", unsafe_allow_html=True)
            m3.markdown(f"<div class='metric-card'><div class='metric-title'>Expenses <span style='color:#f9744b;'>↘</span></div><div class='metric-value'>฿{exp:,.0f}</div><div class='metric-currency'>THB</div></div>", unsafe_allow_html=True)
            
            loan_badge = f"<div style='font-size:11px;color:#f9744b;font-weight:600;margin-top:2px;'>⚠️ หนี้ค้าง: ฿{outstanding_loan:,.0f}</div>" if outstanding_loan > 0 else ""
            m4.markdown(f"<div class='metric-card'><div class='metric-title'>Savings <span style='color:#457b9d;'>↗</span></div><div class='metric-value'>฿{sav_flow:,.0f}</div><div style='font-size:11px;color:#457b9d;font-weight:600;margin-top:2px;'>🏦 คลังรวม: ฿{total_sav_now:,.0f}</div>{loan_badge}</div>", unsafe_allow_html=True)
            m5.markdown(f"<div class='metric-card'><div class='metric-title'>Investments <span style='color:#e9c46a;'>↗</span></div><div class='metric-value'>฿{inv:,.0f}</div><div class='metric-currency'>THB</div></div>", unsafe_allow_html=True)
            
            st.markdown("---")
            
            with st.expander("🔄 เปิด Circle ใหม่ & สั่งตัดรอบบัญชี (Cycle Control)", expanded=False):
                st.write(f"📌 รอบบัญชีที่กำลังใช้งานอยู่ตอนนี้คือ: **{active_cycle_name}**")
                with st.form("new_cycle_form"):
                    new_circle_name = st.text_input("ชื่อ Circle ใหม่ที่จะเปิด (เช่น September 2026)", placeholder="ระบุชื่อรอบเดือนใหม่...")
                    if st.form_submit_button("⏹️ จบรอบปัจจุบัน & เริ่ม Circle ใหม่ทันที", use_container_width=True):
                        now_timestamp = datetime.datetime.now(TZ_TH).strftime('%Y-%m-%d %H:%M:%S')
                        if active_row_idx:
                            cycle_sheet.update_cell(int(active_row_idx), 3, now_timestamp)
                            cycle_sheet.update_cell(int(active_row_idx), 4, "CLOSED")
                            cycle_sheet.update_cell(int(active_row_idx), 6, float(kt_balance))
                        
                        final_name = new_circle_name if new_circle_name.strip() else f"Circle {datetime.datetime.now(TZ_TH).strftime('%B %Y')}"
                        start_carry = float(selected_kt_real) if selected_kt_real > 0 else float(kt_balance)
                        cycle_sheet.append_row([final_name, now_timestamp, "", "ACTIVE", start_carry, start_carry])
                        fetch_cycles.clear()
                        st.success(f"จบรอบเดิมและเริ่ม {final_name} พร้อมตั้งต้นด้วยยอดจริง ฿{start_carry:,.2f} เรียบร้อยครับ!")
                        st.rerun()

            st.markdown("---")

            col_trend_title, col_trend_filter = st.columns([1.5, 2])
            with col_trend_title:
                st.markdown("<p class='quick-add-text' style='margin-top:5px;'>Trend Analysis (Stock Style)</p>", unsafe_allow_html=True)
            with col_trend_filter:
                c_tf, c_ms = st.columns([1.2, 2])
                time_frame = c_tf.selectbox("Timeframe:", ["รายวัน (1D)", "รายสัปดาห์ (1W)", "รายเดือน (1M)", "รายปี (1Y)", "ราย 5 ปี (5Y)"], label_visibility="collapsed")
                visible_metrics = c_ms.multiselect("เลือกเส้นวิเคราะห์คงเหลือ:", ["รายรับ", "รายจ่าย", "เงินออม", "เงินลงทุน", "เงินสุทธิ"], default=["รายรับ", "รายจ่าย", "เงินสุทธิ"])
            
            today = datetime.datetime.now(TZ_TH).date()
            df_trend = df_dash.copy()
            df_trend = df_trend.dropna(subset=['วันเวลา'])
            df_trend = df_trend.sort_values(by='วันเวลา')
            
            if not df_trend.empty:
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
            else:
                st.info("ไม่มีข้อมูลการเงินบันทึกไว้ในช่วงไทม์เฟรมนี้")
            
            st.markdown("---")
            
            expense_df = df_dash[df_dash['ประเภท'] == 'รายจ่าย']
            col_exp_title, col_exp_filter = st.columns([2, 1.5])
            with col_exp_title:
                st.markdown("<p class='quick-add-text'>Expense Analysis</p>", unsafe_allow_html=True)
            with col_exp_filter:
                if not expense_df.empty:
                    raw_cats = expense_df['หมวดหมู่หลัก'].unique()
                    all_main_cats = sorted([str(c).strip() for c in raw_cats if pd.notnull(c) and str(c).strip() != ""])
                    selected_main_filter = st.multiselect("🔎 ติ๊กเลือกหมวดหมู่ที่ต้องการดู:", all_main_cats, default=all_main_cats)

            col_chart1, col_chart2 = st.columns([1, 1.2])
            if not expense_df.empty:
                filtered_expense_df = expense_df[expense_df['หมวดหมู่หลัก'].isin(selected_main_filter)] if selected_main_filter else pd.DataFrame(columns=expense_df.columns)
                all_pie_data = expense_df.groupby('หมวดหมู่หลัก')['จำนวนเงิน'].sum().reset_index()
                unique_main_cats = all_pie_data['หมวดหมู่หลัก'].tolist()
                cat_color_map = {cat: SUB_CAT_PALETTE[i % len(SUB_CAT_PALETTE)] for i, cat in enumerate(unique_main_cats)}
                
                with col_chart1:
                    if not filtered_expense_df.empty:
                        pie_data = filtered_expense_df.groupby('หมวดหมู่หลัก')['จำนวนเงิน'].sum().reset_index()
                        pie_data['จำนวนเงิน'] = pie_data['จำนวนเงิน'].abs() 
                        
                        fig_pie = px.pie(pie_data, values='จำนวนเงิน', names='หมวดหมู่หลัก', hole=0.78, color='หมวดหมู่หลัก', color_discrete_map=cat_color_map)
                        fig_pie.update_traces(textposition='outside', textinfo='percent+label', marker=dict(line=dict(width=0)), textfont=dict(family='Poppins', size=11))
                        fig_pie.update_layout(plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', showlegend=False, margin=dict(t=30, b=30, l=30, r=30))
                        st.plotly_chart(fig_pie, use_container_width=True, theme="streamlit")
                    else:
                        st.info("กรุณาติ๊กเลือกอย่างน้อย 1 หมวดหมู่")
                        
                with col_chart2:
                    if not filtered_expense_df.empty:
                        sub_data = filtered_expense_df.groupby(['หมวดหมู่หลัก', 'หมวดหมู่ย่อย'])['จำนวนเงิน'].sum().reset_index()
                        sub_data['จำนวนเงิน'] = sub_data['จำนวนเงิน'].abs()
                        
                        if len(selected_main_filter) == len(all_main_cats):
                            sub_data = sub_data.sort_values(by="จำนวนเงิน", ascending=False).head(8)
                        else:
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

    # ==========================================
    # 🤝 Tab 3: ลูกหนี้ & หารบิล
    # ==========================================
    with tab3:
        st.markdown("<p class='quick-add-text' style='font-size: 22px;'>🤝 ระบบหารค่าใช้จ่าย & คนติดเงิน (Receivables Tracker)</p>", unsafe_allow_html=True)
        st.caption("💡 สำหรับจดเวลาหารค่าข้าวกับแฟนหรือเพื่อน โดยระบบจะหักยอดเงินสดออกก่อน และคืนเข้า Net Balance เมื่อกดยืนยันว่าคืนเงินแล้ว (ไม่ปนกับรายจ่ายจริง)")
        
        debt_data = fetch_receivables()
        df_debt = pd.DataFrame(debt_data) if debt_data else pd.DataFrame(columns=["ID", "ชื่อคนติดเงิน", "รายการ/รายละเอียด", "จำนวนเงิน", "กระเป๋าที่จ่าย", "วันที่สร้าง", "สถานะ", "วันที่คืน"])
        
        pending_df = df_debt[df_debt['สถานะ'] == '⏳ รอคืนเงิน'] if not df_debt.empty else pd.DataFrame()
        paid_df = df_debt[df_debt['สถานะ'] == '✅ คืนแล้ว'] if not df_debt.empty else pd.DataFrame()
        
        total_pending = pd.to_numeric(pending_df['จำนวนเงิน'], errors='coerce').sum() if not pending_df.empty else 0.0
        total_paid = pd.to_numeric(paid_df['จำนวนเงิน'], errors='coerce').sum() if not paid_df.empty else 0.0
        
        m_d1, m_d2 = st.columns(2)
        m_d1.markdown(f"<div class='metric-card'><div class='metric-title'>⏳ ยอดเงินที่คนอื่นยังติดอยู่ (รอคืน)</div><div class='metric-value' style='color:#f9744b;'>฿{total_pending:,.2f}</div></div>", unsafe_allow_html=True)
        m_d2.markdown(f"<div class='metric-card'><div class='metric-title'>✅ ได้รับคืนมาแล้วทั้งหมด</div><div class='metric-value' style='color:#2a9d8f;'>฿{total_paid:,.2f}</div></div>", unsafe_allow_html=True)
        
        st.markdown("---")
        
        c_add_debt, c_confirm_debt = st.columns([1.2, 1])
        with c_add_debt:
            with st.expander("➕ เพิ่มรายการคนติดเงิน / หารค่าข้าว (Add Receivable)", expanded=True):
                with st.form("add_debt_form", clear_on_submit=True):
                    d_who = st.text_input("ชื่อคนติดเงิน (เช่น แฟน, เพื่อน A)", placeholder="ระบุชื่อ...")
                    d_desc = st.text_input("รายการ (เช่น หารค่าข้าวเย็น, ค่าชาบู)", placeholder="รายละเอียดบิล...")
                    d_amt = st.number_input("จำนวนเงินที่เขาต้องจ่ายคืน (บาท)", min_value=1.0, step=10.0, format="%.2f", value=None, placeholder="0.00")
                    d_wallet = st.selectbox("ทดจ่ายออกจากกระเป๋าไหน?", ["🏦 กรุงไทย", "📱 TrueMoney Wallet"])
                    
                    if st.form_submit_button("💾 บันทึกคนติดเงิน (ตัดเงินทดจ่าย)", use_container_width=True) and d_amt is not None and d_amt > 0:
                        new_id = len(df_debt) + 1
                        now_str = datetime.datetime.now(TZ_TH).strftime('%Y-%m-%d %H:%M:%S')
                        
                        debt_sheet.append_row([new_id, d_who, d_desc, d_amt, d_wallet, now_str, "⏳ รอคืนเงิน", ""])
                        sheet.append_row([now_str, "🤝 เงินทดจ่าย", f"ลูกหนี้: {d_who}", d_amt, f"หารบิล: {d_desc}", d_wallet])
                        
                        fetch_receivables.clear()
                        fetch_main_data.clear()
                        st.toast("บันทึกรายการหารบิลและหักยอดทดจ่ายเรียบร้อย! ✨")
                        st.rerun()

        with c_confirm_debt:
            with st.expander("🎉 ยืนยันรับคืนเงินแล้ว (Mark as Paid)", expanded=True):
                if not pending_df.empty:
                    pending_options = []
                    for idx, r in pending_df.iterrows():
                        pending_options.append((f"#{r['ID']} - {r['ชื่อคนติดเงิน']}: {r['รายการ/รายละเอียด']} (฿{r['จำนวนเงิน']})", idx + 2, r['จำนวนเงิน'], r['ชื่อคนติดเงิน'], r['กระเป๋าที่จ่าย']))
                    
                    selected_debt_label = st.selectbox("เลือกรายการที่ได้รับเงินคืนแล้ว:", [opt[0] for opt in pending_options])
                    return_wallet = st.selectbox("รับเงินคืนเข้ากระเป๋าไหน?", ["🏦 กรุงไทย", "📱 TrueMoney Wallet"], key="ret_wallet_select")
                    
                    if st.button("✅ ยืนยันว่าคืนเงินแล้ว (บวกกลับเข้า Net Balance)", use_container_width=True):
                        for label, r_idx, d_val, d_name, _ in pending_options:
                            if label == selected_debt_label:
                                now_str = datetime.datetime.now(TZ_TH).strftime('%Y-%m-%d %H:%M:%S')
                                debt_sheet.update_cell(int(r_idx), 7, "✅ คืนแล้ว")
                                debt_sheet.update_cell(int(r_idx), 8, now_str)
                                sheet.append_row([now_str, "🤝 รับคืนเงินทดจ่าย", f"ลูกหนี้: {d_name}", float(d_val), "ได้รับคืนเงินที่ทดจ่ายไปก่อน", return_wallet])
                                
                                fetch_receivables.clear()
                                fetch_main_data.clear()
                                st.success("อัปเดตรับคืนเงินและเพิ่มยอดใน Net Balance เรียบร้อยครับ!")
                                st.rerun()
                else:
                    st.info("ไม่มีรายการคนติดเงินค้างอยู่ตอนนี้ครับ 🎉")
        
        st.markdown("---")
        st.markdown("<p class='quick-add-text'>📋 ตารางประวัติหารค่าใช้จ่ายและคนติดเงินทั้งหมด</p>", unsafe_allow_html=True)
        if not df_debt.empty:
            st.dataframe(df_debt, use_container_width=True, hide_index=True)
        else:
            st.info("ยังไม่มีประวัติคนติดเงินในระบบครับ")

    # ==========================================
    # 🎯 Tab 4: Goals (พิมพ์ชื่อพร้อมไอคอนในช่องเดียว)
    # ==========================================
    with tab4:
        st.markdown("<p class='quick-add-text' style='font-size: 22px;'>🎯 ระบบจัดสรรเป้าหมายออมเงิน (Dynamic Goals)</p>", unsafe_allow_html=True)
        st.caption("💡 สำหรับแบ่งกระเป๋าเงินออมออกเป็นหลายๆ เป้าหมาย เช่น ค่าสอบ GRE, กองทุนเที่ยว, กองทุนฉุกเฉิน")
        
        total_allocated_sav = pd.to_numeric(df_goals["สะสมแล้ว (บาท)"], errors="coerce").sum() if not df_goals.empty else 0.0
        remaining_unallocated_sav = total_sav_now - total_allocated_sav

        g_m1, g_m2, g_m3 = st.columns(3)
        g_m1.markdown(f"<div class='metric-card'><div class='metric-title'>💰 คลังเงินออมจริง (Savings Total)</div><div class='metric-value' style='color:#457b9d;'>฿{total_sav_now:,.2f}</div></div>", unsafe_allow_html=True)
        g_m2.markdown(f"<div class='metric-card'><div class='metric-title'>📦 จัดสรรเข้าเป้าหมายแล้ว</div><div class='metric-value' style='color:#2a9d8f;'>฿{total_allocated_sav:,.2f}</div></div>", unsafe_allow_html=True)
        
        rem_color = "#f9744b" if remaining_unallocated_sav < 0 else "#8ab17d"
        g_m3.markdown(f"<div class='metric-card'><div class='metric-title'>⚖️ เงินออมที่ยังไม่จัดสรร</div><div class='metric-value' style='color:{rem_color};'>฿{remaining_unallocated_sav:,.2f}</div></div>", unsafe_allow_html=True)
        
        st.markdown("---")
        
        if not df_goals.empty:
            for idx, row in df_goals.iterrows():
                g_icon = str(row["ไอคอน"]).strip()
                g_name = str(row["ชื่อเป้าหมาย"]).strip()
                title_text = f"{g_icon} {g_name}".strip()
                g_target = float(row["เป้าหมาย (บาท)"]) if pd.notnull(row["เป้าหมาย (บาท)"]) and float(row["เป้าหมาย (บาท)"]) > 0 else 1.0
                g_saved = float(row["สะสมแล้ว (บาท)"]) if pd.notnull(row["สะสมแล้ว (บาท)"]) else 0.0
                
                pct = min(g_saved / g_target, 1.0)
                pct_display = (g_saved / g_target) * 100
                
                st.markdown(f"**{title_text}** — `฿{g_saved:,.2f} / ฿{g_target:,.2f} ({pct_display:.1f}%)`")
                st.progress(pct)
                st.markdown("<div style='margin-bottom: 12px;'></div>", unsafe_allow_html=True)
        else:
            st.info("ยังไม่มีเป้าหมายการออม ลองเพิ่มเป้าหมายแรกด้านล่างได้เลยครับ! ✨")

        st.markdown("---")

        c_goal_add, c_goal_edit = st.columns([1, 1.5])
        with c_goal_add:
            with st.expander("➕ เพิ่มเป้าหมายออมเงินใหม่ (Add New Goal)", expanded=True):
                with st.form("add_goal_form", clear_on_submit=True):
                    # 💡 พิมพ์ไอคอน Emoji รวมเข้ากับชื่อเป้าหมายได้เลยในช่องเดียว
                    new_g_name = st.text_input("ชื่อเป้าหมาย (เช่น ✈️ GRE Fund, 💻 ซื้อ iPad)", placeholder="พิมพ์ชื่อเป้าหมายพร้อมไอคอนได้เลย...")
                    new_g_target = st.number_input("จำนวนเงินเป้าหมาย (บาท)", min_value=100.0, step=1000.0, format="%.2f", value=None, placeholder="0.00")
                    new_g_saved = st.number_input("เงินออมเริ่มต้นในกระเป๋านี้ (บาท)", min_value=0.0, step=500.0, format="%.2f", value=0.0)
                    
                    if st.form_submit_button("💾 เพิ่มเป้าหมายลงระบบ", use_container_width=True):
                        if new_g_name.strip() and new_g_target is not None and new_g_target > 0:
                            goal_sheet.append_row(["", new_g_name.strip(), float(new_g_target), float(new_g_saved)])
                            fetch_goals.clear()
                            st.success(f"เพิ่มเป้าหมาย '{new_g_name}' สำเร็จ! ✨")
                            st.rerun()
                        else:
                            st.warning("กรุณาระบุชื่อเป้าหมายและจำนวนเงินให้ถูกต้องครับ")

        with c_goal_edit:
            st.markdown("### ✏️ ตารางแก้ไข / ลบเป้าหมายการออม")
            st.caption("💡 สามารถคลิกเปลี่ยนชื่อ ปรับตัวเลข หรือ **'เลือกแถวแล้วกดปุ่ม Delete บนคีย์บอร์ด'** เพื่อลบเป้าหมายออกได้ทันทีครับ")
            
            edited_goals = st.data_editor(
                df_goals, 
                use_container_width=True, 
                num_rows="dynamic", 
                key="editor_goals_v30"
            )
            
            if st.button("💾 บันทึกการเปลี่ยนแปลงเป้าหมาย (Save Goals)", use_container_width=True):
                goal_sheet.clear()
                goal_sheet.update(range_name="A1", values=[edited_goals.columns.values.tolist()] + edited_goals.values.tolist())
                fetch_goals.clear()
                st.success("อัปเดตรายการเป้าหมายออมเงินเรียบร้อย! ✨")
                st.rerun()

    # ==========================================
    # ⚙️ Tab 5: Settings (Categories, QuickAdds & Raw Data)
    # ==========================================
    with tab5:
        st.subheader("📁 Categories Editor")
        edited_cat = st.data_editor(cat_raw_df, use_container_width=True, num_rows="dynamic", key="editor_cat_v30")
        if st.button("💾 Save Categories", use_container_width=True):
            cat_sheet.clear()
            cat_sheet.update(range_name="A1", values=[edited_cat.columns.values.tolist()] + edited_cat.values.tolist())
            fetch_categories.clear()
            st.success("Categories updated! ✨")
            st.rerun()

        st.markdown("---")
        st.subheader("⚡ Quick Adds Editor")
        edited_qa = st.data_editor(qa_df, use_container_width=True, num_rows="dynamic", key="editor_qa_v30")
        if st.button("💾 Save Quick Adds", use_container_width=True):
            qa_sheet.clear()
            qa_sheet.update(range_name="A1", values=[edited_qa.columns.values.tolist()] + edited_qa.values.tolist())
            fetch_quick_adds.clear()
            st.success("Quick adds updated!")
            st.rerun()
            
        st.markdown("---")
        st.subheader("✏️ Raw Data Editor")
        if not df.empty:
            clean_df_edit = df[["วันที่", "ประเภท", "หมวดหมู่", "จำนวนเงิน", "รายละเอียด", "กระเป๋า"]]
            edited_df = st.data_editor(clean_df_edit, use_container_width=True, num_rows="dynamic", key="editor_finance_v30")
            if st.button("💾 Save Data to Cloud", use_container_width=True):
                sheet.clear()
                edited_df['วันที่'] = edited_df['วันที่'].astype(str)
                sheet.update(range_name="A1", values=[edited_df.columns.values.tolist()] + edited_df.values.tolist())
                fetch_main_data.clear()
                st.success("Data updated!")
                st.rerun()

    # ==========================================
    # 🏦 Tab 6: Loan Simulator
    # ==========================================
    with tab6:
        st.markdown("<p class='quick-add-text' style='font-size: 22px;'>🏦 เครื่องจำลองสินเชื่อระบบคลาวด์ถาวร (EMI Lock)</p>", unsafe_allow_html=True)
        st.caption("💡 ระบบผูกเข้ากับคลังเงินออมอัตโนมัติ ทุกการกู้หรือคืนเงินจะสะท้อนผลไปที่ Dashboard ทันที")
        
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
                    
                    now_str = datetime.datetime.now(TZ_TH).strftime('%Y-%m-%d %H:%M:%S')
                    sheet.append_row([now_str, "กู้เงินออม", "บริหารเงินออม: กู้เงินออม", new_p, f"เปิดสัญญาเงินกู้ระบบจำลอง ฿{new_p:,.0f}", "🏦 กรุงไทย"])
                    
                    fetch_loans.clear()
                    fetch_main_data.clear()
                    st.success("เปิดสัญญาเงินกู้ฉบับใหม่และบันทึกลงคลังเรียบร้อยครับ!")
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
                    
                    now_str = datetime.datetime.now(TZ_TH).strftime('%Y-%m-%d %H:%M:%S')
                    sheet.append_row([now_str, "คืนเงินกู้ออม", "บริหารเงินออม: คืนเงินกู้ออม", emi_amount, f"ชำระค่างวดสินเชื่อจำลอง งวดที่ {next_paid_count}/{db_months}", "🏦 กรุงไทย"])
                    
                    fetch_loans.clear()
                    fetch_main_data.clear()
                    st.toast(f"ชำระงวดที่ {next_paid_count} สำเร็จ! ระบบสแกนหักยอดและคืนคลังเรียบร้อย ✨")
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
