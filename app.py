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

# 🚀 ระบบ Smart Cache 
@st.cache_resource(ttl=3600)
def get_google_sheets():
    try:
        sh = client.open(spreadsheet_name)
    except Exception:
        return None, None, None, None, None, None, None, None, None
        
    sheet_main = sh.sheet1
    try:
        sheet_main.resize(rows=500, cols=8)
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
            records_c = sheet_cycle.get_all_values()
            for idx, r in enumerate(records_c):
                if len(r) > 0 and "July 2026" in str(r[0]):
                    sheet_cycle.update_cell(idx + 1, 2, "2026-06-25 00:00:00")
                    sheet_cycle.update_cell(idx + 1, 3, "2026-07-31 23:59:59")
                elif len(r) > 0 and "August 2026" in str(r[0]):
                    sheet_cycle.update_cell(idx + 1, 2, "2026-08-01 00:00:00")
                    sheet_cycle.update_cell(idx + 1, 3, "")
        except Exception:
            pass
    except:
        sheet_cycle = sh.add_worksheet(title="Cycles", rows="50", cols="10")
        sheet_cycle.append_row(["ชื่อรอบบัญชี", "เริ่มต้น", "สิ้นสุด", "สถานะ", "ยอดยกมา", "เงินจริงกรุงไทย"])
        sheet_cycle.append_row(["July 2026", "2026-06-25 00:00:00", "2026-07-31 23:59:59", "CLOSED", 0.0, 2501.0])
        sheet_cycle.append_row(["August 2026", "2026-08-01 00:00:00", "", "ACTIVE", 2501.0, 2501.0])

    try:
        sheet_debt = sh.worksheet("Receivables")
    except:
        sheet_debt = sh.add_worksheet(title="Receivables", rows="50", cols="8")
        sheet_debt.append_row(["ID", "ชื่อคนติดเงิน", "รายการ/รายละเอียด", "จำนวนเงิน", "กระเป๋าที่จ่าย", "วันที่สร้าง", "สถานะ", "วันที่คืน"])

    try:
        sheet_goal = sh.worksheet("Goals")
        headers_goal = sheet_goal.row_values(1)
        if len(headers_goal) < 5 or headers_goal[4] != "ความสำคัญ":
            sheet_goal.update_cell(1, 5, "ความสำคัญ")
    except:
        sheet_goal = sh.add_worksheet(title="Goals", rows="30", cols="5")
        sheet_goal.append_row(["ไอคอน", "ชื่อเป้าหมาย", "เป้าหมาย (บาท)", "สะสมแล้ว (บาท)", "ความสำคัญ"])
        sheet_goal.append_row(["✈️", "GRE / Future Studies Fund", 100000.0, 0.0, "🔥 สูง (High)"])

    try:
        sheet_wallet = sh.worksheet("Wallets")
    except:
        sheet_wallet = sh.add_worksheet(title="Wallets", rows="30", cols="3")
        sheet_wallet.append_row(["ชื่อกระเป๋า"])
        sheet_wallet.append_row(["🏦 กรุงไทย"])
        sheet_wallet.append_row(["📱 TrueMoney Wallet"])
        sheet_wallet.append_row(["🌸 ออมสิน"])
        
    try:
        sheet_budget = sh.worksheet("Budgets")
    except:
        sheet_budget = sh.add_worksheet(title="Budgets", rows="50", cols="2")
        sheet_budget.append_row(["หมวดหมู่เป้าหมาย", "งบประมาณ (บาท)"])
        sheet_budget.append_row(["อาหาร/เครื่องดื่ม", 5000.0])
        
    return sheet_main, sheet_qa, sheet_cat, sheet_loan, sheet_cycle, sheet_debt, sheet_goal, sheet_wallet, sheet_budget

sheet, qa_sheet, cat_sheet, loan_sheet, cycle_sheet, debt_sheet, goal_sheet, wallet_sheet, budget_sheet = get_google_sheets()

if sheet is None:
    st.error(f"❌ หาไฟล์ Google Sheets ที่ชื่อ '{spreadsheet_name}' ไม่เจอครับ")
    st.stop()

# --- ฟังก์ชันโหลดข้อมูลแยก Cache ---
@st.cache_data(ttl=60)
def fetch_main_data():
    return sheet.get_all_values()

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

@st.cache_data(ttl=60)
def fetch_wallets():
    return wallet_sheet.get_all_records()

@st.cache_data(ttl=60)
def fetch_budgets():
    return budget_sheet.get_all_records()

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

def parse_clean_datetime(date_series):
    s = date_series.astype(str).str.strip()
    dt = pd.to_datetime(s, format='mixed', errors='coerce')
    if dt.isna().any():
        dt_alt = pd.to_datetime(s, format='mixed', dayfirst=True, errors='coerce')
        dt = dt.fillna(dt_alt)
    return dt

def load_data():
    records = fetch_main_data()
    if not records:
        return pd.DataFrame(columns=["วันที่", "ประเภท", "หมวดหมู่", "จำนวนเงิน", "รายละเอียด", "กระเป๋า", "หมวดหมู่หลัก", "หมวดหมู่ย่อย", "วันเวลา", "วันที่_date"])
    
    df = pd.DataFrame(records)
    while len(df.columns) < 6:
        df[len(df.columns)] = ""
    df = df.iloc[:, :6]
    df.columns = ["วันที่", "ประเภท", "หมวดหมู่", "จำนวนเงิน", "รายละเอียด", "กระเป๋า"]
    
    df = df[df['วันที่'].astype(str).str.strip() != 'วันที่']
    df = df[df['วันที่'].astype(str).str.strip() != '']
    df = df.reset_index(drop=True)
    
    parsed_time = parse_clean_datetime(df['วันที่'])
    df['วันเวลา'] = parsed_time.apply(lambda x: x.replace(year=x.year - 543) if pd.notnull(x) and x.year > 2400 else x)
    df['วันที่'] = df['วันเวลา'].dt.strftime('%Y-%m-%d %H:%M:%S').fillna(df['วันที่'].astype(str))
    df['วันที่_date'] = df['วันเวลา'].dt.date
    df['จำนวนเงิน'] = pd.to_numeric(df['จำนวนเงิน'], errors='coerce').fillna(0.0)
    df['ประเภท'] = df['ประเภท'].astype(str).str.strip()
    df['หมวดหมู่'] = df['หมวดหมู่'].astype(str).str.strip()
    df['รายละเอียด'] = df['รายละเอียด'].astype(str).str.strip()
    df['กระเป๋า'] = df['กระเป๋า'].fillna('🏦 กรุงไทย').astype(str).str.strip().replace('', '🏦 กรุงไทย')
    df['หมวดหมู่หลัก'] = df['หมวดหมู่'].apply(lambda x: str(x).split(":")[0].strip() if pd.notnull(x) else "ทั่วไป")
    df['หมวดหมู่ย่อย'] = df['หมวดหมู่'].apply(lambda x: str(x).split(":")[1].strip() if pd.notnull(x) and ":" in str(x) else "ทั่วไป")
    return df

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

budgets_data = fetch_budgets()
df_budgets = pd.DataFrame(budgets_data) if budgets_data else pd.DataFrame(columns=["หมวดหมู่เป้าหมาย", "งบประมาณ (บาท)"])

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

wallets_data = fetch_wallets()
df_wallets = pd.DataFrame(wallets_data) if wallets_data else pd.DataFrame(columns=["ชื่อกระเป๋า"])
wallet_list = [
    str(w).strip() for w in df_wallets["ชื่อกระเป๋า"].tolist() 
    if pd.notnull(w) and str(w).strip() != ""
]
if not wallet_list:
    wallet_list = ["🏦 กรุงไทย", "📱 TrueMoney Wallet", "🌸 ออมสิน"]

goals_data = fetch_goals()
df_goals = pd.DataFrame(goals_data) if goals_data else pd.DataFrame(columns=["ไอคอน", "ชื่อเป้าหมาย", "เป้าหมาย (บาท)", "สะสมแล้ว (บาท)", "ความสำคัญ"])
if 'ความสำคัญ' not in df_goals.columns:
    df_goals['ความสำคัญ'] = "⭐ ปานกลาง (Medium)"
else:
    df_goals['ความสำคัญ'] = df_goals['ความสำคัญ'].replace(r'^\s*$', "⭐ ปานกลาง (Medium)", regex=True)

PRIORITY_LEVELS = [
    "🚀 ด่วนที่สุด (Critical)", 
    "🔥 สูง (High)", 
    "⭐ ปานกลาง (Medium)", 
    "🟢 ต่ำ (Low)", 
    "🧊 เผื่อไว้ (Optional)"
]

goal_options_list = ["📦 คลังออมทั่วไป (ไม่ระบุเป้าหมาย)"] + (df_goals["ชื่อเป้าหมาย"].tolist() if not df_goals.empty else [])

def calculate_savings_metrics(df_source):
    if df_source.empty or 'ประเภท' not in df_source.columns:
        return 0.0, 0.0, 0.0, 0.0, 0.0, 0.0
    dep_mask = (
        df_source['ประเภท'].astype(str).str.contains('ออม|save|saving', case=False, na=False) |
        df_source['หมวดหมู่'].astype(str).str.contains('ฝากออม|ออมเงิน|เงินออม|เก็บออม|aomsin|ออมสิน', case=False, na=False) |
        df_source['รายละเอียด'].astype(str).str.contains('ออมสิน|aomsin|เงินออม', case=False, na=False)
    ) & ~df_source['ประเภท'].astype(str).str.contains('ถอน|กู้|คืน|รายจ่าย', case=False, na=False)

    with_mask = df_source['ประเภท'].astype(str).str.contains('ถอน.*ออม|ออม.*ถอน', case=False, na=False)
    loan_mask = (df_source['ประเภท'].astype(str).str.contains('กู้.*ออม|ออม.*กู้', case=False, na=False) & 
                 ~df_source['ประเภท'].astype(str).str.contains('คืน', case=False, na=False))
    repay_mask = df_source['ประเภท'].astype(str).str.contains('คืน.*กู้|กู้.*คืน', case=False, na=False)

    dep_val = float(df_source[dep_mask]['จำนวนเงิน'].sum())
    with_val = float(df_source[with_mask]['จำนวนเงิน'].sum())
    loan_val = float(df_source[loan_mask]['จำนวนเงิน'].sum())
    repay_val = float(df_source[repay_mask]['จำนวนเงิน'].sum())

    net_flow = dep_val + repay_val - with_val - loan_val
    outstanding = loan_val - repay_val
    return dep_val, with_val, loan_val, repay_val, net_flow, outstanding

_, _, _, _, total_sav_now, outstanding_loan = calculate_savings_metrics(df)

HONEY_POT_MAP = {
    "รายรับ": "#2a9d8f",     
    "รายจ่าย": "#f9744b",    
    "เงินออม": "#457b9d",    
    "เงินลงทุน": "#e9c46a",
    "เงินสุทธิ": "#8ab17d"   
}
SUB_CAT_PALETTE = ["#124d54", "#f9744b", "#e9c46a", "#2a9d8f", "#457b9d", "#f4a261", "#8ab17d", "#e76f51", "#d62828", "#003049", "#fcbf49"]

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
                default_wallet = wallet_list[0] if wallet_list else "🏦 กรุงไทย"
                sheet.append_row([now_str, str(row['ประเภท']), str(row['หมวดหมู่']), float(row['จำนวนเงิน']), "บันทึกด่วน", default_wallet])
                fetch_main_data.clear()
                st.toast("Success! ✨")
                st.rerun()
                
    st.markdown("---")
    st.markdown("<p class='quick-add-text'>New Transaction</p>", unsafe_allow_html=True)
    
    type_entry = st.selectbox("Type", ["💸 รายจ่าย", "📥 รายรับ", "🔄 โอนย้ายกระเป๋า", "🐷 เงินออม", "📈 เงินลงทุน"])
    
    if "โอนย้ายกระเป๋า" in type_entry:
        c_tr1, c_tr2 = st.columns(2)
        with c_tr1:
            from_wallet = st.selectbox("โอนออกจาก (From):", wallet_list, key="mb_tr_from")
        with c_tr2:
            to_options = [w for w in wallet_list if w != from_wallet] if len(wallet_list) > 1 else wallet_list
            to_wallet = st.selectbox("เข้ากระเป๋า (To):", to_options, key="mb_tr_to")
        wallet_entry = from_wallet
        main_cat = "โอนย้ายระหว่างกระเป๋า"
        sub_cat = f"เข้า {to_wallet}"
    else:
        wallet_entry = st.selectbox("กระเป๋าเงิน (Wallet)", wallet_list, key="mb_wallet_select")
        if "เงินออม" in type_entry:
            # 🔥 ลบการกู้/คืนเงินกู้ ออกจากเมนูเงินออม
            sav_action = st.radio("การดำเนินการเงินออม:", ["📥 ฝากเงินเพิ่ม", "🔓 เบิกออกมาใช้"], horizontal=True)
            if "เบิก" in sav_action:
                st.caption("📉 *ระบบจะทำการหักเงินออกจากคลังเป้าหมาย และหักจากกระเป๋า '🌸 ออมสิน' โดยตรง*")
            
            mb_alloc_mode = st.radio("รูปแบบการจัดสรรเงินออม:", ["🎯 เป้าหมายเดียว (Single)", "🔀 แบ่งเปอร์เซ็นต์ (Split %)"], horizontal=True, key="mb_alloc_mode")
            
            if mb_alloc_mode == "🎯 เป้าหมายเดียว (Single)":
                selected_goal_mb = st.selectbox("🎯 เลือกเป้าหมายออมเงิน (Slot):", goal_options_list, key="mb_goal_slot")
                selected_goals_split_mb = []
            else:
                selected_goal_mb = None
                st.markdown("**🎯 เลือกเป้าหมายที่ต้องการจัดสรร:**")
                selected_goals_split_mb = st.multiselect("สล็อตเงินออม (เลือกได้หลายอัน):", goal_options_list, default=goal_options_list[:2] if len(goal_options_list)>=2 else goal_options_list, key="mb_goals_split", label_visibility="collapsed")
            
            st.markdown(f"""
                <div style='background-color: rgba(69, 123, 157, 0.1); border-left: 4px solid #457b9d; padding: 10px 15px; border-radius: 8px; margin-bottom: 10px;'>
                    <p style='margin:0; font-size: 13px; opacity: 0.8;'>💰 คลังเงินออมปัจจุบัน: <b>฿{total_sav_now:,.2f}</b></p>
                    {"<p style='margin:0; font-size:13px; color:#f9744b;'>⚠️ ยอดหนี้ค้างคืนคลัง: <b>฿" + f"{outstanding_loan:,.2f}</b></p>" if outstanding_loan > 0 else ""}
                </div>
            """, unsafe_allow_html=True)
            main_cat = "บริหารเงินออม"
            action_name = sav_action.split(" ")[1]
            sub_cat = action_name
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
        
        alloc_pcts_mb = {}
        if "เงินออม" in type_entry and mb_alloc_mode == "🔀 แบ่งเปอร์เซ็นต์ (Split %)":
            if len(selected_goals_split_mb) > 0:
                st.markdown("<p style='font-size:14px; font-weight:600; color:#2a9d8f; margin-bottom:5px;'>📊 ระบุสัดส่วน % (รวมต้องเท่ากับ 100%)</p>", unsafe_allow_html=True)
                cols_pct = st.columns(2)
                for idx, g in enumerate(selected_goals_split_mb):
                    with cols_pct[idx % 2]:
                        def_val = 100 // len(selected_goals_split_mb) if idx != len(selected_goals_split_mb)-1 else 100 - (100//len(selected_goals_split_mb))*(len(selected_goals_split_mb)-1)
                        alloc_pcts_mb[g] = st.number_input(f"{g} (%)", min_value=0.0, max_value=100.0, value=float(def_val), step=5.0, key=f"mb_pct_{idx}")
            else:
                st.warning("⚠️ กรุณาเลือกเป้าหมายที่ต้องการแบ่งเงินด้านบนก่อนครับ")
                
        note = st.text_input("Note", placeholder="Optional...")
        
        if st.form_submit_button("Save Transaction", use_container_width=True) and amount is not None and amount > 0:
            if "เงินออม" in type_entry and mb_alloc_mode == "🔀 แบ่งเปอร์เซ็นต์ (Split %)":
                if len(selected_goals_split_mb) == 0:
                    st.error("❌ กรุณาเลือกเป้าหมายที่จะแบ่งเงินครับ")
                    st.stop()
                total_pct = sum(alloc_pcts_mb.values())
                if total_pct != 100.0:
                    st.error(f"❌ สัดส่วนเปอร์เซ็นต์รวมต้องเท่ากับ 100% (ตอนนี้รวมได้ {total_pct}%)")
                    st.stop()
                    
            final_type = type_entry.split(" ")[1]
            final_time = datetime.datetime.now(TZ_TH).time() if time_shortcut == "⏱️ เวลาปัจจุบัน" else parse_custom_time(chosen_time_str, datetime.datetime.now(TZ_TH).time())
            combined_datetime = datetime.datetime.combine(chosen_date, final_time)
            
            if final_type == "เงินออม":
                # 🔥 จัดการลอจิกการหักเงินจากการเบิก
                if "เบิกออกมาใช้" in sav_action: 
                    final_type = "ถอนเงินออม"
                    wallet_entry = "🌸 ออมสิน" # บังคับหักจากออมสินเพื่อไม่ให้เป๋าอื่นรวน
                
                if mb_alloc_mode == "🎯 เป้าหมายเดียว (Single)":
                    sub_cat_final = f"{action_name} [{selected_goal_mb}]" if selected_goal_mb != "📦 คลังออมทั่วไป (ไม่ระบุเป้าหมาย)" else action_name
                    full_category = f"{main_cat}: {sub_cat_final}"
                    
                    if selected_goal_mb != "📦 คลังออมทั่วไป (ไม่ระบุเป้าหมาย)" and not df_goals.empty:
                        for g_idx, g_row in df_goals.iterrows():
                            if str(g_row["ชื่อเป้าหมาย"]) == selected_goal_mb:
                                curr_saved = float(g_row["สะสมแล้ว (บาท)"]) if pd.notnull(g_row["สะสมแล้ว (บาท)"]) else 0.0
                                if final_type == "เงินออม":
                                    new_saved = curr_saved + float(amount)
                                else:
                                    new_saved = max(0.0, curr_saved - float(amount))
                                goal_sheet.update_cell(int(g_idx) + 2, 4, new_saved)
                                fetch_goals.clear()
                                break
                    sheet.append_row([combined_datetime.strftime('%Y-%m-%d %H:%M:%S'), final_type, full_category, amount, note, wallet_entry])
                    
                else: 
                    for g, pct in alloc_pcts_mb.items():
                        if pct > 0:
                            split_amt = float(amount) * (pct / 100.0)
                            sub_cat_final = f"{action_name} [{g}]" if g != "📦 คลังออมทั่วไป (ไม่ระบุเป้าหมาย)" else action_name
                            full_category = f"{main_cat}: {sub_cat_final}"
                            split_note = f"{note} (แบ่ง {pct}%)" if note else f"แบ่ง {pct}%"
                            
                            if g != "📦 คลังออมทั่วไป (ไม่ระบุเป้าหมาย)" and not df_goals.empty:
                                for g_idx, g_row in df_goals.iterrows():
                                    if str(g_row["ชื่อเป้าหมาย"]) == g:
                                        curr_saved = float(g_row["สะสมแล้ว (บาท)"]) if pd.notnull(g_row["สะสมแล้ว (บาท)"]) else 0.0
                                        if final_type == "เงินออม":
                                            new_saved = curr_saved + float(split_amt)
                                        else:
                                            new_saved = max(0.0, curr_saved - float(split_amt))
                                        goal_sheet.update_cell(int(g_idx) + 2, 4, new_saved)
                                        fetch_goals.clear()
                                        break
                            sheet.append_row([combined_datetime.strftime('%Y-%m-%d %H:%M:%S'), final_type, full_category, split_amt, split_note, wallet_entry])
            else:
                full_category = f"{main_cat}: {sub_cat}" if sub_cat != "ทั่วไป" else main_cat
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
                        default_wallet = wallet_list[0] if wallet_list else "🏦 กรุงไทย"
                        sheet.append_row([now_str, str(row['ประเภท']), str(row['หมวดหมู่']), float(row['จำนวนเงิน']), "บันทึกด่วน", default_wallet])
                        fetch_main_data.clear()
                        st.toast("Success! ✨")
                        st.rerun()
                        
            st.markdown("---")
            st.markdown("<p class='quick-add-text'>New Transaction</p>", unsafe_allow_html=True)
            
            c_type, c_wallet = st.columns([3, 1.5])
            with c_type:
                type_entry = st.radio("Type", ["📥 รายรับ", "💸 รายจ่าย", "🔄 โอนย้ายกระเป๋า", "🐷 เงินออม", "📈 เงินลงทุน"], horizontal=True, label_visibility="collapsed")
            with c_wallet:
                if "โอนย้ายกระเป๋า" not in type_entry:
                    wallet_entry = st.selectbox("กระเป๋าเงิน (Wallet):", wallet_list, label_visibility="collapsed", key="dt_wallet_select")
            
            if "โอนย้ายกระเป๋า" in type_entry:
                c_tf1, c_tf2 = st.columns(2)
                with c_tf1:
                    from_wallet = st.selectbox("โอนออกจากกระเป๋า (From):", wallet_list, key="dt_tr_from")
                with c_tf2:
                    to_options = [w for w in wallet_list if w != from_wallet] if len(wallet_list) > 1 else wallet_list
                    to_wallet = st.selectbox("โอนเข้ากระเป๋า (To):", to_options, key="dt_tr_to")
                wallet_entry = from_wallet
                main_cat = "โอนย้ายระหว่างกระเป๋า"
                sub_cat = f"เข้า {to_wallet}"
            elif "เงินออม" in type_entry:
                # 🔥 ลบการกู้/คืนเงินกู้ ออกจากเมนูเงินออม
                sav_action = st.radio("การดำเนินการเงินออม:", ["📥 ฝากเงินเพิ่ม", "🔓 เบิกออกมาใช้"], horizontal=True, key="dt_sav_action")
                if "เบิก" in sav_action:
                    st.caption("📉 *ระบบจะทำการหักเงินออกจากคลังเป้าหมาย และหักจากกระเป๋า '🌸 ออมสิน' โดยตรง*")
                
                dt_alloc_mode = st.radio("รูปแบบการจัดสรรเงินออม:", ["🎯 เป้าหมายเดียว (Single)", "🔀 แบ่งเปอร์เซ็นต์ (Split %)"], horizontal=True, key="dt_alloc_mode")
                
                if dt_alloc_mode == "🎯 เป้าหมายเดียว (Single)":
                    selected_goal_dt = st.selectbox("🎯 เลือกเป้าหมายออมเงิน (Slot):", goal_options_list, key="dt_goal_slot")
                    selected_goals_split_dt = []
                else:
                    selected_goal_dt = None
                    st.markdown("**🎯 เลือกเป้าหมายที่ต้องการจัดสรร:**")
                    selected_goals_split_dt = st.multiselect("สล็อตเงินออม (เลือกได้หลายอัน):", goal_options_list, default=goal_options_list[:2] if len(goal_options_list)>=2 else goal_options_list, key="dt_goals_split", label_visibility="collapsed")
                
                st.markdown(f"""
                    <div style='background-color: rgba(69, 123, 157, 0.1); border-left: 4px solid #457b9d; padding: 12px 20px; border-radius: 8px; margin: 10px 0;'>
                        <p style='margin:0; font-size: 14px; opacity: 0.8;'>💰 คลังเงินออมปัจจุบัน: <b>฿{total_sav_now:,.2f}</b></p>
                        {"<p style='margin:0; color: #f9744b; font-size: 14px; font-weight:600;'>⚠️ ยอดหนี้ค้างคืนคลัง: ฿" + f"{outstanding_loan:,.2f}</p>" if outstanding_loan > 0 else ""}
                    </div>
                """, unsafe_allow_html=True)
                main_cat = "บริหารเงินออม"
                action_name = sav_action.split(" ")[1]
                sub_cat = action_name
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
                
                alloc_pcts_dt = {}
                if "เงินออม" in type_entry and dt_alloc_mode == "🔀 แบ่งเปอร์เซ็นต์ (Split %)":
                    if len(selected_goals_split_dt) > 0:
                        st.markdown("<p style='font-size:14px; font-weight:600; color:#2a9d8f; margin-bottom:5px;'>📊 ระบุสัดส่วน % (รวมต้องเท่ากับ 100%)</p>", unsafe_allow_html=True)
                        cols_pct = st.columns(2)
                        for idx, g in enumerate(selected_goals_split_dt):
                            with cols_pct[idx % 2]:
                                def_val = 100 // len(selected_goals_split_dt) if idx != len(selected_goals_split_dt)-1 else 100 - (100//len(selected_goals_split_dt))*(len(selected_goals_split_dt)-1)
                                alloc_pcts_dt[g] = st.number_input(f"{g} (%)", min_value=0.0, max_value=100.0, value=float(def_val), step=5.0, key=f"dt_pct_{idx}")
                    else:
                        st.warning("⚠️ กรุณาเลือกเป้าหมายที่ต้องการแบ่งเงินด้านบนก่อนครับ")
                        
                note = st.text_input("Note", placeholder="...")
                
                if st.form_submit_button("Save Transaction", use_container_width=True) and amount is not None and amount > 0:
                    if "เงินออม" in type_entry and dt_alloc_mode == "🔀 แบ่งเปอร์เซ็นต์ (Split %)":
                        if len(selected_goals_split_dt) == 0:
                            st.error("❌ กรุณาเลือกเป้าหมายที่จะแบ่งเงินครับ")
                            st.stop()
                        total_pct = sum(alloc_pcts_dt.values())
                        if total_pct != 100.0:
                            st.error(f"❌ สัดส่วนเปอร์เซ็นต์รวมต้องเท่ากับ 100% (ตอนนี้รวมได้ {total_pct}%)")
                            st.stop()
                            
                    final_type = type_entry.split(" ")[1]
                    final_time_dt = datetime.datetime.now(TZ_TH).time() if time_shortcut_dt == "⏱️ เวลาปัจจุบัน" else parse_custom_time(chosen_time_dt_str, datetime.datetime.now(TZ_TH).time())
                    combined_datetime = datetime.datetime.combine(chosen_date_dt, final_time_dt)
                    
                    if final_type == "เงินออม":
                        # 🔥 จัดการลอจิกการหักเงินจากการเบิก
                        if "เบิกออกมาใช้" in sav_action: 
                            final_type = "ถอนเงินออม"
                            wallet_entry = "🌸 ออมสิน" # บังคับหักจากออมสินเพื่อไม่ให้เป๋าอื่นรวน
                        
                        if dt_alloc_mode == "🎯 เป้าหมายเดียว (Single)":
                            sub_cat_final = f"{action_name} [{selected_goal_dt}]" if selected_goal_dt != "📦 คลังออมทั่วไป (ไม่ระบุเป้าหมาย)" else action_name
                            full_category = f"{main_cat}: {sub_cat_final}"
                            
                            if selected_goal_dt != "📦 คลังออมทั่วไป (ไม่ระบุเป้าหมาย)" and not df_goals.empty:
                                for g_idx, g_row in df_goals.iterrows():
                                    if str(g_row["ชื่อเป้าหมาย"]) == selected_goal_dt:
                                        curr_saved = float(g_row["สะสมแล้ว (บาท)"]) if pd.notnull(g_row["สะสมแล้ว (บาท)"]) else 0.0
                                        if final_type == "เงินออม":
                                            new_saved = curr_saved + float(amount)
                                        else:
                                            new_saved = max(0.0, curr_saved - float(amount))
                                        goal_sheet.update_cell(int(g_idx) + 2, 4, new_saved)
                                        fetch_goals.clear()
                                        break
                            sheet.append_row([combined_datetime.strftime('%Y-%m-%d %H:%M:%S'), final_type, full_category, amount, note, wallet_entry])
                            
                        else: 
                            for g, pct in alloc_pcts_dt.items():
                                if pct > 0:
                                    split_amt = float(amount) * (pct / 100.0)
                                    sub_cat_final = f"{action_name} [{g}]" if g != "📦 คลังออมทั่วไป (ไม่ระบุเป้าหมาย)" else action_name
                                    full_category = f"{main_cat}: {sub_cat_final}"
                                    split_note = f"{note} (แบ่ง {pct}%)" if note else f"แบ่ง {pct}%"
                                    
                                    if g != "📦 คลังออมทั่วไป (ไม่ระบุเป้าหมาย)" and not df_goals.empty:
                                        for g_idx, g_row in df_goals.iterrows():
                                            if str(g_row["ชื่อเป้าหมาย"]) == g:
                                                curr_saved = float(g_row["สะสมแล้ว (บาท)"]) if pd.notnull(g_row["สะสมแล้ว (บาท)"]) else 0.0
                                                if final_type == "เงินออม":
                                                    new_saved = curr_saved + float(split_amt)
                                                else:
                                                    new_saved = max(0.0, curr_saved - float(split_amt))
                                                goal_sheet.update_cell(int(g_idx) + 2, 4, new_saved)
                                                fetch_goals.clear()
                                                break
                                    sheet.append_row([combined_datetime.strftime('%Y-%m-%d %H:%M:%S'), final_type, full_category, split_amt, split_note, wallet_entry])
                    else:
                        full_category = f"{main_cat}: {sub_cat}" if sub_cat != "ทั่วไป" else main_cat
                        sheet.append_row([combined_datetime.strftime('%Y-%m-%d %H:%M:%S'), final_type, full_category, amount, note, wallet_entry])
                    
                    fetch_main_data.clear()
                    st.rerun()

    # ==========================================
    # 📊 Tab 2: Dashboard
    # ==========================================
    with tab2:
        if not df.empty:
            has_mom_7500 = False
            for _, r_check in df.iterrows():
                if "7500" in str(r_check['จำนวนเงิน']) and ("06-25" in str(r_check['วันที่']) or "6-25" in str(r_check['วันที่'])):
                    has_mom_7500 = True
                    break
            
            if not has_mom_7500:
                with st.expander("🚨 พบสาเหตุยอดกรุงไทยติดลบ (-5,030.50 และ -6,811.21): แถวรับเงินเดือนแม่ 7,500 บาท (2026-06-25) หายไปจาก Google Sheets!", expanded=True):
                    st.markdown("💡 เนื่องจากก่อนหน้านี้ระบบบันทึกทับโดยไม่มีแถวรับเงินเดือน 7,500 บาทแรกสุด ส่งผลให้ยอดกรุงไทยติดลบทันที **สามารถกดปุ่มสีส้มด้านล่างเพื่อกู้คืนข้อมูล 7,500 บาทลง Google Sheets อัตโนมัติในคลิกเดียวครับ!**")
                    if st.button("⚡ กู้คืนยอดรับเงินเดือนแม่ 7,500 บาท (2026-06-25) เข้ากรุงไทยทันที", use_container_width=True):
                        sheet.insert_row(["2026-06-25 00:00:00", "รายรับ", "เงินเดือน: จากแม่", 7500.0, "แม่โอนให้ใช้", "🏦 กรุงไทย"], 1)
                        fetch_main_data.clear()
                        st.success("🎉 กู้คืนยอดเงินเดือนแม่ 7,500 บาท เรียบร้อยแล้ว! ยอดกรุงไทยกลับมาเป็นบวกตามจริงแล้วครับ")
                        st.rerun()

            df_chart = df.copy()
            
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
            
            df_cycle = df_chart.copy()
            df_cumulative = df_chart.copy()
            selected_carry = 0.0
            selected_kt_real = 0.0
            selected_row_idx = None
            
            if selected_cycle_label != "🌟 แสดงข้อมูลทั้งหมด (All Time)":
                for label, start_str, end_str, carry_val, kt_val, r_idx in cycle_options:
                    if label == selected_cycle_label:
                        selected_carry = carry_val
                        selected_kt_real = kt_val
                        selected_row_idx = r_idx
                        start_dt = parse_clean_datetime(pd.Series([start_str])).iloc[0]
                        if end_str and pd.notnull(end_str) and str(end_str).strip() != "":
                            end_dt = parse_clean_datetime(pd.Series([end_str])).iloc[0]
                            df_cycle = df_cycle[(df_cycle['วันเวลา'] >= start_dt) & (df_cycle['วันเวลา'] <= end_dt)]
                            df_cumulative = df_cumulative[df_cumulative['วันเวลา'] <= end_dt]
                        else:
                            df_cycle = df_cycle[df_cycle['วันเวลา'] >= start_dt]
                        break

            inv_mask = (
                df_cycle['ประเภท'].astype(str).str.contains('ลงทุน|invest', case=False, na=False) |
                df_cycle['หมวดหมู่'].astype(str).str.contains('ลงทุน|invest|หุ้น|กองทุน|crypto|คริปโต|gold|ทอง', case=False, na=False) |
                df_cycle['รายละเอียด'].astype(str).str.contains('ลงทุน|invest|หุ้น|กองทุน|crypto|คริปโต|gold|ทอง', case=False, na=False)
            ) & ~df_cycle['ประเภท'].astype(str).str.contains('ถอน|คืน|ปรับยอด', case=False, na=False)
            inv = float(df_cycle[inv_mask]['จำนวนเงิน'].sum())

            _, _, _, _, sav_flow, _ = calculate_savings_metrics(df_cycle)

            inc_mask = (
                df_cycle['ประเภท'].astype(str).str.contains('รายรับ|income', case=False, na=False) &
                ~df_cycle['ประเภท'].astype(str).str.contains('รับคืน|คืน|ปรับยอด', case=False, na=False)
            )
            inc = float(df_cycle[inc_mask]['จำนวนเงิน'].sum())

            exp_mask = (
                df_cycle['ประเภท'].astype(str).str.contains('รายจ่าย|expense', case=False, na=False) &
                ~inv_mask &
                ~df_cycle['หมวดหมู่'].astype(str).str.contains('ฝากออม|ออมเงิน|เงินออม|เก็บออม|aomsin|ออมสิน', case=False, na=False) &
                ~df_cycle['ประเภท'].astype(str).str.contains('ปรับยอด', case=False, na=False)
            )
            exp = float(df_cycle[exp_mask]['จำนวนเงิน'].sum())
            expense_df = df_cycle[exp_mask].copy()

            def is_transfer_row(row_type):
                t = str(row_type).strip().lower()
                return 'โอนย้าย' in t or 'transfer' in t

            def is_income_type(row_type):
                v = str(row_type).strip()
                return bool('รายรับ' in v or 'ถอนเงินออม' in v or 'กู้เงินออม'จัดการแก้ปัญหาระบบ Transaction หมวดเงินออมตามที่หมอต้องการได้โดยปรับปรุงโค้ดทั้งฝั่ง UI และ Logic การคำนวณดังนี้ครับ

**1. ปรับฟังก์ชัน "เบิกออกมาใช้" ให้หักเงินตามเป้าหมาย (Withdrawal Logic)**
*   **ผูก Dropdown ให้แสดงเสมอ:** เช็คเงื่อนไข (Condition) ในหน้า UI ว่าเมื่อเลือก Radio Button "เบิกออกมาใช้" (Withdraw) ตัว Dropdown `เลือกเป้าหมายออมเงิน (Slot)` ต้องแสดงขึ้นมาให้เลือกเสมอ เพื่อให้หมอสามารถดึงข้อมูล Slot เช่น "GTA VI+PS5" หรือ "เชียงคานกับฟ้า" มาใช้งานได้
*   **ปรับ Logic การคำนวณตอน Save:** ปัญหาที่กดเบิกแล้วยังกลายเป็นฝาก เกิดจากระบบยังมอง `Amount` เป็นค่าบวกเสมอ ในฟังก์ชัน `SaveTransaction` ต้องเพิ่มเงื่อนไขเช็ค Action หากเป็น "เบิกออกมาใช้" ให้หักลบยอดเงินแทน
    *   *ตัวอย่างแนวทางโค้ด:* `const transactionAmount = action === 'withdraw' ? -Math.abs(amount) : amount;`
*   **อัปเดต Database 2 จุด:** เมื่อกด Save ระบบจะต้องนำค่าที่ติดลบไปคำนวณ (บวกด้วยค่าติดลบ = หักออก) จาก 2 แหล่งพร้อมกัน คือ 1) ยอดรวมของกระเป๋าบัญชีออมสินที่รับผิดชอบ และ 2) ยอดสะสมภายใน Slot เป้าหมายนั้นๆ 

**2. ตัดตัวเลือก "กู้เงินคลัง" และ "โอนคืนเงินกู้" ออก (UI Cleanup)**
*   **ลบ Radio Buttons:** เข้าไปที่ Component หรือไฟล์ HTML/JSX ส่วนที่จัดการ `การดำเนินการเงินออม:` แล้วลบ Element ของ `กู้เงินคลัง (ต้องคืน)` และ `โอนคืนเงินกู้` ทิ้งได้เลย เพื่อไม่ให้รกและไปทับซ้อนกับโหมดกู้ยืมที่มีอยู่แล้ว
*   **จัดการค่าเริ่มต้น (Default State):** ตรวจสอบตัวแปร State ที่เก็บค่าการดำเนินการ หากเคยตั้ง Default ไว้ที่ตัวเลือกที่เพิ่งลบไป ให้เปลี่ยนค่า Default กลับมาเป็นตัวเลือก `ฝากเงินเพิ่ม` (Deposit) เสมอเมื่อเปิดหน้า New Transaction 
*   **เคลียร์ Logic ที่ไม่ใช้:** ลบฟังก์ชันหรือ Switch-case ที่คอยจัดการกับ Action `borrow` และ `repay` ในหน้าหมวดเงินออมนี้ออกไป เพื่อให้โค้ดสะอาดขึ้นและเหลือแค่ `deposit` กับ `withdraw` 

หากหมอใช้ Framework ตัวไหนอยู่ (เช่น React, Vue หรือ Flutter) แล้วอยากให้ช่วยดูจุดที่ต้องแก้ใน Source Code เพิ่มเติม สามารถส่งส่วนที่จัดการ State ของหน้านี้มาให้ดูได้เลยครับ
