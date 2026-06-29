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

# 🔤 CSS สไตล์ Soft UI ที่รองรับทั้ง Light & Dark Mode อัตโนมัติ
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;500;600;700&family=Prompt:wght@300;400;500;600&display=swap');
    
    html, body, [class*="css"] { 
        font-family: 'Poppins', 'Prompt', sans-serif !important; 
    }
    
    h1, h2, h3 { font-weight: 700; color: var(--text-color); }
    
    /* แต่งปุ่มสไตล์คลีน */
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
    
    /* 📌 การ์ดแสดงผลตัวเลข (Metric Card) ที่ปรับตัวตามโหมดมืดสว่าง */
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
    
    /* ซ่อนลูกศรในช่องกรอกตัวเลข */
    input[type=number]::-webkit-inner-spin-button, input[type=number]::-webkit-outer-spin-button { -webkit-appearance: none; margin: 0; }
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

# 🎨 ชุดสี Honey Pot & Soft Tones (อัปเดตเพิ่มเส้นเงินสุทธิ)
HONEY_POT_MAP = {
    "รายรับ": "#2a9d8f",     
    "รายจ่าย": "#f9744b",    
    "เงินออม": "#457b9d",    
    "เงินลงทุน": "#e9c46a",
    "เงินสุทธิ": "#8ab17d"   # เพิ่มสีเขียวละมุน (Soft Green) สำหรับเส้นเงินสุทธิ
}

# ชุดสีคุมโทน สำหรับกราฟโดนัทและแท่งให้ล้อกัน
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

    # --- Tab 2: Dashboard ---
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
            net_title_class = "metric-title" if net >= 0 else "metric-title-alert"
            m1.markdown(f"<div class='metric-card'><div class='{net_title_class}'>Net Balance</div><div class='metric-value'>฿{net:,.0f}</div><div class='metric-currency'>THB</div></div>", unsafe_allow_html=True)
            m2.markdown(f"<div class='metric-card'><div class='metric-title'>Income <span style='color:#2a9d8f;'>↗</span></div><div class='metric-value'>฿{inc:,.0f}</div><div class='metric-currency'>THB</div></div>", unsafe_allow_html=True)
            m3.markdown(f"<div class='metric-card'><div class='metric-title'>Expenses <span style='color:#f9744b;'>↘</span></div><div class='metric-value'>฿{exp:,.0f}</div><div class='metric-currency'>THB</div></div>", unsafe_allow_html=True)
            m4.markdown(f"<div class='metric-card'><div class='metric-title'>Savings <span style='color:#457b9d;'>↗</span></div><div class='metric-value'>฿{sav:,.0f}</div><div class='metric-currency'>THB</div></div>", unsafe_allow_html=True)
            m5.markdown(f"<div class='metric-card'><div class='metric-title'>Investments <span style='color:#e9c46a;'>↗</span></div><div class='metric-value'>฿{inv:,.0f}</div><div class='metric-currency'>THB</div></div>", unsafe_allow_html=True)
            
            # --- 📈 กราฟเส้นแบบ Minimal (เพิ่มเส้นเงินสุทธิ) ---
            st.markdown("<p class='quick-add-text' style='margin-top: 20px;'>Trend Analysis</p>", unsafe_allow_html=True)
            time_frame = st.radio("Timeframe:", ["รายวัน", "รายเดือน", "รายปี"], horizontal=True, label_visibility="collapsed")
            
            df_trend = df_chart.copy()
            if time_frame == "รายปี":
                df_trend['เวลา'] = df_trend['วันที่'].dt.strftime('%Y')
            elif time_frame == "รายเดือน":
                df_trend['เวลา'] = df_trend['วันที่'].dt.strftime('%Y-%m')
            else:
                df_trend['เวลา'] = df_trend['วันที่'].dt.strftime('%Y-%m-%d')
                
            trend_data_raw = df_trend.groupby(['เวลา', 'ประเภท'])['จำนวนเงิน'].sum().reset_index()
            
            # 💡 กลไกคำนวณเงินสุทธิ (Net Balance) ประจำช่วงเวลา
            pivot_trend = trend_data_raw.pivot(index='เวลา', columns='ประเภท', values='จำนวนเงิน').fillna(0)
            
            # ตรวจสอบว่าคอลัมน์ไหนหายไปก็เติม 0 ไว้กัน Error
            for col in ['รายรับ', 'รายจ่าย', 'เงินออม', 'เงินลงทุน']:
                if col not in pivot_trend.columns:
                    pivot_trend[col] = 0
            
            # คำนวณสมการ: เงินสุทธิ = รายรับ - (รายจ่าย + ออม + ลงทุน)
            pivot_trend['เงินสุทธิ'] = pivot_trend['รายรับ'] - (pivot_trend['รายจ่าย'] + pivot_trend['เงินออม'] + pivot_trend['เงินลงทุน'])
            
            # แปลงตารางกลับมาให้ Plotly สามารถนำไปวาดเส้นได้
            trend_data = pivot_trend.reset_index().melt(id_vars='เวลา', var_name='ประเภท', value_name='จำนวนเงิน')
            
            fig_trend = px.line(trend_data, x='เวลา', y='จำนวนเงิน', color='ประเภท', 
                                color_discrete_map=HONEY_POT_MAP, markers=True, line_shape='spline')
            
            fig_trend.update_traces(line=dict(width=3), marker=dict(size=8, line=dict(width=1, color="white")))
            fig_trend.update_layout(
                plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', 
                xaxis=dict(showgrid=False, title="", showline=False, tickfont=dict(family='Poppins')),
                yaxis=dict(showgrid=True, gridcolor='rgba(128, 128, 128, 0.1)', title="", zeroline=False, tickfont=dict(family='Poppins')),
                legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="center", x=0.5, title="", font=dict(family='Poppins')),
                hovermode="x unified", margin=dict(t=10, b=0, l=0, r=0)
            )
            st.plotly_chart(fig_trend, use_container_width=True, theme="streamlit")
            
            st.markdown("---")
            
            # --- ⭕ Infographic สัดส่วนรายจ่าย (จับคู่สีให้ตรงกัน) ---
            col_chart1, col_chart2 = st.columns([1, 1.2])
            expense_df = df_chart[df_chart['ประเภท'] == 'รายจ่าย']
            
            with col_chart1:
                st.markdown("<p class='quick-add-text'>Expense Breakdown</p>", unsafe_allow_html=True)
                if not expense_df.empty:
                    pie_data = expense_df.groupby('หมวดหมู่หลัก')['จำนวนเงิน'].sum().reset_index()
                    
                    unique_main_cats = pie_data['หมวดหมู่หลัก'].tolist()
                    cat_color_map = {cat: SUB_CAT_PALETTE[i % len(SUB_CAT_PALETTE)] for i, cat in enumerate(unique_main_cats)}
                    
                    fig_pie = px.pie(pie_data, values='จำนวนเงิน', names='หมวดหมู่หลัก', hole=0.75, 
                                     color='หมวดหมู่หลัก', color_discrete_map=cat_color_map)
                    
                    fig_pie.update_traces(textposition='outside', textinfo='percent+label', marker=dict(line=dict(width=0)), textfont=dict(family='Poppins'))
                    fig_pie.update_layout(plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)', showlegend=False, margin=dict(t=20, b=20, l=20, r=20))
                    
                    st.plotly_chart(fig_pie, use_container_width=True, theme="streamlit")
                else:
                    st.info("No expense data.")
                    
            with col_chart2:
                st.markdown("<p class='quick-add-text'>Top Sub-categories</p>", unsafe_allow_html=True)
                if not expense_df.empty:
                    sub_data = expense_df.groupby(['หมวดหมู่หลัก', 'หมวดหมู่ย่อย'])['จำนวนเงิน'].sum().reset_index()
                    sub_data = sub_data.sort_values(by="จำนวนเงิน", ascending=False).head(8) 
                    
                    fig_bar = px.bar(sub_data, x='จำนวนเงิน', y='หมวดหมู่ย่อย', color='หมวดหมู่หลัก', orientation='h',
                                     color_discrete_map=cat_color_map) 
                    
                    fig_bar.update_traces(marker_line_width=0, opacity=0.9, texttemplate='฿%{x:,.0f}', textposition='outside', textfont=dict(family='Poppins'))
                    fig_bar.update_layout(
                        plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)',
                        xaxis=dict(showgrid=False, title="", zeroline=False, showticklabels=False),
                        yaxis=dict(showgrid=False, title="", autorange="reversed", tickfont=dict(family='Poppins')),
                        showlegend=False, margin=dict(t=20, b=20, l=0, r=20)
                    )
                    
                    st.plotly_chart(fig_bar, use_container_width=True, theme="streamlit")
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
        edited_cat = st.data_editor(cat_raw_df, use_container_width=True, num_rows="dynamic", key="editor_cat_v9")
        if st.button("💾 Save Categories", use_container_width=True):
            cat_sheet.clear()
            cat_sheet.update(range_name="A1", values=[edited_cat.columns.values.tolist()] + edited_cat.values.tolist())
            st.success("Categories updated! ✨")
            st.rerun()

        st.markdown("---")
        st.subheader("⚡ Quick Adds Editor")
        edited_qa = st.data_editor(qa_df, use_container_width=True, num_rows="dynamic", key="editor_qa_v9")
        if st.button("💾 Save Quick Adds", use_container_width=True):
            qa_sheet.clear()
            qa_sheet.update(range_name="A1", values=[edited_qa.columns.values.tolist()] + edited_qa.values.tolist())
            st.success("Quick adds updated!")
            st.rerun()
            
        st.markdown("---")
        st.subheader("✏️ Raw Data Editor")
        if not df.empty:
            clean_df_edit = df[["วันที่", "ประเภท", "หมวดหมู่", "จำนวนเงิน", "รายละเอียด"]]
            edited_df = st.data_editor(clean_df_edit, use_container_width=True, num_rows="dynamic", key="editor_finance_v9")
            if st.button("💾 Save Data to Cloud", use_container_width=True):
                sheet.clear()
                edited_df['วันที่'] = edited_df['วันที่'].astype(str)
                sheet.update(range_name="A1", values=[edited_df.columns.values.tolist()] + edited_df.values.tolist())
                st.success("Data updated!")
                st.rerun()
