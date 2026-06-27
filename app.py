import streamlit as st
import pandas as pd
import datetime
import os
import plotly.express as px

st.set_page_config(page_title="Minimal Finance Pro", layout="wide")

st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Prompt:wght@300;400;500;700&display=swap');
    h1, h2, h3, h4, h5, h6, p, label, input, select, .stButton>button {
        font-family: 'Prompt', sans-serif !important;
    }
    h1 { font-weight: 700; color: #2C3E50; text-align: center; margin-bottom: 1.5rem; }
    .stButton>button { border-radius: 8px; font-weight: 500; }
    .quick-add-text { font-size: 16px; font-weight: bold; color: #2C3E50; margin-bottom: 5px; }
    .metric-card { background-color: #F8F9FA; padding: 15px; border-radius: 10px; text-align: center; box-shadow: 0 2px 5px rgba(0,0,0,0.05); }
    </style>
""", unsafe_allow_html=True)

st.title("☁️ Minimal Finance Pro")

DATA_FILE = "my_finance_data.csv"
QUICK_ADD_FILE = "quick_adds.csv"

def load_data():
    if os.path.exists(DATA_FILE):
        df = pd.read_csv(DATA_FILE)
        df['วันที่'] = pd.to_datetime(df['วันที่']).dt.date
        return df
    return pd.DataFrame(columns=["วันที่", "ประเภท", "หมวดหมู่", "จำนวนเงิน", "รายละเอียด"])

def load_quick_adds():
    if os.path.exists(QUICK_ADD_FILE):
        return pd.read_csv(QUICK_ADD_FILE)
    return pd.DataFrame([
        {"ชื่อปุ่ม": "🍛 อาหาร 50฿", "ประเภท": "รายจ่าย", "หมวดหมู่": "อาหาร/เครื่องดื่ม", "จำนวนเงิน": 50.0},
        {"ชื่อปุ่ม": "🚌 เดินทางไปหา Alice", "ประเภท": "รายจ่าย", "หมวดหมู่": "ค่าเดินทาง", "จำนวนเงิน": 150.0},
        {"ชื่อปุ่ม": "📚 เก็บเงินสอบ GRE", "ประเภท": "เงินออม", "หมวดหมู่": "ออมเพื่อเรียนต่อ/อนาคต", "จำนวนเงิน": 500.0},
        {"ชื่อปุ่ม": "💻 ซื้ออุปกรณ์", "ประเภท": "รายจ่าย", "หมวดหมู่": "อุปกรณ์ไอที/เขียนงาน", "จำนวนเงิน": 200.0}
    ])

def save_data(df):
    df.to_csv(DATA_FILE, index=False)

def save_quick_adds(df):
    df.to_csv(QUICK_ADD_FILE, index=False)

if 'finance_data' not in st.session_state:
    st.session_state.finance_data = load_data()
if 'quick_adds' not in st.session_state:
    st.session_state.quick_adds = load_quick_adds()

def add_entry(date, entry_type, category, amount, note):
    new_data = pd.DataFrame([{"วันที่": date, "ประเภท": entry_type, "หมวดหมู่": category, "จำนวนเงิน": amount, "รายละเอียด": note}])
    st.session_state.finance_data = pd.concat([st.session_state.finance_data, new_data], ignore_index=True)
    save_data(st.session_state.finance_data)
    st.toast(f"บันทึก {amount} บาท สำเร็จ! ✨")

tab1, tab2, tab3, tab4 = st.tabs(["✨ บันทึกเงิน", "📊 วิเคราะห์ Infographic", "🎯 เป้าหมาย", "⚙️ ตั้งค่า/จัดการ"])

# --- Tab 1: หน้าบันทึกเงิน ---
with tab1:
    col_main, col_space = st.columns([2, 1])
    with col_main:
        st.markdown("<p class='quick-add-text'>⚡ บันทึกด่วน</p>", unsafe_allow_html=True)
        qa_df = st.session_state.quick_adds
        if not qa_df.empty:
            cols = st.columns(4)
            for i, row in qa_df.iterrows():
                col = cols[i % 4]
                
                # --- จุดแก้ไขความปลอดภัยข้อมูล (ป้องกัน Error ตัวหนังสือซ้อนหรือช่องว่าง) ---
                button_label = str(row['ชื่อปุ่ม']) if pd.notna(row['ชื่อปุ่ม']) else f"ปุ่มลัด {i+1}"
                btn_type = str(row['ประเภท']) if pd.notna(row['ประเภท']) else "รายจ่าย"
                btn_cat = str(row['หมวดหมู่']) if pd.notna(row['หมวดหมู่']) else "อื่นๆ"
                btn_amt = float(row['จำนวนเงิน']) if pd.notna(row['จำนวนเงิน']) else 0.0
                
                if col.button(button_label, use_container_width=True):
                    if btn_amt > 0:
                        add_entry(datetime.date.today(), btn_type, btn_cat, btn_amt, "บันทึกด่วน")
                        st.rerun()
        else:
            st.caption("ยังไม่มีปุ่มบันทึกด่วน สามารถเพิ่มได้ที่แท็บ '⚙️ ตั้งค่าและจัดการ'")

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

        with st.form("detail_form", clear_on_submit=True):
            amount = st.number_input("จำนวนเงิน (บาท)", min_value=0.0, step=50.0, format="%.2f")
            category = st.selectbox("หมวดหมู่", categories)
            note = st.text_input("รายละเอียด", placeholder="เช่น เติมน้ำมัน, ค่าหนังสือ")
            date = st.date_input("วันที่", datetime.date.today())
            if st.form_submit_button("บันทึกรายการ", use_container_width=True) and amount > 0:
                add_entry(date, clean_type, category, amount, note)
                st.rerun()

# --- Tab 2: หน้าวิเคราะห์ Infographic ---
with tab2:
    df = st.session_state.finance_data.copy()
    if not df.empty:
        df['วันที่'] = pd.to_datetime(df['วันที่'])
        df['ปี'] = df['วันที่'].dt.year
        df['เดือน'] = df['วันที่'].dt.month
        df['ชื่อเดือน'] = df['วันที่'].dt.strftime('%b')
        
        st.markdown("### 🔍 ตัวกรองข้อมูล")
        f_col1, f_col2, f_col3 = st.columns(3)
        year_list = ["ภาพรวมทุกปี"] + sorted(list(df['ปี'].unique()), reverse=True)
        selected_year = f_col1.selectbox("เลือกปี", year_list)
        
        if selected_year != "ภาพรวมทุกปี":
            df_filtered = df[df['ปี'] == selected_year]
            month_list = ["ภาพรวมทั้งปี"] + sorted(list(df_filtered['เดือน'].unique()))
            selected_month = f_col2.selectbox("เลือกเดือน", month_list)
            if selected_month != "ภาพรวมทั้งปี":
                df_filtered = df_filtered[df_filtered['เดือน'] == selected_month]
        else:
            df_filtered = df
            selected_month = "ภาพรวมทั้งปี"
            
        st.markdown("---")
        
        inc = df_filtered[df_filtered['ประเภท'] == 'รายรับ']['จำนวนเงิน'].sum()
        exp = df_filtered[df_filtered['ประเภท'] == 'รายจ่าย']['จำนวนเงิน'].sum()
        sav = df_filtered[df_filtered['ประเภท'] == 'เงินออม']['จำนวนเงิน'].sum()
        inv = df_filtered[df_filtered['ประเภท'] == 'เงินลงทุน']['จำนวนเงิน'].sum()
        net = inc - (exp + sav + inv)

        m1, m2, m3, m4, m5 = st.columns(5)
        m1.markdown(f"<div class='metric-card'><p style='margin:0;color:#7f8c8d;'>เงินคงเหลือสุทธิ</p><h3 style='margin:0;color:{'#2ECC71' if net >= 0 else '#E74C3C'};'>฿{net:,.0f}</h3></div>", unsafe_allow_html=True)
        m2.markdown(f"<div class='metric-card'><p style='margin:0;color:#7f8c8d;'>📥 รายรับ</p><h3 style='margin:0;color:#3498db;'>฿{inc:,.0f}</h3></div>", unsafe_allow_html=True)
        m3.markdown(f"<div class='metric-card'><p style='margin:0;color:#7f8c8d;'>💸 รายจ่าย</p><h3 style='margin:0;color:#e74c3c;'>฿{exp:,.0f}</h3></div>", unsafe_allow_html=True)
        m4.markdown(f"<div class='metric-card'><p style='margin:0;color:#7f8c8d;'>🐷 เงินออม</p><h3 style='margin:0;color:#f1c40f;'>฿{sav:,.0f}</h3></div>", unsafe_allow_html=True)
        m5.markdown(f"<div class='metric-card'><p style='margin:0;color:#7f8c8d;'>📈 ลงทุน</p><h3 style='margin:0;color:#9b59b6;'>฿{inv:,.0f}</h3></div>", unsafe_allow_html=True)
        
        st.write("")
        expense_df = df_filtered[df_filtered['ประเภท'] == 'รายจ่าย']
        
        col_chart1, col_chart2 = st.columns(2)
        
        with col_chart1:
            st.markdown("#### 🍕 สัดส่วนการใช้เงิน (Donut Chart)")
            if not expense_df.empty:
                pie_data = expense_df.groupby('หมวดหมู่')['จำนวนเงิน'].sum().reset_index()
                fig_pie = px.pie(pie_data, values='จำนวนเงิน', names='หมวดหมู่', hole=0.4,
                                 color_discrete_sequence=px.colors.qualitative.Pastel)
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
                    trend_data.rename(columns={'วันที่': 'วันที่ (ของเดือน)'}, inplace=True)
                    fig_line = px.line(trend_data, x='วันที่ (ของเดือน)', y='จำนวนเงิน', markers=True,
                                       line_shape='spline', color_discrete_sequence=['#e74c3c'])
                else:
                    trend_data = expense_df.groupby('ชื่อเดือน')['จำนวนเงิน'].sum().reset_index()
                    months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
                    trend_data['ชื่อเดือน'] = pd.Categorical(trend_data['ชื่อเดือน'], categories=months, ordered=True)
                    trend_data = trend_data.sort_values('ชื่อเดือน')
                    
                    fig_line = px.bar(trend_data, x='ชื่อเดือน', y='จำนวนเงิน', text='จำนวนเงิน',
                                      color_discrete_sequence=['#e74c3c'])
                    fig_line.update_traces(texttemplate='฿%{text:,.0f}', textposition='outside')
                
                fig_line.update_layout(margin=dict(t=0, b=0, l=0, r=0), xaxis_title="", yaxis_title="จำนวนเงิน (บาท)")
                st.plotly_chart(fig_line, use_container_width=True)
            else:
                st.info("ไม่มีข้อมูลเพียงพอสร้างกราฟ")
                
        st.markdown("---")
        
        st.markdown("#### 🏆 Top 3 หมวดหมู่ที่จ่ายหนักที่สุด")
        if not expense_df.empty:
            top_expenses = expense_df.groupby('หมวดหมู่')['จำนวนเงิน'].sum().sort_values(ascending=False).head(3)
            medals = ["🥇", "🥈", "🥉"]
            cols = st.columns(3)
            for i, (cat, amt) in enumerate(top_expenses.items()):
                cols[i].markdown(f"<div style='text-align:center; padding:10px; background-color:#fff3cd; border-radius:10px;'><h4>{medals[i]} {cat}</h4><h3 style='color:#d35400;'>฿{amt:,.0f}</h3></div>", unsafe_allow_html=True)
    else:
        st.info("ยังไม่มีข้อมูล กรุณาบันทึกข้อมูลในแท็บ 'บันทึกเงิน'")

# --- Tab 3: หน้าติดตามเป้าหมาย ---
with tab3:
    st.subheader("🎯 เป้าหมายการเงิน")
    df = st.session_state.finance_data.copy()
    total_study_savings = df[(df['ประเภท'] == 'เงินออม') & (df['หมวดหมู่'] == 'ออมเพื่อเรียนต่อ/อนาคต')]['จำนวนเงิน'].sum() if not df.empty else 0
    GOAL_STUDY = 100000 
    progress_percent = min(total_study_savings / GOAL_STUDY, 1.0)
    st.write("✈️ **กองทุนเพื่ออนาคต (เรียนต่อ/สอบ)**")
    st.progress(progress_percent)
    st.caption(f"สะสมแล้ว ฿{total_study_savings:,.2f} จากเป้าหมาย ฿{GOAL_STUDY:,.2f} ({progress_percent*100:.1f}%)")

# --- Tab 4: ตั้งค่าและจัดการ ---
with tab4:
    st.subheader("⚡ จัดการปุ่มบันทึกด่วน")
    edited_qa = st.data_editor(st.session_state.quick_adds, use_container_width=True, num_rows="dynamic", key="editor_qa")
    if st.button("💾 บันทึกปุ่มด่วน", use_container_width=True):
        st.session_state.quick_adds = edited_qa
        save_quick_adds(edited_qa)
        st.success("อัปเดตปุ่มสำเร็จ!")
        st.rerun()
    st.markdown("---")
    st.subheader("✏️ แก้ไข/ลบ ประวัติทั้งหมด")
    if not st.session_state.finance_data.empty:
        edited_df = st.data_editor(st.session_state.finance_data, use_container_width=True, num_rows="dynamic", key="editor_finance")
        if st.button("💾 บันทึกประวัติ", use_container_width=True):
            st.session_state.finance_data = edited_df
            save_data(edited_df)
            st.success("อัปเดตสำเร็จ!")
            st.rerun()