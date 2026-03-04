import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime
import re

# 1. Page Configuration
st.set_page_config(page_title="CapEx 2026 Executive Dashboard", layout="wide")

# CSS เพื่อเลียนแบบดีไซน์จากภาพสกรีนแคปจริง
st.markdown("""
    <style>
    header {visibility: hidden;}
    .block-container { padding-top: 1rem; background-color: #f8fafd; }
    
    /* ส่วนหัวชื่อ Dashboard */
    .main-title {
        font-size: 2.2rem;
        font-weight: 800;
        color: #2c3e50;
        margin-bottom: 5px;
        display: flex;
        align-items: center;
    }

    /* กล่องวันที่: ขนาดใหญ่ แยกบรรทัด อยู่มุมขวา */
    .date-box {
        background-color: #ffffff;
        padding: 10px 20px;
        border-radius: 10px;
        border: 1px solid #e0e6ed;
        text-align: center;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
    }
    .date-label { font-size: 0.8rem; color: #7f8c8d; font-weight: bold; }
    .date-value { 
        font-size: 2.2rem; 
        color: #1e3a8a; 
        font-weight: 800; 
        display: block;
        line-height: 1.1;
    }

    /* Progress Card แบบในรูปภาพล่าสุด (ขาว-น้ำเงิน) */
    .metric-card {
        background-color: #ffffff;
        padding: 20px;
        border-radius: 12px;
        text-align: center;
        border: 1px solid #eef2f6;
        box-shadow: 0 4px 6px rgba(0,0,0,0.02);
        height: 100%;
    }
    .metric-label { font-size: 1rem; font-weight: 700; color: #64748b; text-transform: uppercase; }
    .metric-value { font-size: 2.8rem; font-weight: 850; color: #1e40af; margin: 10px 0; }
    .metric-budget { font-size: 0.9rem; color: #94a3b8; margin-bottom: 15px; }
    
    /* Progress Bar ใต้ตัวเลขแบบในรูป */
    .bar-bg { background-color: #f1f5f9; border-radius: 10px; height: 8px; width: 80%; margin: 0 auto; overflow: hidden; }
    .bar-fill { background: linear-gradient(90deg, #3b82f6, #60a5fa); height: 100%; border-radius: 10px; }
    </style>
    """, unsafe_allow_html=True)

# 2. Data Loading
@st.cache_data(ttl=60)
def load_data():
    try:
        df = pd.read_excel("CapEx_2026.xlsx")
        df.columns = df.columns.str.strip()
        column_map = {
            'โครงการ': 'Project', 'งบประมาณ': 'Budget', 'โรงงาน': 'Factory',
            'ประเภท': 'Category', '%ความคืบหน้า': 'Progress'
        }
        df = df.rename(columns=column_map)
        df['Factory'] = df['Factory'].astype(str).str.strip()
        df['Budget'] = pd.to_numeric(df['Budget'], errors='coerce').fillna(0)
        df['Progress'] = pd.to_numeric(df['Progress'], errors='coerce').fillna(0)
        return df
    except Exception as e:
        st.error(f"Error: {e}")
        return None

df = load_data()

if df is not None:
    # --- Top Row: Title & Date ---
    t_col1, t_col2 = st.columns([3, 1])
    with t_col1:
        st.markdown('<div class="main-title">🏛️ CapEx 2026 Investment Progress</div>', unsafe_allow_html=True)
    with t_col2:
        current_date = datetime.now().strftime("%d %B %Y").upper()
        st.markdown(f"""
            <div class="date-box">
                <div class="date-label">DATA REFRESHED AS OF</div>
                <div class="date-value">{current_date}</div>
            </div>
        """, unsafe_allow_html=True)

    st.write("") # Spacer

    # --- Section 1: Progress Cards (แบบภาพล่าสุด) ---
    def get_factory_stats(name_regex):
        subset = df[df['Factory'].str.contains(name_regex, case=False, na=False)]
        if not subset.empty:
            return subset['Progress'].mean(), subset['Budget'].sum()
        return 0, 0

    stats = [
        ("OVERALL", df['Progress'].mean(), df['Budget'].sum()),
        ("PK", *get_factory_stats("PK")),
        ("DC", *get_factory_stats("DC")),
        ("KS/KN", *get_factory_stats("KS|KN")),
        ("MCE", *get_factory_stats("MCE"))
    ]

    cols = st.columns(5)
    for i, (label, prog, bud) in enumerate(stats):
        with cols[i]:
            # ปรับงบประมาณให้แสดงเป็นหน่วยล้าน (M) เพื่อความสะอาดตาแบบในรูป
            bud_display = f"฿{bud/1e6:.1f}M" if bud >= 1e6 else f"฿{bud:,.0f}"
            st.markdown(f"""
                <div class="metric-card">
                    <div class="metric-label">{label}</div>
                    <div class="metric-value">{prog:.2f}%</div>
                    <div class="metric-budget">{bud_display}</div>
                    <div class="bar-bg"><div class="bar-fill" style="width:{prog}%"></div></div>
                </div>
            """, unsafe_allow_html=True)

    # --- Section 2: Visual Analytics ---
    st.markdown("<br>", unsafe_allow_html=True)
    c1, c2 = st.columns([1, 1.2])

    with c1:
        st.subheader("📊 Budget Allocation by Factory")
        fig_pie = px.pie(df, values='Budget', names='Factory', hole=0.5,
                         color_discrete_sequence=px.colors.qualitative.Pastel)
        fig_pie.update_layout(margin=dict(t=30, b=10, l=10, r=10))
        st.plotly_chart(fig_pie, use_container_width=True)

    with c2:
        st.subheader("🎯 Progress vs. Budget Matrix")
        fig_scatter = px.scatter(df, x="Budget", y="Progress", size="Budget", color="Factory",
                                 hover_name="Project", template="plotly_white")
        fig_scatter.update_traces(marker=dict(line=dict(width=1, color='DarkSlateGrey')))
        st.plotly_chart(fig_scatter, use_container_width=True)

    # --- Section 3: Horizontal Bar Chart ---
    st.subheader("📈 Project-Specific Progress Tracking")
    df_sorted = df.sort_values("Progress", ascending=True)
    fig_bar = px.bar(df_sorted, x="Progress", y="Project", orientation='h',
                     color="Progress", color_continuous_scale="Blues",
                     text_auto='.2f') # ทศนิยม 2 ตำแหน่ง
    fig_bar.update_layout(height=max(400, len(df)*25), plot_bgcolor="white")
    st.plotly_chart(fig_bar, use_container_width=True)

    # --- Section 4: Table ---
    with st.expander("📋 View Full Project Inventory"):
        st.dataframe(df[['Factory', 'Project', 'Budget', 'Progress']].style.format({
            'Progress': '{:.2f}%', 
            'Budget': '{:,.0f}'
        }), use_container_width=True)

else:
    st.error("ไม่สามารถโหลดข้อมูลจาก CapEx_2026.xlsx ได้")