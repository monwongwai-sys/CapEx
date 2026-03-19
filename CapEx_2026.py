import streamlit as st
import pandas as pd
import plotly.express as px
from datetime import datetime
import re

# 1. Page Configuration
st.set_page_config(page_title="CapEx 2026 Executive Dashboard", layout="wide")

# CSS สำหรับดีไซน์และวันที่ขนาดใหญ่
st.markdown("""
    <style>
    header {visibility: hidden;}
    .block-container { padding-top: 1rem; background-color: #f8fafd; }
    
    .main-title {
        font-size: 2.2rem;
        font-weight: 800;
        color: #2c3e50;
        margin-bottom: 5px;
    }

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
        color: #1e40af; 
        font-weight: 800; 
        display: block;
        line-height: 1.1;
    }

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
    # --- Top Row ---
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

    # --- Section 1: Progress Cards ---
    def get_stats(reg):
        sub = df[df['Factory'].str.contains(reg, case=False, na=False)]
        return (sub['Progress'].mean(), sub['Budget'].sum()) if not sub.empty else (0, 0)

    stats = [
        ("OVERALL", df['Progress'].mean(), df['Budget'].sum()),
        ("PK", *get_stats("PK")), ("DC", *get_stats("DC")),
        ("KS/KN", *get_stats("KS|KN")), ("MCE", *get_stats("MCE"))
    ]

    cols = st.columns(5)
    for i, (label, prog, bud) in enumerate(stats):
        with cols[i]:
            bud_txt = f"฿{bud/1e6:.1f}M" if bud >= 1e6 else f"฿{bud:,.0f}"
            st.markdown(f"""
                <div class="metric-card">
                    <div class="metric-label">{label}</div>
                    <div class="metric-value">{prog:.2f}%</div>
                    <div class="metric-budget">{bud_txt}</div>
                    <div class="bar-bg"><div class="bar-fill" style="width:{prog}%"></div></div>
                </div>
            """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)

    # --- Section 2: Visual Analytics ---
    c1, c2 = st.columns([1, 1.2])
    with c1:
        st.subheader("📊 Budget Allocation by Factory")
        # --- ปรับแต่ง Pie Chart: ตัวอักษรสีขาวขนาดใหญ่ ---
        fig_pie = px.pie(df, values='Budget', names='Factory', hole=0.4,
                         color_discrete_sequence=px.colors.qualitative.Bold)
        
        fig_pie.update_traces(
            textposition='inside', 
            textinfo='percent+label',
            insidetextfont=dict(size=18, color="white") # ปรับขนาด 18 และสีขาว
        )
        fig_pie.update_layout(showlegend=False, margin=dict(t=0, b=0, l=0, r=0))
        st.plotly_chart(fig_pie, use_container_width=True)

    with c2:
        st.subheader("🎯 Progress vs. Budget Matrix")
        fig_scatter = px.scatter(df, x="Budget", y="Progress", size="Budget", color="Factory",
                                 hover_name="Project", template="plotly_white")
        st.plotly_chart(fig_scatter, use_container_width=True)

    # --- Section 3: กราฟแท่งที่เลือกโรงงานได้ ---
    st.markdown("---")
    st.subheader("📈 Project-Specific Progress Tracking")
    
    factory_list = ["Show All"] + sorted(df['Factory'].unique().tolist())
    selected_fac = st.selectbox("🎯 Filter Projects by Factory:", factory_list)

    bar_df = df if selected_fac == "Show All" else df[df['Factory'] == selected_fac]

    if not bar_df.empty:
        fig_bar = px.bar(bar_df.sort_values("Progress", ascending=True), 
                         x="Progress", y="Project", orientation='h',
                         color="Progress", color_continuous_scale="Blues",
                         text_auto='.2f')
        # ปรับสีตัวอักษรบนกราฟแท่งให้เป็นสีขาวด้วยถ้าต้องการความเข้ากัน
        fig_bar.update_traces(textfont=dict(color="white", size=12))
        fig_bar.update_layout(height=max(400, len(bar_df)*30), plot_bgcolor="white")
        st.plotly_chart(fig_bar, use_container_width=True)

    with st.expander("📋 View Full Project Inventory"):
        st.dataframe(bar_df[['Factory', 'Project', 'Budget', 'Progress']].style.format({
            'Progress': '{:.2f}%', 'Budget': '{:,.0f}'
        }), use_container_width=True)

else:
    st.error("Cannot load data from CapEx_2026.xlsx")