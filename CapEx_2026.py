import streamlit as st
import pandas as pd
import plotly.express as px
import re
from datetime import datetime

# 1. Page Configuration
st.set_page_config(page_title="CapEx 2026 Executive Dashboard", layout="wide")

# CSS เพื่อซ่อน Header (GitHub icon, Menu) และปรับแต่ง UI
st.markdown("""
    <style>
    header {visibility: hidden;}
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    .block-container { padding-top: 2rem; }
    .stMetric { 
        background-color: #ffffff; 
        padding: 15px; 
        border-radius: 10px; 
        box-shadow: 0 2px 4px rgba(0,0,0,0.05); 
        border-left: 5px solid #007bff;
    }
    .update-box {
        background-color: #f8f9fa;
        padding: 10px 20px;
        border-radius: 5px;
        border-left: 4px solid #0056b3;
        margin-bottom: 25px;
    }
    .update-title { color: #6c757d; font-size: 0.8rem; font-weight: bold; margin-bottom: 2px; }
    .update-date { color: #0056b3; font-size: 1.2rem; font-weight: bold; }
    </style>
    """, unsafe_allow_html=True)

# 2. Data Loading & Mapping
@st.cache_data
def load_data():
    try:
        df = pd.read_excel("CapEx_2026.xlsx")
        df.columns = df.columns.str.strip()
        column_map = {
            'โครงการ': 'Project', 'งบประมาณ': 'Budget', 'โรงงาน': 'Factory',
            'ประเภท': 'Category', 'วัตถุประสงค์': 'Objective',
            'ประโยชน์ที่ได้รับ': 'Benefits', '%IRR/NPV/PB': 'Financial_Data'
        }
        df = df.rename(columns=column_map)
        
        # คลีนข้อมูล: ตัดช่องว่างหน้า-หลังชื่อโรงงานและประเภท เพื่อป้องกันชื่อซ้ำ
        df['Factory'] = df['Factory'].astype(str).str.strip()
        df['Category'] = df['Category'].astype(str).str.strip()
        
        df['Budget'] = pd.to_numeric(df['Budget'], errors='coerce').fillna(0)
        
        def extract_irr(val):
            if pd.isna(val) or val == '-': return None
            match = re.search(r"[-+]?\d*\.\d+|\d+", str(val))
            return float(match.group()) if match else None
        
        df['IRR_Value'] = df['Financial_Data'].apply(extract_irr)
        df['IRR_Value'] = pd.to_numeric(df['IRR_Value'], errors='coerce') 
        return df
    except Exception as e:
        st.error(f"Error: {e}")
        return None

df = load_data()

if df is not None:
    # --- Top Section ---
    st.title("🏛️ CapEx 2026 Investment Executive Dashboard")
    
    current_date = datetime.now().strftime("%d %B %Y").upper()
    st.markdown(f"""
        <div class="update-box">
            <div class="update-title">LATEST DATA UPDATE AS OF</div>
            <div class="update-date">{current_date}</div>
        </div>
    """, unsafe_allow_html=True)

    # --- Dashboard Controls (ติ๊กถูกใน Expander) ---
    with st.expander("⚙️ Dashboard Controls", expanded=True):
        c1, c2 = st.columns(2)
        
        with c1:
            st.write("**Select Factory**")
            # ดึง Unique Factory ที่คลีนแล้วมาแสดง
            factories = sorted([f for f in df["Factory"].unique() if f != 'nan'])
            selected_factories = [f for f in factories if st.checkbox(f, value=True, key=f"f_{f}")]
            
        with c2:
            st.write("**Select Category**")
            categories = sorted([c for c in df["Category"].unique() if c != 'nan'])
            selected_categories = [c for c in categories if st.checkbox(c, value=True, key=f"c_{c}")]

    df_filtered = df[(df["Factory"].isin(selected_factories)) & (df["Category"].isin(selected_categories))]

    if not df_filtered.empty:
        # --- Metrics ---
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Total Investment", f"฿{df_filtered['Budget'].sum():,.0f}")
        m2.metric("Project Count", f"{len(df_filtered)}")
        avg_irr = df_filtered['IRR_Value'].mean()
        m3.metric("Avg. IRR (%)", f"{avg_irr:.2f}%" if pd.notnull(avg_irr) else "N/A")
        
        top_spender_group = df_filtered.groupby("Factory")["Budget"].sum()
        if not top_spender_group.empty:
            m4.metric("Top Spender", top_spender_group.idxmax())

        st.markdown("---")
        
        # --- Charts ---
        col1, col2 = st.columns([2, 1])
        with col1:
            st.write("**Budget by Project**")
            df_bar = df_filtered.sort_values("Budget", ascending=False)
            fig_bar = px.bar(df_bar,
                             x="Project", y="Budget", color="Category",
                             hover_name="Project",
                             hover_data={
                                 "Project": False,
                                 "Budget": ":,.2f",
                                 "Category": True,
                                 "Factory": True,
                                 "Objective": True,
                                 "Benefits": True
                             },
                             template="plotly_white")
            
            fig_bar.update_layout(
                xaxis=dict(
                    title=dict(text="Project Name", standoff=35),
                    tickmode='array',
                    tickvals=df_bar["Project"],
                    ticktext=[str(p)[:15] + ".." if len(str(p)) > 15 else p for p in df_bar["Project"]],
                    tickangle=-45,
                    automargin=True
                ),
                margin=dict(b=120, r=20),
                showlegend=True,
                legend=dict(orientation="v", yanchor="top", y=1, xanchor="left", x=1.02)
            )
            st.plotly_chart(fig_bar, use_container_width=True)

        with col2:
            st.write("**Budget Distribution**")
            fig_pie = px.pie(df_filtered, values='Budget', names='Factory', hole=0.4)
            st.plotly_chart(fig_pie, use_container_width=True)

        col3, col4 = st.columns(2)
        with col3:
            st.write("**Budget vs IRR Matrix**")
            df_scatter = df_filtered.dropna(subset=['IRR_Value'])
            if not df_scatter.empty:
                fig_scatter = px.scatter(df_scatter, 
                                         x="Budget", y="IRR_Value",
                                         size="Budget", color="Factory", hover_name="Project",
                                         hover_data={"Budget": ":,.2f", "IRR_Value": ":.2f", "Objective": True, "Benefits": True, "Factory": True})
                st.plotly_chart(fig_scatter, use_container_width=True)

        with col4:
            st.write("**Budget by Category**")
            df_cat_summary = df_filtered.groupby("Category")["Budget"].sum().reset_index()
            fig_cat = px.bar(df_cat_summary, y="Category", x="Budget", orientation='h', color="Category")
            fig_cat.update_layout(
                showlegend=True,
                yaxis=dict(showticklabels=False, title=None, automargin=True),
                xaxis_title="Budget (THB)",
                margin=dict(l=20) 
            )
            st.plotly_chart(fig_cat, use_container_width=True)

        with st.expander("📋 Click to view/hide Project Inventory Table", expanded=True):
            st.dataframe(df_filtered.drop(columns=['IRR_Value']).fillna('-'), use_container_width=True, hide_index=True)

    else:
        st.warning("No data matches selected filters.")