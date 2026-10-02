import streamlit as st
import pandas as pd
import io
import os
import glob
import re
from cpr_scanner import CPRScanner

# Page Configuration
st.set_page_config(
    page_title="CPR Option Scanner | Central Pivot Range",
    page_icon="🎯",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom Styling
st.markdown("""
<style>
    /* Theme overrides & custom layout */
    .stApp {
        background-color: #0e1117;
        color: #e0e6ed;
    }
    .header-banner {
        background: linear-gradient(135deg, #102a43 0%, #1f4e78 50%, #002244 100%);
        padding: 1.8rem 2rem;
        border-radius: 12px;
        box-shadow: 0 4px 20px rgba(0,0,0,0.4);
        margin-bottom: 2rem;
        border: 1px solid #2b4c7e;
    }
    .header-title {
        color: #ffffff;
        font-size: 2.2rem;
        font-weight: 800;
        margin: 0;
        letter-spacing: -0.5px;
    }
    .header-subtitle {
        color: #9fb3c8;
        font-size: 1.05rem;
        margin-top: 0.4rem;
        margin-bottom: 0;
    }
    .kpi-card {
        background: #161b22;
        border: 1px solid #30363d;
        border-radius: 10px;
        padding: 1.2rem;
        text-align: center;
        box-shadow: 0 2px 8px rgba(0,0,0,0.2);
        transition: transform 0.2s ease, border-color 0.2s ease;
    }
    .kpi-card:hover {
        transform: translateY(-2px);
        border-color: #388bfd;
    }
    .kpi-value {
        font-size: 1.8rem;
        font-weight: 800;
        color: #58a6ff;
        margin-bottom: 0.2rem;
    }
    .kpi-label {
        font-size: 0.85rem;
        font-weight: 600;
        color: #8b949e;
        text-transform: uppercase;
        letter-spacing: 0.5px;
    }
    .stButton>button {
        background: linear-gradient(135deg, #1f4e78 0%, #2b6cb0 100%);
        color: white;
        font-weight: 700;
        font-size: 1.05rem;
        height: 3.2rem;
        border: none;
        border-radius: 8px;
        box-shadow: 0 4px 12px rgba(31, 78, 120, 0.4);
        transition: all 0.2s ease;
    }
    .stButton>button:hover {
        background: linear-gradient(135deg, #2b6cb0 0%, #3182ce 100%);
        box-shadow: 0 6px 16px rgba(49, 130, 206, 0.5);
        color: white;
    }
    .download-btn button {
        background: linear-gradient(135deg, #238636 0%, #2ea043 100%) !important;
        box-shadow: 0 4px 12px rgba(35, 134, 54, 0.4) !important;
    }
    .tag-ce {
        background-color: #13231b;
        color: #3fb950;
        padding: 2px 8px;
        border-radius: 4px;
        font-weight: 600;
        border: 1px solid #238636;
    }
    .tag-pe {
        background-color: #331518;
        color: #f85149;
        padding: 2px 8px;
        border-radius: 4px;
        font-weight: 600;
        border: 1px solid #da3633;
    }
</style>
""", unsafe_allow_html=True)

# Helper function to generate Excel in memory
def generate_cpr_excel_bytes(df, top_n=5):
    output = io.BytesIO()
    cpr_scanner = CPRScanner()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        cpr_scanner.generate_excel_writer(df, writer, top_n=top_n)
    return output.getvalue()

# Header Banner
st.markdown("""
<div class="header-banner">
    <div style="display: flex; align-items: center; justify-content: space-between;">
        <div>
            <h1 class="header-title">🎯 CPR Option Scanner</h1>
            <p class="header-subtitle">Central Pivot Range EOD Analysis & Breakout Scanner for Stock & Index Options</p>
        </div>
        <div style="text-align: right;">
            <span style="background: rgba(56, 139, 253, 0.15); border: 1px solid #388bfd; color: #58a6ff; padding: 6px 14px; border-radius: 20px; font-weight: 600; font-size: 0.9rem;">
                ⚡ Dedicated CPR App
            </span>
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

# Sidebar
st.sidebar.title("⚙️ CPR Scanner Settings")
st.sidebar.markdown("---")

# Look for default local files in current or parent directory
local_zips = sorted(glob.glob("*.zip") + glob.glob("../*.zip"))
today_default = None
yest_default = None
if len(local_zips) >= 2:
    today_default = local_zips[-1]
    yest_default = local_zips[-2]

st.sidebar.subheader("📁 Data Source Selection")
use_local = False
if today_default and yest_default:
    use_local = st.sidebar.checkbox(f"Use auto-detected local Zip files ({os.path.basename(yest_default)} & {os.path.basename(today_default)})", value=True)

today_file = None
yest_file = None

if use_local:
    today_file = today_default
    yest_file = yest_default
    st.sidebar.info(f"✅ Using local BhavCopies:\n- **Today**: `{os.path.basename(today_default)}`\n- **Yesterday**: `{os.path.basename(yest_default)}`")
else:
    today_uploaded = st.sidebar.file_uploader("Upload Today's BhavCopy (.zip)", type=["zip"], key="today_zip")
    yest_uploaded = st.sidebar.file_uploader("Upload Yesterday's BhavCopy (.zip)", type=["zip"], key="yest_zip")
    if today_uploaded and yest_uploaded:
        today_file = today_uploaded
        yest_file = yest_uploaded

top_n_val = st.sidebar.slider("Top N Leaderboard Size", min_value=3, max_value=25, value=5, step=1)

st.sidebar.markdown("---")
st.sidebar.markdown("""
### 💡 CPR Setup Quick Guide
- **Narrow CPR**: Contraction phase $\\rightarrow$ High probability breakout setup.
- **Inside CPR**: CPR inside yesterday's CPR range $\\rightarrow$ Strong momentum expected.
- **Higher Value CPR**: Bullish continuation bias.
- **Lower Value CPR**: Bearish continuation bias.
""")

# Run Scanner Action
if today_file and yest_file:
    if st.sidebar.button("🚀 RUN CPR SCANNER", use_container_width=True) or 'cpr_df' in st.session_state:
        if 'cpr_df' not in st.session_state or st.sidebar.button("🔄 Re-Run Scan"):
            with st.spinner("Processing NSE Option BhavCopy & Calculating Central Pivot Ranges..."):
                scanner = CPRScanner()
                df_results = scanner.process_data(today_file, yest_file)
                st.session_state['cpr_df'] = df_results

        df = st.session_state.get('cpr_df')

        if df is None or df.empty:
            st.error("Failed to process data or no valid option data found. Please check uploaded files.")
        else:
            # Summary Metrics Row
            total_symbols = df['Symbol'].nunique()
            total_records = len(df)
            narrow_count = df[df['Is_Narrow_CPR'] == True]['Symbol'].nunique()
            wide_count = df[df['Is_Wide_CPR'] == True]['Symbol'].nunique()
            inside_count = df[df['Is_Inside_CPR'] == True]['Symbol'].nunique()
            outside_count = df[df['Is_Outside_CPR'] == True]['Symbol'].nunique()
            hv_count = df[df['Is_Higher_Value_CPR'] == True]['Symbol'].nunique()
            lv_count = df[df['Is_Lower_Value_CPR'] == True]['Symbol'].nunique()

            m1, m2, m3, m4, m5, m6, m7 = st.columns(7)
            with m1:
                st.markdown(f'<div class="kpi-card"><div class="kpi-value">{total_symbols}</div><div class="kpi-label">Symbols</div></div>', unsafe_allow_html=True)
            with m2:
                st.markdown(f'<div class="kpi-card"><div class="kpi-value">{narrow_count}</div><div class="kpi-label">Narrow CPR</div></div>', unsafe_allow_html=True)
            with m3:
                st.markdown(f'<div class="kpi-card"><div class="kpi-value">{inside_count}</div><div class="kpi-label">Inside CPR</div></div>', unsafe_allow_html=True)
            with m4:
                st.markdown(f'<div class="kpi-card"><div class="kpi-value">{hv_count}</div><div class="kpi-label">Higher Value</div></div>', unsafe_allow_html=True)
            with m5:
                st.markdown(f'<div class="kpi-card"><div class="kpi-value">{lv_count}</div><div class="kpi-label">Lower Value</div></div>', unsafe_allow_html=True)
            with m6:
                st.markdown(f'<div class="kpi-card"><div class="kpi-value">{wide_count}</div><div class="kpi-label">Wide CPR</div></div>', unsafe_allow_html=True)
            with m7:
                st.markdown(f'<div class="kpi-card"><div class="kpi-value">{outside_count}</div><div class="kpi-label">Outside CPR</div></div>', unsafe_allow_html=True)

            st.markdown("<br>", unsafe_allow_html=True)

            # Excel Download Button Header Row
            d_col1, d_col2 = st.columns([3, 1])
            with d_col1:
                st.subheader("📋 Scan Results & Technical Categorization")
            with d_col2:
                excel_data = generate_cpr_excel_bytes(df, top_n=top_n_val)
                st.download_button(
                    label="📥 Download CPR Excel Report",
                    data=excel_data,
                    file_name=f"CPR_Scanner_Report.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

            # Main Application Tabs
            tab_dash, tab_main, tab_patterns, tab_leaders, tab_guide = st.tabs([
                "🎯 Dashboard & Filters",
                "📊 Main CPR Data",
                "⚡ CPR Pattern Categories",
                "🔥 Market Leaders (Top N)",
                "📖 CPR Formula Guide"
            ])

            # Tab 1: Dashboard & Filters
            with tab_dash:
                st.markdown("##### 🔍 Quick Search & Dynamic Filter")
                fc1, fc2, fc3 = st.columns(3)
                with fc1:
                    symbol_search = st.text_input("Search Symbol (e.g. NIFTY, RELIANCE)", "").upper()
                with fc2:
                    opt_filter = st.selectbox("Option Type", ["All", "CE", "PE"])
                with fc3:
                    pattern_filter = st.selectbox("CPR Pattern Filter", [
                        "All Patterns",
                        "Narrow CPR",
                        "Wide CPR",
                        "Inside CPR",
                        "Outside CPR",
                        "Higher Value CPR",
                        "Lower Value CPR"
                    ])

                filtered_df = df.copy()
                if symbol_search:
                    filtered_df = filtered_df[filtered_df['Symbol'].str.contains(symbol_search, na=False)]
                if opt_filter != "All":
                    filtered_df = filtered_df[filtered_df['Option_Type'] == opt_filter]
                
                if pattern_filter == "Narrow CPR":
                    filtered_df = filtered_df[filtered_df['Is_Narrow_CPR'] == True]
                elif pattern_filter == "Wide CPR":
                    filtered_df = filtered_df[filtered_df['Is_Wide_CPR'] == True]
                elif pattern_filter == "Inside CPR":
                    filtered_df = filtered_df[filtered_df['Is_Inside_CPR'] == True]
                elif pattern_filter == "Outside CPR":
                    filtered_df = filtered_df[filtered_df['Is_Outside_CPR'] == True]
                elif pattern_filter == "Higher Value CPR":
                    filtered_df = filtered_df[filtered_df['Is_Higher_Value_CPR'] == True]
                elif pattern_filter == "Lower Value CPR":
                    filtered_df = filtered_df[filtered_df['Is_Lower_Value_CPR'] == True]

                display_cols = ['Symbol', 'Expiry', 'Spot_Close', 'ATM_Strike', 'Option_Type', 
                                'Today_P', 'Today_TC', 'Today_BC', 'Today_CPR_Width', 'Today_CPR_Width_Pct',
                                'Is_Narrow_CPR', 'Is_Inside_CPR', 'Is_Higher_Value_CPR', 'Is_Lower_Value_CPR']
                
                st.dataframe(filtered_df[display_cols], use_container_width=True, height=450)

            # Tab 2: Main CPR Data
            with tab_main:
                st.markdown("##### 📄 Complete CPR Data Records")
                st.dataframe(df, use_container_width=True, height=500)

            # Tab 3: CPR Pattern Categories
            with tab_patterns:
                cat_tabs = st.tabs([
                    "Narrow CPR",
                    "Wide CPR",
                    "Inside CPR",
                    "Outside CPR",
                    "Higher Value CPR",
                    "Lower Value CPR"
                ])

                def render_side_by_side(df_source, filter_col, title):
                    sub_df = df_source[df_source[filter_col] == True].copy()
                    if sub_df.empty:
                        st.info(f"No records matching {title}.")
                        return
                    
                    show_cols = ['Symbol', 'Spot_Close', 'ATM_Strike', 'Today_P', 'Today_TC', 'Today_BC', 'Today_CPR_Width']
                    ce_df = sub_df[sub_df['Option_Type'] == 'CE'][show_cols].reset_index(drop=True)
                    pe_df = sub_df[sub_df['Option_Type'] == 'PE'][show_cols].reset_index(drop=True)

                    c1, c2 = st.columns(2)
                    with c1:
                        st.markdown(f"###### 🟢 {title} - CE Options ({len(ce_df)})")
                        st.dataframe(ce_df, use_container_width=True, height=400)
                    with c2:
                        st.markdown(f"###### 🔴 {title} - PE Options ({len(pe_df)})")
                        st.dataframe(pe_df, use_container_width=True, height=400)

                with cat_tabs[0]:
                    render_side_by_side(df, 'Is_Narrow_CPR', 'Narrow CPR')
                with cat_tabs[1]:
                    render_side_by_side(df, 'Is_Wide_CPR', 'Wide CPR')
                with cat_tabs[2]:
                    render_side_by_side(df, 'Is_Inside_CPR', 'Inside CPR')
                with cat_tabs[3]:
                    render_side_by_side(df, 'Is_Outside_CPR', 'Outside CPR')
                with cat_tabs[4]:
                    render_side_by_side(df, 'Is_Higher_Value_CPR', 'Higher Value CPR')
                with cat_tabs[5]:
                    render_side_by_side(df, 'Is_Lower_Value_CPR', 'Lower Value CPR')

            # Tab 4: Market Leaders (Top N)
            with tab_leaders:
                st.markdown(f"##### 🔥 Top {top_n_val} Market Activity Leaders")
                metrics = [
                    ('OpnIntrst', 'Open Interest (OI) Leaders'),
                    ('ChngInOpnIntrst', 'Change in Open Interest Leaders'),
                    ('TtlTradgVol', 'Trading Volume Leaders'),
                    ('TtlNbOfTxsExctd', 'Transaction Count Leaders')
                ]

                col_a, col_b = st.columns(2)
                for i, (m_col, m_title) in enumerate(metrics):
                    df[m_col] = pd.to_numeric(df[m_col], errors='coerce').fillna(0)
                    top_df = df.sort_values(by=m_col, ascending=False).head(top_n_val).copy()
                    display_top = top_df[['Symbol', 'Option_Type', 'ATM_Strike', 'Spot_Close', m_col]].reset_index(drop=True)

                    target_col = col_a if i % 2 == 0 else col_b
                    with target_col:
                        st.markdown(f"###### 🏆 Top {top_n_val} {m_title}")
                        st.dataframe(display_top, use_container_width=True)

            # Tab 5: CPR Formula Guide
            with tab_guide:
                st.markdown("""
                ### 📖 CPR (Central Pivot Range) Calculation & Setup Guide
                
                #### 1. Core CPR Formulas
                - **Pivot Point (P)** = $\\frac{\\text{High} + \\text{Low} + \\text{Close}}{3}$
                - **Bottom Central Pivot (BC)** = $\\frac{\\text{High} + \\text{Low}}{2}$
                - **Top Central Pivot (TC)** = $(2 \\times P) - BC$
                - **Top CPR** = $\\max(TC, BC)$
                - **Bottom CPR** = $\\min(TC, BC)$
                - **CPR Width** = $\\text{Top CPR} - \\text{Bottom CPR}$
                - **CPR Width %** = $\\frac{\\text{CPR Width}}{P} \\times 100$

                #### 2. Standard Pivot Support & Resistance
                - **R1** = $(2 \\times P) - \\text{Low}$ | **S1** = $(2 \\times P) - \\text{High}$
                - **R2** = $P + (\\text{High} - \\text{Low})$ | **S2** = $P - (\\text{High} - \\text{Low})$
                - **R3** = $\\text{High} + 2 \\times (P - \\text{Low})$ | **S3** = $\\text{Low} - 2 \\times (\\text{High} - P)$

                #### 3. Key Trading Setups
                - **Narrow CPR**: Indicates low volatility on previous day $\\rightarrow$ High chance of strong directional trend breakout today.
                - **Inside CPR**: Today's CPR is fully inside Yesterday's CPR range $\\rightarrow$ Major breakout candidate.
                - **Higher Value CPR**: Today's Bottom CPR > Yesterday's Top CPR $\\rightarrow$ Strong Bullish Bias.
                - **Lower Value CPR**: Today's Top CPR < Yesterday's Bottom CPR $\\rightarrow$ Strong Bearish Bias.
                - **Wide CPR**: Indicates high volatility on previous day $\\rightarrow$ Sideways / Range-bound movement expected.
                """)
else:
    st.info("👈 Please select local ZIP files or upload Today & Yesterday NSE Option BhavCopies in the sidebar to start scanning.")
