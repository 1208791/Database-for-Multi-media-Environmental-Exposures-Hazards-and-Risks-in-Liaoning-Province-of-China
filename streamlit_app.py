import streamlit as st
import pandas as pd
import os
st.set_page_config(layout="wide")

# -------------------------- 全局配置 --------------------------
EXCEL_FILE_NAME = "data.xlsx"
EXCEL_FILE_PATH = os.path.join(os.path.dirname(__file__), EXCEL_FILE_NAME)
QUERY_COLUMN = 'CAS'

# -------------------------- 页面样式（全局）- 所有字体≥36px --------------------------
def set_global_styles():
    st.markdown("""
        <style>
        /* =============== 全局基础样式 =============== */
        * {
            font-size: 36px !important;
            line-height: 1.4 !important;
        }
        html, body, [class*="css"] {
            font-size: 36px !important;
        }
        /* Streamlit组件字体强化 */
        .stAlert, .stAlert p, 
        .stSpinner, .stSpinner p,
        .stMarkdown, .stMarkdown p,
        .stDataFrame, .stTable,
        .stDownloadButton, .stButton,
        div[data-testid="stMarkdownContainer"],
        div[data-testid="stText"] {
            font-size: 36px !important;
        }
        
        /* =============== 表格专项 =============== */
        .dataframe, .dataframe th, .dataframe td, 
        table, th, td, 
        .stDataFrame table, 
        .stDataFrame th, 
        .stDataFrame td {
            font-size: 36px !important;
            padding: 12px 15px !important;
            line-height: 1.3 !important;
        }
        /* 表格标题加粗 */
        .dataframe th {
            font-weight: 600 !important;
        }
        
        /* =============== 按钮统一规范 =============== */
        button, .stButton button, .stDownloadButton button {
            font-size: 36px !important;
            height: 80px !important;
            line-height: 80px !important;
            min-width: 180px !important;
            padding: 0 25px !important;
            border-radius: 8px !important;
            font-weight: 600 !important;
        }
        /* 按钮悬停效果 */
        button:hover, .stButton button:hover, .stDownloadButton button:hover {
            background-color: #f0f8ff !important;
            border-color: #1f77b4 !important;
        }
        
        /* =============== 输入框强化 =============== */
        input[type="text"] {
            font-size: 36px !important;
            height: 80px !important;
            line-height: 80px !important;
            padding: 0 25px !important;
        }
        input::placeholder {
            font-size: 36px !important;
            color: #999999 !important;
            opacity: 1 !important;
        }
        
        /* =============== 原有样式调整（确保≥36px） =============== */
        .main {
            background-color: #ffffff;
            font-family: "Segoe UI", Roboto, "Microsoft YaHei", sans-serif;
        }
        .search-container {
            display: flex;
            justify-content: center;
            gap: 0;
            margin-bottom: 80px;
        }
        .search-input {
            width: 800px;
            height: 80px;
            padding: 0 25px !important;
            border: 2px solid #1f77b4;
            border-right: none;
            border-radius: 8px 0 0 8px;
            font-size: 36px !important;
            line-height: 80px !important;
            box-sizing: border-box;
        }
        .search-button {
            height: 80px;
            width: 180px;
            background-color: #ffffff;
            color: #1f77b4;
            border: 2px solid #1f77b4;
            border-radius: 0 8px 8px 0;
            font-size: 36px !important;
            font-weight: 600;
            line-height: 80px !important;
            padding: 0 !important;
            box-sizing: border-box;
        }
        .footer {
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding: 0 20px;
            position: fixed;
            bottom: 20px;
            width: 95%;
            font-size: 36px !important;
        }
        .footer-text {
            font-size: 36px !important;
            font-weight: 500;
        }
        .result-page-title {
            font-size: 60px !important;
            color: #1f77b4;
            text-align: center;
            font-weight: 600;
            margin: 20px 0;
            line-height: 1.2 !important;
        }
        .cas-info-row {
            display: flex;
            justify-content: center;
            align-items: center;
            gap: 40px;
            margin: 30px 0 40px 0;
            flex-wrap: wrap;
            font-size: 36px !important;
        }
        .cas-item {
            font-size: 36px !important;
            color: #1f77b4;
            font-weight: 600;
            padding: 0 15px;
        }
        .section-divider {
            border: none;
            height: 3px;
            background-color: #1f77b4;
            margin: 40px 0;
        }
        .section-title {
            font-size: 36px !important;
            color: #1f77b4;
            font-weight: 600;
            margin: 45px 0 25px 0;
            text-align: left;
            padding-left: 10px;
            border-left: 4px solid #1f77b4;
        }
        .download-btn-container {
            display: flex;
            justify-content: center;
            margin: 60px 0 90px 0;
        }
        
        /* =============== 提示信息强化 =============== */
        .stAlert {
            padding: 25px !important;
            border-radius: 10px !important;
            margin: 20px 0 !important;
        }
        .stAlert p {
            margin: 0 !important;
            padding: 5px 0 !important;
        }
        
        /* =============== 响应式优化（大字体适配） =============== */
        @media (max-width: 1200px) {
            .search-input { width: 650px !important; }
            .cas-info-row { flex-direction: column; gap: 20px; }
            .footer { flex-direction: column; gap: 15px; text-align: center; }
        }
        
        /* =============== Help 按钮样式 =============== */
        .help-button {
            font-size: 36px !important;
            background-color: #ffffff !important;
            color: #1f77b4 !important;
            border: 2px solid #1f77b4 !important;
            border-radius: 8px !important;
            height: 80px !important;
            line-height: 80px !important;
            min-width: 150px !important;
            padding: 0 30px !important;
            box-sizing: border-box !important;
            font-weight: 600 !important;
        }
        .help-button:hover {
            background-color: #f0f8ff !important;
            transform: translateY(-2px) !important;
            box-shadow: 0 4px 12px rgba(31, 119, 180, 0.2) !important;
        }
        
        /* =============== 标题样式 =============== */
        .main-title {
            text-align: center;
            font-size: 60px !important;
            color: #1f77b4;
            font-weight: 600;
            margin: 40px 0 20px 0;
            line-height: 1.2 !important;
        }
        </style>
        <script>
        // 强制应用字体大小（双重保障）
        document.addEventListener('DOMContentLoaded', function() {
            const style = document.createElement('style');
            style.textContent = `
                * { font-size: 36px !important; }
                button, input, select, textarea { font-size: 36px !important; height: auto !important; }
                .stDataFrame div { font-size: 36px !important; }
            `;
            document.head.appendChild(style);
        });
        </script>
    """, unsafe_allow_html=True)

# -------------------------- 数据加载函数 --------------------------
def load_all_data_on_query():
    try:
        with st.spinner("Querying data..."):
            df_sheet1 = pd.read_excel(EXCEL_FILE_PATH, sheet_name=0)
            df_sheet2 = pd.read_excel(EXCEL_FILE_PATH, sheet_name=1)
            
            for df, sheet_name in [(df_sheet1, "第一个"), (df_sheet2, "第二个")]:
                if QUERY_COLUMN not in df.columns:
                    st.error(f"Excel file's {sheet_name} worksheet does not contain '{QUERY_COLUMN}' column!")
                    return None, None
                df[QUERY_COLUMN] = df[QUERY_COLUMN].astype(str).str.strip()
        
        return df_sheet1, df_sheet2
    except FileNotFoundError:
        st.error(f"File not found: {EXCEL_FILE_PATH}. Please confirm the file is in the current directory and the name is correct!")
        return None, None
    except Exception as e:
        st.error(f"Failed to load data: {str(e)}")
        return None, None

# -------------------------- 页面渲染函数 --------------------------
def render_home_page():
    """Render the home page (query input page)"""
    # Database title
    st.markdown('<div class="main-title">Database for Multi-media Environmental Exposures, Hazards, and Risks in Liaoning Province of China</div>', unsafe_allow_html=True)
    
    # Search area
    st.markdown('<div class="search-container">', unsafe_allow_html=True)
    col1, col2 = st.columns([5, 1], gap="small", vertical_alignment="center")
    with col1:
        cas_input = st.text_input(
            "",
            placeholder="Enter CAS number to search",
            label_visibility="collapsed"
        )
        st.markdown("""
            <script>
            const input = document.querySelector('input[placeholder="Enter CAS number to search"]');
            if (input) input.classList.add('search-input');
            </script>
        """, unsafe_allow_html=True)
    with col2:
        search_btn = st.button("Search", use_container_width=True)
        st.markdown("""
            <script>
            const btn = document.querySelector('button[kind="primary"]');
            if (btn) btn.classList.add('search-button');
            </script>
        """, unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

    # Footer area - Using simple button with modal
    col11_1, col11_2, col11_3 = st.columns([1, 5, 6], vertical_alignment="center")
    with col11_1:
        try:
            st.image("1.png", width=150)  # Increased image size
        except:
            st.markdown('<div style="width:150px;height:150px;background:#e9ecef;border-radius:8px"></div>', unsafe_allow_html=True)
    with col11_2:
        st.markdown('<span class="footer-text">Dalian University of Technology</span>', unsafe_allow_html=True)
    with col11_3:
        # Simple help button without dropdown
        if st.button("Help", key="help_btn", use_container_width=True):
            st.markdown("""
                Enter a CAS Number and click "Search"
            """, unsafe_allow_html=True)
            
            # Add a close button
            st.button("Close", key="close_help", use_container_width=True)

    if search_btn:
        if not cas_input.strip():
            st.warning("Please enter a valid CAS number!")
        else:
            cas_input = cas_input.strip()
            df_sheet1, df_sheet2 = load_all_data_on_query()
            
            if df_sheet1 is not None and df_sheet2 is not None:
                result_sheet1 = df_sheet1[df_sheet1[QUERY_COLUMN] == cas_input]
                result_sheet2 = df_sheet2[df_sheet2[QUERY_COLUMN] == cas_input]
                
                st.session_state['cas_input'] = cas_input
                st.session_state['result_sheet1'] = result_sheet1
                st.session_state['result_sheet2'] = result_sheet2
                st.session_state['page'] = 'result'
                st.rerun()

def render_result_page():
    """Render the results page"""
    cas_input = st.session_state.get('cas_input', '')
    result_sheet1 = st.session_state.get('result_sheet1', pd.DataFrame())
    result_sheet2 = st.session_state.get('result_sheet2', pd.DataFrame())

    st.markdown("""
        <div class="main-title">Database for Multi-media Environmental Exposures, Hazards, and Risks in Liaoning Province of China</div>
    """, unsafe_allow_html=True)

    chemical_name = ""
    smiles = ""
    if len(result_sheet1) > 0:
        first_row = result_sheet1.iloc[0]
        chemical_name = first_row.get('Chemical name', '')
        # smiles = first_row.get('SMILES', '')

    # st.markdown(f"""
    #     <div class="cas-info-row">
    #         <span class="cas-item">CAS: {cas_input}</span>
    #         <span class="cas-item">Name: {chemical_name}</span>
    #         <span class="cas-item">SMILES: {smiles}</span>
    #     </div>
    # """, unsafe_allow_html=True)
    st.markdown(f"""
        <div class="cas-info-row">
            <span class="cas-item">CAS: {cas_input}</span>
            <span class="cas-item">Name: {chemical_name}</span>
        </div>
    """, unsafe_allow_html=True)
    # st.markdown('<hr class="section-divider">', unsafe_allow_html=True)

    st.markdown('<h2 class="section-title">Exposure data</h2>', unsafe_allow_html=True)
    if len(result_sheet1) > 0:
        st.dataframe(
            result_sheet1,
            use_container_width=True,
            hide_index=True
        )
    else:
        st.info("No exposure data found for this CAS number")

    # st.markdown('<hr class="section-divider">', unsafe_allow_html=True)

    st.markdown('<h2 class="section-title">Species sensitivity distribution curve</h2>', unsafe_allow_html=True)
    if len(result_sheet2) > 0:
        display_cols = ['Chemical name', 'CAS', 'Environmental medium','μ (log10 μg/L)',
                       'σ (log10 μg/L)', 'HC5 (ng/L)']
        valid_cols = [col for col in display_cols if col in result_sheet2.columns]
        if valid_cols:
            st.dataframe(
                result_sheet2[valid_cols].drop_duplicates(),
                use_container_width=True,
                hide_index=True
            )
        else:
            st.info("No SSD data available")
    else:
        st.info("No SSD data found for this CAS number")

    st.markdown('<h2 class="section-title">Risk assessment</h2>', unsafe_allow_html=True)
    if len(result_sheet2) > 0:
        display_cols = ['Chemical name', 'CAS', 'PAF',
                       'Environmental medium']
        valid_cols = [col for col in display_cols if col in result_sheet2.columns]
        if valid_cols:
            st.dataframe(
                result_sheet2[valid_cols].drop_duplicates(),
                use_container_width=True,
                hide_index=True
            )
        else:
            st.info("No risk assessment data available")
    else:
        st.info("No risk assessment data found for this CAS number")


    st.markdown('<div class="download-btn-container">', unsafe_allow_html=True)
    if st.download_button(
        label="Download",
        data=pd.concat([result_sheet1, result_sheet2]).to_csv(index=False, encoding='gbk').encode('gbk'),
        file_name=f"{cas_input}_data.csv",
        mime="text/csv",
        key="download_btn"
    ):
        # st.success("Data successfully downloaded!")
        pass
    st.markdown('</div>', unsafe_allow_html=True)

    if st.button("Back", key="back_btn"):
        for key in ['cas_input', 'result_sheet1', 'result_sheet2', 'page']:
            st.session_state.pop(key, None)
        st.session_state['page'] = 'home'
        st.rerun()

# -------------------------- Main program entry --------------------------
def main():
    if 'page' not in st.session_state:
        st.session_state['page'] = 'home'
    
    set_global_styles()

    if st.session_state['page'] == 'home':
        render_home_page()
    elif st.session_state['page'] == 'result':
        render_result_page()

if __name__ == "__main__":
    main()