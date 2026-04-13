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

        /* =============== 下拉框样式 =============== */
        select {
            font-size: 36px !important;
            height: 80px !important;
            line-height: 80px !important;
            padding: 0 25px !important;
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
            margin-bottom: 40px;
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
        .module-selector {
            display: flex;
            justify-content: center;
            gap: 20px;
            margin: 30px 0 50px 0;
        }
        .module-btn {
            font-size: 36px !important;
            padding: 15px 40px !important;
            min-width: 200px !important;
            height: 70px !important;
            border-radius: 8px !important;
            font-weight: 600 !important;
        }
        .module-btn.active {
            background-color: #1f77b4 !important;
            color: #ffffff !important;
        }
        .module-btn.inactive {
            background-color: #ffffff !important;
            color: #1f77b4 !important;
            border: 2px solid #1f77b4 !important;
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
            .module-selector { flex-direction: column; align-items: center; }
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
def load_data_for_module(module_name):
    """Load data for specified module (Exposure or Hazard)"""
    try:
        sheet_name = 'Exposure' if module_name == 'Exposure' else 'Hazard & Risk'
        with st.spinner(f"Loading {module_name} data..."):
            df = pd.read_excel(EXCEL_FILE_PATH, sheet_name=sheet_name)

            # Validate required columns
            if QUERY_COLUMN not in df.columns:
                st.error(f"Excel file's '{sheet_name}' worksheet does not contain '{QUERY_COLUMN}' column!")
                return None

            # Define filter columns for each module
            filter_cols = []
            if module_name == 'Exposure':
                filter_cols = ['City/Region', 'Environmental medium']
            elif module_name == 'Hazard & Risk':
                filter_cols = ['Environmental medium']

            # Check if filter columns exist
            for col in filter_cols:
                if col not in df.columns:
                    st.warning(f"Column '{col}' not found in '{sheet_name}' sheet. Filter options will be limited.")

            # Clean CAS column
            df[QUERY_COLUMN] = df[QUERY_COLUMN].astype(str).str.strip()

            return df
    except FileNotFoundError:
        st.error(f"File not found: {EXCEL_FILE_PATH}. Please confirm the file exists!")
        return None
    except Exception as e:
        st.error(f"Failed to load data: {str(e)}")
        return None


def get_filter_options(df, module_name):
    """Get unique values for filter dropdowns"""
    options = {}

    if module_name == 'Exposure':
        for col in ['City/Region', 'Environmental medium']:
            if col in df.columns:
                # Get unique non-null values, convert to string, filter empty
                unique_vals = df[col].dropna().astype(str).str.strip().unique()
                unique_vals = [v for v in unique_vals if v and v.lower() != 'nan']
                options[col] = ['All'] + sorted(unique_vals)
    elif module_name == 'Hazard & Risk':
        col = 'Environmental medium'
        if col in df.columns:
            unique_vals = df[col].dropna().astype(str).str.strip().unique()
            unique_vals = [v for v in unique_vals if v and v.lower() != 'nan']
            options[col] = ['All'] + sorted(unique_vals)

    return options


# -------------------------- 页面渲染函数 --------------------------
def render_home_page():
    """Render the home page with module selection and query inputs"""

    # Database title
    st.markdown(
        '<div class="main-title">Database for Multi-media Environmental Exposures, Hazards,and Risks in Liaoning Province of China</div>',
        unsafe_allow_html=True)

    # Module selector
    st.markdown('<div class="module-selector">', unsafe_allow_html=True)
    col_m1, col_m2 = st.columns([1, 1])
    with col_m1:
        if st.button("Exposure", key="btn_exposure", use_container_width=True):
            st.session_state['selected_module'] = 'Exposure'
            st.session_state['page'] = 'home'
            st.rerun()
    with col_m2:
        if st.button("Hazard & Risk", key="btn_hazard", use_container_width=True):
            st.session_state['selected_module'] = 'Hazard & Risk'
            st.session_state['page'] = 'home'
            st.rerun()
    st.markdown('</div>', unsafe_allow_html=True)

    # Highlight active module
    active_module = st.session_state.get('selected_module', 'Exposure')
    st.markdown(f"""
        <script>
        const expBtn = document.querySelector('button[key="btn_exposure"]');
        const hazBtn = document.querySelector('button[key="btn_hazard"]');
        if (expBtn && hazBtn) {{
            if ('{active_module}' === 'Exposure') {{
                expBtn.classList.add('active');
                expBtn.classList.remove('inactive');
                hazBtn.classList.add('inactive');
                hazBtn.classList.remove('active');
            }} else {{
                hazBtn.classList.add('active');
                hazBtn.classList.remove('inactive');
                expBtn.classList.add('inactive');
                expBtn.classList.remove('active');
            }}
        }}
        </script>
    """, unsafe_allow_html=True)

    # Query area
    # st.markdown('<div style="background:#f8f9fa;padding:30px;border-radius:12px;margin:20px 0;">',
    #             unsafe_allow_html=True)

    # CAS input (always present)
    cas_input = st.text_input(
        "CAS Number",
        placeholder="Enter CAS number to search",
        key="cas_query_input",
        label_visibility="collapsed"
    )

    # Module-specific filters
    filter_values = {}

    if active_module == 'Exposure':
        # st.markdown('<h3 style="color:#1f77b4;margin:25px 0 15px 0;">Exposure Filters</h3>', unsafe_allow_html=True)
        df_exposure = load_data_for_module('Exposure')
        if df_exposure is not None:
            options = get_filter_options(df_exposure, 'Exposure')
            col_f1, col_f2 = st.columns(2)
            with col_f1:
                if 'City/Region' in options:
                    filter_values['City/Region'] = st.selectbox(
                        "City/Region",
                        options['City/Region'],
                        key="filter_city_region",
                        label_visibility="collapsed" if 'City/Region' not in options else "visible"
                    )
            with col_f2:
                if 'Environmental medium' in options:
                    filter_values['Environmental medium'] = st.selectbox(
                        "Environmental medium",
                        options['Environmental medium'],
                        key="filter_sample_type",
                        label_visibility="collapsed" if 'Environmental medium' not in options else "visible"
                    )

    elif active_module == 'Hazard & Risk':
        # st.markdown('<h3 style="color:#1f77b4;margin:25px 0 15px 0;">Hazard Filters</h3>', unsafe_allow_html=True)
        df_hazard = load_data_for_module('Hazard & Risk')
        if df_hazard is not None:
            options = get_filter_options(df_hazard, 'Hazard & Risk')
            if 'Environmental medium' in options:
                filter_values['Environmental medium'] = st.selectbox(
                    "Environmental medium",
                    options['Environmental medium'],
                    key="filter_media"
                )

    st.markdown('</div>', unsafe_allow_html=True)

    # Search button
    st.markdown('<div class="search-container">', unsafe_allow_html=True)
    col1, col2 = st.columns([5, 1], gap="small", vertical_alignment="center")
    with col1:
        st.markdown('<div style="font-size:28px;color:#666;padding:10px 25px;">Press Search to query data</div>',
                    unsafe_allow_html=True)
    with col2:
        search_btn = st.button("Search", use_container_width=True, key="main_search_btn")
    st.markdown('</div>', unsafe_allow_html=True)

    # Footer area
    col11_1, col11_2, col11_3 = st.columns([1, 5, 6], vertical_alignment="center")
    with col11_1:
        try:
            st.image("1.png", width=150)
        except:
            st.markdown('<div style="width:150px;height:150px;background:#e9ecef;border-radius:8px"></div>',
                        unsafe_allow_html=True)
    with col11_2:
        st.markdown('<span class="footer-text">Dalian University of Technology</span>', unsafe_allow_html=True)
    with col11_3:
        if st.button("Help", key="help_btn", use_container_width=True):
            if active_module == 'Exposure':
                help_text = """
                **Exposure Module Usage:**<br>
                1. Enter a CAS Number<br>
                2. (Optional) Select City/Region from dropdown<br>
                3. (Optional) Select Environmental medium from dropdown<br>
                4. Click "Search" to view matching exposure data
                """
            else:
                help_text = """
                **Hazard & Risk Module Usage:**<br>
                1. Enter a CAS Number<br>
                2. (Optional) Select Environmental medium from dropdown<br>
                3. Click "Search" to view matching species sensitivity distribution curve data, hazardous concentration for 5% of species and potentially affected fraction
                """
            st.markdown(f"""
                <div style="background:#e7f3fe;border-left:4px solid #1f77b4;padding:20px;border-radius:8px;margin:15px 0;">
                {help_text}
                </div>
            """, unsafe_allow_html=True)
            st.button("Close Help", key="close_help_btn")

    # Handle search
    if search_btn:
        if not cas_input.strip():
            st.warning("Please enter a valid CAS number!")
        else:
            cas_input = cas_input.strip()
            module = st.session_state.get('selected_module', 'Exposure')
            df = load_data_for_module(module)

            if df is not None:
                # Apply filters
                mask = df[QUERY_COLUMN] == cas_input

                # Apply dropdown filters
                for col, val in filter_values.items():
                    if col in df.columns and val != 'All':
                        mask = mask & (df[col].astype(str).str.strip() == val)

                result_df = df[mask]

                # Store in session state
                st.session_state['cas_input'] = cas_input
                st.session_state['selected_module'] = module
                st.session_state['result_df'] = result_df
                st.session_state['filter_values'] = filter_values
                st.session_state['page'] = 'result'
                st.rerun()


def render_result_page():
    """Render the results page"""
    cas_input = st.session_state.get('cas_input', '')
    module = st.session_state.get('selected_module', 'Exposure')
    result_df = st.session_state.get('result_df', pd.DataFrame())
    filter_values = st.session_state.get('filter_values', {})

    # Title
    st.markdown(f"""
        <h1 class="main-title">{module} Module - Query Results</h1>
    """, unsafe_allow_html=True)

    # Query info
    filter_info = " | ".join([f"{k}: {v}" for k, v in filter_values.items() if v != 'All'])
    filter_display = f" | {filter_info}" if filter_info else ""

    st.markdown(f"""
        <div class="cas-info-row">
            <span class="cas-item">CAS: {cas_input}</span>
            <span class="cas-item">Module: {module}</span>
            <span class="cas-item">Records Found: {len(result_df)}</span>
        </div>
        <div style="text-align:center;font-size:28px;color:#666;">{filter_display}</div>
    """, unsafe_allow_html=True)

    st.markdown('<hr class="section-divider">', unsafe_allow_html=True)

    # Results table
    st.markdown(f'<h2 class="section-title">{module} Data Results</h2>', unsafe_allow_html=True)

    if len(result_df) > 0:
        st.dataframe(
            result_df,
            use_container_width=True,
            hide_index=True
        )
    else:
        st.info(f"No {module.lower()} data found for CAS '{cas_input}' with the selected filters.")

    # Download button
    st.markdown('<div class="download-btn-container">', unsafe_allow_html=True)
    if len(result_df) > 0:
        csv_data = result_df.to_csv(index=False, encoding='utf-8-sig')
        st.download_button(
            label="Download CSV",
            data=csv_data,
            file_name=f"{cas_input}_{module.lower()}_data.csv",
            mime="text/csv",
            key="download_btn"
        )
    st.markdown('</div>', unsafe_allow_html=True)

    # Back button
    if st.button("← Back to Search", key="back_btn"):
        # Clear result-related session state but keep module selection
        for key in ['cas_input', 'result_df', 'filter_values', 'page']:
            st.session_state.pop(key, None)
        st.session_state['page'] = 'home'
        st.rerun()


# -------------------------- Main program entry --------------------------
def main():
    # Initialize session state
    if 'page' not in st.session_state:
        st.session_state['page'] = 'home'
    if 'selected_module' not in st.session_state:
        st.session_state['selected_module'] = 'Exposure'

    # Apply global styles
    set_global_styles()

    # Route to appropriate page
    if st.session_state['page'] == 'home':
        render_home_page()
    elif st.session_state['page'] == 'result':
        render_result_page()


if __name__ == "__main__":
    main()