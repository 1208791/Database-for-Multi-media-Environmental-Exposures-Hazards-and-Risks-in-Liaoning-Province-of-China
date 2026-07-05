import streamlit as st
import pandas as pd
import os

st.set_page_config(layout="wide", page_title="Ecological Hazard Database", page_icon="🌊")

# -------------------------- 全局配置 --------------------------
EXCEL_FILE_NAME = "data.xlsx"
EXCEL_FILE_PATH = os.path.join(os.path.dirname(__file__), EXCEL_FILE_NAME)
QUERY_COLUMN = 'CAS'

# 🎨 主题配色
THEME = {
    'primary': '#1f77b4',  # 主蓝色
    'primary_dark': '#155a8a',  # 深蓝
    'secondary': '#2ca02c',  # 成功绿
    'warning': '#ff7f0e',  # 警告橙
    'danger': '#d62728',  # 危险红
    'bg_light': '#f8f9fa',  # 浅灰背景
    'card_shadow': '0 4px 12px rgba(31, 119, 180, 0.15)'
}


# -------------------------- 页面样式（全局）- 所有字体≥36px --------------------------
def set_global_styles():
    st.markdown(f"""
        <style>
        /* =============== 全局基础样式 =============== */
        * {{ font-size: 36px !important; line-height: 1.4 !important; box-sizing: border-box; }}
        html, body, [class*="css"] {{ font-size: 36px !important; }}

        /* Streamlit组件字体强化 */
        .stAlert, .stAlert p, .stSpinner, .stSpinner p,
        .stMarkdown, .stMarkdown p, .stDataFrame, .stTable,
        .stDownloadButton, .stButton,
        div[data-testid="stMarkdownContainer"], div[data-testid="stText"] {{
            font-size: 36px !important;
        }}

        /* =============== 卡片式模块容器 =============== */
        .module-card {{
            background: white;
            border-radius: 16px;
            padding: 25px 30px;
            margin: 20px 0;
            box-shadow: {THEME['card_shadow']};
            border-left: 5px solid {THEME['primary']};
            transition: transform 0.2s ease, box-shadow 0.2s ease;
        }}
        .module-card:hover {{
            transform: translateY(-3px);
            box-shadow: 0 8px 25px rgba(31, 119, 180, 0.25);
        }}

        /* =============== 模块标题样式（带图标） =============== */
        .module-title {{
            display: flex;
            align-items: center;
            gap: 15px;
            font-size: 42px !important;
            color: {THEME['primary']};
            font-weight: 700;
            margin: 0 0 20px 0;
            padding-bottom: 15px;
            border-bottom: 2px solid #e9ecef;
        }}
        .module-icon {{
            font-size: 48px;
            display: inline-flex;
            align-items: center;
            justify-content: center;
            width: 60px;
            height: 60px;
            background: linear-gradient(135deg, {THEME['primary']}, {THEME['primary_dark']});
            border-radius: 12px;
            color: white;
            box-shadow: 0 3px 10px rgba(31, 119, 180, 0.3);
        }}

        /* =============== 表格专项 =============== */
        .dataframe, .dataframe th, .dataframe td,
        table, th, td, .stDataFrame table {{
            font-size: 36px !important;
            padding: 12px 15px !important;
            line-height: 1.3 !important;
        }}
        .dataframe th {{
            font-weight: 600 !important;
            background-color: #f1f8ff !important;
            color: {THEME['primary_dark']} !important;
            border-bottom: 2px solid {THEME['primary']} !important;
        }}
        .dataframe td {{ border-bottom: 1px solid #e9ecef !important; }}
        .dataframe tr:hover {{ background-color: #f8fbff !important; }}

        /* =============== 按钮统一规范 =============== */
        button, .stButton button, .stDownloadButton button {{
            font-size: 36px !important;
            height: 80px !important;
            line-height: 80px !important;
            min-width: 180px !important;
            padding: 0 25px !important;
            border-radius: 12px !important;
            font-weight: 600 !important;
            transition: all 0.2s ease !important;
            box-shadow: 0 2px 8px rgba(0,0,0,0.1) !important;
        }}
        button:hover, .stButton button:hover {{
            background-color: {THEME['primary_dark']} !important;
            border-color: {THEME['primary_dark']} !important;
            transform: translateY(-2px) !important;
            box-shadow: 0 6px 20px rgba(31, 119, 180, 0.3) !important;
        }}

        /* =============== 输入框强化 =============== */
        input[type="text"] {{
            font-size: 36px !important;
            height: 80px !important;
            padding: 0 25px !important;
            border-radius: 12px 0 0 12px !important;
            border: 2px solid {THEME['primary']} !important;
            transition: border-color 0.2s ease;
        }}
        input[type="text"]:focus {{
            border-color: {THEME['primary_dark']} !important;
            box-shadow: 0 0 0 3px rgba(31, 119, 180, 0.15) !important;
            outline: none !important;
        }}
        input::placeholder {{
            font-size: 36px !important;
            color: #999999 !important;
            opacity: 1 !important;
        }}

        /* =============== 搜索区域样式 =============== */
        .search-container {{
            display: flex;
            justify-content: center;
            gap: 0;
            margin: 30px 0 50px 0;
            background: white;
            padding: 8px;
            border-radius: 16px;
            box-shadow: {THEME['card_shadow']};
            max-width: 1000px;
            margin-left: auto;
            margin-right: auto;
        }}
        .search-input {{
            width: 100%;
            height: 80px;
            padding: 0 25px !important;
            border: none !important;
            border-radius: 12px 0 0 12px !important;
            font-size: 36px !important;
            box-sizing: border-box;
        }}
        .search-button {{
            height: 80px;
            width: 200px;
            background: linear-gradient(135deg, {THEME['primary']}, {THEME['primary_dark']}) !important;
            color: white !important;
            border: none !important;
            border-radius: 0 12px 12px 0 !important;
            font-size: 36px !important;
            font-weight: 600;
            margin: 0 !important;
        }}

        /* =============== CAS信息卡片 =============== */
        .cas-card {{
            background: linear-gradient(135deg, #f8fbff, #e9f2fa);
            border-radius: 16px;
            padding: 25px 30px;
            margin: 25px 0;
            border: 2px solid #d4e6f1;
        }}
        .cas-info-row {{
            display: flex;
            justify-content: center;
            align-items: center;
            gap: 30px;
            flex-wrap: wrap;
            font-size: 36px !important;
        }}
        .cas-item {{
            display: flex;
            align-items: center;
            gap: 10px;
            font-size: 36px !important;
            color: #333;
            font-weight: 500;
            padding: 10px 20px;
            background: white;
            border-radius: 10px;
            box-shadow: 0 2px 8px rgba(0,0,0,0.08);
        }}
        .cas-item strong {{ color: {THEME['primary']}; font-weight: 700; }}

        /* =============== 页脚样式 =============== */
        .footer {{
            display: flex;
            justify-content: space-between;
            align-items: center;
            padding: 20px 30px;
            position: fixed;
            bottom: 20px;
            width: 95%;
            max-width: 1400px;
            background: white;
            border-radius: 16px;
            box-shadow: {THEME['card_shadow']};
            z-index: 100;
            margin: 0 auto;
            left: 50%;
            transform: translateX(-50%);
        }}
        .footer-logo {{ display: flex; align-items: center; gap: 15px; }}
        .footer-text {{ font-size: 36px !important; font-weight: 500; color: #555; }}
        .help-btn {{
            background: white !important;
            color: {THEME['primary']} !important;
            border: 2px solid {THEME['primary']} !important;
        }}

        /* =============== 主标题 =============== */
        .main-title {{
            text-align: center;
            font-size: 56px !important;
            color: {THEME['primary']};
            font-weight: 700;
            margin: 30px 0 40px 0;
            line-height: 1.2 !important;
            text-shadow: 0 2px 4px rgba(31, 119, 180, 0.1);
            display: flex;
            align-items: center;
            justify-content: center;
            gap: 20px;
        }}
        .main-title-icon {{
            font-size: 64px;
            animation: wave 3s ease-in-out infinite;
        }}
        @keyframes wave {{
            0%, 100% {{ transform: translateY(0); }}
            50% {{ transform: translateY(-10px); }}
        }}

        /* =============== 提示信息 =============== */
        .stAlert {{
            padding: 25px !important;
            border-radius: 12px !important;
            margin: 20px 0 !important;
            border-left: 5px solid;
        }}
        .stAlert p {{ margin: 0 !important; font-size: 36px !important; }}

        /* =============== 下载按钮容器 =============== */
        .download-container {{
            display: flex;
            justify-content: center;
            margin: 40px 0 60px 0;
        }}
        .download-btn {{
            background: linear-gradient(135deg, {THEME['secondary']}, #228b22) !important;
            color: white !important;
            border: none !important;
        }}
        .download-btn:hover {{
            background: linear-gradient(135deg, #228b22, #1e7b1e) !important;
            transform: translateY(-2px) !important;
        }}

        /* =============== 返回按钮 =============== */
        .back-btn {{
            background: white !important;
            color: {THEME['primary']} !important;
            border: 2px solid {THEME['primary']} !important;
        }}

        /* =============== 空状态提示 =============== */
        .empty-state {{
            text-align: center;
            padding: 40px;
            color: #888;
            font-style: italic;
            background: #f8f9fa;
            border-radius: 12px;
            border: 2px dashed #dee2e6;
        }}

        /* =============== 响应式优化 =============== */
        @media (max-width: 1200px) {{
            .search-container {{ flex-direction: column; padding: 15px; }}
            .search-input {{ border-radius: 12px !important; margin-bottom: 10px; }}
            .search-button {{ border-radius: 12px !important; width: 100%; }}
            .cas-info-row {{ flex-direction: column; gap: 15px; }}
            .footer {{ flex-direction: column; gap: 15px; text-align: center; }}
            .module-title {{ font-size: 38px !important; }}
            .module-icon {{ font-size: 42px; width: 52px; height: 52px; }}
        }}

        /* =============== Font Awesome 图标支持 =============== */
        .fa {{ font-size: 48px; margin-right: 10px; }}
        </style>

        <!-- Font Awesome CDN -->
        <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.4.0/css/all.min.css">

        <script>
        document.addEventListener('DOMContentLoaded', function() {{
            // 动态应用搜索框样式
            setTimeout(() => {{
                const input = document.querySelector('input[placeholder="Enter CAS number to search"]');
                if (input) {{
                    input.classList.add('search-input');
                    input.parentElement.classList.add('search-wrapper');
                }}
                const btn = document.querySelector('button:has(span:contains("Search"))');
                if (btn) btn.classList.add('search-button');
            }}, 500);
        }});
        </script>
    """, unsafe_allow_html=True)


# -------------------------- 辅助函数：带图标的模块标题 --------------------------
def module_header(icon: str, title: str, color: str = None):
    """生成带图标的模块标题"""
    color_style = f"color: {color};" if color else ""
    return f"""
    <div class="module-title">
        <span class="module-icon">{icon}</span>
        <span style="{color_style}">{title}</span>
    </div>
    """


# -------------------------- 数据加载函数 --------------------------
def load_all_data_on_query():
    try:
        with st.spinner("🔍 Querying data..."):
            df_sheet1 = pd.read_excel(EXCEL_FILE_PATH, sheet_name=0)
            df_sheet2 = pd.read_excel(EXCEL_FILE_PATH, sheet_name=1)

            for df, sheet_name in [(df_sheet1, "Sheet1"), (df_sheet2, "Sheet2")]:
                if QUERY_COLUMN not in df.columns:
                    st.error(f"❌ Worksheet '{sheet_name}' missing '{QUERY_COLUMN}' column!")
                    return None, None
                df[QUERY_COLUMN] = df[QUERY_COLUMN].astype(str).str.strip()

        return df_sheet1, df_sheet2
    except FileNotFoundError:
        st.error(
            f"❌ File not found: `{EXCEL_FILE_NAME}`\n\nPlease ensure the file is in the same directory as this script.")
        return None, None
    except Exception as e:
        st.error(f"❌ Failed to load data: {str(e)}")
        return None, None


# -------------------------- 页面渲染函数 --------------------------
def render_home_page():
    """Render the home page (query input page)"""

    # 🌊 主标题（带动画图标）
    st.markdown(
        '<div class="main-title"><span class="main-title-icon">🌊</span>Ecological Hazard Database of Emerging Contaminants in Liaoning Province<br></div>',
        unsafe_allow_html=True)

    # 🔍 搜索区域（卡片式）
    col1, col2 = st.columns([5, 1], gap="small", vertical_alignment="center")
    with col1:
        cas_input = st.text_input(
            "",
            placeholder="🔢 Enter CAS number to search",
            label_visibility="collapsed",
            key="cas_search_input"
        )
    with col2:
        search_btn = st.button("Search", use_container_width=True)
    st.markdown('</div>', unsafe_allow_html=True)


    # 👇 页脚
    col11_1, col11_2, col11_3 = st.columns([1, 5, 6], vertical_alignment="center")
    with col11_1:
        try:
            st.image("1.png", width=120)
        except:
            st.markdown(
                '<div style="width:120px;height:120px;background:linear-gradient(135deg,#1f77b4,#155a8a);border-radius:16px;display:flex;align-items:center;justify-content:center;color:white;font-size:48px">🏛️</div>',
                unsafe_allow_html=True)
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

    # 🔎 搜索逻辑
    if search_btn:
        if not cas_input.strip():
            st.warning("⚠️ Please enter a valid CAS number!")
        else:
            cas_input = cas_input.strip()
            df_sheet1, df_sheet2 = load_all_data_on_query()

            if df_sheet1 is not None and df_sheet2 is not None:
                result_sheet1 = df_sheet1[df_sheet1[QUERY_COLUMN] == cas_input]
                result_sheet2 = df_sheet2[df_sheet2[QUERY_COLUMN] == cas_input]

                # if len(result_sheet1) == 0 and len(result_sheet2) == 0:
                #     st.error(f"❌ No data found for CAS: `{cas_input}`\n\nPlease check the number and try again.")
                # else:
                st.session_state.update({
                    'cas_input': cas_input,
                    'result_sheet1': result_sheet1,
                    'result_sheet2': result_sheet2,
                    'page': 'result'
                })
                st.rerun()


def render_result_page():
    """Render the results page with enhanced UI"""
    cas_input = st.session_state.get('cas_input', '')
    result_sheet1 = st.session_state.get('result_sheet1', pd.DataFrame())
    result_sheet2 = st.session_state.get('result_sheet2', pd.DataFrame())

    # 🌊 主标题
    st.markdown('<div class="main-title"><span class="main-title-icon">🌊</span>Ecological Hazard Database of Emerging Contaminants in Liaoning Province</div>',
                unsafe_allow_html=True)

    # 🧪 CAS 信息卡片
    chemical_name = result_sheet1.iloc[0].get('Chemical name', 'N/A') if len(result_sheet1) > 0 else 'N/A'

    st.markdown(f"""
    <div class="cas-card">
        <div class="cas-info-row">
            <div class="cas-item">🔢 <strong>CAS:</strong> {cas_input}</div>
            <div class="cas-item">🧪 <strong>Name:</strong> {chemical_name}</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # 📊 模块 1: Exposure Data
    st.markdown(module_header("📥", "Exposure Data"), unsafe_allow_html=True)
    if len(result_sheet1) > 0:
        st.dataframe(result_sheet1, use_container_width=True, hide_index=True)
    else:
        st.markdown('<div class="empty-state">🔍 No exposure data found for this CAS number</div>',
                    unsafe_allow_html=True)

    # 📈 模块 2: Species Sensitivity Distribution
    st.markdown(module_header("📈", "Species Sensitivity Distribution", color=THEME['secondary']),
                unsafe_allow_html=True)
    if len(result_sheet2) > 0:
        display_cols = ['Chemical name', 'CAS', 'Mean of log toxicity value', 'Standard deviation', 'HC5 (ng/L)',
                        'Environmental medium']
        valid_cols = [col for col in display_cols if col in result_sheet2.columns]
        if valid_cols:
            st.dataframe(result_sheet2[valid_cols].drop_duplicates(), use_container_width=True, hide_index=True)
        else:
            st.markdown('<div class="empty-state">📋 No SSD data columns available</div>', unsafe_allow_html=True)
    else:
        st.markdown('<div class="empty-state">🔍 No SSD data found for this CAS number</div>', unsafe_allow_html=True)

    # ⚠️ 模块 3: Risk Assessment
    st.markdown(module_header("⚠️", "Risk Assessment", color=THEME['warning']), unsafe_allow_html=True)
    if len(result_sheet2) > 0:
        display_cols = ['Chemical name', 'CAS', 'PAF', 'Environmental medium']
        valid_cols = [col for col in display_cols if col in result_sheet2.columns]
        if valid_cols:
            st.dataframe(result_sheet2[valid_cols].drop_duplicates(), use_container_width=True, hide_index=True)
        else:
            st.markdown('<div class="empty-state">📋 No risk assessment columns available</div>', unsafe_allow_html=True)
    else:
        st.markdown('<div class="empty-state">🔍 No risk assessment data found for this CAS number</div>',
                    unsafe_allow_html=True)
    st.markdown('</div>', unsafe_allow_html=True)

    # 💾 下载按钮
    st.markdown('<div class="download-container">', unsafe_allow_html=True)
    if not result_sheet1.empty or not result_sheet2.empty:
        combined_df = pd.concat([result_sheet1, result_sheet2]).drop_duplicates()
        csv_data = combined_df.to_csv(index=False, encoding='utf-8').encode('utf-8')

        if st.download_button(
                label="💾 Download CSV",
                data=csv_data,
                file_name=f"{cas_input}_ecological_data.csv",
                mime="text/csv",
                key="download_btn",
                # type="primary"
        ):
            st.toast("✅ Data downloaded successfully!", icon="🎉")
    st.markdown('</div>', unsafe_allow_html=True)

    # ↩️ 返回按钮
    col_back, col_space = st.columns([1, 4])
    with col_back:
        if st.button("↩️ Back", key="back_btn", type="secondary", use_container_width=True):
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
