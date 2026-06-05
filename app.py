import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import re

# ==========================================
# 🤖 新增：智慧尋找標題列 (取代容易失效的 skiprows)
# ==========================================
def load_excel_robust(file, sheet_name):
    try:
        # 讀取整個表格，暫不預設標題列
        df = pd.read_excel(file, sheet_name=sheet_name, header=None)
        
        # 由上往下掃描，尋找同時包含 'date' 與 'tester' 的列作為標題
        header_row = 0
        for i, row in df.iterrows():
            row_str = [str(x).strip().lower() for x in row.values]
            if 'date' in row_str and any('tester' in x for x in row_str):
                header_row = i
                break
                
        df.columns = df.iloc[header_row]
        df = df.iloc[header_row+1:].reset_index(drop=True)
        df.columns = [str(c).strip() for c in df.columns]
        df = df.dropna(how='all')
        return df
    except Exception as e:
        return pd.DataFrame()

# ==========================================
# 🛡️ 新增：欄位自動修復引擎 (Auto-Healing Columns)
# ==========================================
def standardize_columns(df, sheet_type):
    df = df.copy()
    if df.empty: return df
    df.columns = [str(c).strip() for c in df.columns]
    
    # 統一 Date
    for c in list(df.columns):
        if str(c).strip().lower() == 'date':
            df.rename(columns={c: 'Date'}, inplace=True)
            break
    if 'Date' not in df.columns: df['Date'] = None

    # 統一任務說明
    for c in list(df.columns):
        if 'purpose' in c.lower() and 'description' in c.lower():
            df.rename(columns={c: 'Task Details'}, inplace=True)
            break
    if 'Task Details' not in df.columns: df['Task Details'] = "N/A"
        
    # 統一 TEMP
    for c in list(df.columns):
        if c.lower() == 'temp':
            df.rename(columns={c: 'TEMP'}, inplace=True)
            break
    if 'TEMP' not in df.columns: df['TEMP'] = "Unknown"
        
    # 統一 Customer Requestor
    for c in list(df.columns):
        if c.lower() == 'customer requestor':
            df.rename(columns={c: 'Customer Requestor'}, inplace=True)
            break
    if 'Customer Requestor' not in df.columns: df['Customer Requestor'] = "Unknown"

    if sheet_type == 'Tester':
        # 統一 Tester #
        for c in list(df.columns):
            if c.lower() in ['tester', 'tester#', 'tester #']:
                df.rename(columns={c: 'Tester #'}, inplace=True)
                break
        if 'Tester #' not in df.columns: df['Tester #'] = "Unknown"
        
        # 尋找數值時數 (避開文字時間)
        found = False
        for c in list(df.columns):
            if c.lower() == 'tester hours':
                if 'Tester Total Hours' in df.columns and c != 'Tester Total Hours':
                    df = df.drop(columns=['Tester Total Hours'])
                df.rename(columns={c: 'Tester Total Hours'}, inplace=True)
                found = True
                break
        if not found:
            if 'Tester Total Hours' not in df.columns: df['Tester Total Hours'] = 0.0

    elif sheet_type == 'Engineering':
        # 統一 Tester
        for c in list(df.columns):
            if c.lower() in ['tester', 'tester#', 'tester #']:
                df.rename(columns={c: 'Tester'}, inplace=True)
                break
        if 'Tester' not in df.columns: df['Tester'] = "Unknown"
        
        # 尋找 Engineering 時數
        found = False
        for c in list(df.columns):
            if c.lower() == 'eng hours2':
                if 'Engineering Support Hours' in df.columns:
                    df.drop(columns=['Engineering Support Hours'], inplace=True)
                df.rename(columns={c: 'Engineering Support Hours'}, inplace=True)
                found = True
                break
        if not found:
            for c in list(df.columns):
                if c.lower() == 'eng hours':
                    if 'Engineering Support Hours' in df.columns:
                        df.drop(columns=['Engineering Support Hours'], inplace=True)
                    df.rename(columns={c: 'Engineering Support Hours'}, inplace=True)
                    found = True
                    break
        if not found: df['Engineering Support Hours'] = 0.0

        # 如果檔案漏填工程師姓名 (Name)，自動修補
        found_name = False
        for c in list(df.columns):
            if c.lower() == 'name':
                df.rename(columns={c: 'Name'}, inplace=True)
                found_name = True
                break
        if not found_name: df['Name'] = "Unknown"
            
    return df

# ==========================================
# 共用函數區
# ==========================================
def split_and_distribute(df, target_col, hours_col):
    df = df.copy()
    if target_col not in df.columns: return df
    df[target_col] = df[target_col].astype(str).replace(['nan', 'None', ''], 'Unknown')
    def safe_split(val):
        if val == 'Unknown': return ['Unknown']
        parts = re.split(r'[/,;\n\r]+', str(val))
        parts = [p.strip() for p in parts if p.strip()]
        return parts if parts else ['Unknown']
    df['__split_list'] = df[target_col].apply(safe_split)
    df['__split_count'] = df['__split_list'].apply(len)
    df[hours_col] = df[hours_col] / df['__split_count']
    df = df.explode('__split_list')
    df[target_col] = df['__split_list']
    df = df.drop(columns=['__split_list', '__split_count'])
    return df

def aggregate_data(df, group_by_cols, hours_col, show_breakdown=True, is_eng=False):
    if isinstance(group_by_cols, str): group_by_cols = [group_by_cols]
    
    txt_click_expand = "(Click to expand)" if is_eng else "(點擊展開)"
    txt_total = "Total" if is_eng else "總計"
    txt_no_detail = "No details provided (N/A)" if is_eng else "無詳細說明 (N/A)"
    txt_team_detail_prefix = "🏢 [{team}] Task Details:\n" if is_eng else "🏢 【{team}】 任務明細：\n"
    
    breakdown_col = f"⏱️ {hours_col} {txt_click_expand}"
    
    columns_to_return = group_by_cols + [hours_col, 'Task Details']
    if show_breakdown:
        columns_to_return.insert(-1, breakdown_col)
        
    if df.empty: 
        return pd.DataFrame(columns=columns_to_return)
        
    res_hours = df.groupby(group_by_cols)[hours_col].sum().reset_index()
    res = res_hours
    
    if show_breakdown:
        def format_hours(g):
            total = g[hours_col].sum()
            if 'Team' not in g.columns:
                return f"{txt_total}: {total:.2f} hrs"
                
            cso_hrs = g[g['Team'] == 'CSO'][hours_col].sum()
            gchip_hrs = g[g['Team'] == 'Gchip'][hours_col].sum()
            other_hrs = g[g['Team'] == 'Other / Unassigned'][hours_col].sum()
            
            details = [f"{txt_total}: {total:.2f} hrs"]
            if cso_hrs > 0: details.append(f"  ├ CSO: {cso_hrs:.2f} hrs")
            if gchip_hrs > 0: details.append(f"  ├ Gchip: {gchip_hrs:.2f} hrs")
            if other_hrs > 0: details.append(f"  └ Other: {other_hrs:.2f} hrs")
            
            return '\n'.join(details)
            
        res_breakdown = df.groupby(group_by_cols).apply(format_hours).reset_index(name=breakdown_col)
        res = pd.merge(res_hours, res_breakdown, on=group_by_cols)

    def format_details(g):
        details_str = []
        if 'Team' not in g.columns:
            if 'Task Details' not in g.columns: return txt_no_detail
            tasks = g['Task Details'].unique()
            valid_items = [str(i).strip() for i in tasks if pd.notna(i) and str(i).strip() != '']
            return '\n'.join([f"  • {item}" for item in valid_items]) if valid_items else txt_no_detail
        
        for team_name in ['CSO', 'Gchip', 'Other / Unassigned']:
            team_rows = g[g['Team'] == team_name]
            if team_rows.empty or 'Task Details' not in team_rows.columns: continue
            tasks = team_rows['Task Details'].unique()
            valid_items = [str(i).strip() for i in tasks if pd.notna(i) and str(i).strip() != '']
            if valid_items:
                team_str = txt_team_detail_prefix.format(team=team_name) + '\n'.join([f"  • {item}" for item in valid_items])
                details_str.append(team_str)
        return '\n\n----------------------------------------\n\n'.join(details_str) if details_str else txt_no_detail
        
    res_details = df.groupby(group_by_cols).apply(format_details).reset_index(name='Task Details')
    
    res = pd.merge(res, res_details, on=group_by_cols)
    res[hours_col] = res[hours_col].round(2)
    
    if 'Month' not in group_by_cols: res = res.sort_values(hours_col, ascending=False)
    return res

# ==========================================
# 網頁主程式開始
# ==========================================
st.set_page_config(page_title="Hours Analysis Dashboard", layout="wide")

st.markdown("""
<style>
    div.row-widget.stRadio > div {
        flex-direction: row;
        gap: 20px;
        padding: 10px 0;
    }
    div.row-widget.stRadio label {
        font-size: 1.1rem !important;
        font-weight: 600 !important;
        cursor: pointer;
    }
    .stDataFrame {
        background-color: transparent !important;
    }
</style>
""", unsafe_allow_html=True)

# ==========================================
# 🌐 多國語系 (i18n) 開關與文字變數定義區
# ==========================================
is_eng = st.sidebar.toggle("🌐 Switch to English (全英文介面)", value=False)

TXT_APP_TITLE = "📊 Tester & Engineer Hours Advanced Dashboard" if is_eng else "📊 機台與工程師時數進階分析儀表板"
TXT_REL_NOTES_TITLE = "🚀 Release Notes (Click to Expand)" if is_eng else "🚀 版本更新紀錄 / Release Notes (點擊展開)"
TXT_REL_NOTES_CONTENT = """
* **v31 (最新版)**: 🤖 **智慧相容升級**！導入「AI尋標引擎」與「欄位自動修復機制」，自動適應 Excel 報表格式的變動（如標題行位置改變、工程師姓名缺漏、欄位名稱微調），從此告別上傳格式錯誤。
* **v30**: 🛡️ 欄位自動修復引擎建置，解決 Excel 實際欄位名稱微調問題。
* **v29**: 📅 **月份篩選功能**！支援多檔案的月份擷取與自訂時間過濾。
* **v28**: 📂 多檔案合併分析支援！上傳區塊支援一次選取多個 Excel 檔案合併。
* **v27**: 🐛 多國語系 Bug 修復。
* **v26**: 🌐 雙語系支援 (i18n)！全站字串抽離，支援中英無縫切換。
* **v25**: 🌟 標題與維度微調！
* **v1~v24**: 包含圖表渲染優化、防呆機制、拖曳上傳、均分邏輯等核心功能建置。
""" if not is_eng else """
* **v31 (Latest)**: 🤖 **Smart Compatibility Upgrade**! Introduced Auto-Healing Engine. It dynamically finds the correct headers and fixes missing/renamed columns (e.g., missing 'Name' column, header row shifting) across different Excel versions automatically.
* **v30**: 🛡️ Auto-Healing Columns logic added.
* **v29**: 📅 **Month Filter added**!
* **v28**: 📂 Multiple File Upload & Merge Support!
* **v27**: 🐛 i18n Bug Fix.
* **v26**: 🌐 Bilingual Support (i18n)!
* **v1~v25**: Core functionalities including algorithms, drag-and-drop upload, layout, and stability optimizations.
"""

TXT_SIDEBAR_CTRL = "⚙️ Control Panel" if is_eng else "⚙️ 控制面板 (Control Panel)"
TXT_SIDEBAR_GUIDE = """
**💡 Quick Guide:**
1. Upload **one or more** Excel files below.
2. Select Months, set KPI targets & Team definitions.
3. Select analysis views on the right.
""" if is_eng else """
**💡 操作指南：**
1. 於下方上傳 **一個或多個** Excel 檔案。
2. 選擇統計月份、設定 KPI 與團隊成員。
3. 在右側主畫面選擇不同分析維度。
"""
TXT_UPLOAD_FILE = "📂 Upload Excel File(s)" if is_eng else "📂 上傳 Excel 紀錄表 (支援多選)"
TXT_MONTH_FILTER = "📅 Select Months to Analyze" if is_eng else "📅 選擇統計月份"
TXT_KPI_SETTING = "🎯 KPI Target Settings" if is_eng else "🎯 KPI 目標設定"
TXT_TESTER_COUNT_LABEL = "Set total testers (for Min Target Hours)" if is_eng else "設定機台總數量 (供計算最低標時數)"
TXT_TEAM_DEF = "👥 Team Members Definition" if is_eng else "👥 團隊成員定義"
TXT_DEF_CSO = "Define CSO Members" if is_eng else "定義 CSO 成員"
TXT_DEF_GCHIP = "Define Gchip Members" if is_eng else "定義 Gchip 成員"

TXT_KPI_SUMMARY = "📌 Executive Summary" if is_eng else "📌 核心數據總覽 (Executive Summary)"
TXT_KPI_TOTAL_TESTER = "🖥️ Total Tester Hours" if is_eng else "🖥️ 總機台使用時數"
TXT_KPI_TOTAL_ENG = "🧑‍🔧 Total Eng Hours" if is_eng else "🧑‍🔧 總工程支援時數"
TXT_KPI_TOP_TESTER = "🔥 Top Usage Tester" if is_eng else "🔥 最高用量機台"

TXT_SWITCH_VIEW = "### 🔍 Switch Analysis View" if is_eng else "### 🔍 切換分析視角"
TXT_SELECT_DIM = "Select Analysis Dimension" if is_eng else "選擇分析維度"
TXT_NAV_TEAM = "🏢 Team Analysis" if is_eng else "🏢 團隊歸屬分析 (Team)"
TXT_NAV_MONTHLY = "📅 Monthly Trends" if is_eng else "📅 每月趨勢分析 (Monthly)"
TXT_NAV_TEMP = "🌡️ Advanced Dimensions (TEMP/ENG Member)" if is_eng else "🌡️ 進階維度分析 (TEMP/ENG Member)"
TXT_NAV_REQ = "👤 Requestor Analysis" if is_eng else "👤 客戶需求者分析 (Requestor)"

TXT_NO_FILE = "👈 Please upload Excel file(s) from the left sidebar to begin." if is_eng else "👈 請於左側邊欄 (Sidebar) 上傳 Excel 檔案以開始分析。"
TXT_ERROR = "Error occurred during execution:" if is_eng else "執行時發生錯誤:"
TXT_NO_DATA = "No data to display." if is_eng else "無資料可顯示。"
TXT_FILTER = "🔽 Filter" if is_eng else "🔽 篩選"
TXT_TASK_COL = "📋 Task Details (Click to expand)" if is_eng else "📋 任務說明 (點擊展開)"
TXT_TASK_HELP = "Click cell to view complete task details split by CSO and Gchip" if is_eng else "點擊儲存格，即可查看區分 CSO 與 Gchip 的完整工作內容"
TXT_BREAKDOWN_HELP = "Click cell to view CSO / Gchip hour breakdown" if is_eng else "點擊儲存格，即可查看該時數的 CSO / Gchip 貢獻拆分"

UI_TESTER_TEAM = "🟦 [Tester Hours] Total by Team" if is_eng else "🟦 [Tester Hours] 依團隊統計"
CHART_TESTER_TEAM = "[Tester Hours] Total by Team"
UI_ENG_TEAM = "🟧 [Engineering Hours] Total by Team" if is_eng else "🟧 [Engineering Hours] 依團隊統計"
CHART_ENG_TEAM = "[Engineering Hours] Total by Team"

UI_TESTER_MONTHLY = "🟦 [Tester Hours] Monthly by Tester" if is_eng else "🟦 [Tester Hours] 每月機台時數"
CHART_TESTER_MONTHLY = "[Tester Hours] Monthly by Tester"
UI_ENG_TESTER = "🟧 [Engineering Hours] Monthly by Tester" if is_eng else "🟧 [Engineering Hours] 每月依機台 (Tester) 統計"
CHART_ENG_TESTER = "[Engineering Hours] Monthly by Tester"

UI_TESTER_TEMP = "🟦 [Tester Hours] Total by TEMP" if is_eng else "🟦 [Tester Hours] 依溫度 (TEMP) 統計"
CHART_TESTER_TEMP = "[Tester Hours] Total by TEMP"
UI_ENG_MONTHLY_ENG = "🟧 [Engineering Hours] Monthly by Engineer" if is_eng else "🟧 [Engineering Hours] 每月工程師時數"
CHART_ENG_MONTHLY_ENG = "[Engineering Hours] Monthly by Engineer"

UI_TESTER_REQ = "🟦 [Tester Hours] Total by Requestor" if is_eng else "🟦 [Tester Hours] 依客戶統計"
CHART_TESTER_REQ = "[Tester Hours] Total by Requestor"
UI_ENG_REQ = "🟧 [Engineering Hours] Total by Requestor" if is_eng else "🟧 [Engineering Hours] 依客戶統計"
CHART_ENG_REQ = "[Engineering Hours] Total by Requestor"

# ==========================================
# 介面開始繪製
# ==========================================
st.title(TXT_APP_TITLE)

with st.expander(TXT_REL_NOTES_TITLE):
    st.markdown(TXT_REL_NOTES_CONTENT)

with st.sidebar:
    st.header(TXT_SIDEBAR_CTRL)
    st.info(TXT_SIDEBAR_GUIDE)
    uploaded_files = st.file_uploader(TXT_UPLOAD_FILE, type=["xlsx", "xls"], accept_multiple_files=True)

if uploaded_files:
    try:
        list_df_tester_raw = []
        list_df_eng_raw = []
        
        for file in uploaded_files:
            # 🌟 透過動態尋找標題列讀取
            temp_tester = load_excel_robust(file, "Tester Hours")
            # 🌟 透過欄位自癒引擎整理對應
            temp_tester = standardize_columns(temp_tester, 'Tester')
            list_df_tester_raw.append(temp_tester)
            
            temp_eng = load_excel_robust(file, "Engineering Hours")
            temp_eng = standardize_columns(temp_eng, 'Engineering')
            list_df_eng_raw.append(temp_eng)
            
        df_tester_raw = pd.concat(list_df_tester_raw, ignore_index=True)
        df_eng_raw = pd.concat(list_df_eng_raw, ignore_index=True)
        
        # 篩選所需的必備欄位 (由於自動修復引擎已確保欄位存在，可安心選取)
        df_tester = df_tester_raw[['Date', 'Tester #', 'Tester Total Hours', 'TEMP', 'Customer Requestor', 'Task Details']].copy()
        df_tester.dropna(subset=['Date', 'Tester #', 'Tester Total Hours'], how='all', inplace=True)
        df_tester['Date'] = pd.to_datetime(df_tester['Date'], errors='coerce')
        df_tester.dropna(subset=['Date'], inplace=True)
        df_tester['Month'] = df_tester['Date'].dt.to_period('M').astype(str)
        df_tester['Tester Total Hours'] = pd.to_numeric(df_tester['Tester Total Hours'], errors='coerce').fillna(0)

        df_eng = df_eng_raw[['Date', 'Name', 'Engineering Support Hours', 'Tester', 'Customer Requestor', 'Task Details']].copy()
        df_eng.dropna(subset=['Date', 'Name', 'Engineering Support Hours'], how='all', inplace=True)
        df_eng['Date'] = pd.to_datetime(df_eng['Date'], errors='coerce')
        df_eng.dropna(subset=['Date'], inplace=True)
        df_eng['Month'] = df_eng['Date'].dt.to_period('M').astype(str)
        df_eng['Engineering Support Hours'] = pd.to_numeric(df_eng['Engineering Support Hours'], errors='coerce').fillna(0)

        # --- 側邊欄：月份篩選 ---
        with st.sidebar:
            st.divider()
            all_months = sorted(list(set(df_tester.get('Month', [])).union(set(df_eng.get('Month', [])))))
            selected_months = st.multiselect(TXT_MONTH_FILTER, options=all_months, default=all_months)
            
            if selected_months:
                if not df_tester.empty: df_tester = df_tester[df_tester['Month'].isin(selected_months)]
                if not df_eng.empty: df_eng = df_eng[df_eng['Month'].isin(selected_months)]
            else:
                df_tester = df_tester.iloc[0:0] 
                df_eng = df_eng.iloc[0:0]
        
        # 進行時數均分邏輯
        for col in ['Tester #', 'TEMP', 'Customer Requestor']:
            if col in df_tester.columns:
                df_tester = split_and_distribute(df_tester, target_col=col, hours_col='Tester Total Hours')
        for col in ['Name', 'Tester', 'Customer Requestor']:
            if col in df_eng.columns:
                df_eng = split_and_distribute(df_eng, target_col=col, hours_col='Engineering Support Hours')

        # --- 側邊欄：KPI 設定與團隊定義 ---
        with st.sidebar:
            st.divider()
            st.subheader(TXT_KPI_SETTING)
            tester_count = st.number_input(TXT_TESTER_COUNT_LABEL, min_value=1, value=10, step=1)
            
            st.divider()
            st.subheader(TXT_TEAM_DEF)
            all_reqs_t = df_tester['Customer Requestor'].unique() if 'Customer Requestor' in df_tester.columns else []
            all_reqs_e = df_eng['Customer Requestor'].unique() if 'Customer Requestor' in df_eng.columns else []
            all_requestors = sorted(list(set(all_reqs_t) | set(all_reqs_e)))
            
            if 'cso_selection' not in st.session_state:
                st.session_state.cso_selection = [n for n in ['Alec'] if n in all_requestors]
            if 'gchip_selection' not in st.session_state:
                st.session_state.gchip_selection = [n for n in ['Rajesh', 'Louis', 'Chi-Chang'] if n in all_requestors]

            avail_for_cso = [x for x in all_requestors if x not in st.session_state.gchip_selection]
            avail_for_gchip = [x for x in all_requestors if x not in st.session_state.cso_selection]

            cso_members = st.multiselect(TXT_DEF_CSO, options=avail_for_cso, key="cso_selection")
            gchip_members = st.multiselect(TXT_DEF_GCHIP, options=avail_for_gchip, key="gchip_selection")

            def map_team(name):
                if name in cso_members: return 'CSO'
                elif name in gchip_members: return 'Gchip'
                else: return 'Other / Unassigned'

            if not df_tester.empty: df_tester['Team'] = df_tester['Customer Requestor'].apply(map_team)
            if not df_eng.empty: df_eng['Team'] = df_eng['Customer Requestor'].apply(map_team)

        # ==========================================
        # 📈 主畫面：頂部 KPI
        # ==========================================
        st.subheader(TXT_KPI_SUMMARY)
        
        total_tester_hrs = df_tester['Tester Total Hours'].sum() if not df_tester.empty else 0
        
        total_days = 0
        for m in selected_months:
            try: total_days += pd.Period(m).days_in_month
            except: pass
        if total_days == 0: total_days = 30
            
        target_utilization = 0.5
        min_required_hours = total_days * 24 * tester_count * target_utilization
        delta_val = total_tester_hrs - min_required_hours
        
        total_eng_hrs = df_eng['Engineering Support Hours'].sum() if not df_eng.empty else 0
        top_tester = df_tester.groupby('Tester #')['Tester Total Hours'].sum().idxmax() if not df_tester.empty else "N/A"
        
        TXT_KPI_TARGET = f"🎯 Target Min Hours ({tester_count} units/50%)" if is_eng else f"🎯 最低標使用時數 ({tester_count}台/50%)"
        
        kpi1, kpi2, kpi3, kpi4 = st.columns(4)
        kpi1.metric(label=TXT_KPI_TOTAL_TESTER, value=f"{total_tester_hrs:,.1f} hrs")
        kpi2.metric(
            label=TXT_KPI_TARGET, 
            value=f"{min_required_hours:,.0f} hrs", 
            delta=f"{delta_val:,.1f} hrs", 
            delta_color="normal"
        )
        kpi3.metric(label=TXT_KPI_TOTAL_ENG, value=f"{total_eng_hrs:,.1f} hrs")
        kpi4.metric(label=TXT_KPI_TOP_TESTER, value=f"{top_tester}")
        
        st.divider()

        # ==========================================
        # 🎨 圖表設定與排版函數
        # ==========================================
        plt.style.use('default')
        corporate_params = {
            "font.sans-serif": ["Microsoft JhengHei", "PingFang TC", "Arial Unicode MS", "SimHei", "sans-serif"],
            "axes.unicode_minus": False, "figure.facecolor": "#FFFFFF", "axes.facecolor": "#F8F9FA",
            "grid.color": "#DEE2E6", "grid.linestyle": "-", "grid.alpha": 0.8,
            "text.color": "#212529", "axes.labelcolor": "#495057", "xtick.color": "#6C757D", "ytick.color": "#6C757D",
        }
        sns.set_theme(style="whitegrid", rc=corporate_params)

        def render_table_and_chart(ui_title, chart_title, df, x_col, y_col, hue_col=None, filter_col=None, custom_palette=None, show_breakdown=True):
            st.markdown(f"#### {ui_title}")
            col_data, col_chart = st.columns([1, 2])
            with col_data:
                filtered_df = df
                if filter_col and not df.empty and filter_col in df.columns:
                    unique_items = sorted(df[filter_col].unique().tolist())
                    selected_items = st.multiselect(f"{TXT_FILTER} {filter_col}", options=unique_items, default=unique_items, key=f"filter_{chart_title}")
                    filtered_df = df[df[filter_col].isin(selected_items)]
                
                column_config = {
                    "Task Details": st.column_config.TextColumn(
                        TXT_TASK_COL, 
                        help=TXT_TASK_HELP,
                        width="medium"
                    )
                }
                
                if show_breakdown and y_col in filtered_df.columns:
                    txt_click_expand = "(Click to expand)" if is_eng else "(點擊展開)"
                    breakdown_col_name = f"⏱️ {y_col} {txt_click_expand}"
                    column_config[y_col] = None  
                    column_config[breakdown_col_name] = st.column_config.TextColumn(
                        breakdown_col_name,
                        help=TXT_BREAKDOWN_HELP,
                        width="medium"
                    )
                
                st.dataframe(
                    filtered_df, 
                    use_container_width=True, 
                    hide_index=True,  
                    column_config=column_config
                )
                
            with col_chart:
                if filtered_df.empty: st.warning(TXT_NO_DATA)
                else:
                    fig, ax = plt.subplots(figsize=(10, 4.5))
                    if hue_col and hue_col in filtered_df.columns:
                        sns.barplot(data=filtered_df, x=x_col, y=y_col, hue=hue_col, ax=ax, palette=custom_palette, edgecolor="#FFFFFF", linewidth=1.2)
                        legend = ax.legend(title=hue_col, bbox_to_anchor=(1.05, 1), loc='upper left', frameon=False)
                        plt.setp(legend.get_texts(), color='#495057'); plt.setp(legend.get_title(), color='#212529', fontweight='bold')
                    else:
                        sns.barplot(data=filtered_df, x=x_col, y=y_col, ax=ax, palette=custom_palette, edgecolor="#FFFFFF", linewidth=1.2)
                    ax.spines['top'].set_visible(False); ax.spines['right'].set_visible(False)
                    ax.spines['left'].set_color('#CED4DA'); ax.spines['bottom'].set_color('#CED4DA')
                    ax.set_title(chart_title, fontweight='bold', pad=15, color='#212529')
                    ax.set_xlabel(x_col, labelpad=10); ax.set_ylabel(y_col, labelpad=10)
                    plt.xticks(rotation=45, ha='right'); plt.tight_layout()
                    st.pyplot(fig)
            st.divider()

        # ==========================================
        # 📑 視角切換導覽列
        # ==========================================
        st.markdown(TXT_SWITCH_VIEW)
        
        selected_view = st.radio(
            label=TXT_SELECT_DIM,
            options=[TXT_NAV_TEAM, TXT_NAV_MONTHLY, TXT_NAV_TEMP, TXT_NAV_REQ],
            horizontal=True,
            label_visibility="collapsed" 
        )
        
        st.markdown("<br>", unsafe_allow_html=True)

        if selected_view == TXT_NAV_TEAM:
            team_tester_hours = aggregate_data(df_tester, 'Team', 'Tester Total Hours', show_breakdown=False, is_eng=is_eng)
            team_eng_hours = aggregate_data(df_eng, 'Team', 'Engineering Support Hours', show_breakdown=False, is_eng=is_eng)
            render_table_and_chart(UI_TESTER_TEAM, CHART_TESTER_TEAM, team_tester_hours, 'Team', 'Tester Total Hours', filter_col='Team', custom_palette=['#2B5B84', '#E67E22', '#95A5A6'], show_breakdown=False)
            render_table_and_chart(UI_ENG_TEAM, CHART_ENG_TEAM, team_eng_hours, 'Team', 'Engineering Support Hours', filter_col='Team', custom_palette=['#2980B9', '#D35400', '#7F8C8D'], show_breakdown=False)

        elif selected_view == TXT_NAV_MONTHLY:
            monthly_tester_hours = aggregate_data(df_tester, ['Month', 'Tester #'], 'Tester Total Hours', is_eng=is_eng)
            eng_tester_hours = aggregate_data(df_eng, ['Month', 'Tester'], 'Engineering Support Hours', is_eng=is_eng)
            render_table_and_chart(UI_TESTER_MONTHLY, CHART_TESTER_MONTHLY, monthly_tester_hours, 'Month', 'Tester Total Hours', hue_col='Tester #', filter_col='Tester #', custom_palette='deep')
            render_table_and_chart(UI_ENG_TESTER, CHART_ENG_TESTER, eng_tester_hours, 'Month', 'Engineering Support Hours', hue_col='Tester', filter_col='Tester', custom_palette='Oranges_r')

        elif selected_view == TXT_NAV_TEMP:
            temp_hours = aggregate_data(df_tester, 'TEMP', 'Tester Total Hours', is_eng=is_eng)
            monthly_eng_hours = aggregate_data(df_eng, ['Month', 'Name'], 'Engineering Support Hours', is_eng=is_eng)
            render_table_and_chart(UI_TESTER_TEMP, CHART_TESTER_TEMP, temp_hours, 'TEMP', 'Tester Total Hours', filter_col='TEMP', custom_palette='Blues_r')
            render_table_and_chart(UI_ENG_MONTHLY_ENG, CHART_ENG_MONTHLY_ENG, monthly_eng_hours, 'Month', 'Engineering Support Hours', hue_col='Name', filter_col='Name', custom_palette='muted')

        elif selected_view == TXT_NAV_REQ:
            tester_req_hours = aggregate_data(df_tester, 'Customer Requestor', 'Tester Total Hours', show_breakdown=False, is_eng=is_eng)
            eng_req_hours = aggregate_data(df_eng, 'Customer Requestor', 'Engineering Support Hours', show_breakdown=False, is_eng=is_eng)
            render_table_and_chart(UI_TESTER_REQ, CHART_TESTER_REQ, tester_req_hours, 'Customer Requestor', 'Tester Total Hours', filter_col='Customer Requestor', custom_palette='Set2', show_breakdown=False)
            render_table_and_chart(UI_ENG_REQ, CHART_ENG_REQ, eng_req_hours, 'Customer Requestor', 'Engineering Support Hours', filter_col='Customer Requestor', custom_palette='Set1', show_breakdown=False)

    except Exception as e:
        st.error(f"{TXT_ERROR} {e}")

else:
    st.info(TXT_NO_FILE)
