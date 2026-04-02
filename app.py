import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
import calendar

# ==========================================
# 1. Page Configuration & Setup
# ==========================================
st.set_page_config(page_title="Tester & Engineering Dashboard", layout="wide")

# 解決 Matplotlib 中文顯示問題 (優先採用系統內建中文字型)
plt.rcParams['font.sans-serif'] = ['Microsoft JhengHei', 'PingFang TC', 'SimHei', 'Arial Unicode MS', 'sans-serif']
plt.rcParams['axes.unicode_minus'] = False

# Apply Professional Corporate Theme for Seaborn
sns.set_theme(style="whitegrid", palette="muted")
plt.rcParams.update({
    "axes.edgecolor": "#E0E0E0",
    "axes.linewidth": 1,
    "grid.color": "#E0E0E0",
    "grid.linestyle": "--",
    "grid.alpha": 0.7,
    "axes.labelcolor": "#495057",
    "text.color": "#212529",
    "xtick.color": "#495057",
    "ytick.color": "#495057"
})

# ==========================================
# 2. Language Toggle System
# ==========================================
if "lang" not in st.session_state:
    st.session_state.lang = "中文" # 預設語言

def t(en_text, zh_text):
    """Dynamic translation helper function"""
    return en_text if st.session_state.lang == "English" else zh_text

# ==========================================
# 3. Sidebar & Setup
# ==========================================
st.sidebar.radio(
    "🌐 Language / 語言", 
    ["中文", "English"], 
    key="lang", 
    horizontal=True
)
st.sidebar.divider()

st.sidebar.title(t("⚙️ Control Panel", "⚙️ 控制面板"))
uploaded_file = st.sidebar.file_uploader(
    t("Upload Excel File", "上傳 Excel 檔案"), 
    type=["xlsx", "xls"]
)

st.sidebar.divider()
st.sidebar.subheader(t("🎯 KPI Target Setting", "🎯 KPI 目標設定"))
target_testers = st.sidebar.number_input(
    t("Number of Testers", "機台數量"), 
    min_value=1, value=10, step=1
)

# ==========================================
# 4. Core Functions
# ==========================================
def split_and_distribute(df, col_to_split, hours_col):
    """Splits multiple entities separated by /, ;, ,, or \n and distributes hours equally."""
    if col_to_split not in df.columns:
        return df
    
    df[col_to_split] = df[col_to_split].astype(str).str.replace(r'[/;\n]', ',', regex=True)
    
    new_rows = []
    for _, row in df.iterrows():
        # [V28 FIX] Force string conversion and handle 'nan' to prevent 'float object has no attribute split'
        val = str(row[col_to_split]).strip()
        if val.lower() == 'nan' or not val:
            continue
            
        entities = [e.strip() for e in val.split(',') if e.strip()]
        if not entities:
            continue
            
        distributed_hours = row[hours_col] / len(entities)
        for entity in entities:
            new_row = row.copy()
            new_row[col_to_split] = entity
            new_row[hours_col] = distributed_hours
            new_rows.append(new_row)
            
    return pd.DataFrame(new_rows)

def aggregate_data(df, group_col, hours_col, show_breakdown=False):
    """Aggregates data, generates task descriptions, and optionally breaks down hours by team."""
    if df.empty:
        return pd.DataFrame()

    if 'Team' not in df.columns:
        df['Team'] = 'Other/Unassigned'

    desc_cols = [c for c in ['Lot #wafer', 'Purpose', 'Description'] if c in df.columns]
    if desc_cols:
        df['Task_Detail'] = df[desc_cols].astype(str).agg(' | '.join, axis=1)
    else:
        df['Task_Detail'] = t("No description available", "無可用說明")

    grouped = df.groupby(group_col)
    
    # Define dynamic column names
    task_col_name = t("📋 Task Description (Click to Expand)", "📋 任務說明 (點擊展開)")
    breakdown_col_name = t("⏱️ Hours Breakdown (Click to Expand)", "⏱️ 時數明細 (點擊展開)")
    
    results = []
    for name, group in grouped:
        total_hrs = group[hours_col].sum()
        
        task_texts = []
        for team in ['CSO', 'Gchip', 'Other/Unassigned']:
            team_data = group[group['Team'] == team]
            if not team_data.empty:
                unique_tasks = team_data['Task_Detail'].unique()
                task_texts.append(f"----- {team} -----")
                task_texts.extend([f"• {t_desc}" for t_desc in unique_tasks])
        
        full_task_desc = "\n".join(task_texts)
        
        row_data = {
            group_col: name,
            hours_col: round(total_hrs, 2),
            task_col_name: full_task_desc
        }
        
        if show_breakdown:
            cso_hrs = group[group['Team'] == 'CSO'][hours_col].sum()
            gchip_hrs = group[group['Team'] == 'Gchip'][hours_col].sum()
            row_data[breakdown_col_name] = f"CSO: {round(cso_hrs, 2)} hrs\nGchip: {round(gchip_hrs, 2)} hrs"
            
        results.append(row_data)

    res_df = pd.DataFrame(results).sort_values(by=hours_col, ascending=False).reset_index(drop=True)
    return res_df

def render_table_and_chart(df, ui_title, chart_title, group_col, hours_col, show_breakdown=False, is_engineering=False):
    """Renders the UI layout with a table on the left and a chart on the right."""
    st.subheader(ui_title)
    
    warning_msg = t("⚠️ All items excluded. No data available for charting.", "⚠️ 已排除所有項目，無資料可供繪圖。")
    
    if df.empty:
        st.warning(warning_msg)
        return

    unique_items = sorted(df[group_col].astype(str).unique())
    filter_label = t(f"Filter {group_col}:", f"篩選 {group_col}:")
    selected_items = st.multiselect(filter_label, options=unique_items, default=unique_items, key=f"filter_{chart_title}")
    
    filtered_df = df[df[group_col].isin(selected_items)]
    
    if filtered_df.empty:
        st.warning(warning_msg)
        return

    agg_df = aggregate_data(filtered_df, group_col, hours_col, show_breakdown=show_breakdown)

    col1, col2 = st.columns([1, 2])
    
    task_col_name = t("📋 Task Description (Click to Expand)", "📋 任務說明 (點擊展開)")
    breakdown_col_name = t("⏱️ Hours Breakdown (Click to Expand)", "⏱️ 時數明細 (點擊展開)")

    with col1:
        st.dataframe(
            agg_df, 
            use_container_width=True,
            hide_index=True,
            column_config={
                task_col_name: st.column_config.TextColumn(width="large"),
                breakdown_col_name: st.column_config.TextColumn(width="medium") if show_breakdown else None
            }
        )
        
    with col2:
        fig, ax = plt.subplots(figsize=(10, 5))
        palette = "Oranges_r" if is_engineering else "Blues_r"
        
        if group_col == 'Team':
            color_map = {'CSO': '#1F4E79', 'Gchip': '#F47D20', 'Other/Unassigned': '#A6A6A6'}
            sns.barplot(data=agg_df, x=group_col, y=hours_col, palette=color_map, ax=ax, edgecolor="white")
        else:
            sns.barplot(data=agg_df, x=group_col, y=hours_col, palette=palette, ax=ax, edgecolor="white")
            
        ax.set_title(chart_title, fontweight='bold', pad=15)
        plt.xticks(rotation=45, ha='right')
        ax.set_ylabel(t("Total Hours", "總時數"))
        ax.set_xlabel("")
        sns.despine(left=True, bottom=True)
        st.pyplot(fig)

# ==========================================
# 5. Main App Logic
# ==========================================
st.title(t("📊 Engineering & Tester Operations Dashboard", "📊 工程與機台營運儀表板"))

guide_en = """
**User Guide:**
1. **Upload:** Use the sidebar panel to upload your data file.
2. **Define Teams:** Assign members to CSO or Gchip below.
3. **Navigate & Filter:** Use the tabs to switch views and drop-down menus to filter data.
"""
guide_zh = """
**操作指南：**
1. **上傳資料：** 請使用左側控制面板上傳您的 Excel 檔案。
2. **定義團隊：** 在下方動態定義 CSO 與 Gchip 的成員。
3. **瀏覽與篩選：** 點擊分頁切換視角，並使用表格上方的選單進行資料篩選。
"""
st.info(t(guide_en, guide_zh))

if uploaded_file is None:
    st.write(t("👈 Please upload an Excel file from the sidebar to begin analysis.", "👈 請先從左側面板上傳 Excel 檔案以開始分析。"))
else:
    try:
        # Load data
        df_tester_raw = pd.read_excel(uploaded_file, sheet_name="Tester Hours", skiprows=3)
        df_eng_raw = pd.read_excel(uploaded_file, sheet_name="Engineering Hours")
        
        # Clean Tester Data
        df_tester = df_tester_raw.copy()
        df_tester['Date'] = pd.to_datetime(df_tester['Date'], errors='coerce')
        df_tester = df_tester.dropna(subset=['Date'])
        df_tester['Month'] = df_tester['Date'].dt.to_period('M').astype(str)
        df_tester = split_and_distribute(df_tester, 'Tester #', 'Tester Total Hours')
        df_tester = split_and_distribute(df_tester, 'TEMP', 'Tester Total Hours')
        df_tester = split_and_distribute(df_tester, 'Customer Requestor', 'Tester Total Hours')
        
        # Clean Engineering Data
        df_eng = df_eng_raw.copy()
        df_eng['Date'] = pd.to_datetime(df_eng['Date'], errors='coerce')
        df_eng = df_eng.dropna(subset=['Date'])
        df_eng['Month'] = df_eng['Date'].dt.to_period('M').astype(str)
        df_eng = split_and_distribute(df_eng, 'Name', 'Engineering Support Hours')
        df_eng = split_and_distribute(df_eng, 'Customer Requestor', 'Engineering Support Hours')
        df_eng = split_and_distribute(df_eng, 'Tester', 'Engineering Support Hours')

        # KPI Calculation
        unique_months = pd.concat([df_tester['Date'], df_eng['Date']]).dt.to_period('M').unique()
        total_days = sum([calendar.monthrange(m.year, m.month)[1] for m in unique_months])
        kpi_target = total_days * 24 * target_testers * 0.5

        total_tester_hrs = df_tester['Tester Total Hours'].sum()
        total_eng_hrs = df_eng['Engineering Support Hours'].sum()
        kpi_variance = total_tester_hrs - kpi_target

        col_k1, col_k2, col_k3 = st.columns(3)
        col_k1.metric(t("🟦 Total Tester Hours", "🟦 總機台時數"), f"{total_tester_hrs:,.1f} Hrs")
        col_k2.metric(t("🟧 Total Engineering Hours", "🟧 總工程時數"), f"{total_eng_hrs:,.1f} Hrs")
        col_k3.metric(t("🎯 Minimum Target Hours (KPI)", "🎯 最低標使用時數 (KPI)"), f"{kpi_target:,.1f} Hrs", delta=f"{kpi_variance:,.1f} Hrs", delta_color="normal")
        st.divider()

        # Dynamic Team Assignment
        st.subheader(t("👥 Team Member Definition", "👥 團隊成員定義"))
        all_requestors = set(df_tester['Customer Requestor'].dropna().astype(str).unique()) | set(df_eng['Customer Requestor'].dropna().astype(str).unique())
        all_requestors = sorted(list(all_requestors))

        if "cso_selection" not in st.session_state:
            st.session_state["cso_selection"] = [n for n in ['Alec'] if n in all_requestors]
        if "gchip_selection" not in st.session_state:
            st.session_state["gchip_selection"] = [n for n in ['Rajesh', 'Louis', 'Chi-Chang'] if n in all_requestors]

        avail_for_cso = [n for n in all_requestors if n not in st.session_state["gchip_selection"]]
        avail_for_gchip = [n for n in all_requestors if n not in st.session_state["cso_selection"]]

        t_col1, t_col2 = st.columns(2)
        with t_col1:
            st.multiselect(t("Select CSO Members:", "選擇 CSO 成員："), options=avail_for_cso, key="cso_selection")
        with t_col2:
            st.multiselect(t("Select Gchip Members:", "選擇 Gchip 成員："), options=avail_for_gchip, key="gchip_selection")

        def map_team(name):
            name_str = str(name).strip()
            if name_str in st.session_state["cso_selection"]:
                return 'CSO'
            elif name_str in st.session_state["gchip_selection"]:
                return 'Gchip'
            else:
                return 'Other/Unassigned'

        df_tester['Team'] = df_tester['Customer Requestor'].apply(map_team)
        df_eng['Team'] = df_eng['Customer Requestor'].apply(map_team)

        # Tabs Navigation Translation
        tabs_en = ["Team Analysis", "Monthly Trend Analysis", "Advanced Dimension Analysis (TEMP/ENG Member)", "Customer Requestor Analysis"]
        tabs_zh = ["團隊歸屬分析", "每月趨勢分析", "進階維度分析 (TEMP/ENG Member)", "客戶需求者分析"]
        current_tabs = tabs_en if st.session_state.lang == "English" else tabs_zh

        active_tab = st.radio("Navigation", current_tabs, horizontal=True, label_visibility="collapsed")
        st.markdown("<br>", unsafe_allow_html=True)

        if active_tab in ["Team Analysis", "團隊歸屬分析"]:
            render_table_and_chart(df_tester, t("🟦 [Tester Hours] Team Allocation", "🟦 [Tester Hours] 團隊歸屬分配"), "[Tester Hours] Allocation by Team", "Team", "Tester Total Hours", show_breakdown=False)
            st.divider()
            render_table_and_chart(df_eng, t("🟧 [Engineering Hours] Team Allocation", "🟧 [Engineering Hours] 團隊歸屬分配"), "[Engineering Hours] Allocation by Team", "Team", "Engineering Support Hours", show_breakdown=False, is_engineering=True)

        elif active_tab in ["Monthly Trend Analysis", "每月趨勢分析"]:
            render_table_and_chart(df_tester, t("🟦 [Tester Hours] Monthly Tester Usage", "🟦 [Tester Hours] 每月機台使用時數"), "[Tester Hours] Monthly Usage by Tester", "Month", "Tester Total Hours", show_breakdown=True)
            st.divider()
            df_eng_month_tester = df_eng.copy()
            df_eng_month_tester['Month_Tester'] = df_eng_month_tester['Month'] + " / " + df_eng_month_tester['Tester'].astype(str)
            render_table_and_chart(df_eng_month_tester, t("🟧 [Engineering Hours] Monthly Support by Tester", "🟧 [Engineering Hours] 依機台與月份統計"), "[Engineering Hours] Monthly Support by Tester", "Month_Tester", "Engineering Support Hours", show_breakdown=True, is_engineering=True)

        elif active_tab in ["Advanced Dimension Analysis (TEMP/ENG Member)", "進階維度分析 (TEMP/ENG Member)"]:
            render_table_and_chart(df_tester, t("🟦 [Tester Hours] Usage by Temperature", "🟦 [Tester Hours] 依溫度 (TEMP) 統計"), "[Tester Hours] Usage by TEMP", "TEMP", "Tester Total Hours", show_breakdown=True)
            st.divider()
            render_table_and_chart(df_eng, t("🟧 [Engineering Hours] Support by Engineer", "🟧 [Engineering Hours] 依工程師個人統計"), "[Engineering Hours] Monthly Support by Engineer", "Name", "Engineering Support Hours", show_breakdown=True, is_engineering=True)

        elif active_tab in ["Customer Requestor Analysis", "客戶需求者分析"]:
            render_table_and_chart(df_tester, t("🟦 [Tester Hours] Usage by Customer Requestor", "🟦 [Tester Hours] 依客戶需求者統計"), "[Tester Hours] Usage by Requestor", "Customer Requestor", "Tester Total Hours", show_breakdown=False)
            st.divider()
            render_table_and_chart(df_eng, t("🟧 [Engineering Hours] Support by Customer Requestor", "🟧 [Engineering Hours] 依客戶需求者統計"), "[Engineering Hours] Support by Requestor", "Customer Requestor", "Engineering Support Hours", show_breakdown=False, is_engineering=True)

    except Exception as e:
        error_msg_en = f"Read Error: Please ensure the uploaded Excel file contains 'Tester Hours' and 'Engineering Hours' sheets.\n\nSystem details: {e}"
        error_msg_zh = f"讀取失敗：請確認上傳的 Excel 檔案內包含 'Tester Hours' 與 'Engineering Hours' 分頁。\n\n詳細錯誤： {e}"
        st.error(t(error_msg_en, error_msg_zh))

# ==========================================
# 6. Release Notes
# ==========================================
with st.expander(t("📝 Release Notes (Version History)", "📝 版本紀錄 (Release Notes)")):
    st.markdown(t("""
    * **v28:** Fixed a Pandas `float` type conversion bug in the data distribution algorithm when encountering blank cells.
    * **v27:** Added interactive English/Chinese language toggle in the sidebar.
    * **v26:** Translated the entire user interface and text elements to English for corporate standardization.
    * **v25:** Adjusted dimensional names and aligned engineering tester data with a monthly timeline.
    * **v24:** Reorganized tabs for better logical grouping.
    * **v23:** Added dynamic KPI Target Settings (Number of Testers) to the sidebar.
    * **v22:** Focused the Hours Breakdown feature strictly on Trend and Advanced Dimension tabs.
    * **v21:** Implemented expandable 'Hours Breakdown' feature showing CSO/Gchip ratios.
    * **v20:** Replaced custom CSS tabs with native horizontal radio buttons.
    * **v19:** Styled tabs with professional corporate colors and added KPI logic framework.
    * **v18:** Added UI CSS enhancements for visual tabs and restored full release notes.
    * **v17:** Restructured layout with Sidebar, Top KPIs, and Tab Navigation.
    * **v16:** Professional corporate styling (white/grey background, corporate blue/orange palettes).
    * **v15:** Added distinct CSO/Gchip separating lines within the expandable task description window.
    * **v14:** Created 'Task Description' aggregation algorithm with interactive expanding columns.
    * **v1 - v13:** Core setup, aggregation, charting, and multi-unit time-splitting logic.
    """, """
    * **v28:** 修復 Pandas 讀取空值與數字時自動轉型為 float 導致 `.split()` 報錯的邊界錯誤。
    * **v27:** 於左側控制面板新增中英文雙語切換功能，支援即時翻譯。
    * **v26:** 將所有 UI 元素強制轉為英文以符合外商標準（現已整合進 v27 雙語系統）。
    * **v25:** 微調分頁名稱為進階維度分析，並對齊工程師與機台的月份時間軸。
    * **v24:** 排版邏輯優化，將機台趨勢與人員維度分頁進行板塊重組。
    * **v23:** 新增動態 KPI 目標設定，使用者可自定義機台數量以連動計算稼動率。
    * **v22:** 聚焦時數明細展開功能，僅在趨勢與維度分析中啟用，保持介面清爽。
    * **v21:** 實裝表格展開功能，可即時查看該數據中 CSO/Gchip 的貢獻時數明細。
    * **v20:** 棄用高風險 CSS，全面改用水平 Radio 按鈕作為原生分頁導覽列。
    * **v19:** 加入最低標 KPI 公式計算，並優化頁籤的企業風格底色。
    * **v18:** 注入卡片式頁籤視覺強化，並完整修復先前的版本紀錄。
    * **v17:** 導入專業 BI 儀表板排版 (側邊欄、頂部大 KPI 指標、水平頁籤切換)。
    * **v16:** 全面套用低飽和度專業商務配色與極簡白底框線。
    * **v15:** 在「任務明細」彈出視窗中加入 CSO 與 Gchip 專屬的分隔線。
    * **v14:** 開發「文字聚合演算法」，讓使用者可直接從表格點擊展開觀看 Lot/Purpose 明細。
    * **v1 - v13:** 基礎架構建立、Matplotlib 繪圖修復、時數切分均分邏輯開發。
    """))
