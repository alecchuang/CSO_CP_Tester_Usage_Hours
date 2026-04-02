import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import numpy as np
import calendar

# ==========================================
# 1. 頁面與雙語翻譯設定
# ==========================================
st.set_page_config(page_title="Tester & Engineering Dashboard", layout="wide")

LANG_DICT = {
    "TW": {
        "lang_selector": "🌐 選擇語言 / Select Language",
        "sidebar_settings": "⚙️ 控制面板",
        "upload_file": "📂 1. 上傳資料",
        "kpi_settings": "🎯 2. KPI 目標設定",
        "machine_count": "設定機台總數量 (預設 10 台)：",
        "team_def": "👥 3. 團隊成員定義",
        "cso_members": "選擇 CSO 成員：",
        "gchip_members": "選擇 Gchip 成員：",
        "guide_title": "💡 操作指南",
        "guide_1": "請先從左側上傳包含 'Tester Hours' 與 'Engineering Hours' 兩個分頁的 Excel 檔案。",
        "guide_2": "可於左側動態定義 CSO 與 Gchip 的團隊成員，預設值已自動帶入。",
        "guide_3": "左側可設定總機台數量，系統將自動依據當月天數計算最低標使用時數。",
        "guide_4": "點擊表格最右側的「任務說明」可展開檢視詳細工作項目。",
        "upload_msg": "👈 請先從左側上傳 Excel 檔案以開始分析。",
        "total_tester_hrs": "總機台使用時數",
        "target_hrs": "🎯 最低標使用時數 (Target)",
        "total_eng_hrs": "總工程師支援時數",
        "tab_team": "🏢 團隊歸屬分析",
        "tab_monthly": "📅 每月趨勢分析",
        "tab_adv": "🌡️ 進階維度分析 (TEMP/ENG Member)",
        "tab_req": "👤 客戶需求者分析",
        "filter_label": "🔍 篩選資料：",
        "no_data": "⚠️ 無資料可供顯示。",
        "col_task": "📋 任務說明 (點擊展開)",
        "col_detail": "⏱️ 時數明細 (點擊展開)",
        "cso_section": "--- CSO 任務 ---",
        "gchip_section": "--- Gchip 任務 ---",
        "error_sheet": "讀取失敗：請確認上傳的 Excel 檔案內包含 'Tester Hours' 與 'Engineering Hours' 這兩個分頁。",
        "release_notes": "📝 版本更新紀錄 (Release Notes)",
        "ui_team_tester": "依團隊 (Team) 統計",
        "ui_team_eng": "依團隊 (Team) 統計",
        "ui_monthly_tester": "每月機台總時數",
        "ui_monthly_eng": "每月依機台 (Tester) 統計",
        "ui_adv_tester": "依溫度 (TEMP) 統計",
        "ui_adv_eng": "每月工程師時數",
        "ui_req_tester": "依客戶需求者統計",
        "ui_req_eng": "依客戶需求者統計"
    },
    "EN": {
        "lang_selector": "🌐 Select Language / 選擇語言",
        "sidebar_settings": "⚙️ Control Panel",
        "upload_file": "📂 1. Upload Data",
        "kpi_settings": "🎯 2. KPI Target Settings",
        "machine_count": "Total Tester Count (Default 10):",
        "team_def": "👥 3. Team Definition",
        "cso_members": "Select CSO Members:",
        "gchip_members": "Select Gchip Members:",
        "guide_title": "💡 Operation Guide",
        "guide_1": "Upload an Excel file containing 'Tester Hours' and 'Engineering Hours' sheets from the left.",
        "guide_2": "Dynamically define CSO and Gchip members on the left. Defaults are loaded automatically.",
        "guide_3": "Set the total machine count on the left; the system calculates target hours based on days in the month.",
        "guide_4": "Click 'Task Description' in the tables to expand and view detailed tasks.",
        "upload_msg": "👈 Please upload an Excel file from the left sidebar to begin analysis.",
        "total_tester_hrs": "Total Tester Hours",
        "target_hrs": "🎯 Target Hours",
        "total_eng_hrs": "Total Engineering Hours",
        "tab_team": "🏢 Team Analysis",
        "tab_monthly": "📅 Monthly Trend",
        "tab_adv": "🌡️ Advanced Dimension (TEMP/ENG Member)",
        "tab_req": "👤 Customer Requestor Analysis",
        "filter_label": "🔍 Filter Data:",
        "no_data": "⚠️ No data to display.",
        "col_task": "📋 Task Description (Click to Expand)",
        "col_detail": "⏱️ Hours Breakdown (Click to Expand)",
        "cso_section": "--- CSO Tasks ---",
        "gchip_section": "--- Gchip Tasks ---",
        "error_sheet": "Read Error: Ensure the Excel file contains 'Tester Hours' and 'Engineering Hours' sheets.",
        "release_notes": "📝 Release Notes",
        "ui_team_tester": "Statistics by Team",
        "ui_team_eng": "Statistics by Team",
        "ui_monthly_tester": "Total Tester Hours by Month",
        "ui_monthly_eng": "Monthly Statistics by Tester",
        "ui_adv_tester": "Statistics by TEMP",
        "ui_adv_eng": "Monthly Engineering Hours",
        "ui_req_tester": "Statistics by Customer Requestor",
        "ui_req_eng": "Statistics by Customer Requestor"
    }
}

def _t(key):
    return LANG_DICT[st.session_state.lang].get(key, key)

# ==========================================
# 2. 核心處理函數 
# ==========================================
def split_and_distribute(df, col_name, value_col):
    if col_name not in df.columns:
        return df
    
    rows = []
    for _, row in df.iterrows():
        val = str(row[col_name])
        for sep in [',', ';', '\n', '/']:
            val = val.replace(sep, '|')
        
        parts = [p.strip() for p in val.split('|') if p.strip()]
        if not parts:
            parts = ['Unknown']
            
        divided_val = row[value_col] / len(parts) if pd.notnull(row[value_col]) else 0
        
        for p in parts:
            new_row = row.copy()
            new_row[col_name] = p
            new_row[value_col] = divided_val
            rows.append(new_row)
            
    return pd.DataFrame(rows)

def aggregate_data(df, group_col, val_col, show_breakdown=False):
    if df.empty:
        return pd.DataFrame()
        
    def custom_agg(x):
        d = {}
        d[val_col] = x[val_col].sum()
        
        cso_tasks, gchip_tasks = set(), set()
        for _, row in x.iterrows():
            desc = f"{row.get('Lot #wafer', '')} / {row.get('Purpose', '')} / {row.get('Description', '')}".strip(" /")
            if not desc: continue
            
            team = row.get('Team', 'Other')
            if team == 'CSO': cso_tasks.add(f"• {desc}")
            else: gchip_tasks.add(f"• {desc}")
            
        combined_desc = ""
        if cso_tasks: combined_desc += _t("cso_section") + "\n" + "\n".join(cso_tasks) + "\n\n"
        if gchip_tasks: combined_desc += _t("gchip_section") + "\n" + "\n".join(gchip_tasks)
        d[_t('col_task')] = combined_desc.strip()
        
        if show_breakdown:
            cso_hrs = x[x.get('Team', 'Other') == 'CSO'][val_col].sum()
            gchip_hrs = x[x.get('Team', 'Other') != 'CSO'][val_col].sum()
            d[_t('col_detail')] = f"CSO: {cso_hrs:.2f} hrs\nGchip: {gchip_hrs:.2f} hrs"
            
        return pd.Series(d)

    agg_df = df.groupby(group_col, dropna=False).apply(custom_agg).reset_index()
    agg_df[val_col] = agg_df[val_col].round(2)
    return agg_df.sort_values(by=val_col, ascending=False)

def get_target_hours(df, machine_count):
    if df.empty or 'Date' not in df.columns:
        return 0
    unique_months = df['Month'].dropna().unique()
    total_days = 0
    for m in unique_months:
        try:
            year, month = int(str(m)[:4]), int(str(m)[5:7])
            _, days = calendar.monthrange(year, month)
            total_days += days
        except:
            total_days += 30
    return total_days * 24 * machine_count * 0.5

def render_table_and_chart(df, group_col, val_col, ui_title_key, chart_title, source_mark, palette="Blues_r", show_breakdown=False):
    ui_title = _t(ui_title_key)
    st.subheader(f"{source_mark} {ui_title}")
    
    if df.empty:
        st.warning(_t("no_data"))
        st.markdown("---")
        return
        
    unique_vals = df[group_col].dropna().unique().tolist()
    
    # 建立左右分欄
    col1, col2 = st.columns([1, 2])
    
    with col1:
        # 將篩選器完美置入左側表格的正上方
        selected_vals = st.multiselect(f"{_t('filter_label')} {ui_title}", unique_vals, default=unique_vals, key=f"ms_{chart_title}")
        filtered_df = df[df[group_col].isin(selected_vals)]
        
        if filtered_df.empty:
            st.warning(_t("no_data"))
        else:
            agg_df = aggregate_data(filtered_df, group_col, val_col, show_breakdown=show_breakdown)
            cfg = {
                _t('col_task'): st.column_config.TextColumn(width="medium"),
                _t('col_detail'): st.column_config.TextColumn(width="small")
            }
            st.dataframe(agg_df, use_container_width=True, hide_index=True, column_config=cfg)
        
    with col2:
        if filtered_df.empty:
            st.info(_t("no_data"))
        else:
            sns.set_theme(style="whitegrid")
            fig, ax = plt.subplots(figsize=(8, 4))
            sns.barplot(data=agg_df, x=group_col, y=val_col, hue=group_col, palette=palette, ax=ax, legend=False, edgecolor="#FFFFFF")
            ax.set_title(chart_title, fontweight='bold', color="#212529", pad=15)
            ax.set_xlabel(group_col, color="#212529", labelpad=10)
            ax.set_ylabel(val_col, color="#212529", labelpad=10)
            plt.xticks(rotation=45, ha='right')
            
            ax.spines['top'].set_visible(False)
            ax.spines['right'].set_visible(False)
            
            st.pyplot(fig)
            plt.close(fig)
        
    st.markdown("---")

# ==========================================
# 3. 網頁主程式與側邊欄
# ==========================================
if "lang" not in st.session_state:
    st.session_state.lang = "TW"

with st.sidebar:
    st.radio("🌐 Language Selector", ["TW", "EN"], key="lang", horizontal=True)
    st.header(_t("sidebar_settings"))
    
    st.subheader(_t("upload_file"))
    uploaded_file = st.file_uploader("", type=["xlsx", "xls"])
    
    st.subheader(_t("kpi_settings"))
    machine_count = st.number_input(_t("machine_count"), min_value=1, value=10, step=1)

title_text = "📊 設備與工程時數營運儀表板" if st.session_state.lang == "TW" else "📊 Tester & Engineering Hours Dashboard"
st.title(title_text)

st.info(f"""
**{_t("guide_title")}**
* {_t("guide_1")}
* {_t("guide_2")}
* {_t("guide_3")}
* {_t("guide_4")}
""")

if uploaded_file is None:
    st.warning(_t("upload_msg"))
else:
    try:
        df_tester = pd.read_excel(uploaded_file, sheet_name="Tester Hours", skiprows=3)
        df_eng = pd.read_excel(uploaded_file, sheet_name="Engineering Hours")
        
        df_tester['Date'] = pd.to_datetime(df_tester['Date'], errors='coerce')
        df_tester['Month'] = df_tester['Date'].dt.to_period('M').astype(str)
        df_eng['Date'] = pd.to_datetime(df_eng['Date'], errors='coerce')
        df_eng['Month'] = df_eng['Date'].dt.to_period('M').astype(str)
        
        df_tester = split_and_distribute(df_tester, 'Tester #', 'Tester Total Hours')
        df_tester = split_and_distribute(df_tester, 'TEMP', 'Tester Total Hours')
        df_tester = split_and_distribute(df_tester, 'Customer Requestor', 'Tester Total Hours')
        
        df_eng = split_and_distribute(df_eng, 'Tester', 'Engineering Support Hours')
        df_eng = split_and_distribute(df_eng, 'Customer Requestor', 'Engineering Support Hours')
        df_eng = split_and_distribute(df_eng, 'Name', 'Engineering Support Hours')

        all_req_tester = df_tester['Customer Requestor'].dropna().astype(str).unique().tolist()
        all_req_eng = df_eng['Customer Requestor'].dropna().astype(str).unique().tolist()
        all_requestors = sorted(list(set(all_req_tester + all_req_eng)))
        
        if "cso_selection" not in st.session_state:
            st.session_state.cso_selection = [n for n in all_requestors if "Alec" in n]
        if "gchip_selection" not in st.session_state:
            st.session_state.gchip_selection = [n for n in all_requestors if n in ["Rajesh", "Louis", "Chi-Chang"]]
            
        avail_for_cso = [n for n in all_requestors if n not in st.session_state.gchip_selection]
        avail_for_gchip = [n for n in all_requestors if n not in st.session_state.cso_selection]
        
        with st.sidebar:
            st.subheader(_t("team_def"))
            cso_team = st.multiselect(_t("cso_members"), avail_for_cso, key="cso_selection")
            gchip_team = st.multiselect(_t("gchip_members"), avail_for_gchip, key="gchip_selection")
            
        df_tester['Team'] = df_tester['Customer Requestor'].apply(lambda x: 'CSO' if x in cso_team else ('Gchip' if x in gchip_team else 'Other'))
        df_eng['Team'] = df_eng['Customer Requestor'].apply(lambda x: 'CSO' if x in cso_team else ('Gchip' if x in gchip_team else 'Other'))

        actual_tester = df_tester['Tester Total Hours'].sum()
        actual_eng = df_eng['Engineering Support Hours'].sum()
        target_tester = get_target_hours(df_tester, machine_count)
        diff_target = actual_tester - target_tester
        
        kpi1, kpi2, kpi3 = st.columns(3)
        kpi1.metric(_t("total_tester_hrs"), f"{actual_tester:,.2f} hrs")
        kpi2.metric(_t("target_hrs"), f"{target_tester:,.2f} hrs", f"{diff_target:,.2f} hrs", delta_color="normal")
        kpi3.metric(_t("total_eng_hrs"), f"{actual_eng:,.2f} hrs")
        st.markdown("<br>", unsafe_allow_html=True)

        selected_tab = st.radio("", [
            _t("tab_team"), 
            _t("tab_monthly"), 
            _t("tab_adv"), 
            _t("tab_req")
        ], horizontal=True, label_visibility="collapsed")
        st.markdown("---")

        if selected_tab == _t("tab_team"):
            render_table_and_chart(df_tester, 'Team', 'Tester Total Hours', 
                                  "ui_team_tester", "[Tester Hours] by Team", "🟦", palette=["#1F449C", "#F05A28", "#A0A0A0"])
            render_table_and_chart(df_eng, 'Team', 'Engineering Support Hours', 
                                  "ui_team_eng", "[Engineering Hours] by Team", "🟧", palette=["#1F449C", "#F05A28", "#A0A0A0"])

        elif selected_tab == _t("tab_monthly"):
            render_table_and_chart(df_tester, 'Month', 'Tester Total Hours', 
                                  "ui_monthly_tester", "[Tester Hours] Total by Month", "🟦", show_breakdown=True)
            
            df_eng['Month_Tester'] = df_eng['Month'] + " - " + df_eng['Tester'].astype(str)
            render_table_and_chart(df_eng, 'Month_Tester', 'Engineering Support Hours', 
                                  "ui_monthly_eng", "[Engineering Hours] by Month & Tester", "🟧", show_breakdown=True)

        elif selected_tab == _t("tab_adv"):
            render_table_and_chart(df_tester, 'TEMP', 'Tester Total Hours', 
                                  "ui_adv_tester", "[Tester Hours] by TEMP", "🟦", show_breakdown=True)
            render_table_and_chart(df_eng, 'Month', 'Engineering Support Hours', 
                                  "ui_adv_eng", "[Engineering Hours] Total by Month", "🟧", palette="Oranges_r", show_breakdown=True)

        elif selected_tab == _t("tab_req"):
            render_table_and_chart(df_tester, 'Customer Requestor', 'Tester Total Hours', 
                                  "ui_req_tester", "[Tester Hours] by Customer Requestor", "🟦", show_breakdown=False)
            render_table_and_chart(df_eng, 'Customer Requestor', 'Engineering Support Hours', 
                                  "ui_req_eng", "[Engineering Hours] by Customer Requestor", "🟧", palette="Oranges_r", show_breakdown=False)

    except ValueError as ve:
        st.error(f"{_t('error_sheet')}\n\n詳細錯誤：{ve}")
    except Exception as e:
        st.error(f"發生未知的錯誤：{e}")

# ==========================================
# 4. 版本更新紀錄
# ==========================================
st.markdown("<br><br>", unsafe_allow_html=True)
with st.expander(_t("release_notes")):
    st.markdown("""
    * **V27**: 修復介面排版，將多選篩選器重新收納至左側數據表格的正上方，還原最直覺的資料檢視動線。
    * **V26**: 加入了雙語介面切換 (支援中文與英文)，保證底層邏輯與原始版本完全一致，僅翻譯文字。
    * **V25**: 加入時數明細展開功能，升級表格 UI，使用 `st.column_config`。
    * **V24**: 將介面設計翻新為專業商務風 (Professional Corporate Theme)，提升易讀性。
    * **V23**: 加入側邊欄動態 KPI 目標機台數設定，能依據天數自動計算稼動率目標。
    * **V22**: 修復了團隊成員設定互斥的問題，導入 `st.session_state` 防止人員重複歸屬。
    * **V21**: 加入跨團隊的時數拆解視窗，主管能直接從表格查看 CSO 與 Gchip 的佔比。
    * **V20**: 導入原生 `st.radio` 的水平選單，徹底解決 CSS 與深淺色模式的白底衝突。
    * **V18**: 實作了側邊欄收納設定區，以及使用 Tabs 進行畫面切換，大幅減少頁面滾動。
    * **V15**: 新增文字聚合功能，能將任務說明合併去重，並畫出明顯的分隔線。
    * **V1-V14**: 完成核心資料讀取、切割平分演算法 (`/`, `,`, `;`, `\n`)、圖表 100% 英文防亂碼設定。
    """)
