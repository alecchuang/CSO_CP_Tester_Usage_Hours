import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import re

# ==========================================
# 共用函數區
# ==========================================
def split_and_distribute(df, target_col, hours_col):
    df = df.copy()
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

# ==========================================
# 🌟 聚合函數 (支援時數結構展開)
# ==========================================
def aggregate_data(df, group_by_cols, hours_col, show_breakdown=True):
    if isinstance(group_by_cols, str): group_by_cols = [group_by_cols]
    
    breakdown_col = f"⏱️ {hours_col} (Open)"
    
    # 決定要回傳哪些欄位
    columns_to_return = group_by_cols + [hours_col, 'Task Details']
    if show_breakdown:
        columns_to_return.insert(-1, breakdown_col)
        
    if df.empty: 
        return pd.DataFrame(columns=columns_to_return)
        
    # 1. 取得基本加總時數 
    res_hours = df.groupby(group_by_cols)[hours_col].sum().reset_index()
    res = res_hours
    
    # 2. 如果允許展開，則產生時數分配明細字串
    if show_breakdown:
        def format_hours(g):
            total = g[hours_col].sum()
            if 'Team' not in g.columns:
                return f"總計: {total:.2f} hrs"
                
            cso_hrs = g[g['Team'] == 'CSO'][hours_col].sum()
            gchip_hrs = g[g['Team'] == 'Gchip'][hours_col].sum()
            other_hrs = g[g['Team'] == 'Other / Unassigned'][hours_col].sum()
            
            details = [f"總計: {total:.2f} hrs"]
            if cso_hrs > 0: details.append(f"  ├ CSO: {cso_hrs:.2f} hrs")
            if gchip_hrs > 0: details.append(f"  ├ Gchip: {gchip_hrs:.2f} hrs")
            if other_hrs > 0: details.append(f"  └ Other: {other_hrs:.2f} hrs")
            
            return '\n'.join(details)
            
        res_breakdown = df.groupby(group_by_cols).apply(format_hours).reset_index(name=breakdown_col)
        res = pd.merge(res_hours, res_breakdown, on=group_by_cols)

    # 3. 任務文字說明聚合
    def format_details(g):
        details_str = []
        if 'Team' not in g.columns:
            tasks = g['Task Details'].unique()
            valid_items = [str(i).strip() for i in tasks if pd.notna(i) and str(i).strip() != '']
            return '\n'.join([f"  • {item}" for item in valid_items]) if valid_items else "無詳細說明 (N/A)"
        
        for team_name in ['CSO', 'Gchip', 'Other / Unassigned']:
            team_rows = g[g['Team'] == team_name]
            if team_rows.empty: continue
            tasks = team_rows['Task Details'].unique()
            valid_items = [str(i).strip() for i in tasks if pd.notna(i) and str(i).strip() != '']
            if valid_items:
                team_str = f"🏢 【{team_name}】 任務明細：\n" + '\n'.join([f"  • {item}" for item in valid_items])
                details_str.append(team_str)
        return '\n\n----------------------------------------\n\n'.join(details_str) if details_str else "無詳細說明 (N/A)"
        
    res_details = df.groupby(group_by_cols).apply(format_details).reset_index(name='Task Details')
    
    # 將資料合併並排序
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

# --- 主畫面標題 ---
st.title("📊 機台與工程師時數進階分析儀表板")

with st.expander("🚀 版本更新紀錄 / Release Notes (點擊展開)"):
    st.markdown("""
    * **v25 (最新版)**: 🌟 **標題與維度微調**！將「進階維度分析」重新命名為更貼切的「進階維度分析 (TEMP/ENG Member)」，並為「[Engineering Hours] 依機台統計」補上月份 (Month) 維度，讓分析更具時間脈絡。
    * **v24**: 🌟 排版邏輯調整！將「每月工程師時數」移至「進階維度分析」，並將「依機台統計」移至「每月趨勢分析」。
    * **v23**: 🌟 動態 KPI 目標設定！最低標時數的「機台數量」設定移至左側，可自行調整並即時同步。
    * **v22**: 🌟 優化明細展開範圍！專注於「每月趨勢」與「進階維度」，移除不必要的頁面展開。
    * **v21**: 🌟 時數結構展開功能！總時數欄位升級為「點擊展開」格式，可直接檢視 CSO 與 Gchip 貢獻明細。
    * **v20**: 🌟 無縫導覽與KPI升級！改用原生水平導覽列，新增動態計算的「最低標使用時數」指標。
    * **v19**: 將超大頁籤底色更改為專業的淺灰藍色。
    * **v18**: 頁籤視覺強化與修復。
    * **v17**: UX 介面大改版！導入側邊欄收納設定、主畫面頂部加入 KPI 看板。
    * **v16**: 🛡️ 系統穩定度升級！修復 aggregate_data KeyError。
    * **v15**: 📝 任務明細自動使用分隔線將 CSO 與 Gchip 拆分顯示。
    * **v14**: 新增「📋 任務說明」展開查詢功能。
    * **v13**: 視覺風格優化！轉換為專業風格 (Professional Corporate Theme)。
    * **v12**: 導入深色科技感主題 (Dark Tech Theme)。
    * **v11**: 新增版本紀錄摺疊面板。
    * **v10**: 加入「團隊成員自定義」功能。
    * **v9**: 加入「動態篩選器 (Multiselect)」。
    * **v8**: 解決中文亂碼問題，圖表純英文。
    * **v7**: 介面大改版，左表格、右圖表。
    * **v6**: 導入「多單位分割與時數均分邏輯」。
    * **v5**: 擴充分析維度。
    * **v4**: 加入檔案上傳功能。
    * **v3**: 轉換為 Streamlit Web App。
    * **v2**: 加入 Engineering Hours 分頁。
    * **v1**: 初始版本。
    """)

# ==========================================
# 👈 左側邊欄 (Sidebar)
# ==========================================
with st.sidebar:
    st.header("⚙️ 控制面板 (Control Panel)")
    
    st.info("""
    **💡 操作指南：**
    1. 於下方上傳 Excel 檔案。
    2. 設定 KPI 的機台總數與團隊成員。
    3. 在右側主畫面選擇不同分析維度。
    """)
    
    uploaded_file = st.file_uploader("📂 上傳 Excel 紀錄表", type=["xlsx", "xls"])
    
    st.divider()
    st.subheader("🎯 KPI 目標設定")
    tester_count = st.number_input("設定機台總數量 (供計算最低標時數)", min_value=1, value=10, step=1)

if uploaded_file is not None:
    try:
        # --- 資料預處理 ---
        target_detail_col = 'Lot #wafer / Purpose /Description'
        
        df_tester_raw = pd.read_excel(uploaded_file, sheet_name="Tester Hours", skiprows=3)
        df_tester = df_tester_raw[['Date', 'Tester #', 'Tester Total Hours', 'TEMP', 'Customer Requestor', target_detail_col]].copy()
        df_tester.rename(columns={target_detail_col: 'Task Details'}, inplace=True)
        df_tester.dropna(subset=['Date', 'Tester #', 'Tester Total Hours'], how='all', inplace=True)
        df_tester['Date'] = pd.to_datetime(df_tester['Date'], errors='coerce')
        df_tester.dropna(subset=['Date'], inplace=True)
        df_tester['Month'] = df_tester['Date'].dt.to_period('M').astype(str)
        df_tester['Tester Total Hours'] = pd.to_numeric(df_tester['Tester Total Hours'], errors='coerce').fillna(0)

        df_eng_raw = pd.read_excel(uploaded_file, sheet_name="Engineering Hours")
        df_eng = df_eng_raw[['Date', 'Name', 'Engineering Support Hours', 'Tester', 'Customer Requestor', target_detail_col]].copy()
        df_eng.rename(columns={target_detail_col: 'Task Details'}, inplace=True)
        df_eng.dropna(subset=['Date', 'Name', 'Engineering Support Hours'], how='all', inplace=True)
        df_eng['Date'] = pd.to_datetime(df_eng['Date'], errors='coerce')
        df_eng.dropna(subset=['Date'], inplace=True)
        df_eng['Month'] = df_eng['Date'].dt.to_period('M').astype(str)
        df_eng['Engineering Support Hours'] = pd.to_numeric(df_eng['Engineering Support Hours'], errors='coerce').fillna(0)

        for col in ['Tester #', 'TEMP', 'Customer Requestor']:
            df_tester = split_and_distribute(df_tester, target_col=col, hours_col='Tester Total Hours')
        for col in ['Name', 'Tester', 'Customer Requestor']:
            df_eng = split_and_distribute(df_eng, target_col=col, hours_col='Engineering Support Hours')

        # --- 側邊欄：團隊成員自定義區塊 ---
        with st.sidebar:
            st.divider()
            st.subheader("👥 團隊成員定義")
            all_requestors = sorted(list(set(df_tester['Customer Requestor'].unique()) | set(df_eng['Customer Requestor'].unique())))
            
            if 'cso_selection' not in st.session_state:
                st.session_state.cso_selection = [n for n in ['Alec'] if n in all_requestors]
            if 'gchip_selection' not in st.session_state:
                st.session_state.gchip_selection = [n for n in ['Rajesh', 'Louis', 'Chi-Chang'] if n in all_requestors]

            avail_for_cso = [x for x in all_requestors if x not in st.session_state.gchip_selection]
            avail_for_gchip = [x for x in all_requestors if x not in st.session_state.cso_selection]

            cso_members = st.multiselect("定義 CSO 成員", options=avail_for_cso, key="cso_selection")
            gchip_members = st.multiselect("定義 Gchip 成員", options=avail_for_gchip, key="gchip_selection")

            def map_team(name):
                if name in cso_members: return 'CSO'
                elif name in gchip_members: return 'Gchip'
                else: return 'Other / Unassigned'

            df_tester['Team'] = df_tester['Customer Requestor'].apply(map_team)
            df_eng['Team'] = df_eng['Customer Requestor'].apply(map_team)

        # ==========================================
        # 📈 主畫面：頂部 KPI 總覽看板
        # ==========================================
        st.subheader("📌 核心數據總覽 (Executive Summary)")
        
        total_tester_hrs = df_tester['Tester Total Hours'].sum()
        
        unique_months = df_tester['Month'].dropna().unique()
        total_days = 0
        for m in unique_months:
            try:
                total_days += pd.Period(m).days_in_month
            except:
                pass
        
        if total_days == 0: total_days = 30
            
        target_utilization = 0.5
        min_required_hours = total_days * 24 * tester_count * target_utilization
        delta_val = total_tester_hrs - min_required_hours
        
        total_eng_hrs = df_eng['Engineering Support Hours'].sum()
        top_tester = df_tester.groupby('Tester #')['Tester Total Hours'].sum().idxmax() if not df_tester.empty else "N/A"
        
        kpi1, kpi2, kpi3, kpi4 = st.columns(4)
        kpi1.metric(label="🖥️ 總機台使用時數", value=f"{total_tester_hrs:,.1f} hrs")
        kpi2.metric(
            label=f"🎯 最低標使用時數 ({tester_count}台/50%)", 
            value=f"{min_required_hours:,.0f} hrs", 
            delta=f"{delta_val:,.1f} hrs", 
            delta_color="normal"
        )
        kpi3.metric(label="🧑‍🔧 總工程支援時數", value=f"{total_eng_hrs:,.1f} hrs")
        kpi4.metric(label="🔥 最高用量機台", value=f"{top_tester}")
        
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
                if filter_col:
                    unique_items = sorted(df[filter_col].unique().tolist())
                    selected_items = st.multiselect(f"🔽 篩選 {filter_col}", options=unique_items, default=unique_items, key=f"filter_{chart_title}")
                    filtered_df = df[df[filter_col].isin(selected_items)]
                
                # 動態配置表格欄位屬性
                column_config = {
                    "Task Details": st.column_config.TextColumn(
                        "📋 任務說明 (點擊展開)", 
                        help="點擊儲存格，即可查看區分 CSO 與 Gchip 的完整工作內容",
                        width="medium"
                    )
                }
                
                # 只有當 show_breakdown 為 True 時，才隱藏原數字欄並顯示文字明細欄
                if show_breakdown:
                    breakdown_col_name = f"⏱️ {y_col} (點擊展開)"
                    column_config[y_col] = None  # 隱藏原數字欄
                    column_config[breakdown_col_name] = st.column_config.TextColumn(
                        breakdown_col_name,
                        help="點擊儲存格，即可查看該時數的 CSO / Gchip 貢獻拆分",
                        width="medium"
                    )
                
                st.dataframe(
                    filtered_df, 
                    use_container_width=True, 
                    hide_index=True,  
                    column_config=column_config
                )
                
            with col_chart:
                if filtered_df.empty: st.warning("無資料可顯示。")
                else:
                    fig, ax = plt.subplots(figsize=(10, 4.5))
                    if hue_col:
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
        # 📑 分析視角切換導覽列
        # ==========================================
        st.markdown("### 🔍 切換分析視角")
        
        # 🌟 變更選項名稱
        selected_view = st.radio(
            label="選擇分析維度",
            options=[
                "🏢 團隊歸屬分析 (Team)", 
                "📅 每月趨勢分析 (Monthly)", 
                "🌡️ 進階維度分析 (TEMP/ENG Member)", 
                "👤 客戶需求者分析 (Requestor)"
            ],
            horizontal=True,
            label_visibility="collapsed" 
        )
        
        st.markdown("<br>", unsafe_allow_html=True)

        if selected_view == "🏢 團隊歸屬分析 (Team)":
            team_tester_hours = aggregate_data(df_tester, 'Team', 'Tester Total Hours', show_breakdown=False)
            team_eng_hours = aggregate_data(df_eng, 'Team', 'Engineering Support Hours', show_breakdown=False)
            render_table_and_chart("🟦 [Tester Hours] 依團隊統計", "[Tester Hours] Total by Team", team_tester_hours, 'Team', 'Tester Total Hours', filter_col='Team', custom_palette=['#2B5B84', '#E67E22', '#95A5A6'], show_breakdown=False)
            render_table_and_chart("🟧 [Engineering Hours] 依團隊統計", "[Engineering Hours] Total by Team", team_eng_hours, 'Team', 'Engineering Support Hours', filter_col='Team', custom_palette=['#2980B9', '#D35400', '#7F8C8D'], show_breakdown=False)

        elif selected_view == "📅 每月趨勢分析 (Monthly)":
            # [Tester Hours] 每月機台時數
            monthly_tester_hours = aggregate_data(df_tester, ['Month', 'Tester #'], 'Tester Total Hours')
            render_table_and_chart("🟦 [Tester Hours] 每月機台時數", "[Tester Hours] Monthly by Tester", monthly_tester_hours, 'Month', 'Tester Total Hours', hue_col='Tester #', filter_col='Tester #', custom_palette='deep')
            
            # 🌟 修改：加入 Month 欄位進行雙維度聚合
            eng_tester_hours = aggregate_data(df_eng, ['Month', 'Tester'], 'Engineering Support Hours')
            render_table_and_chart("🟧 [Engineering Hours] 每月依機台 (Tester) 統計", "[Engineering Hours] Monthly by Tester", eng_tester_hours, 'Month', 'Engineering Support Hours', hue_col='Tester', filter_col='Tester', custom_palette='Oranges_r')

        elif selected_view == "🌡️ 進階維度分析 (TEMP/ENG Member)":
            # [Tester Hours] 依溫度 (TEMP) 統計
            temp_hours = aggregate_data(df_tester, 'TEMP', 'Tester Total Hours')
            render_table_and_chart("🟦 [Tester Hours] 依溫度 (TEMP) 統計", "[Tester Hours] Total by TEMP", temp_hours, 'TEMP', 'Tester Total Hours', filter_col='TEMP', custom_palette='Blues_r')
            
            # [Engineering Hours] 每月工程師時數
            monthly_eng_hours = aggregate_data(df_eng, ['Month', 'Name'], 'Engineering Support Hours')
            render_table_and_chart("🟧 [Engineering Hours] 每月工程師時數", "[Engineering Hours] Monthly by Engineer", monthly_eng_hours, 'Month', 'Engineering Support Hours', hue_col='Name', filter_col='Name', custom_palette='muted')

        elif selected_view == "👤 客戶需求者分析 (Requestor)":
            tester_req_hours = aggregate_data(df_tester, 'Customer Requestor', 'Tester Total Hours', show_breakdown=False)
            eng_req_hours = aggregate_data(df_eng, 'Customer Requestor', 'Engineering Support Hours', show_breakdown=False)
            render_table_and_chart("🟦 [Tester Hours] 依客戶統計", "[Tester Hours] Total by Requestor", tester_req_hours, 'Customer Requestor', 'Tester Total Hours', filter_col='Customer Requestor', custom_palette='Set2', show_breakdown=False)
            render_table_and_chart("🟧 [Engineering Hours] 依客戶統計", "[Engineering Hours] Total by Requestor", eng_req_hours, 'Customer Requestor', 'Engineering Support Hours', filter_col='Customer Requestor', custom_palette='Set1', show_breakdown=False)

    except Exception as e:
        st.error(f"執行時發生錯誤: {e}")

else:
    st.info("👈 請於左側邊欄 (Sidebar) 上傳 Excel 檔案以開始分析。")
