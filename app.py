import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import re

# ==========================================
# 定義：時數分割與均分函數
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
# 定義：資料與任務說明聚合函數 (New Feature!)
# ==========================================
def aggregate_data(df, group_by_cols, hours_col):
    """將時數進行加總，並將對應的 Description 文字合併去重"""
    if isinstance(group_by_cols, str):
        group_by_cols = [group_by_cols]
        
    def join_unique_details(x):
        # 過濾空值並去除前後空白，避免重複
        valid_items = [str(i).strip() for i in x.unique() if pd.notna(i) and str(i).strip() != '']
        if not valid_items:
            return "無詳細說明 (N/A)"
        return '\n'.join([f"• {item}" for item in valid_items])
        
    agg_dict = {
        hours_col: 'sum',
        'Task Details': join_unique_details  # 聚合任務說明
    }
    
    res = df.groupby(group_by_cols).agg(agg_dict).reset_index()
    res[hours_col] = res[hours_col].round(2)
    
    # 若不是以月份排序，則以時數高低排序
    if 'Month' not in group_by_cols:
        res = res.sort_values(hours_col, ascending=False)
        
    return res

# ==========================================
# 網頁主程式開始
# ==========================================
st.set_page_config(page_title="Hours Analysis Dashboard", layout="wide")

st.title("📊 機台與工程師時數進階分析儀表板")

# --- 版本說明 (Release Notes) ---
with st.expander("🚀 版本更新紀錄 / Release Notes (點擊展開)"):
    st.markdown("""
    * **v14 (最新版)**: 📝 **新增「任務明細查詢」功能**！在所有數據表格的右側加入「📋 任務說明」欄位，點擊即可展開查看該列數據對應的所有原始工作內容 (Lot / Purpose / Description)。
    * **v13**: 🏢 視覺風格優化！移除高對比科技風，轉換為乾淨、明亮且具備商務質感的專業風格 (Professional Corporate Theme)。
    * **v12**: 導入深色科技感主題 (Dark Tech Theme)，加入高對比螢光配色與極簡網格。
    * **v11**: 新增版本紀錄摺疊面板，優化 UI 引導說明。
    * **v10**: 加入「團隊成員自定義」功能，支援 CSO/Gchip 互斥選擇與預設人員自動偵測。
    * **v9**: 在每個數據表格上方加入「動態篩選器 (Multiselect)」，圖表隨篩選結果即時連動。
    * **v8**: 解決中文亂碼問題，將圖表內部文字（標題、軸線）統一轉換為純英文顯示。
    * **v7**: 介面大改版，採用「左表格、右圖表」的並排設計，提升數據核對效率。
    * **v6**: 核心演算法更新，導入「多單位分割與時數均分邏輯 (處理 / , ; \\n 等符號)」。
    * **v5**: 擴充分析維度，加入 TEMP、Customer Requestor、Tester 等進階統計。
    * **v4**: 加入檔案上傳功能 (File Uploader)，支援使用者自行上傳 Excel 進行分析。
    * **v3**: 轉換為 Streamlit Web App 互動式架構。
    * **v2**: 加入 Engineering Hours 分頁數據解析。
    * **v1**: 初始版本，解析 Tester Hours 並產生基礎月度統計長條圖。
    """)

st.info("""
**💡 操作指南 (Quick Guide)：**
1. **上傳檔案**：點擊下方按鈕上傳 Excel 檔案。
2. **定義團隊**：在下方「團隊成員定義」區塊，將人員分配至 **CSO** 或 **Gchip**。
3. **查看工作內容**：點擊表格內 **「📋 任務說明」** 的儲存格，即可展開閱讀完整的工作細節。
4. **篩選數據**：在表格上方點擊 **「x」** 排除不需要的項目，右側圖表會同步更新。
""")

uploaded_file = st.file_uploader("請上傳您的 Excel 時數紀錄表", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        df_tester_raw = pd.read_excel(uploaded_file, sheet_name="Tester Hours", skiprows=3)
        df_eng_raw = pd.read_excel(uploaded_file, sheet_name="Engineering Hours")

        # --- 資料預處理 ---
        # 確保擷取 'Lot #wafer / Purpose /Description' 欄位
        target_detail_col = 'Lot #wafer / Purpose /Description'
        
        df_tester = df_tester_raw[['Date', 'Tester #', 'Tester Total Hours', 'TEMP', 'Customer Requestor', target_detail_col]].copy()
        df_tester.rename(columns={target_detail_col: 'Task Details'}, inplace=True)
        df_tester.dropna(subset=['Date', 'Tester #', 'Tester Total Hours'], how='all', inplace=True)
        df_tester['Date'] = pd.to_datetime(df_tester['Date'], errors='coerce')
        df_tester.dropna(subset=['Date'], inplace=True)
        df_tester['Month'] = df_tester['Date'].dt.to_period('M').astype(str)
        df_tester['Tester Total Hours'] = pd.to_numeric(df_tester['Tester Total Hours'], errors='coerce').fillna(0)

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

        st.divider()

        # ==========================================
        # 團隊成員自定義區塊
        # ==========================================
        st.subheader("⚙️ 團隊成員定義 / Team Definitions")
        all_requestors = sorted(list(set(df_tester['Customer Requestor'].unique()) | set(df_eng['Customer Requestor'].unique())))
        
        if 'cso_selection' not in st.session_state:
            st.session_state.cso_selection = [n for n in ['Alec'] if n in all_requestors]
        if 'gchip_selection' not in st.session_state:
            st.session_state.gchip_selection = [n for n in ['Rajesh', 'Louis', 'Chi-Chang'] if n in all_requestors]

        avail_for_cso = [x for x in all_requestors if x not in st.session_state.gchip_selection]
        avail_for_gchip = [x for x in all_requestors if x not in st.session_state.cso_selection]

        col_cso_config, col_gchip_config = st.columns(2)
        with col_cso_config:
            cso_members = st.multiselect("定義 CSO 成員 (CSO Members)", options=avail_for_cso, key="cso_selection")
        with col_gchip_config:
            gchip_members = st.multiselect("定義 Gchip 成員 (Gchip Members)", options=avail_for_gchip, key="gchip_selection")

        def map_team(name):
            if name in cso_members: return 'CSO'
            elif name in gchip_members: return 'Gchip'
            else: return 'Other / Unassigned'

        df_tester['Team'] = df_tester['Customer Requestor'].apply(map_team)
        df_eng['Team'] = df_eng['Customer Requestor'].apply(map_team)

        # ==========================================
        # 🏢 專業商務圖表主題設定 (Professional Corporate Theme)
        # ==========================================
        plt.style.use('default')
        corporate_params = {
            "font.sans-serif": ["Microsoft JhengHei", "PingFang TC", "Arial Unicode MS", "SimHei", "sans-serif"],
            "axes.unicode_minus": False,
            "figure.facecolor": "#FFFFFF",
            "axes.facecolor": "#F8F9FA",
            "grid.color": "#DEE2E6",
            "grid.linestyle": "-",
            "grid.alpha": 0.8,
            "text.color": "#212529",
            "axes.labelcolor": "#495057",
            "xtick.color": "#6C757D",
            "ytick.color": "#6C757D",
        }
        sns.set_theme(style="whitegrid", rc=corporate_params)

        def render_table_and_chart(ui_title, chart_title, df, x_col, y_col, hue_col=None, filter_col=None, custom_palette=None):
            st.markdown(f"#### {ui_title}")
            col_data, col_chart = st.columns([1, 2])
            with col_data:
                if filter_col:
                    unique_items = sorted(df[filter_col].unique().tolist())
                    selected_items = st.multiselect(
                        f"🔽 Filter {filter_col}", options=unique_items, default=unique_items, key=f"filter_{chart_title}"
                    )
                    filtered_df = df[df[filter_col].isin(selected_items)]
                else:
                    filtered_df = df
                
                # --- ✨ 表格設定優化：加入互動式 Task Details 欄位 ---
                st.dataframe(
                    filtered_df, 
                    use_container_width=True,
                    hide_index=True,  # 隱藏醜醜的數字編號
                    column_config={
                        "Task Details": st.column_config.TextColumn(
                            "📋 任務說明 (點擊展開)",
                            help="點擊此欄位的儲存格，即可查看該項目的完整工作內容",
                            width="medium" # 設定較寬的欄位
                        )
                    }
                )
            
            with col_chart:
                if filtered_df.empty:
                    st.warning("No data to display.")
                else:
                    fig, ax = plt.subplots(figsize=(10, 4.5))
                    
                    if hue_col:
                        sns.barplot(data=filtered_df, x=x_col, y=y_col, hue=hue_col, ax=ax, palette=custom_palette, edgecolor="#FFFFFF", linewidth=1.2)
                        legend = ax.legend(title=hue_col, bbox_to_anchor=(1.05, 1), loc='upper left', frameon=False)
                        plt.setp(legend.get_texts(), color='#495057')
                        plt.setp(legend.get_title(), color='#212529', fontweight='bold')
                    else:
                        sns.barplot(data=filtered_df, x=x_col, y=y_col, ax=ax, palette=custom_palette, edgecolor="#FFFFFF", linewidth=1.2)
                    
                    ax.spines['top'].set_visible(False)
                    ax.spines['right'].set_visible(False)
                    ax.spines['left'].set_color('#CED4DA')
                    ax.spines['bottom'].set_color('#CED4DA')

                    ax.set_title(chart_title, fontweight='bold', pad=15, color='#212529')
                    ax.set_xlabel(x_col, labelpad=10)
                    ax.set_ylabel(y_col, labelpad=10)
                    plt.xticks(rotation=45, ha='right')
                    plt.tight_layout()
                    st.pyplot(fig)
            st.divider()

        # ==========================================
        # 繪製圖表 (改用強化的聚合函數 aggregate_data)
        # ==========================================
        
        st.subheader("🏢 團隊歸屬分析 / Team Analysis")
        team_tester_hours = aggregate_data(df_tester, 'Team', 'Tester Total Hours')
        team_eng_hours = aggregate_data(df_eng, 'Team', 'Engineering Support Hours')
        render_table_and_chart("🟦 [Tester Hours] 依團隊統計", "[Tester Hours] Total by Team", team_tester_hours, 'Team', 'Tester Total Hours', filter_col='Team', custom_palette=['#2B5B84', '#E67E22', '#95A5A6'])
        render_table_and_chart("🟧 [Engineering Hours] 依團隊統計", "[Engineering Hours] Total by Team", team_eng_hours, 'Team', 'Engineering Support Hours', filter_col='Team', custom_palette=['#2980B9', '#D35400', '#7F8C8D'])

        st.subheader("📅 每月趨勢分析 / Monthly Trends")
        monthly_tester_hours = aggregate_data(df_tester, ['Month', 'Tester #'], 'Tester Total Hours')
        monthly_eng_hours = aggregate_data(df_eng, ['Month', 'Name'], 'Engineering Support Hours')
        render_table_and_chart("🟦 [Tester Hours] 每月機台時數", "[Tester Hours] Monthly by Tester", monthly_tester_hours, 'Month', 'Tester Total Hours', hue_col='Tester #', filter_col='Tester #', custom_palette='deep')
        render_table_and_chart("🟧 [Engineering Hours] 每月工程師時數", "[Engineering Hours] Monthly by Engineer", monthly_eng_hours, 'Month', 'Engineering Support Hours', hue_col='Name', filter_col='Name', custom_palette='muted')

        st.subheader("🔍 進階維度分析 / Advanced Dimensions")
        temp_hours = aggregate_data(df_tester, 'TEMP', 'Tester Total Hours')
        eng_tester_hours = aggregate_data(df_eng, 'Tester', 'Engineering Support Hours')
        render_table_and_chart("🟦 [Tester Hours] 依溫度 (TEMP) 統計", "[Tester Hours] Total by TEMP", temp_hours, 'TEMP', 'Tester Total Hours', filter_col='TEMP', custom_palette='Blues_r')
        render_table_and_chart("🟧 [Engineering Hours] 依機台 (Tester) 統計", "[Engineering Hours] Total by Tester", eng_tester_hours, 'Tester', 'Engineering Support Hours', filter_col='Tester', custom_palette='Oranges_r')

        st.subheader("👤 客戶需求者分析 / Requestor Analysis")
        tester_req_hours = aggregate_data(df_tester, 'Customer Requestor', 'Tester Total Hours')
        eng_req_hours = aggregate_data(df_eng, 'Customer Requestor', 'Engineering Support Hours')
        render_table_and_chart("🟦 [Tester Hours] 依客戶統計", "[Tester Hours] Total by Requestor", tester_req_hours, 'Customer Requestor', 'Tester Total Hours', filter_col='Customer Requestor', custom_palette='Set2')
        render_table_and_chart("🟧 [Engineering Hours] 依客戶統計", "[Engineering Hours] Total by Requestor", eng_req_hours, 'Customer Requestor', 'Engineering Support Hours', filter_col='Customer Requestor', custom_palette='Set1')

    except Exception as e:
        st.error(f"執行時發生錯誤: {e}")

else:
    st.info("👈 請先上傳 Excel 檔案以開始分析。")
