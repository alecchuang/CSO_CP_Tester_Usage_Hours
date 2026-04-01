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

def aggregate_data(df, group_by_cols, hours_col):
    if isinstance(group_by_cols, str): group_by_cols = [group_by_cols]
    if df.empty: return pd.DataFrame(columns=group_by_cols + [hours_col, 'Task Details'])
    res_hours = df.groupby(group_by_cols)[hours_col].sum().reset_index()
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
    res = pd.merge(res_hours, res_details, on=group_by_cols)
    res[hours_col] = res[hours_col].round(2)
    if 'Month' not in group_by_cols: res = res.sort_values(hours_col, ascending=False)
    return res

# ==========================================
# 網頁主程式開始
# ==========================================
st.set_page_config(page_title="Hours Analysis Dashboard", layout="wide")

# ==========================================
# 🎨 注入自訂 CSS 來強化 Tabs (頁籤) 的視覺效果
# ==========================================
# 更改為更專業的顏色：淺灰藍色 (#eef2f5) 取代白色 (#ffffff)
st.markdown("""
<style>
    /* 讓整個 Tab 列表區塊有背景色和圓角 */
    div[data-baseweb="tab-list"] {
        gap: 8px;
        background-color: #f1f3f5;
        padding: 10px 10px 0 10px;
        border-radius: 12px 12px 0 0;
        border-bottom: 3px solid #dee2e6;
    }
    /* 強化每個獨立 Tab 的樣式 - 替換為專業的淺灰藍色 */
    div[data-baseweb="tab"] {
        height: 55px;
        padding: 0 25px;
        background-color: #eef2f5; 
        border: 1px solid #dee2e6;
        border-bottom: none;
        border-radius: 10px 10px 0 0;
        font-size: 18px !important;
        font-weight: bold;
        color: #495057;
        transition: all 0.2s ease-in-out;
    }
    /* 游標移上去時的特效 */
    div[data-baseweb="tab"]:hover {
        background-color: #dbe4eb;
        color: #000000;
    }
    /* 當 Tab 被選中時的高亮特效 (商務藍) */
    div[data-baseweb="tab"][aria-selected="true"] {
        background-color: #0056b3;
        color: #ffffff !important;
        border-color: #0056b3;
    }
    /* 隱藏 Streamlit 預設的細細底線 */
    div[data-baseweb="tab-highlight"] {
        display: none;
    }
</style>
""", unsafe_allow_html=True)

# --- 主畫面標題 ---
st.title("📊 機台與工程師時數進階分析儀表板")

with st.expander("🚀 版本更新紀錄 / Release Notes (點擊展開)"):
    st.markdown("""
    * **v19 (最新版)**: 🌟 **介面與數據優化**！將超大頁籤底色更改為專業的淺灰藍色；並在核心數據總覽新增「最低標使用時數」指標 (機台數預設 10 台，每月基本要求 50% 稼動率)。
    * **v18**: 🌟 頁籤視覺強化與修復！透過自訂 CSS 大幅突顯 Tab 選擇區塊，增強點擊引導；並完整修復、保留 v1 至 v18 的完整版本更新紀錄。
    * **v17**: 🌟 UX 介面大改版！導入側邊欄 (Sidebar) 收納設定、主畫面頂部加入 KPI 數據看板，並使用「頁籤 (Tabs)」分類圖表。
    * **v16**: 🛡️ 系統穩定度升級！修復 aggregate_data KeyError，強化空資料防呆機制。
    * **v15**: 📝 任務明細自動使用分隔線將 CSO 與 Gchip 拆分顯示，權責更清晰。
    * **v14**: 新增「📋 任務說明」展開查詢功能，可直接在表格內查看原始工作內容。
    * **v13**: 視覺風格優化！轉換為乾淨、明亮且具備商務質感的專業風格 (Professional Corporate Theme)。
    * **v12**: 導入深色科技感主題 (Dark Tech Theme)。
    * **v11**: 新增版本紀錄摺疊面板，優化 UI 引導說明。
    * **v10**: 加入「團隊成員自定義」功能，支援 CSO/Gchip 互斥選擇與預設人員自動偵測。
    * **v9**: 加入「動態篩選器 (Multiselect)」，圖表隨篩選結果即時連動。
    * **v8**: 解決中文亂碼問題，圖表內部文字統一純英文。
    * **v7**: 介面大改版，採用「左表格、右圖表」並排設計。
    * **v6**: 核心演算法更新，導入「多單位分割與時數均分邏輯」。
    * **v5**: 擴充分析維度，加入 TEMP、Customer Requestor、Tester 等進階統計。
    * **v4**: 加入檔案上傳功能 (File Uploader)。
    * **v3**: 轉換為 Streamlit Web App 互動式架構。
    * **v2**: 加入 Engineering Hours 分頁數據解析。
    * **v1**: 初始版本，解析 Tester Hours 並產生基礎月度統計長條圖。
    """)

# ==========================================
# 👈 左側邊欄 (Sidebar)：收納設定與上傳
# ==========================================
with st.sidebar:
    st.header("⚙️ 控制面板 (Control Panel)")
    
    st.info("""
    **💡 操作指南：**
    1. 於下方上傳 Excel 檔案。
    2. 定義 CSO 與 Gchip 團隊成員。
    3. 在右側主畫面的**超大頁籤**切換不同分析視角。
    """)
    
    uploaded_file = st.file_uploader("📂 上傳 Excel 紀錄表", type=["xlsx", "xls"])

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
        
        # 1. 總機台使用時數
        total_tester_hrs = df_tester['Tester Total Hours'].sum()
        
        # 2. 最低標使用時數計算邏輯
        unique_months = df_tester['Month'].dropna().unique()
        total_days = 0
        for m in unique_months:
            try:
                # 取得該月份(例如 '2026-02')的總天數(28天)並累加
                total_days += pd.Period(m).days_in_month
            except:
                pass
        
        # 防呆機制：若抓不到月份天數，則預設給 30 天
        if total_days == 0: total_days = 30
            
        tester_count = 10      # 暫訂 10 台機台
        target_utilization = 0.5 # 目標稼動率 50%
        # 公式: 總天數 * 24小時 * 10台 * 50%
        min_required_hours = total_days * 24 * tester_count * target_utilization
        
        # 計算與達標時數的差異
        delta_val = total_tester_hrs - min_required_hours
        
        # 3. 其他指標
        total_eng_hrs = df_eng['Engineering Support Hours'].sum()
        top_tester = df_tester.groupby('Tester #')['Tester Total Hours'].sum().idxmax() if not df_tester.empty else "N/A"
        
        # 呈現 4 個指標
        kpi1, kpi2, kpi3, kpi4 = st.columns(4)
        
        kpi1.metric(label="🖥️ 總機台使用時數", value=f"{total_tester_hrs:,.1f} hrs")
        
        # 使用 delta 屬性，當總機台使用時數 >= 最低標時，右邊會顯示綠色的正值；反之顯示紅色的負值
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

        def render_table_and_chart(ui_title, chart_title, df, x_col, y_col, hue_col=None, filter_col=None, custom_palette=None):
            st.markdown(f"#### {ui_title}")
            col_data, col_chart = st.columns([1, 2])
            with col_data:
                filtered_df = df
                if filter_col:
                    unique_items = sorted(df[filter_col].unique().tolist())
                    selected_items = st.multiselect(f"🔽 篩選 {filter_col}", options=unique_items, default=unique_items, key=f"filter_{chart_title}")
                    filtered_df = df[df[filter_col].isin(selected_items)]
                st.dataframe(
                    filtered_df, use_container_width=True, hide_index=True,  
                    column_config={"Task Details": st.column_config.TextColumn("📋 任務說明 (點擊展開)", width="medium")}
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

        # ==========================================
        # 📑 強化版頁籤切換區塊 (Tabs)
        # ==========================================
        st.markdown("<br>", unsafe_allow_html=True) 
        
        tab1, tab2, tab3, tab4 = st.tabs([
            "🏢 團隊歸屬分析", 
            "📅 每月趨勢分析", 
            "🔍 進階維度分析", 
            "👤 客戶需求者分析"
        ])

        with tab1:
            st.markdown("<br>", unsafe_allow_html=True)
            team_tester_hours = aggregate_data(df_tester, 'Team', 'Tester Total Hours')
            team_eng_hours = aggregate_data(df_eng, 'Team', 'Engineering Support Hours')
            render_table_and_chart("🟦 [Tester Hours] 依團隊統計", "[Tester Hours] Total by Team", team_tester_hours, 'Team', 'Tester Total Hours', filter_col='Team', custom_palette=['#2B5B84', '#E67E22', '#95A5A6'])
            st.divider()
            render_table_and_chart("🟧 [Engineering Hours] 依團隊統計", "[Engineering Hours] Total by Team", team_eng_hours, 'Team', 'Engineering Support Hours', filter_col='Team', custom_palette=['#2980B9', '#D35400', '#7F8C8D'])

        with tab2:
            st.markdown("<br>", unsafe_allow_html=True)
            monthly_tester_hours = aggregate_data(df_tester, ['Month', 'Tester #'], 'Tester Total Hours')
            monthly_eng_hours = aggregate_data(df_eng, ['Month', 'Name'], 'Engineering Support Hours')
            render_table_and_chart("🟦 [Tester Hours] 每月機台時數", "[Tester Hours] Monthly by Tester", monthly_tester_hours, 'Month', 'Tester Total Hours', hue_col='Tester #', filter_col='Tester #', custom_palette='deep')
            st.divider()
            render_table_and_chart("🟧 [Engineering Hours] 每月工程師時數", "[Engineering Hours] Monthly by Engineer", monthly_eng_hours, 'Month', 'Engineering Support Hours', hue_col='Name', filter_col='Name', custom_palette='muted')

        with tab3:
            st.markdown("<br>", unsafe_allow_html=True)
            temp_hours = aggregate_data(df_tester, 'TEMP', 'Tester Total Hours')
            eng_tester_hours = aggregate_data(df_eng, 'Tester', 'Engineering Support Hours')
            render_table_and_chart("🟦 [Tester Hours] 依溫度 (TEMP) 統計", "[Tester Hours] Total by TEMP", temp_hours, 'TEMP', 'Tester Total Hours', filter_col='TEMP', custom_palette='Blues_r')
            st.divider()
            render_table_and_chart("🟧 [Engineering Hours] 依機台 (Tester) 統計", "[Engineering Hours] Total by Tester", eng_tester_hours, 'Tester', 'Engineering Support Hours', filter_col='Tester', custom_palette='Oranges_r')

        with tab4:
            st.markdown("<br>", unsafe_allow_html=True)
            tester_req_hours = aggregate_data(df_tester, 'Customer Requestor', 'Tester Total Hours')
            eng_req_hours = aggregate_data(df_eng, 'Customer Requestor', 'Engineering Support Hours')
            render_table_and_chart("🟦 [Tester Hours] 依客戶統計", "[Tester Hours] Total by Requestor", tester_req_hours, 'Customer Requestor', 'Tester Total Hours', filter_col='Customer Requestor', custom_palette='Set2')
            st.divider()
            render_table_and_chart("🟧 [Engineering Hours] 依客戶統計", "[Engineering Hours] Total by Requestor", eng_req_hours, 'Customer Requestor', 'Engineering Support Hours', filter_col='Customer Requestor', custom_palette='Set1')

    except Exception as e:
        st.error(f"執行時發生錯誤: {e}")

else:
    st.info("👈 請於左側邊欄 (Sidebar) 上傳 Excel 檔案以開始分析。")
