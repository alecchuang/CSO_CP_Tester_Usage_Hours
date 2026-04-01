import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
import re

# ==========================================
# 定義：時數分割與均分函數
# ==========================================
def split_and_distribute(df, target_col, hours_col):
    """將指定欄位中的多個數值(用/,;\n隔開)拆分，並將該列時數平均分配"""
    df = df.copy()
    df[target_col] = df[target_col].astype(str).replace(['nan', 'None', ''], 'Unknown')
    
    def safe_split(val):
        if val == 'Unknown':
            return ['Unknown']
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
# 網頁主程式開始
# ==========================================
st.set_page_config(page_title="Hours Analysis Dashboard", layout="wide")
st.title("📊 機台與工程師時數進階分析儀表板")

uploaded_file = st.file_uploader("請上傳您的 Excel 時數紀錄表", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        st.info("檔案上傳成功！正在初步解析數據...")
        
        # 讀取檔案
        df_tester_raw = pd.read_excel(uploaded_file, sheet_name="Tester Hours", skiprows=3)
        df_eng_raw = pd.read_excel(uploaded_file, sheet_name="Engineering Hours")

        # --- 資料預處理 (Tester Hours) ---
        df_tester = df_tester_raw[['Date', 'Tester #', 'Tester Total Hours', 'TEMP', 'Customer Requestor']].copy()
        df_tester.dropna(subset=['Date', 'Tester #', 'Tester Total Hours'], how='all', inplace=True)
        df_tester['Date'] = pd.to_datetime(df_tester['Date'], errors='coerce')
        df_tester.dropna(subset=['Date'], inplace=True)
        df_tester['Month'] = df_tester['Date'].dt.to_period('M').astype(str)
        df_tester['Tester Total Hours'] = pd.to_numeric(df_tester['Tester Total Hours'], errors='coerce').fillna(0)

        # --- 資料預處理 (Engineering Hours) ---
        df_eng = df_eng_raw[['Date', 'Name', 'Engineering Support Hours', 'Tester', 'Customer Requestor']].copy()
        df_eng.dropna(subset=['Date', 'Name', 'Engineering Support Hours'], how='all', inplace=True)
        df_eng['Date'] = pd.to_datetime(df_eng['Date'], errors='coerce')
        df_eng.dropna(subset=['Date'], inplace=True)
        df_eng['Month'] = df_eng['Date'].dt.to_period('M').astype(str)
        df_eng['Engineering Support Hours'] = pd.to_numeric(df_eng['Engineering Support Hours'], errors='coerce').fillna(0)

        # 套用均分邏輯
        for col in ['Tester #', 'TEMP', 'Customer Requestor']:
            df_tester = split_and_distribute(df_tester, target_col=col, hours_col='Tester Total Hours')
        for col in ['Name', 'Tester', 'Customer Requestor']:
            df_eng = split_and_distribute(df_eng, target_col=col, hours_col='Engineering Support Hours')

        st.divider()

        # ==========================================
        # 🌟 團隊成員自定義區塊 (加入互斥防呆邏輯)
        # ==========================================
        st.subheader("⚙️ 團隊成員定義 / Team Definitions")
        
        # 找出數據中出現過的所有 Requestor 名稱
        all_requestors = sorted(list(set(df_tester['Customer Requestor'].unique()) | set(df_eng['Customer Requestor'].unique())))
        
        # 1. 透過 session_state 初始化預設名單 (只在第一次載入時執行)
        if 'cso_selection' not in st.session_state:
            st.session_state.cso_selection = [name for name in ['Alec'] if name in all_requestors]
        if 'gchip_selection' not in st.session_state:
            st.session_state.gchip_selection = [name for name in ['Rajesh', 'Louis', 'Chi-Chang'] if name in all_requestors]

        # 2. 動態計算可選清單 (排除對方已經選擇的人員)
        avail_for_cso = [x for x in all_requestors if x not in st.session_state.gchip_selection]
        avail_for_gchip = [x for x in all_requestors if x not in st.session_state.cso_selection]

        col_cso_config, col_gchip_config = st.columns(2)
        
        with col_cso_config:
            # 透過 key 綁定 session_state，不需要再寫 default
            cso_members = st.multiselect(
                "定義 CSO 成員 (CSO Members)", 
                options=avail_for_cso, 
                key="cso_selection"
            )
            
        with col_gchip_config:
            # 透過 key 綁定 session_state
            gchip_members = st.multiselect(
                "定義 Gchip 成員 (Gchip Members)", 
                options=avail_for_gchip, 
                key="gchip_selection"
            )

        # 套用團隊歸屬邏輯
        def map_team(name):
            if name in cso_members:
                return 'CSO'
            elif name in gchip_members:
                return 'Gchip'
            else:
                return 'Other / Unassigned'

        df_tester['Team'] = df_tester['Customer Requestor'].apply(map_team)
        df_eng['Team'] = df_eng['Customer Requestor'].apply(map_team)

        # ==========================================
        # 繪圖排版函數
        # ==========================================
        sns.set_theme(style="whitegrid")

        def render_table_and_chart(ui_title, chart_title, df, x_col, y_col, hue_col=None, filter_col=None, palette=None):
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
                st.dataframe(filtered_df, use_container_width=True)
            with col_chart:
                if filtered_df.empty:
                    st.warning("No data to display.")
                else:
                    fig, ax = plt.subplots(figsize=(10, 4.5))
                    if hue_col:
                        sns.barplot(data=filtered_df, x=x_col, y=y_col, hue=hue_col, ax=ax, palette=palette)
                        ax.legend(title=hue_col, bbox_to_anchor=(1.05, 1), loc='upper left')
                    else:
                        sns.barplot(data=filtered_df, x=x_col, y=y_col, ax=ax, palette=palette)
                    ax.set_title(chart_title, fontweight='bold')
                    ax.set_xlabel(x_col); ax.set_ylabel(y_col)
                    plt.xticks(rotation=45, ha='right'); plt.tight_layout()
                    st.pyplot(fig)
            st.divider()

        # ==========================================
        # 繪製各區塊 
        # ==========================================
        
        # --- 1. 團隊歸屬分析 (CSO vs Gchip) ---
        st.subheader("🏢 團隊歸屬分析 / Team Analysis (CSO vs Gchip)")
        team_tester_hours = df_tester.groupby('Team')['Tester Total Hours'].sum().round(2).reset_index().sort_values('Tester Total Hours', ascending=False)
        team_eng_hours = df_eng.groupby('Team')['Engineering Support Hours'].sum().round(2).reset_index().sort_values('Engineering Support Hours', ascending=False)

        render_table_and_chart(
            ui_title="🟦 [Tester Hours] 依團隊統計機台時數", 
            chart_title="[Tester Hours] Total Hours by Team",
            df=team_tester_hours, x_col='Team', y_col='Tester Total Hours', filter_col='Team', palette='Set1'
        )
        render_table_and_chart(
            ui_title="🟧 [Engineering Hours] 依團隊統計工程師時數", 
            chart_title="[Engineering Hours] Total Hours by Team",
            df=team_eng_hours, x_col='Team', y_col='Engineering Support Hours', filter_col='Team', palette='Dark2'
        )

        # --- 2. 每月趨勢分析 ---
        st.subheader("📅 每月趨勢分析 / Monthly Trends")
        monthly_tester_hours = df_tester.groupby(['Month', 'Tester #'])['Tester Total Hours'].sum().round(2).reset_index()
        monthly_eng_hours = df_eng.groupby(['Month', 'Name'])['Engineering Support Hours'].sum().round(2).reset_index()

        render_table_and_chart(
            ui_title="🟦 [Tester Hours] 每月機台總使用時數", 
            chart_title="[Tester Hours] Monthly Total by Tester",
            df=monthly_tester_hours, x_col='Month', y_col='Tester Total Hours', hue_col='Tester #', filter_col='Tester #'
        )
        render_table_and_chart(
            ui_title="🟧 [Engineering Hours] 每月工程師支援時數", 
            chart_title="[Engineering Hours] Monthly Total by Engineer",
            df=monthly_eng_hours, x_col='Month', y_col='Engineering Support Hours', hue_col='Name', filter_col='Name'
        )

        # --- 3. 進階維度分析 ---
        st.subheader("🔍 進階維度分析 / Advanced Dimensions")
        temp_hours = df_tester.groupby('TEMP')['Tester Total Hours'].sum().round(2).reset_index().sort_values('Tester Total Hours', ascending=False)
        eng_tester_hours = df_eng.groupby('Tester')['Engineering Support Hours'].sum().round(2).reset_index().sort_values('Engineering Support Hours', ascending=False)

        render_table_and_chart(
            ui_title="🟦 [Tester Hours] 依溫度 (TEMP) 統計機台時數", 
            chart_title="[Tester Hours] Total Hours by TEMP",
            df=temp_hours, x_col='TEMP', y_col='Tester Total Hours', filter_col='TEMP', palette='Set2'
        )
        render_table_and_chart(
            ui_title="🟧 [Engineering Hours] 依機台 (Tester) 統計工程師時數", 
            chart_title="[Engineering Hours] Total Hours by Tester",
            df=eng_tester_hours, x_col='Tester', y_col='Engineering Support Hours', filter_col='Tester', palette='magma'
        )

        # --- 4. 客戶需求者分析 ---
        st.subheader("👤 客戶需求者分析 / Customer Requestor Analysis")
        tester_req_hours = df_tester.groupby('Customer Requestor')['Tester Total Hours'].sum().round(2).reset_index().sort_values('Tester Total Hours', ascending=False)
        eng_req_hours = df_eng.groupby('Customer Requestor')['Engineering Support Hours'].sum().round(2).reset_index().sort_values('Engineering Support Hours', ascending=False)

        render_table_and_chart(
            ui_title="🟦 [Tester Hours] 依客戶需求者統計", 
            chart_title="[Tester Hours] Total Hours by Customer Requestor",
            df=tester_req_hours, x_col='Customer Requestor', y_col='Tester Total Hours', filter_col='Customer Requestor', palette='viridis'
        )
        render_table_and_chart(
            ui_title="🟧 [Engineering Hours] 依客戶需求者統計", 
            chart_title="[Engineering Hours] Total Hours by Customer Requestor",
            df=eng_req_hours, x_col='Customer Requestor', y_col='Engineering Support Hours', filter_col='Customer Requestor', palette='rocket'
        )

    except Exception as e:
        st.error(f"執行時發生錯誤: {e}")

else:
    st.info("👈 請將您的 Excel 檔案拖曳到上方，或點擊 Browse files 選擇檔案。")
