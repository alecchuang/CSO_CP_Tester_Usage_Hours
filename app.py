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
        st.info("檔案上傳成功！正在為您自動拆解多單位並重新計算時數...")
        
        df_tester_raw = pd.read_excel(uploaded_file, sheet_name="Tester Hours", skiprows=3)
        df_eng_raw = pd.read_excel(uploaded_file, sheet_name="Engineering Hours")

        # ==========================================
        # 處理 [Tester Hours] 資料
        # ==========================================
        df_tester = df_tester_raw[['Date', 'Tester #', 'Tester Total Hours', 'TEMP', 'Customer Requestor']].copy()
        df_tester.dropna(subset=['Date', 'Tester #', 'Tester Total Hours'], how='all', inplace=True)
        df_tester['Date'] = pd.to_datetime(df_tester['Date'], errors='coerce')
        df_tester.dropna(subset=['Date'], inplace=True)
        df_tester['Month'] = df_tester['Date'].dt.to_period('M').astype(str)
        df_tester['Tester Total Hours'] = pd.to_numeric(df_tester['Tester Total Hours'], errors='coerce').fillna(0)

        for col in ['Tester #', 'TEMP', 'Customer Requestor']:
            df_tester = split_and_distribute(df_tester, target_col=col, hours_col='Tester Total Hours')

        # 基礎維度匯總
        monthly_tester_hours = df_tester.groupby(['Month', 'Tester #'])['Tester Total Hours'].sum().round(2).reset_index()
        temp_hours = df_tester.groupby('TEMP')['Tester Total Hours'].sum().round(2).reset_index().sort_values('Tester Total Hours', ascending=False)
        tester_req_hours = df_tester.groupby('Customer Requestor')['Tester Total Hours'].sum().round(2).reset_index().sort_values('Tester Total Hours', ascending=False)

        # 🌟 新增：團隊歸屬分析 (CSO vs Gchip)
        # 只要名字是 Alec 就歸類為 CSO，其餘皆為 Gchip
        df_tester['Team'] = df_tester['Customer Requestor'].apply(lambda x: 'CSO' if str(x).strip() == 'Alec' else 'Gchip')
        team_tester_hours = df_tester.groupby('Team')['Tester Total Hours'].sum().round(2).reset_index().sort_values('Tester Total Hours', ascending=False)

        # ==========================================
        # 處理 [Engineering Hours] 資料
        # ==========================================
        df_eng = df_eng_raw[['Date', 'Name', 'Engineering Support Hours', 'Tester', 'Customer Requestor']].copy()
        df_eng.dropna(subset=['Date', 'Name', 'Engineering Support Hours'], how='all', inplace=True)
        df_eng['Date'] = pd.to_datetime(df_eng['Date'], errors='coerce')
        df_eng.dropna(subset=['Date'], inplace=True)
        df_eng['Month'] = df_eng['Date'].dt.to_period('M').astype(str)
        df_eng['Engineering Support Hours'] = pd.to_numeric(df_eng['Engineering Support Hours'], errors='coerce').fillna(0)

        for col in ['Name', 'Tester', 'Customer Requestor']:
            df_eng = split_and_distribute(df_eng, target_col=col, hours_col='Engineering Support Hours')

        # 基礎維度匯總
        monthly_eng_hours = df_eng.groupby(['Month', 'Name'])['Engineering Support Hours'].sum().round(2).reset_index()
        eng_tester_hours = df_eng.groupby('Tester')['Engineering Support Hours'].sum().round(2).reset_index().sort_values('Engineering Support Hours', ascending=False)
        eng_req_hours = df_eng.groupby('Customer Requestor')['Engineering Support Hours'].sum().round(2).reset_index().sort_values('Engineering Support Hours', ascending=False)

        # 🌟 新增：團隊歸屬分析 (CSO vs Gchip)
        df_eng['Team'] = df_eng['Customer Requestor'].apply(lambda x: 'CSO' if str(x).strip() == 'Alec' else 'Gchip')
        team_eng_hours = df_eng.groupby('Team')['Engineering Support Hours'].sum().round(2).reset_index().sort_values('Engineering Support Hours', ascending=False)

        st.success("✅ 資料解析與均分完成！")
        st.divider()

        # ==========================================
        # 繪圖排版與動態篩選函數 (UI中文 / Chart英文)
        # ==========================================
        sns.set_theme(style="whitegrid")

        def render_table_and_chart(ui_title, chart_title, df, x_col, y_col, hue_col=None, filter_col=None, palette=None):
            st.markdown(f"#### {ui_title}")
            
            col_data, col_chart = st.columns([1, 2])
            
            with col_data:
                # 互動式篩選器
                if filter_col:
                    unique_items = sorted(df[filter_col].unique().tolist())
                    selected_items = st.multiselect(
                        f"🔽 篩選 {filter_col} (點擊 'x' 以排除特定項目)", 
                        options=unique_items, 
                        default=unique_items,
                        key=f"filter_{chart_title}"
                    )
                    filtered_df = df[df[filter_col].isin(selected_items)]
                else:
                    filtered_df = df
                
                st.dataframe(filtered_df, use_container_width=True)
                
            with col_chart:
                if filtered_df.empty:
                    st.warning("⚠️ 已排除所有項目，無資料可供繪圖。")
                else:
                    fig, ax = plt.subplots(figsize=(10, 4.5))
                    if hue_col:
                        sns.barplot(data=filtered_df, x=x_col, y=y_col, hue=hue_col, ax=ax, palette=palette)
                        ax.legend(title=hue_col, bbox_to_anchor=(1.05, 1), loc='upper left')
                    else:
                        sns.barplot(data=filtered_df, x=x_col, y=y_col, ax=ax, palette=palette)
                    
                    # 強制英文標題避免亂碼
                    ax.set_title(chart_title, fontweight='bold')
                    ax.set_xlabel(x_col)
                    ax.set_ylabel(y_col)
                    plt.xticks(rotation=45, ha='right')
                    plt.tight_layout()
                    st.pyplot(fig)
                
            st.divider()

        # ==========================================
        # 繪製各區塊
        # ==========================================
        
        # --- 區塊 1：每月趨勢 ---
        st.subheader("📅 每月趨勢分析 / Monthly Trends")
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

        # --- 區塊 2：團隊歸屬分析 (新加入的 2 張圖表) ---
        st.subheader("🏢 團隊歸屬分析 / Team Analysis (CSO vs Gchip)")
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

        # --- 區塊 3：進階維度 (溫度/機台) ---
        st.subheader("🔍 進階維度分析 / Advanced Dimensions")
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

        # --- 區塊 4：客戶需求者 ---
        st.subheader("👤 客戶需求者分析 / Customer Requestor Analysis")
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

    except ValueError as ve:
        st.error(f"讀取失敗：請確認上傳的 Excel 檔案內包含所需的欄位與分頁。\n\n詳細錯誤：{ve}")
    except KeyError as ke:
        st.error(f"找不到特定欄位：請確認 Excel 內有這個欄位 {ke}。")
    except Exception as e:
        st.error(f"處理檔案時發生未預期的錯誤: {e}")

else:
    st.info("👈 請將您的 Excel 檔案拖曳到上方，或點擊 Browse files 選擇檔案。")
