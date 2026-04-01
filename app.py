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

        # [匯總資料] 確保數字為小數點後兩位，方便表格顯示
        monthly_tester_hours = df_tester.groupby(['Month', 'Tester #'])['Tester Total Hours'].sum().round(2).reset_index()
        temp_hours = df_tester.groupby('TEMP')['Tester Total Hours'].sum().round(2).reset_index().sort_values('Tester Total Hours', ascending=False)
        tester_req_hours = df_tester.groupby('Customer Requestor')['Tester Total Hours'].sum().round(2).reset_index().sort_values('Tester Total Hours', ascending=False)

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

        # [匯總資料]
        monthly_eng_hours = df_eng.groupby(['Month', 'Name'])['Engineering Support Hours'].sum().round(2).reset_index()
        eng_tester_hours = df_eng.groupby('Tester')['Engineering Support Hours'].sum().round(2).reset_index().sort_values('Engineering Support Hours', ascending=False)
        eng_req_hours = df_eng.groupby('Customer Requestor')['Engineering Support Hours'].sum().round(2).reset_index().sort_values('Engineering Support Hours', ascending=False)

        st.success("✅ 資料解析與均分完成！")
        st.divider()

        # ==========================================
        # 自動化 UI 排版函數 (左數據、右圖表)
        # ==========================================
        sns.set_theme(style="whitegrid")

        def render_table_and_chart(title, df, x_col, y_col, hue_col=None, palette=None):
            st.markdown(f"#### {title}")
            
            # 將畫面分割為左右兩欄：左邊占 1 份寬度，右邊占 2 份寬度
            col_data, col_chart = st.columns([1, 2])
            
            with col_data:
                # 左側：顯示資料表 (表格)
                st.dataframe(df, use_container_width=True)
                
            with col_chart:
                # 右側：顯示長條圖
                fig, ax = plt.subplots(figsize=(10, 4.5))
                if hue_col:
                    sns.barplot(data=df, x=x_col, y=y_col, hue=hue_col, ax=ax, palette=palette)
                    ax.legend(title=hue_col, bbox_to_anchor=(1.05, 1), loc='upper left')
                else:
                    sns.barplot(data=df, x=x_col, y=y_col, ax=ax, palette=palette)
                
                ax.set_title(title, fontweight='bold')
                ax.set_xlabel(x_col)
                ax.set_ylabel(y_col)
                plt.xticks(rotation=45, ha='right')
                plt.tight_layout()
                st.pyplot(fig)
                
            st.divider() # 每個區塊結束後畫一條分隔線

        # ==========================================
        # 繪製各區塊
        # ==========================================
        st.subheader("📅 每月趨勢分析")
        render_table_and_chart("🟦 [Tester Hours] 每月機台總使用時數", monthly_tester_hours, 'Month', 'Tester Total Hours', hue_col='Tester #')
        render_table_and_chart("🟧 [Engineering Hours] 每月工程師支援時數", monthly_eng_hours, 'Month', 'Engineering Support Hours', hue_col='Name')

        st.subheader("🔍 進階維度分析 (溫度與機台)")
        render_table_and_chart("🟦 [Tester Hours] 依溫度 (TEMP) 統計機台時數", temp_hours, 'TEMP', 'Tester Total Hours', palette='Set2')
        render_table_and_chart("🟧 [Engineering Hours] 依機台 (Tester) 統計工程師時數", eng_tester_hours, 'Tester', 'Engineering Support Hours', palette='magma')

        st.subheader("👤 客戶需求者分析 (Customer Requestor)")
        render_table_and_chart("🟦 [Tester Hours] 依客戶需求者統計", tester_req_hours, 'Customer Requestor', 'Tester Total Hours', palette='viridis')
        render_table_and_chart("🟧 [Engineering Hours] 依客戶需求者統計", eng_req_hours, 'Customer Requestor', 'Engineering Support Hours', palette='rocket')

    except ValueError as ve:
        st.error(f"讀取失敗：請確認上傳的 Excel 檔案內包含所需的欄位與分頁。\n\n詳細錯誤：{ve}")
    except KeyError as ke:
        st.error(f"找不到特定欄位：請確認 Excel 內有這個欄位 {ke}。")
    except Exception as e:
        st.error(f"處理檔案時發生未預期的錯誤: {e}")

else:
    st.info("👈 請將您的 Excel 檔案拖曳到上方，或點擊 Browse files 選擇檔案。")
