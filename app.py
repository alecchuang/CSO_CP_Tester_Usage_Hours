import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# 設定網頁標題與寬度
st.set_page_config(page_title="Hours Analysis", layout="wide")

st.title("📊 機台與工程師時數分析")

# --- 1. 建立檔案上傳區塊 ---
# 限定只能上傳 Excel 檔案
uploaded_file = st.file_uploader("請上傳您的 Excel 時數紀錄表", type=["xlsx", "xls"])

# --- 2. 判斷是否有上傳檔案 ---
# 只有當使用者上傳了檔案，才執行後續的讀取與分析
if uploaded_file is not None:
    try:
        st.info("檔案上傳成功！正在為您解析資料並繪製圖表...")
        
        # 直接讀取上傳的檔案物件 (uploaded_file)
        df_tester = pd.read_excel(uploaded_file, sheet_name="Tester Hours", skiprows=3)
        df_eng = pd.read_excel(uploaded_file, sheet_name="Engineering Hours")

        # --- 3. 處理 Tester Hours 資料 ---
        df_tester = df_tester[['Date', 'Tester #', 'Tester Total Hours']].copy()
        df_tester.dropna(subset=['Date', 'Tester #', 'Tester Total Hours'], how='all', inplace=True)
        df_tester['Date'] = pd.to_datetime(df_tester['Date'], errors='coerce')
        df_tester.dropna(subset=['Date'], inplace=True)
        df_tester['Month'] = df_tester['Date'].dt.to_period('M').astype(str)
        df_tester['Tester Total Hours'] = pd.to_numeric(df_tester['Tester Total Hours'], errors='coerce').fillna(0)
        monthly_tester_hours = df_tester.groupby(['Month', 'Tester #'])['Tester Total Hours'].sum().reset_index()

        # --- 4. 處理 Engineering Hours 資料 ---
        df_eng = df_eng[['Date', 'Name', 'Engineering Support Hours']].copy()
        df_eng.dropna(subset=['Date', 'Name', 'Engineering Support Hours'], how='all', inplace=True)
        df_eng['Date'] = pd.to_datetime(df_eng['Date'], errors='coerce')
        df_eng.dropna(subset=['Date'], inplace=True)
        df_eng['Month'] = df_eng['Date'].dt.to_period('M').astype(str)
        df_eng['Engineering Support Hours'] = pd.to_numeric(df_eng['Engineering Support Hours'], errors='coerce').fillna(0)
        monthly_eng_hours = df_eng.groupby(['Month', 'Name'])['Engineering Support Hours'].sum().reset_index()

        st.success("分析完成！")
        st.divider()

        # --- 5. 繪製圖表 ---
        sns.set_theme(style="whitegrid")

        # 圖表 1: Tester Hours
        st.subheader("🖥️ 每月機台總使用時數 (Tester Hours)")
        fig1, ax1 = plt.subplots(figsize=(12, 6))
        sns.barplot(data=monthly_tester_hours, x='Month', y='Tester Total Hours', hue='Tester #', ax=ax1)
        ax1.set_title('Total Tester Hours per Month by Tester', fontsize=14)
        ax1.set_xlabel('Month')
        ax1.set_ylabel('Total Hours')
        ax1.legend(title='Tester #', bbox_to_anchor=(1.05, 1), loc='upper left')
        plt.tight_layout()
        st.pyplot(fig1) 

        st.divider()

        # 圖表 2: Engineering Hours
        st.subheader("🧑‍🔧 每月工程師支援時數 (Engineering Hours)")
        fig2, ax2 = plt.subplots(figsize=(12, 6))
        sns.barplot(data=monthly_eng_hours, x='Month', y='Engineering Support Hours', hue='Name', ax=ax2)
        ax2.set_title('Total Engineering Hours per Month by Engineer', fontsize=14)
        ax2.set_xlabel('Month')
        ax2.set_ylabel('Total Hours')
        ax2.legend(title='Engineer Name', bbox_to_anchor=(1.05, 1), loc='upper left', ncol=2)
        plt.tight_layout()
        st.pyplot(fig2)

    except ValueError as ve:
        # 捕捉找不到特定 Sheet (分頁) 的錯誤
        st.error(f"讀取失敗：請確認上傳的 Excel 檔案內包含 `Tester Hours` 與 `Engineering Hours` 這兩個分頁。\n\n詳細錯誤：
