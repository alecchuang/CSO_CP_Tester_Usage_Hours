import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# 設定網頁標題與寬度
st.set_page_config(page_title="Hours Analysis", layout="wide")

st.title("📊 機台與工程師時數分析")

# 填入您的 Excel 檔案路徑 (請確保檔案與 app.py 放在同一個資料夾)
file_path = "Google Monthly Engineering ATT_Service_Hours_Wafer Sort_ 2026__Eng_hours_revD.xlsx"

try:
    # --- 1. 讀取資料 ---
    st.info("正在讀取資料，請稍候...")
    df_tester = pd.read_excel(file_path, sheet_name="Tester Hours", skiprows=3)
    df_eng = pd.read_excel(file_path, sheet_name="Engineering Hours")

    # --- 2. 處理 Tester Hours 資料 ---
    df_tester = df_tester[['Date', 'Tester #', 'Tester Total Hours']].copy()
    df_tester.dropna(subset=['Date', 'Tester #', 'Tester Total Hours'], how='all', inplace=True)
    df_tester['Date'] = pd.to_datetime(df_tester['Date'], errors='coerce')
    df_tester.dropna(subset=['Date'], inplace=True)
    df_tester['Month'] = df_tester['Date'].dt.to_period('M').astype(str)
    df_tester['Tester Total Hours'] = pd.to_numeric(df_tester['Tester Total Hours'], errors='coerce').fillna(0)
    monthly_tester_hours = df_tester.groupby(['Month', 'Tester #'])['Tester Total Hours'].sum().reset_index()

    # --- 3. 處理 Engineering Hours 資料 ---
    df_eng = df_eng[['Date', 'Name', 'Engineering Support Hours']].copy()
    df_eng.dropna(subset=['Date', 'Name', 'Engineering Support Hours'], how='all', inplace=True)
    df_eng['Date'] = pd.to_datetime(df_eng['Date'], errors='coerce')
    df_eng.dropna(subset=['Date'], inplace=True)
    df_eng['Month'] = df_eng['Date'].dt.to_period('M').astype(str)
    df_eng['Engineering Support Hours'] = pd.to_numeric(df_eng['Engineering Support Hours'], errors='coerce').fillna(0)
    monthly_eng_hours = df_eng.groupby(['Month', 'Name'])['Engineering Support Hours'].sum().reset_index()

    st.success("資料讀取完成！")
    st.divider() # 畫一條分隔線

    # --- 4. 繪製圖表 ---
    sns.set_theme(style="whitegrid")

    # 圖表 1: Tester Hours
    st.subheader("🖥️ 每月機台總使用時數 (Tester Hours)")
    # 在 Streamlit 中，我們必須先建立畫布 (fig, ax)
    fig1, ax1 = plt.subplots(figsize=(12, 6))
    sns.barplot(data=monthly_tester_hours, x='Month', y='Tester Total Hours', hue='Tester #', ax=ax1)
    ax1.set_title('Total Tester Hours per Month by Tester', fontsize=14)
    ax1.set_xlabel('Month')
    ax1.set_ylabel('Total Hours')
    ax1.legend(title='Tester #', bbox_to_anchor=(1.05, 1), loc='upper left')
    plt.tight_layout()
    # 使用 st.pyplot 將畫布顯示在網頁上
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
    # 同樣使用 st.pyplot 顯示第二張圖
    st.pyplot(fig2)

except FileNotFoundError:
    st.error(f"找不到檔案：`{file_path}`。請確認檔案已上傳且名稱完全相符！")
except Exception as e:
    st.error(f"執行時發生未預期的錯誤: {e}")
