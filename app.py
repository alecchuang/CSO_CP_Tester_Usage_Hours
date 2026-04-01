import streamlit as st
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# 設定網頁標題與寬度
st.set_page_config(page_title="Hours Analysis Dashboard", layout="wide")

st.title("📊 機台與工程師時數進階分析儀表板")

# --- 1. 建立檔案上傳區塊 ---
uploaded_file = st.file_uploader("請上傳您的 Excel 時數紀錄表", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        st.info("檔案上傳成功！正在為您解析資料並繪製圖表...")
        
        # 讀取檔案
        df_tester_raw = pd.read_excel(uploaded_file, sheet_name="Tester Hours", skiprows=3)
        df_eng_raw = pd.read_excel(uploaded_file, sheet_name="Engineering Hours")

        # ==========================================
        # 處理 Tester Hours 資料
        # ==========================================
        # 提取需要的欄位 (包含新加入的 TEMP 與 Customer Requestor)
        df_tester = df_tester_raw[['Date', 'Tester #', 'Tester Total Hours', 'TEMP', 'Customer Requestor']].copy()
        df_tester.dropna(subset=['Date', 'Tester #', 'Tester Total Hours'], how='all', inplace=True)
        df_tester['Date'] = pd.to_datetime(df_tester['Date'], errors='coerce')
        df_tester.dropna(subset=['Date'], inplace=True)
        df_tester['Month'] = df_tester['Date'].dt.to_period('M').astype(str)
        df_tester['Tester Total Hours'] = pd.to_numeric(df_tester['Tester Total Hours'], errors='coerce').fillna(0)

        # 基礎圖表：依月份與機台
        monthly_tester_hours = df_tester.groupby(['Month', 'Tester #'])['Tester Total Hours'].sum().reset_index()
        
        # 新增圖表 1：依 TEMP 累加
        df_tester['TEMP'] = df_tester['TEMP'].astype(str) # 轉字串避免被當作連續數值
        temp_hours = df_tester.groupby('TEMP')['Tester Total Hours'].sum().reset_index()
        temp_hours = temp_hours.sort_values('Tester Total Hours', ascending=False) # 依照時數排序
        
        # 新增圖表 2：依 Customer Requestor 累加
        df_tester['Customer Requestor'] = df_tester['Customer Requestor'].astype(str)
        req_hours = df_tester.groupby('Customer Requestor')['Tester Total Hours'].sum().reset_index()
        req_hours = req_hours.sort_values('Tester Total Hours', ascending=False)

        # ==========================================
        # 處理 Engineering Hours 資料
        # ==========================================
        # 提取需要的欄位 (包含新加入的 Tester)
        df_eng = df_eng_raw[['Date', 'Name', 'Engineering Support Hours', 'Tester']].copy()
        df_eng.dropna(subset=['Date', 'Name', 'Engineering Support Hours'], how='all', inplace=True)
        df_eng['Date'] = pd.to_datetime(df_eng['Date'], errors='coerce')
        df_eng.dropna(subset=['Date'], inplace=True)
        df_eng['Month'] = df_eng['Date'].dt.to_period('M').astype(str)
        df_eng['Engineering Support Hours'] = pd.to_numeric(df_eng['Engineering Support Hours'], errors='coerce').fillna(0)

        # 基礎圖表：依月份與工程師
        monthly_eng_hours = df_eng.groupby(['Month', 'Name'])['Engineering Support Hours'].sum().reset_index()
        
        # 新增圖表 3：依 Tester (機台) 累加
        df_eng['Tester'] = df_eng['Tester'].astype(str)
        eng_tester_hours = df_eng.groupby('Tester')['Engineering Support Hours'].sum().reset_index()
        eng_tester_hours = eng_tester_hours.sort_values('Engineering Support Hours', ascending=False)

        st.success("資料解析完成！圖表已生成。")
        st.divider()

        # ==========================================
        # 繪製圖表區
        # ==========================================
        sns.set_theme(style="whitegrid")

        # 為了讓版面更好看，我們將原本的兩張時間趨勢圖放在全寬度
        st.subheader("📅 每月趨勢分析")
        
        col1, col2 = st.columns(2)
        with col1:
            st.markdown("##### 每月機台總使用時數 (Tester Hours)")
            fig1, ax1 = plt.subplots(figsize=(10, 5))
            sns.barplot(data=monthly_tester_hours, x='Month', y='Tester Total Hours', hue='Tester #', ax=ax1)
            ax1.set_xlabel('Month')
            ax1.set_ylabel('Total Hours')
            ax1.legend(title='Tester #', bbox_to_anchor=(1.05, 1), loc='upper left')
            plt.tight_layout()
            st.pyplot(fig1)

        with col2:
            st.markdown("##### 每月工程師支援時數 (Engineering Hours)")
            fig2, ax2 = plt.subplots(figsize=(10, 5))
            sns.barplot(data=monthly_eng_hours, x='Month', y='Engineering Support Hours', hue='Name', ax=ax2)
            ax2.set_xlabel('Month')
            ax2.set_ylabel('Total Hours')
            ax2.legend(title='Engineer Name', bbox_to_anchor=(1.05, 1), loc='upper left', ncol=2)
            plt.tight_layout()
            st.pyplot(fig2)

        st.divider()

        # 將新的三張圖表排列在下方
        st.subheader("🔍 進階維度分析")

        col3, col4 = st.columns(2)
        
        with col3:
            # 新圖表 1: 依 TEMP
            st.markdown("##### 🌡️ 依溫度 (TEMP) 統計機台總時數")
            fig3, ax3 = plt.subplots(figsize=(10, 5))
            sns.barplot(data=temp_hours, x='TEMP', y='Tester Total Hours', ax=ax3, palette='Set2')
            ax3.set_xlabel('TEMP')
            ax3.set_ylabel('Total Hours')
            plt.tight_layout()
            st.pyplot(fig3)

        with col4:
            # 新圖表 3: 依 Tester 統計工程師時數
            st.markdown("##### 🛠️ 依機台 (Tester) 統計工程師支援總時數")
            fig5, ax5 = plt.subplots(figsize=(10, 5))
            sns.barplot(data=eng_tester_hours, x='Tester', y='Engineering Support Hours', ax=ax5, palette='magma')
            ax5.set_xlabel('Tester')
            ax5.set_ylabel('Total Hours')
            plt.xticks(rotation=45, ha='right') # 旋轉 X 軸標籤以免擠在一起
            plt.tight_layout()
            st.pyplot(fig5)

        # 新圖表 2: 依 Customer Requestor (因為人名可能很長，單獨給他一整行的寬度)
        st.markdown("##### 👤 依客戶需求者 (Customer Requestor) 統計機台總時數")
        fig4, ax4 = plt.subplots(figsize=(12, 5))
        sns.barplot(data=req_hours, x='Customer Requestor', y='Tester Total Hours', ax=ax4, palette='viridis')
        ax4.set_xlabel('Customer Requestor')
        ax4.set_ylabel('Total Hours')
        plt.xticks(rotation=45, ha='right') # 將名字傾斜 45 度避免重疊
        plt.tight_layout()
        st.pyplot(fig4)

    except ValueError as ve:
        st.error(f"讀取失敗：請確認上傳的 Excel 檔案內包含所需的欄位與分頁。\n\n詳細錯誤：{ve}")
    except KeyError as ke:
        st.error(f"找不到特定欄位：請確認 Excel 內有這個欄位 {ke}。")
    except Exception as e:
        st.error(f"處理檔案時發生未預期的錯誤: {e}")

else:
    st.info("👈 請將您的 Excel 檔案拖曳到上方，或點擊 Browse files 選擇檔案。")
