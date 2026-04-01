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
    
    # 確保欄位為字串並處理空值
    df[target_col] = df[target_col].astype(str).replace(['nan', 'None', ''], 'Unknown')
    
    def safe_split(val):
        if val == 'Unknown':
            return ['Unknown']
        # 使用正則表達式分割: 斜線、逗號、分號，以及換行符號(\n, \r)
        parts = re.split(r'[/,;\n\r]+', str(val))
        # 去除前後空白並過濾掉空字串
        parts = [p.strip() for p in parts if p.strip()]
        return parts if parts else ['Unknown']

    # 1. 產生分割後的清單陣列 (例如: ['A', 'B'])
    df['__split_list'] = df[target_col].apply(safe_split)
    
    # 2. 計算這格被切割成多少個單位 (例如: 2)
    df['__split_count'] = df['__split_list'].apply(len)
    
    # 3. 時數平均分配 (總時數 / 單位數)
    # 若原本時數為 20, 則變成 20 / 2 = 10
    df[hours_col] = df[hours_col] / df['__split_count']
    
    # 4. 將陣列展開成多筆資料 (explode)
    # 原本 1 列會變成 2 列，且每列的時數都是均分後的 10
    df = df.explode('__split_list')
    
    # 5. 把展開後的值寫回原本的欄位
    df[target_col] = df['__split_list']
    
    # 刪除暫存用的欄位
    df = df.drop(columns=['__split_list', '__split_count'])
    return df

# ==========================================
# 網頁主程式開始
# ==========================================
# 設定網頁標題與寬度
st.set_page_config(page_title="Hours Analysis Dashboard", layout="wide")

st.title("📊 機台與工程師時數進階分析儀表板 (含時數均分)")

# --- 1. 建立檔案上傳區塊 ---
uploaded_file = st.file_uploader("請上傳您的 Excel 時數紀錄表", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        st.info("檔案上傳成功！正在為您自動拆解多單位並重新計算時數...")
        
        # 讀取檔案
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

        # 🌟 套用「分割與均分」邏輯到需要統計的維度欄位 🌟
        for col in ['Tester #', 'TEMP', 'Customer Requestor']:
            df_tester = split_and_distribute(df_tester, target_col=col, hours_col='Tester Total Hours')

        # 1. 依月份與機台 (已均分)
        monthly_tester_hours = df_tester.groupby(['Month', 'Tester #'])['Tester Total Hours'].sum().reset_index()
        
        # 2. 依 TEMP (已均分)
        temp_hours = df_tester.groupby('TEMP')['Tester Total Hours'].sum().reset_index()
        temp_hours = temp_hours.sort_values('Tester Total Hours', ascending=False)
        
        # 3. 依 Customer Requestor (已均分)
        tester_req_hours = df_tester.groupby('Customer Requestor')['Tester Total Hours'].sum().reset_index()
        tester_req_hours = tester_req_hours.sort_values('Tester Total Hours', ascending=False)

        # ==========================================
        # 處理 [Engineering Hours] 資料
        # ==========================================
        df_eng = df_eng_raw[['Date', 'Name', 'Engineering Support Hours', 'Tester', 'Customer Requestor']].copy()
        df_eng.dropna(subset=['Date', 'Name', 'Engineering Support Hours'], how='all', inplace=True)
        df_eng['Date'] = pd.to_datetime(df_eng['Date'], errors='coerce')
        df_eng.dropna(subset=['Date'], inplace=True)
        df_eng['Month'] = df_eng['Date'].dt.to_period('M').astype(str)
        df_eng['Engineering Support Hours'] = pd.to_numeric(df_eng['Engineering Support Hours'], errors='coerce').fillna(0)

        # 🌟 套用「分割與均分」邏輯到需要統計的維度欄位 🌟
        for col in ['Name', 'Tester', 'Customer Requestor']:
            df_eng = split_and_distribute(df_eng, target_col=col, hours_col='Engineering Support Hours')

        # 1. 依月份與工程師 (已均分)
        monthly_eng_hours = df_eng.groupby(['Month', 'Name'])['Engineering Support Hours'].sum().reset_index()
        
        # 2. 依 Tester (已均分)
        eng_tester_hours = df_eng.groupby('Tester')['Engineering Support Hours'].sum().reset_index()
        eng_tester_hours = eng_tester_hours.sort_values('Engineering Support Hours', ascending=False)

        # 3. 依 Customer Requestor (已均分)
        eng_req_hours = df_eng.groupby('Customer Requestor')['Engineering Support Hours'].sum().reset_index()
        eng_req_hours = eng_req_hours.sort_values('Engineering Support Hours', ascending=False)

        st.success("✅ 資料解析與均分完成！圖表已生成。")
        st.divider()

        # ==========================================
        # 繪製圖表區
        # ==========================================
        sns.set_theme(style="whitegrid")

        # --- Row 1: 每月趨勢 ---
        st.subheader("📅 每月趨勢分析")
        col1, col2 = st.columns(2)
        
        with col1:
            st.markdown("##### 🟦 [Tester Hours] 每月機台總使用時數")
            fig1, ax1 = plt.subplots(figsize=(10, 5))
            sns.barplot(data=monthly_tester_hours, x='Month', y='Tester Total Hours', hue='Tester #', ax=ax1)
            ax1.set_title('[Tester Hours] Monthly Total by Tester', fontweight='bold')
            ax1.set_xlabel('Month')
            ax1.set_ylabel('Total Hours')
            ax1.legend(title='Tester #', bbox_to_anchor=(1.05, 1), loc='upper left')
            plt.tight_layout()
            st.pyplot(fig1)

        with col2:
            st.markdown("##### 🟧 [Engineering Hours] 每月工程師支援時數")
            fig2, ax2 = plt.subplots(figsize=(10, 5))
            sns.barplot(data=monthly_eng_hours, x='Month', y='Engineering Support Hours', hue='Name', ax=ax2)
            ax2.set_title('[Engineering Hours] Monthly Total by Engineer', fontweight='bold')
            ax2.set_xlabel('Month')
            ax2.set_ylabel('Total Hours')
            ax2.legend(title='Engineer Name', bbox_to_anchor=(1.05, 1), loc='upper left', ncol=2)
            plt.tight_layout()
            st.pyplot(fig2)

        st.divider()

        # --- Row 2: 設備與環境維度 ---
        st.subheader("🔍 進階維度分析 (溫度與機台)")
        col3, col4 = st.columns(2)
        
        with col3:
            st.markdown("##### 🟦 [Tester Hours] 依溫度 (TEMP) 統計機台時數")
            fig3, ax3 = plt.subplots(figsize=(10, 5))
            sns.barplot(data=temp_hours, x='TEMP', y='Tester Total Hours', ax=ax3, palette='Set2')
            ax3.set_title('[Tester Hours] Total Hours by TEMP', fontweight='bold')
            ax3.set_xlabel('TEMP')
            ax3.set_ylabel('Total Hours')
            plt.tight_layout()
            st.pyplot(fig3)

        with col4:
            st.markdown("##### 🟧 [Engineering Hours] 依機台 (Tester) 統計工程師時數")
            fig5, ax5 = plt.subplots(figsize=(10, 5))
            sns.barplot(data=eng_tester_hours, x='Tester', y='Engineering Support Hours', ax=ax5, palette='magma')
            ax5.set_title('[Engineering Hours] Total Hours by Tester', fontweight='bold')
            ax5.set_xlabel('Tester')
            ax5.set_ylabel('Total Hours')
            plt.xticks(rotation=45, ha='right')
            plt.tight_layout()
            st.pyplot(fig5)

        st.divider()

        # --- Row 3: 客戶需求者維度 ---
        st.subheader("👤 客戶需求者分析 (Customer Requestor)")
        col5, col6 = st.columns(2)

        with col5:
            st.markdown("##### 🟦 [Tester Hours] 依客戶需求者統計")
            fig4, ax4 = plt.subplots(figsize=(10, 5))
            sns.barplot(data=tester_req_hours, x='Customer Requestor', y='Tester Total Hours', ax=ax4, palette='viridis')
            ax4.set_title('[Tester Hours] Total Hours by Customer Requestor', fontweight='bold')
            ax4.set_xlabel('Customer Requestor')
            ax4.set_ylabel('Total Hours')
            plt.xticks(rotation=45, ha='right')
            plt.tight_layout()
            st.pyplot(fig4)

        with col6:
            st.markdown("##### 🟧 [Engineering Hours] 依客戶需求者統計")
            fig6, ax6 = plt.subplots(figsize=(10, 5))
            sns.barplot(data=eng_req_hours, x='Customer Requestor', y='Engineering Support Hours', ax=ax6, palette='rocket')
            ax6.set_title('[Engineering Hours] Total Hours by Customer Requestor', fontweight='bold')
            ax6.set_xlabel('Customer Requestor')
            ax6.set_ylabel('Total Hours')
            plt.xticks(rotation=45, ha='right')
            plt.tight_layout()
            st.pyplot(fig6)

    except ValueError as ve:
        st.error(f"讀取失敗：請確認上傳的 Excel 檔案內包含所需的欄位與分頁。\n\n詳細錯誤：{ve}")
    except KeyError as ke:
        st.error(f"找不到特定欄位：請確認 Excel 內有這個欄位 {ke}。")
    except Exception as e:
        st.error(f"處理檔案時發生未預期的錯誤: {e}")

else:
    st.info("👈 請將您的 Excel 檔案拖曳到上方，或點擊 Browse files 選擇檔案。")
