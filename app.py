
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

def analyze_hours(excel_file_path):
    # ---------------------------
    # 1. 讀取 Excel 檔案中的分頁
    # ---------------------------
    try:
        # 讀取 Tester Hours (跳過前 3 列的非表頭資訊)
        df_tester = pd.read_excel(excel_file_path, sheet_name="Tester Hours", skiprows=3)
        # 讀取 Engineering Hours (這頁的表頭通常在第一列，所以不用 skiprows，視實際情況調整)
        df_eng = pd.read_excel(excel_file_path, sheet_name="Engineering Hours") 
    except Exception as e:
        print(f"讀取 Excel 發生錯誤: {e}")
        return

    # ---------------------------
    # 2. 處理 Tester Hours 資料
    # ---------------------------
    df_tester = df_tester[['Date', 'Tester #', 'Tester Total Hours']].copy()
    df_tester.dropna(subset=['Date', 'Tester #', 'Tester Total Hours'], how='all', inplace=True)
    df_tester['Date'] = pd.to_datetime(df_tester['Date'], errors='coerce')
    df_tester.dropna(subset=['Date'], inplace=True)
    
    # 提取月份
    df_tester['Month'] = df_tester['Date'].dt.to_period('M').astype(str)
    df_tester['Tester Total Hours'] = pd.to_numeric(df_tester['Tester Total Hours'], errors='coerce').fillna(0)

    # 分組加總 (依月份與機台)
    monthly_tester_hours = df_tester.groupby(['Month', 'Tester #'])['Tester Total Hours'].sum().reset_index()

    # ---------------------------
    # 3. 處理 Engineering Hours 資料
    # ---------------------------
    # 篩選出 Date, Name, Engineering Support Hours 欄位
    df_eng = df_eng[['Date', 'Name', 'Engineering Support Hours']].copy()
    df_eng.dropna(subset=['Date', 'Name', 'Engineering Support Hours'], how='all', inplace=True)
    df_eng['Date'] = pd.to_datetime(df_eng['Date'], errors='coerce')
    df_eng.dropna(subset=['Date'], inplace=True)
    
    # 提取月份
    df_eng['Month'] = df_eng['Date'].dt.to_period('M').astype(str)
    # 確保時數為數值格式
    df_eng['Engineering Support Hours'] = pd.to_numeric(df_eng['Engineering Support Hours'], errors='coerce').fillna(0)

    # 分組加總 (依月份與工程師名字)
    monthly_eng_hours = df_eng.groupby(['Month', 'Name'])['Engineering Support Hours'].sum().reset_index()

    # ---------------------------
    # 4. 繪製圖表 (顯示兩張圖)
    # ---------------------------
    sns.set_theme(style="whitegrid")

    # --- 圖表 1: Tester Hours ---
    plt.figure(figsize=(12, 6))
    sns.barplot(data=monthly_tester_hours, x='Month', y='Tester Total Hours', hue='Tester #')
    plt.title('Total Tester Hours per Month by Tester', fontsize=16)
    plt.xlabel('Month', fontsize=12)
    plt.ylabel('Total Hours', fontsize=12)
    plt.legend(title='Tester #', bbox_to_anchor=(1.05, 1), loc='upper left')
    plt.tight_layout()
    plt.show() # 顯示第一張圖

    # --- 圖表 2: Engineering Hours ---
    plt.figure(figsize=(12, 6))
    sns.barplot(data=monthly_eng_hours, x='Month', y='Engineering Support Hours', hue='Name')
    plt.title('Total Engineering Hours per Month by Engineer', fontsize=16)
    plt.xlabel('Month', fontsize=12)
    plt.ylabel('Total Hours', fontsize=12)
    # 若名字太多，可以微調 legend 位置或大小
    plt.legend(title='Engineer Name', bbox_to_anchor=(1.05, 1), loc='upper left', ncol=2)
    plt.tight_layout()
    plt.show() # 顯示第二張圖

# 執行範例: 將檔案名稱替換為您實際的 Excel 檔案路徑
if __name__ == "__main__":
    file_path = "Google Monthly Engineering ATT_Service_Hours_Wafer Sort_ 2026__Eng_hours_revD.xlsx"
    analyze_hours(file_path)
