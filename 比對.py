import pandas as pd
from fuzzywuzzy import fuzz
import os

# --- 設定檔案路徑和欄位名稱 ---
file_a_path = r'D:\Users\jason.hsieh\Desktop\ifare.xlsx'
file_b_path = r'D:\Users\jason.hsieh\Desktop\1957全國福利.xlsx'

title_col_a = 'Title'
time_col_a = 'UpdateTime'

title_col_b = 'title'
time_col_b = 'time'

output_file_name = '比對結果_相似度分析.xlsx'

# --- 讀取 Excel 檔案 ---
try:
    df_a = pd.read_excel(file_a_path)
    df_b = pd.read_excel(file_b_path)
    print(f"成功讀取檔案 A: {file_a_path}")
    print(f"成功讀取檔案 B: {file_b_path}")
except FileNotFoundError as e:
    print(f"錯誤：找不到檔案。請確認檔案路徑是否正確：{e}")
    # 在實際運行中，這裡可能會退出，但在這裡我們只印出錯誤資訊
    exit()
except KeyError as e:
    print(f"錯誤：指定的欄位名稱不存在。請確認欄位名稱是否正確且區分大小寫：{e}")
    # 提供現有欄位列表幫助使用者除錯
    if 'df_a' in locals():
        print(f"檔案 A 的欄位: {df_a.columns.tolist()}")
    if 'df_b' in locals():
        print(f"檔案 B 的欄位: {df_b.columns.tolist()}")
    exit()
except Exception as e:
    print(f"讀取檔案時發生未預期的錯誤: {e}")
    exit()

# --- 檢查必要的欄位是否存在 ---
if title_col_a not in df_a.columns or time_col_a not in df_a.columns:
    print(f"錯誤：檔案 A '{file_a_path}' 中缺少 '{title_col_a}' 或 '{time_col_a}' 欄位。")
    print(f"檔案 A 的欄位: {df_a.columns.tolist()}")
    exit()

if title_col_b not in df_b.columns or time_col_b not in df_b.columns:
    print(f"錯誤：檔案 B '{file_b_path}' 中缺少 '{title_col_b}' 或 '{time_col_b}' 欄位。")
    print(f"檔案 B 的欄位: {df_b.columns.tolist()}")
    exit()

# --- 儲存比對結果的列表 ---
results = []

# --- 逐一比對標題 ---
print("\n開始進行標題比對，這可能需要一些時間，請稍候...")
# 獲取 File B 的 title 和 time 欄位，並將它們轉換為列表以加速查詢
titles_b = df_b[title_col_b].astype(str).tolist()
times_b = df_b[time_col_b].tolist()

for index_a, row_a in df_a.iterrows():
    # 確保 title1 是字串，如果為 NaN 則轉換為空字串
    title1 = str(row_a[title_col_a]) if pd.notna(row_a[title_col_a]) else ""
    time1 = row_a[time_col_a]

    best_match_title2 = ""
    best_match_time2 = None
    highest_similarity = 0

    # 針對 File B 的每個標題進行比對
    for index_b, title2_val in enumerate(titles_b):
        current_similarity = fuzz.token_set_ratio(title1, title2_val)
 # 計算 Levenshtein 相似度 (0-100)

        if current_similarity > highest_similarity:
            highest_similarity = current_similarity
            best_match_title2 = title2_val
            best_match_time2 = times_b[index_b]

    results.append({
        'title1 (來自 ifare.xlsx)': title1,
        'time1 (來自 ifare.xlsx)': time1,
        'title2 (最相似的，來自 1957全國福利.xlsx)': best_match_title2,
        'time2 (對應的，來自 1957全國福利.xlsx)': best_match_time2,
        '相似度 (Levenshtein)': highest_similarity
    })

# --- 建立結果 DataFrame ---
result_df = pd.DataFrame(results)

# --- 儲存為新的 Excel 檔案 ---
try:
    result_df.to_excel(output_file_name, index=False)
    print(f"\n比對結果已成功儲存至檔案：'{output_file_name}'")
    # 顯示前幾行結果，以便快速檢查
    print("\n比對結果檔案的前 5 行預覽：")
    print(result_df.head())
except Exception as e:
    print(f"儲存結果檔案時發生錯誤：{e}")