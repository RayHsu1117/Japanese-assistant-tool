import os
import pandas as pd
from openpyxl import load_workbook

# 設定包含所有資料夾的主目錄
root_folder = '..\A103001131 (csv-9)'
excel_file = 'output2.xlsx'

# 如果 Excel 檔案存在，則加載；否則直接創建一個新的
if os.path.exists(excel_file):
    try:
        book = load_workbook(excel_file)
        print(f"加載現有的 Excel 檔案: {excel_file}")
    except Exception as e:
        print(f"無法加載現有的 Excel 檔案: {e}")
        book = None
else:
    book = None
    print(f"請創建新的 Excel 檔案: {excel_file}")

# 開始將每個資料夾中的 CSV 檔案匯入 Excel
with pd.ExcelWriter(excel_file, engine='openpyxl', mode='a' if book else 'w',if_sheet_exists='overlay') as writer:
    # if book:
    #     writer.book = book  # 如果 Excel 文件已存在，則加載它

    # 遍歷主資料夾中的每個子資料夾與檔案
    for dirpath, dirnames, filenames in os.walk(root_folder):
        for filename in filenames:
            # 如果該檔案是 CSV 檔案
            if filename.endswith('.csv'):
                # 取得 CSV 檔案的完整路徑
                csv_path = os.path.join(dirpath, filename)

                try:
                    # 讀取 CSV 檔案
                    df = pd.read_csv(csv_path)

                    # 處理退選學生
                    df['Status'] = df['成績(Score)'].apply(lambda x: '退選' if x == '退選' else '正常')

                    # 取得資料夾名稱和檔名，形成工作表名稱
                    folder_name = os.path.basename(dirpath)  # 取得資料夾名稱
                    sheet_name = f"{folder_name}_{os.path.splitext(filename)[0]}"  # 資料夾名稱_檔名

                    # 限制工作表名稱不超過 31 個字元
                    sheet_name = sheet_name[:5]

                    # 如果工作表已經存在，進行比對
                    if sheet_name in writer.book.sheetnames:
                        # print(f"工作表 {sheet_name} 已經存在，開始比對數據")

                        # 讀取現有的 Excel 工作表
                        existing_df = pd.read_excel(excel_file, sheet_name=sheet_name)

                        # 比對新加入的學生
                        new_students = df[~df['學號(Student ID)'].isin(existing_df['學號(Student ID)'])]
                        if not new_students.empty:
                            print(f"發現新加入的學生：\n{new_students}")

                        # 比對已退課的學生，通過 Score 標記為退選
                        removed_students = df[df['Status'] == '退選']
                        if not removed_students.empty:
                            print(f"發現已退課的學生：\n{removed_students}")

                            # 找出棄選的學生：現有名單中有，但新名單中不存在的學生
                        dropped_students = existing_df[~existing_df['學號(Student ID)'].isin(df['學號(Student ID)'])]
                        if not dropped_students.empty:
                            print(f"發現棄選的學生（已從名單消失）：\n{dropped_students}")

                            # 在 Score 欄位標記為 "棄選"
                            existing_df.loc[existing_df['學號(Student ID)'].isin(dropped_students['學號(Student ID)']), '成績(Score)'] = '棄選'

                        # 合併新舊名單，更新狀態並保持退選標記
                        updated_df = pd.concat([existing_df, new_students]).drop_duplicates(subset=['學號(Student ID)'], keep='last')

                        # 更新狀態，將退選學生標註為「退選」
                        for index, row in df.iterrows():
                            if row['Status'] == '退選':
                                updated_df.loc[updated_df['學號(Student ID)'] == row['學號(Student ID)'], 'Status'] = '退選'

                        # 更新工作表
                        updated_df.to_excel(writer, sheet_name=sheet_name, index=False)
                        # print(f"已更新工作表：{sheet_name}")

                    else:
                        # 如果工作表不存在，直接創建並寫入新數據
                        df.to_excel(writer, sheet_name=sheet_name, index=False)
                        print(f"成功匯入 CSV: {csv_path} 到工作表: {sheet_name}")

                except Exception as e:
                    print(f"無法讀取 CSV 檔案: {csv_path}，錯誤: {e}")

# Excel 檔案會自動保存
print(f"所有資料夾中的 CSV 檔案已成功匯入 {excel_file}，每個 CSV 對應一個工作表。")
