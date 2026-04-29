import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from datetime import datetime
import glob
import os

def batch_network_analysis(input_folder="source_data"):
    # 1. 搜尋資料夾內所有 Excel 檔
    files = glob.glob(os.path.join(input_folder, "*.xlsx"))
    
    if not files:
        print(f"⚠️ 找不到檔案！請確保 '{input_folder}' 資料夾內有 .xlsx 檔案。")
        return

    all_data = []
    print(f"🔍 偵測到 {len(files)} 個區域檔案，開始批次處理...")

    # 2. 讀取並合併數據
    for file in files:
        df = pd.read_excel(file)
        # 增加一欄「來源區域」以便區分 (從檔名提取)
        area_name = os.path.basename(file).replace(".xlsx", "")
        df['來源區域'] = area_name
        all_data.append(df)

    final_df = pd.concat(all_data, ignore_index=True)

    # 3. 診斷邏輯
    def diagnose(row):
        if row['介面狀態'] == 'Down': return 'CRITICAL'
        if row['CPU使用率(%)'] > 80: return 'WARNING'
        return 'NORMAL'

    final_df['診斷結果'] = final_df.apply(diagnose, axis=1)

    # 4. 產出報告
    output_filename = f"全區域維運總報告_{datetime.now().strftime('%Y%m%d')}.xlsx"
    writer = pd.ExcelWriter(output_filename, engine='openpyxl')
    
    # 寫入分頁 1：詳細數據
    final_df.to_excel(writer, index=False, sheet_name='詳細巡檢清單')
    
    # 寫入分頁 2：異常統計 (主管最愛看)
    stats = final_df['診斷結果'].value_counts().to_frame()
    stats.to_excel(writer, sheet_name='異常統計總覽')

    # 5. 美化格式 (以第一頁為例)
    ws = writer.sheets['詳細巡檢清單']
    red_fill = PatternFill(start_color="FF9999", end_color="FF9999", fill_type="solid")
    
    for row_idx in range(2, len(final_df) + 2):
        res = ws.cell(row=row_idx, column=final_df.columns.get_loc('診斷結果')+1).value
        if res == 'CRITICAL':
            for col in range(1, len(final_df.columns) + 1):
                ws.cell(row=row_idx, column=col).fill = red_fill

    writer.close()
    print(f"✅ 批次處理完成！總報告已存至：{output_filename}")

if __name__ == "__main__":
    # 記得先建立資料夾並把你的模擬檔丟進去
    if not os.path.exists("source_data"):
        os.makedirs("source_data")
        print("📁 已自動建立 'source_data' 資料夾，請將模擬 Excel 檔移入此處。")
    else:
        batch_network_analysis("source_data")