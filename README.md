網路維運分析報告與自動化模擬系統
本專案旨在結合 AI 模擬技術與 Python 自動化腳本，針對網路庫存資料進行深度分析，並自動生成維運報告。透過 AI 實作模擬，系統能更精準地預測網路設備的需求趨勢，提升運作效率。

🛠 核心功能
AI 數據模擬：利用模擬演算法產生接近真實環境的網路流量與庫存變動數據。

自動化資料擷取：透過 Python 腳本（main.py）自動抓取、整合相關網路維運資訊。

Excel 報表生成：自動產出具備版本控制概念（如 v1.0.0）的維運分析報告，節省人工製表時間。

Git 分支開發模式：採用 Feature Branch 工作流，確保功能更新（如「新增作品功能」）不影響主線穩定性。



📂 專案檔案說明
main.py: 主執行程式，負責 AI 模擬邏輯與資料處理流程。

source_data: 本專案的核心資料來源資料夾，存放用於分析的原始數據。

網路維運分析報告_20260429.xlsx: 系統自動產出的最終視覺化分析結果。

資料來源：https://www.kaggle.com/datasets/freshersstaff/it-system-performance-and-resource-metrics?resource=download
