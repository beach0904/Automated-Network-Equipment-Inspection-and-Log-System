import sqlite3
import pandas as pd
import glob
import os
from datetime import datetime

class NetworkMonitorV110:
    def __init__(self, db_name="NetworkOps_v110.db"):
        self.db_name = db_name
        self.init_db()

    def init_db(self):
        """初始化符合 3NF 設計的資料庫架構"""
        conn = sqlite3.connect(self.db_name)
        cursor = conn.cursor()

        # 1. Areas 表 (第一正規化：確保原子性，抽離重複的區域資訊)
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS Areas (
                area_id INTEGER PRIMARY KEY AUTOINCREMENT,
                area_name TEXT UNIQUE NOT NULL
            )
        ''')

        # 2. Devices 表 (第二正規化：消除部分相依，設備基本資訊與巡檢數據分離)
        # 這裡儲存設備的靜態資訊
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS Devices (
                device_id INTEGER PRIMARY KEY AUTOINCREMENT,
                device_name TEXT UNIQUE NOT NULL,
                ip_address TEXT NOT NULL,
                area_id INTEGER,
                FOREIGN KEY (area_id) REFERENCES Areas(area_id)
            )
        ''')

        # 3. PerformanceLogs 表 (第三正規化：消除遞移相依)
        # 儲存隨時間變化的動態數據
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS PerformanceLogs (
                log_id INTEGER PRIMARY KEY AUTOINCREMENT,
                device_id INTEGER,
                cpu_usage INTEGER,
                mem_usage INTEGER,
                interface_status TEXT,
                diagnosis_result TEXT,
                inspect_time DATETIME,
                FOREIGN KEY (device_id) REFERENCES Devices(device_id)
            )
        ''')

        conn.commit()
        conn.close()

    def process_source_data(self, folder_path="source_data"):
        """讀取 Excel 並寫入 3NF 資料庫"""
        files = glob.glob(os.path.join(folder_path, "*.xlsx"))
        if not files:
            print("⚠️ 找不到來源檔案")
            return

        conn = sqlite3.connect(self.db_name)
        cursor = conn.cursor()

        for file in files:
            area_name = os.path.basename(file).replace(".xlsx", "")
            print(f"正在導入區域：{area_name}...")

            # 插入或取得 Area ID
            cursor.execute("INSERT OR IGNORE INTO Areas (area_name) VALUES (?)", (area_name,))
            cursor.execute("SELECT area_id FROM Areas WHERE area_name = ?", (area_name,))
            area_id = cursor.fetchone()[0]

            df = pd.read_excel(file)

            for _, row in df.iterrows():
                # 插入或取得 Device ID (靜態資訊)
                cursor.execute('''
                    INSERT OR IGNORE INTO Devices (device_name, ip_address, area_id)
                    VALUES (?, ?, ?)
                ''', (row['設備名稱'], row['管理IP'], area_id))
                
                cursor.execute("SELECT device_id FROM Devices WHERE device_name = ?", (row['設備名稱'],))
                device_id = cursor.fetchone()[0]

                # 判定邏輯
                status = row['介面狀態']
                cpu = row['CPU使用率(%)']
                diagnosis = "CRITICAL" if status == "Down" else ("WARNING" if cpu > 80 else "NORMAL")

                # 插入巡檢紀錄 (動態資訊)
                cursor.execute('''
                    INSERT INTO PerformanceLogs (device_id, cpu_usage, mem_usage, interface_status, diagnosis_result, inspect_time)
                    VALUES (?, ?, ?, ?, ?, ?)
                ''', (device_id, cpu, row['記憶體使用率(%)'], status, diagnosis, row['最後巡檢時間']))

        conn.commit()
        conn.close()
        print(f"✅ v1.1.0 數據導入完成，資料庫：{self.db_name}")

    def generate_report(self):
        """透過 SQL Join 產出報表，展現關聯式資料庫優勢"""
        conn = sqlite3.connect(self.db_name)
        query = '''
            SELECT 
                d.device_name, d.ip_address, a.area_name, 
                p.cpu_usage, p.interface_status, p.diagnosis_result, p.inspect_time
            FROM PerformanceLogs p
            JOIN Devices d ON p.device_id = d.device_id
            JOIN Areas a ON d.area_id = a.area_id
            WHERE p.diagnosis_result != 'NORMAL'
            ORDER BY p.inspect_time DESC
        '''
        error_df = pd.read_sql_query(query, conn)
        conn.close()
        
        if not error_df.empty:
            output_file = f"v110_異常追蹤報告_{datetime.now().strftime('%Y%m%d')}.xlsx"
            error_df.to_excel(output_file, index=False)
            print(f"🚩 已產出異常設備歷史報告：{output_file}")
        else:
            print("🙌 目前所有設備狀態正常。")

if __name__ == "__main__":
    monitor = NetworkMonitorV110()
    # 執行導入 (前提是 source_data 資料夾已有檔案)
    if os.path.exists("source_data"):
        monitor.process_source_data()
        monitor.generate_report()
    else:
        print("請先執行之前的腳本生成 source_data 資料夾與模擬數據。")