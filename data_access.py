# -*- coding: utf-8 -*-
"""
士林電機 FA 業務戰情與報價系統 - 核心資料存取唯一窗口 (data_access.py)

【架構職責】
1. 集中管理所有資料庫連線、查詢與異動，前端與計算層絕不直接寫 SQL。
2. 100% 採用參數化查詢 (Parameterized Query)，杜絕 SQL 注入攻擊。
3. 嚴格實作三段式單號生成 (Prefix-Sequence-Revision) 與禁止覆寫舊單 (Append-Only)。
4. 預留環境變數切換機制，未來直升公司 SQL Server Express 時，僅需更動本檔連線引擎。
"""

import os
import sqlite3
from datetime import datetime, date, timedelta
from pathlib import Path
from typing import List, Dict, Any, Optional, Tuple

# 資料庫預設本機路徑 (採用 pathlib 跨平台路徑處理)
DEFAULT_DB_FILE = Path(__file__).resolve().parent / "quotation_system.db"


def get_connection(db_path: Optional[Path] = None) -> sqlite3.Connection:
    """取得資料庫連線，並開啟欄位字典映射與外鍵約束"""
    path = db_path or DEFAULT_DB_FILE
    conn = sqlite3.connect(str(path))
    conn.row_factory = sqlite3.Row  # 允許透過欄位名稱字典存取
    conn.execute("PRAGMA foreign_keys = ON;")  # 開啟外鍵完整性檢查
    return conn


def init_database(db_path: Optional[Path] = None) -> None:
    """
    初始化系統 13 張標準正規化資料表
    相容現階段 SQLite，結構直接符合未來 SQL Server Express 規格
    """
    conn = get_connection(db_path)
    cursor = conn.cursor()

    try:
        # 1. 業務每日日報表 (OGSM 行動層)
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS daily_report (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            report_date TEXT NOT NULL,
            sales_name TEXT NOT NULL,
            customer_id INTEGER,
            customer_name TEXT NOT NULL,
            contact_name TEXT,
            visit_type TEXT NOT NULL,
            content_summary TEXT NOT NULL,
            manager_task_id INTEGER,
            created_at TEXT NOT NULL DEFAULT (datetime('now', 'localtime'))
        );
        """)

        # 2. CRM 客戶商機表 (機會層)
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS crm_opportunity (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            daily_report_id INTEGER,
            customer_id INTEGER,
            customer_name TEXT NOT NULL,
            industry_type TEXT,
            product_series TEXT,
            est_quarter TEXT,
            est_scale REAL DEFAULT 0.0,
            status TEXT NOT NULL DEFAULT '挖掘中',
            created_at TEXT NOT NULL DEFAULT (datetime('now', 'localtime')),
            FOREIGN KEY (daily_report_id) REFERENCES daily_report(id)
        );
        """)

        # 3. Deal 專案案件表 (標案層)
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS deal_case (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            opportunity_id INTEGER,
            customer_id INTEGER,
            customer_name TEXT NOT NULL,
            project_name TEXT NOT NULL,
            project_type TEXT NOT NULL DEFAULT '一次性專案',
            comp_brand TEXT,
            comp_channel TEXT,
            est_deal_date TEXT,
            est_amount REAL DEFAULT 0.0,
            stage TEXT NOT NULL DEFAULT '初訪開拓',
            win_loss_reason TEXT,
            created_at TEXT NOT NULL DEFAULT (datetime('now', 'localtime')),
            FOREIGN KEY (opportunity_id) REFERENCES crm_opportunity(id)
        );
        """)

        # 4. 報價單表頭檔 (三段式單號，Append-Only 禁止 UPDATE 覆寫)
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS quotation_header (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            quote_no TEXT UNIQUE NOT NULL,
            prefix TEXT NOT NULL,
            sequence INTEGER NOT NULL,
            revision INTEGER NOT NULL DEFAULT 0,
            deal_id INTEGER,
            quote_date TEXT NOT NULL,
            valid_date TEXT NOT NULL,
            customer_name TEXT NOT NULL,
            tax_id TEXT,
            contact_name TEXT,
            phone TEXT,
            fax TEXT,
            delivery_address TEXT,
            payment_terms TEXT,
            project_name TEXT NOT NULL,
            project_type TEXT NOT NULL,
            subtotal REAL NOT NULL DEFAULT 0.0,
            tax_rate REAL NOT NULL DEFAULT 5.0,
            tax_amount REAL NOT NULL DEFAULT 0.0,
            grand_total REAL NOT NULL DEFAULT 0.0,
            status TEXT NOT NULL DEFAULT '追蹤中',
            is_deleted INTEGER NOT NULL DEFAULT 0,
            deleted_at TEXT,
            created_at TEXT NOT NULL DEFAULT (datetime('now', 'localtime')),
            updated_at TEXT NOT NULL DEFAULT (datetime('now', 'localtime')),
            FOREIGN KEY (deal_id) REFERENCES deal_case(id)
        );
        """)

        # 5. 報價單明細列檔
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS quotation_item (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            header_id INTEGER NOT NULL,
            quote_no TEXT NOT NULL,
            item_index INTEGER NOT NULL,
            model_no TEXT NOT NULL,
            spec TEXT,
            dealer_price REAL NOT NULL DEFAULT 0.0,
            discount_pct REAL,
            expected_price REAL,
            final_unit_price REAL NOT NULL DEFAULT 0.0,
            quantity REAL NOT NULL DEFAULT 1.0,
            row_subtotal REAL NOT NULL DEFAULT 0.0,
            note_lead_time TEXT,
            FOREIGN KEY (header_id) REFERENCES quotation_header(id) ON DELETE CASCADE
        );
        """)

        # 6. 專案售價申請單表
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS special_price_application (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            deal_id INTEGER,
            quotation_id INTEGER,
            quote_no TEXT NOT NULL,
            applicant_sales TEXT NOT NULL,
            apply_date TEXT NOT NULL,
            apply_reason TEXT NOT NULL,
            comp_price_info TEXT,
            target_subtotal REAL NOT NULL DEFAULT 0.0,
            est_margin_pct REAL,
            review_status TEXT NOT NULL DEFAULT '待主管審核',
            review_comment TEXT,
            created_at TEXT NOT NULL DEFAULT (datetime('now', 'localtime')),
            FOREIGN KEY (deal_id) REFERENCES deal_case(id),
            FOREIGN KEY (quotation_id) REFERENCES quotation_header(id)
        );
        """)

        # 7. 客戶主檔
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS customer_master (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT UNIQUE NOT NULL,
            tax_id TEXT,
            default_contact TEXT,
            phone TEXT,
            fax TEXT,
            address TEXT,
            payment_terms TEXT DEFAULT '月結30天'
        );
        """)

        # 8. 產品主檔 (士林 FA 工控型號庫)
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS product_master (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            model_no TEXT UNIQUE NOT NULL,
            spec TEXT NOT NULL,
            list_price REAL NOT NULL DEFAULT 0.0,
            dealer_price REAL NOT NULL DEFAULT 0.0,
            stock_qty INTEGER NOT NULL DEFAULT 0
        );
        """)

        # 9. 主管指派任務表
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS manager_task (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            title TEXT NOT NULL,
            target_customer TEXT,
            assigned_sales TEXT NOT NULL,
            deadline TEXT,
            status TEXT NOT NULL DEFAULT '待執行',
            created_at TEXT NOT NULL DEFAULT (datetime('now', 'localtime'))
        );
        """)

        # 10. 產業客戶攻略錦囊表
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS industry_strategy (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            industry_name TEXT UNIQUE NOT NULL,
            pain_points TEXT NOT NULL,
            main_products TEXT NOT NULL,
            competitor_defense TEXT NOT NULL
        );
        """)

        # 11. 戰情計分板指標設定檔 (成果 vs 驅動指標)
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS scoreboard_metric (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            sales_name TEXT NOT NULL,
            metric_type TEXT NOT NULL, -- '成果目標' 或 '驅動目標'
            metric_name TEXT NOT NULL,
            target_val REAL NOT NULL,
            actual_val REAL NOT NULL DEFAULT 0.0,
            unit TEXT DEFAULT '',
            week_idx INTEGER DEFAULT 1,
            lamp_status TEXT DEFAULT '綠燈', -- '綠燈', '黃燈', '紅燈'
            updated_at TEXT NOT NULL DEFAULT (datetime('now', 'localtime'))
        );
        """)

        # 12. 系統全域設定檔 (捨入位數、捨入方式、稅率)
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS system_settings (
            setting_key TEXT PRIMARY KEY,
            setting_val TEXT NOT NULL
        );
        """)

        # 14. 產品線領先指標 (驅動目標) 表
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS lead_driver_metric (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            sales_name TEXT NOT NULL,
            product_line TEXT NOT NULL, -- '士林伺服', '士林變頻器', '士林PLC'
            driver_name TEXT NOT NULL,
            target_val REAL NOT NULL,
            actual_val REAL NOT NULL DEFAULT 0.0,
            unit TEXT NOT NULL,
            period_type TEXT NOT NULL DEFAULT '週', -- '週', '月'
            start_date TEXT NOT NULL,
            end_date TEXT,
            status TEXT NOT NULL DEFAULT 'active', -- 'active', 'archived'
            created_at TEXT NOT NULL DEFAULT (datetime('now', 'localtime'))
        );
        """)

        # 15. 驅動目標更換申請表 (防呆審核與策略覆盤)
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS driver_change_request (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            driver_id INTEGER NOT NULL,
            sales_name TEXT NOT NULL,
            product_line TEXT NOT NULL,
            old_driver_name TEXT NOT NULL,
            new_driver_name TEXT NOT NULL,
            new_target_val REAL NOT NULL,
            new_unit TEXT NOT NULL,
            change_reason TEXT NOT NULL,
            apply_date TEXT NOT NULL,
            review_status TEXT NOT NULL DEFAULT '待審核', -- '待審核', '已核准', '已駁回'
            reviewer_comment TEXT,
            reviewed_at TEXT,
            created_at TEXT NOT NULL DEFAULT (datetime('now', 'localtime')),
            FOREIGN KEY (driver_id) REFERENCES lead_driver_metric(id)
        );
        """)

        # 13. 建立常用核心索引 (加速查歷史與過濾)
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_quote_prefix_seq ON quotation_header (prefix, sequence);")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_quote_customer ON quotation_header (customer_name);")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_quote_status_del ON quotation_header (status, is_deleted);")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_item_quote_no ON quotation_item (quote_no);")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_daily_sales_date ON daily_report (sales_name, report_date);")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_driver_sales_status ON lead_driver_metric (sales_name, status);")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_driver_req_sales_status ON driver_change_request (sales_name, review_status);")

        # 寫入預設設定檔與士林種子產品資料 (若為全新建庫)
        cursor.execute("INSERT OR IGNORE INTO system_settings (setting_key, setting_val) VALUES ('round_digits', '0');")
        cursor.execute("INSERT OR IGNORE INTO system_settings (setting_key, setting_val) VALUES ('round_mode', '四捨五入');")
        cursor.execute("INSERT OR IGNORE INTO system_settings (setting_key, setting_val) VALUES ('tax_rate', '5.0');")

        # 植入標準展示客戶與產品種子資料
        cursor.execute("""
        INSERT OR IGNORE INTO customer_master (name, tax_id, default_contact, phone, fax, address, payment_terms)
        VALUES ('赫力電機股份有限公司', '23456789', '陳經理', '02-2888-1234', '02-2888-5678', '台北市內湖區行愛路100號', '月結30天');
        """)
        cursor.execute("""
        INSERT OR IGNORE INTO customer_master (name, tax_id, default_contact, phone, fax, address, payment_terms)
        VALUES ('贊翔實業有限公司', '87654321', '陳姿蓉特助', '03-555-8888', '03-555-9999', '新竹市東區光復路一段200號', '票期60天');
        """)

        seed_products = [
            ("SE3-023-0.4K", "3Φ220V 0.5HP 高機能向量變頻器", 4200.0, 2520.0, 15),
            ("SE3-023-0.75K", "3Φ220V 1HP 高機能向量變頻器", 4800.0, 2880.0, 22),
            ("SE3-023-1.5K", "3Φ220V 2HP 高機能向量變頻器", 5600.0, 3360.0, 18),
            ("SE3-023-11K", "3Φ220V 15HP 高機能向量變頻器", 23500.0, 14100.0, 8),
            ("SS2-023-0.75K", "3Φ220V 1HP 簡易型向量變頻器", 3600.0, 2160.0, 30),
            ("AX5S-32MT-ES", "士林高階可程式控制器 16DI/16DO", 8900.0, 5340.0, 12),
        ]
        for m, s, lp, dp, sq in seed_products:
            cursor.execute("""
            INSERT OR IGNORE INTO product_master (model_no, spec, list_price, dealer_price, stock_qty)
            VALUES (?, ?, ?, ?, ?);
            """, (m, s, lp, dp, sq))

        # 植入標準領先指標 (驅動目標) 種子資料
        d_cnt = cursor.execute("SELECT COUNT(*) AS c FROM lead_driver_metric").fetchone()[0]
        if d_cnt == 0:
            today_d = date.today()
            d35_ago = (today_d - timedelta(days=35)).strftime("%Y-%m-%d")
            d14_ago = (today_d - timedelta(days=14)).strftime("%Y-%m-%d")

            seed_drivers = [
                # 曾維崧 (FA 全區主管，亦為業務) - 已執行 35 天 (滿 4 週)
                ("曾維崧", "士林伺服", "伺服關鍵客群專案洽談", 6.0, 5.0, "場", "月", d35_ago),
                ("曾維崧", "士林變頻器", "大型廠區變頻節能案場勘查", 8.0, 7.0, "場", "月", d35_ago),
                ("曾維崧", "士林PLC", "AX5S 高階 PLC 連線展示", 5.0, 5.0, "場", "月", d35_ago),

                # 曾仁君 (直賣業務主管，亦為業務) - 已執行 35 天 (滿 4 週)
                ("曾仁君", "士林伺服", "半導體設備伺服實機試機", 4.0, 3.0, "場", "週", d35_ago),
                ("曾仁君", "士林變頻器", "直賣變頻器升級汰換拜訪", 6.0, 5.0, "次", "週", d35_ago),
                ("曾仁君", "士林PLC", "專案自動化產線 PLC 規格制定", 3.0, 3.0, "案", "週", d35_ago),

                # 張何達 (經銷業務主管，亦為業務) - 已執行 35 天 (滿 4 週)
                ("張何達", "士林伺服", "經銷商伺服現貨在庫拜訪", 8.0, 8.0, "家", "週", d35_ago),
                ("張何達", "士林變頻器", "盤廠經銷通路變頻專案盤點", 10.0, 9.0, "家", "週", d35_ago),
                ("張何達", "士林PLC", "經銷商推廣 AX 系列技術交流", 6.0, 6.0, "場", "週", d35_ago),

                # 何宛茹 (直賣業務) - 已執行 35 天 (滿 4 週)
                ("何宛茹", "士林伺服", "直賣伺服節能設備實機測試", 4.0, 3.0, "場", "週", d35_ago),
                ("何宛茹", "士林變頻器", "新竹設備廠變頻器推廣洽談", 6.0, 5.0, "家", "週", d35_ago),
                ("何宛茹", "士林PLC", "包裝機設備 PLC 升級專案", 2.0, 2.0, "案", "週", d35_ago),

                # 葉仁豪 (經銷業務) - 已執行 14 天 (未滿 4 週，防呆阻擋展示)
                ("葉仁豪", "士林伺服", "伺服新樣品送樣驗證", 4.0, 1.0, "組", "週", d14_ago),
                ("葉仁豪", "士林變頻器", "變頻器舊換新推廣電訪", 12.0, 10.0, "通", "週", d14_ago),
                ("葉仁豪", "士林PLC", "PLC 搭配變頻配套展示", 4.0, 3.0, "次", "週", d14_ago),
            ]
            for s_name, p_line, d_name, t_val, a_val, unt, p_type, s_date in seed_drivers:
                cursor.execute("""
                INSERT INTO lead_driver_metric (
                    sales_name, product_line, driver_name, target_val, actual_val,
                    unit, period_type, start_date, status
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, 'active');
                """, (s_name, p_line, d_name, t_val, a_val, unt, p_type, s_date))

            # 植入一筆展示用變更申請單 (何宛茹 申請變更士林變頻器指標，供曾維崧/曾仁君測試審核)
            r_row = cursor.execute("""
            SELECT id FROM lead_driver_metric WHERE sales_name = '何宛茹' AND product_line = '士林變頻器' LIMIT 1
            """).fetchone()
            if r_row:
                cursor.execute("""
                INSERT INTO driver_change_request (
                    driver_id, sales_name, product_line, old_driver_name,
                    new_driver_name, new_target_val, new_unit, change_reason,
                    apply_date, review_status
                ) VALUES (?, '何宛茹', '士林變頻器', '新竹設備廠變頻器推廣洽談',
                          '半導體耗電產線大馬力變頻替換', 5.0, '場',
                          '原推廣洽談轉化率佳，配合節能政策聚焦大馬力機型實地能耗檢測。',
                          ?, '待審核');
                """, (r_row[0], str(today_d)))

        conn.commit()
    finally:
        conn.close()


# ==============================================================================
# 客戶與產品基礎主檔存取
# ==============================================================================

def get_customers(keyword: str = "") -> List[Dict[str, Any]]:
    """查詢客戶主檔 (支援關鍵字比對)"""
    conn = get_connection()
    try:
        if keyword:
            like_kw = f"%{keyword.strip()}%"
            rows = conn.execute("""
            SELECT * FROM customer_master
            WHERE name LIKE ? OR tax_id LIKE ? OR default_contact LIKE ?
            ORDER BY name ASC
            """, (like_kw, like_kw, like_kw)).fetchall()
        else:
            rows = conn.execute("SELECT * FROM customer_master ORDER BY name ASC").fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


def search_products(query: str = "", limit: int = 50) -> List[Dict[str, Any]]:
    """搜尋士林產品庫 (依型號或規格)"""
    conn = get_connection()
    try:
        if query:
            like_q = f"%{query.strip()}%"
            rows = conn.execute("""
            SELECT * FROM product_master
            WHERE model_no LIKE ? OR spec LIKE ?
            ORDER BY model_no ASC LIMIT ?
            """, (like_q, like_q, limit)).fetchall()
        else:
            rows = conn.execute("SELECT * FROM product_master ORDER BY model_no ASC LIMIT ?", (limit,)).fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


# ==============================================================================
# 業務日報 (OGSM 行動層) 存取函式
# ==============================================================================

def save_daily_report(report: Dict[str, Any]) -> int:
    """儲存一筆業務日報 (單一寫入點，參數化查詢)"""
    conn = get_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("""
        INSERT INTO daily_report (
            report_date, sales_name, customer_id, customer_name,
            contact_name, visit_type, content_summary, manager_task_id
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?);
        """, (
            report.get("report_date", str(date.today())),
            report["sales_name"],
            report.get("customer_id"),
            report["customer_name"],
            report.get("contact_name", ""),
            report.get("visit_type", "實地拜訪"),
            report["content_summary"],
            report.get("manager_task_id")
        ))
        conn.commit()
        return int(cursor.lastrowid or 0)
    finally:
        conn.close()


def get_daily_reports(start_date: Optional[str] = None, end_date: Optional[str] = None, sales_name: Optional[str] = None) -> List[Dict[str, Any]]:
    """查詢日報清單 (支援日期區間與業務姓名篩選)"""
    conn = get_connection()
    try:
        sql = "SELECT * FROM daily_report WHERE 1=1"
        params: List[Any] = []
        if start_date:
            sql += " AND report_date >= ?"
            params.append(start_date)
        if end_date:
            sql += " AND report_date <= ?"
            params.append(end_date)
        if sales_name:
            sql += " AND sales_name = ?"
            params.append(sales_name)
        sql += " ORDER BY report_date DESC, id DESC"
        rows = conn.execute(sql, params).fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


# ==============================================================================
# CRM 商機 (機會層) 與 Deal 案件 (標案層) 存取
# ==============================================================================

def save_crm_opportunity(opp: Dict[str, Any]) -> int:
    """新增或更新 CRM 商機"""
    conn = get_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("""
        INSERT INTO crm_opportunity (
            daily_report_id, customer_id, customer_name, industry_type,
            product_series, est_quarter, est_scale, status
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?);
        """, (
            opp.get("daily_report_id"),
            opp.get("customer_id"),
            opp["customer_name"],
            opp.get("industry_type", "自動化設備產業"),
            opp.get("product_series", "士林品"),
            opp.get("est_quarter", "Q3"),
            float(opp.get("est_scale", 0.0)),
            opp.get("status", "挖掘中")
        ))
        conn.commit()
        return int(cursor.lastrowid or 0)
    finally:
        conn.close()


def save_deal_case(deal: Dict[str, Any]) -> int:
    """由商機正式建立或直接新增 Deal 專案案件"""
    conn = get_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("""
        INSERT INTO deal_case (
            opportunity_id, customer_id, customer_name, project_name,
            project_type, comp_brand, comp_channel, est_deal_date,
            est_amount, stage, win_loss_reason
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);
        """, (
            deal.get("opportunity_id"),
            deal.get("customer_id"),
            deal["customer_name"],
            deal["project_name"],
            deal.get("project_type", "一次性專案"),
            deal.get("comp_brand", "台達"),
            deal.get("comp_channel", "經銷商"),
            deal.get("est_deal_date", str(date.today())),
            float(deal.get("est_amount", 0.0)),
            deal.get("stage", "初訪開拓"),
            deal.get("win_loss_reason", "")
        ))
        conn.commit()
        return int(cursor.lastrowid or 0)
    finally:
        conn.close()


def get_deal_cases(stage: Optional[str] = None) -> List[Dict[str, Any]]:
    """取得所有進行中或特定階段之 Deal 案件"""
    conn = get_connection()
    try:
        if stage:
            rows = conn.execute("SELECT * FROM deal_case WHERE stage = ? ORDER BY id DESC", (stage,)).fetchall()
        else:
            rows = conn.execute("SELECT * FROM deal_case ORDER BY id DESC").fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


# ==============================================================================
# 三段式報價單編號生成與 Append-Only 存取核心
# ==============================================================================

def generate_next_quote_no(prefix: str, sequence: Optional[int] = None) -> Tuple[str, int, int]:
    """
    依據 Master Prompt 嚴格生成三段式編號 [Prefix]-[Sequence]-[Revision]
    1. 新增報價 (sequence=None)：取該 Prefix 下的最大 sequence + 1，revision 固定為 0
    2. 修改既有單 (給定 sequence)：取該 Prefix+Sequence 下的最大 revision + 1
    回傳值：(完整單號, sequence, revision)
    """
    conn = get_connection()
    try:
        prefix_clean = prefix.strip()
        if sequence is None:
            # 新報價單：找目前同 Prefix 之最大 Sequence
            row = conn.execute("""
            SELECT MAX(sequence) AS max_seq FROM quotation_header WHERE prefix = ?
            """, (prefix_clean,)).fetchone()
            max_seq = row["max_seq"] if (row and row["max_seq"] is not None) else 0
            new_seq = max_seq + 1
            new_rev = 0
        else:
            # 既有報價單修改另存新版次：Sequence 不變，找最大 Revision + 1
            row = conn.execute("""
            SELECT MAX(revision) AS max_rev FROM quotation_header WHERE prefix = ? AND sequence = ?
            """, (prefix_clean, sequence)).fetchone()
            max_rev = row["max_rev"] if (row and row["max_rev"] is not None) else 0
            new_seq = sequence
            new_rev = max_rev + 1

        quote_no = f"{prefix_clean}-{new_seq:03d}-{new_rev:02d}"
        return quote_no, new_seq, new_rev
    finally:
        conn.close()


def save_quotation(header: Dict[str, Any], items: List[Dict[str, Any]]) -> str:
    """
    儲存報價單 (鐵律：Append-Only，嚴禁 UPDATE 覆寫舊報價資料)
    若傳入之單號已存在或要求新版次，自動以 Revision + 1 另存新筆資料
    回傳值：實際存入之唯一報價單號 quote_no
    """
    conn = get_connection()
    try:
        cursor = conn.cursor()
        prefix = header["prefix"].strip()
        req_seq = header.get("sequence")

        # 檢查該單號是否已被使用，若已使用則強制升級為新版次
        if req_seq is not None:
            quote_no, seq, rev = generate_next_quote_no(prefix, sequence=req_seq)
        else:
            quote_no, seq, rev = generate_next_quote_no(prefix)

        # 1. 寫入表頭
        cursor.execute("""
        INSERT INTO quotation_header (
            quote_no, prefix, sequence, revision, deal_id, quote_date, valid_date,
            customer_name, tax_id, contact_name, phone, fax, delivery_address,
            payment_terms, project_name, project_type, subtotal, tax_rate,
            tax_amount, grand_total, status, is_deleted
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 0);
        """, (
            quote_no, prefix, seq, rev,
            header.get("deal_id"),
            header.get("quote_date", str(date.today())),
            header.get("valid_date", str(date.today())),
            header["customer_name"],
            header.get("tax_id", ""),
            header.get("contact_name", ""),
            header.get("phone", ""),
            header.get("fax", ""),
            header.get("delivery_address", ""),
            header.get("payment_terms", ""),
            header["project_name"],
            header.get("project_type", "一次性專案"),
            float(header.get("subtotal", 0.0)),
            float(header.get("tax_rate", 5.0)),
            float(header.get("tax_amount", 0.0)),
            float(header.get("grand_total", 0.0)),
            header.get("status", "追蹤中")
        ))
        header_id = cursor.lastrowid

        # 2. 寫入各產品明細列
        for idx, itm in enumerate(items, start=1):
            cursor.execute("""
            INSERT INTO quotation_item (
                header_id, quote_no, item_index, model_no, spec, dealer_price,
                discount_pct, expected_price, final_unit_price, quantity,
                row_subtotal, note_lead_time
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);
            """, (
                header_id, quote_no, idx,
                itm["model_no"],
                itm.get("spec", ""),
                float(itm.get("dealer_price", 0.0)),
                float(itm["discount_pct"]) if itm.get("discount_pct") is not None else None,
                float(itm["expected_price"]) if itm.get("expected_price") is not None else None,
                float(itm.get("final_unit_price", 0.0)),
                float(itm.get("quantity", 1.0)),
                float(itm.get("row_subtotal", 0.0)),
                itm.get("note_lead_time", "")
            ))

        conn.commit()
        return quote_no
    finally:
        conn.close()


def get_quotation(quote_no: str) -> Optional[Dict[str, Any]]:
    """依三段式單號完整讀取報價單 (含表頭與明細清單)"""
    conn = get_connection()
    try:
        h_row = conn.execute("SELECT * FROM quotation_header WHERE quote_no = ?", (quote_no,)).fetchone()
        if not h_row:
            return None
        header_dict = dict(h_row)
        i_rows = conn.execute("SELECT * FROM quotation_item WHERE quote_no = ? ORDER BY item_index ASC", (quote_no,)).fetchall()
        header_dict["items"] = [dict(r) for r in i_rows]
        return header_dict
    finally:
        conn.close()


def list_quotations(keyword: str = "", status: str = "ALL", include_deleted: bool = False) -> List[Dict[str, Any]]:
    """查詢報價單清單 (支援多欄位關鍵字檢索、狀態篩選與垃圾桶過濾)"""
    conn = get_connection()
    try:
        sql = "SELECT * FROM quotation_header WHERE 1=1"
        params: List[Any] = []

        if not include_deleted:
            sql += " AND is_deleted = 0"
        else:
            sql += " AND is_deleted = 1"

        if status and status != "ALL":
            sql += " AND status = ?"
            params.append(status)

        if keyword:
            like_kw = f"%{keyword.strip()}%"
            sql += " AND (quote_no LIKE ? OR customer_name LIKE ? OR project_name LIKE ? OR contact_name LIKE ?)"
            params.extend([like_kw, like_kw, like_kw, like_kw])

        sql += " ORDER BY id DESC"
        rows = conn.execute(sql, params).fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


def soft_delete_quotation(quote_no: str) -> bool:
    """軟刪除報價單 (移至垃圾桶，保留 30 天防手滑緩衝期)"""
    conn = get_connection()
    try:
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        res = conn.execute("""
        UPDATE quotation_header SET is_deleted = 1, deleted_at = ? WHERE quote_no = ?
        """, (now_str, quote_no))
        conn.commit()
        return res.rowcount > 0
    finally:
        conn.close()


def restore_quotation(quote_no: str) -> bool:
    """自垃圾桶一鍵復原報價單"""
    conn = get_connection()
    try:
        res = conn.execute("""
        UPDATE quotation_header SET is_deleted = 0, deleted_at = NULL WHERE quote_no = ?
        """, (quote_no,))
        conn.commit()
        return res.rowcount > 0
    finally:
        conn.close()


# ==============================================================================
# 專案售價申請單 (連鎖帶出情資與經銷底價)
# ==============================================================================

def save_special_price_application(app_data: Dict[str, Any]) -> int:
    """儲存一筆專案售價申請單"""
    conn = get_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("""
        INSERT INTO special_price_application (
            deal_id, quotation_id, quote_no, applicant_sales, apply_date,
            apply_reason, comp_price_info, target_subtotal, est_margin_pct,
            review_status
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?);
        """, (
            app_data.get("deal_id"),
            app_data.get("quotation_id"),
            app_data["quote_no"],
            app_data["applicant_sales"],
            app_data.get("apply_date", str(date.today())),
            app_data["apply_reason"],
            app_data.get("comp_price_info", ""),
            float(app_data.get("target_subtotal", 0.0)),
            float(app_data.get("est_margin_pct", 0.0)),
            app_data.get("review_status", "待主管審核")
        ))
        conn.commit()
        return int(cursor.lastrowid or 0)
    finally:
        conn.close()


# ==============================================================================
# 產品線領先指標 (驅動目標) 與變更審核管理存取
# ==============================================================================

def get_active_drivers(sales_name: str) -> List[Dict[str, Any]]:
    """
    取得該業務員三大產品線當前生效中的領先指標 (驅動目標)
    計算達成率、對標紅黃綠燈（>=90% 綠燈、75%~89.9% 黃燈、<75% 紅燈）
    與執行天數 (檢驗是否滿 28 天 / 4 週)
    """
    conn = get_connection()
    try:
        rows = conn.execute("""
        SELECT * FROM lead_driver_metric
        WHERE sales_name = ? AND status = 'active'
        ORDER BY CASE product_line
            WHEN '士林伺服' THEN 1
            WHEN '士林變頻器' THEN 2
            WHEN '士林PLC' THEN 3
            ELSE 4 END ASC
        """, (sales_name,)).fetchall()

        # 若該業務員無紀錄，自動以通用範本建立並寫入
        if not rows:
            default_templates = [
                ("士林伺服", "伺服客群實地拜訪與技術交流", 4.0, 3.0, "場", "週"),
                ("士林變頻器", "變頻器節能案場勘查與推廣", 6.0, 5.0, "次", "週"),
                ("士林PLC", "PLC 與自動化連線架構制定", 3.0, 2.0, "案", "週"),
            ]
            init_date = (date.today() - timedelta(days=30)).strftime("%Y-%m-%d")
            for p_line, d_name, t_val, a_val, unt, p_type in default_templates:
                conn.execute("""
                INSERT INTO lead_driver_metric (
                    sales_name, product_line, driver_name, target_val, actual_val,
                    unit, period_type, start_date, status
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, 'active');
                """, (sales_name, p_line, d_name, t_val, a_val, unt, p_type, init_date))
            conn.commit()
            rows = conn.execute("""
            SELECT * FROM lead_driver_metric
            WHERE sales_name = ? AND status = 'active'
            ORDER BY CASE product_line
                WHEN '士林伺服' THEN 1
                WHEN '士林變頻器' THEN 2
                WHEN '士林PLC' THEN 3
                ELSE 4 END ASC
            """, (sales_name,)).fetchall()

        result = []
        today_dt = date.today()
        for r in rows:
            d = dict(r)
            t_val = float(d.get("target_val", 0.0))
            a_val = float(d.get("actual_val", 0.0))
            rate = round((a_val / t_val * 100.0), 1) if t_val > 0 else 0.0
            d["rate_pct"] = rate

            # 依據規則判定燈號：>=90% 綠燈、75%~89.9% 黃燈、<75% 紅燈
            if rate >= 90.0:
                d["traffic_light"] = "green"
                d["light_text"] = "達標"
                d["bar_class"] = "bar-success"
            elif rate >= 75.0:
                d["traffic_light"] = "yellow"
                d["light_text"] = "預警"
                d["bar_class"] = "bar-warning"
            elif rate >= 50.0:
                d["traffic_light"] = "red"
                d["light_text"] = "低標"
                d["bar_class"] = "bar-alert"
            else:
                d["traffic_light"] = "red"
                d["light_text"] = "嚴重落後"
                d["bar_class"] = "bar-danger"

            # 計算自 start_date 起執行天數與週數 (防呆審核判斷)
            try:
                start_dt = datetime.strptime(d["start_date"], "%Y-%m-%d").date()
                days_active = max(0, (today_dt - start_dt).days)
            except Exception:
                days_active = 0

            d["days_active"] = days_active
            d["weeks_active"] = round(days_active / 7.0, 1)
            d["can_change"] = (days_active >= 28)
            d["days_remaining"] = max(0, 28 - days_active)
            result.append(d)

        return result
    finally:
        conn.close()


def apply_driver_change(data: Dict[str, Any]) -> Tuple[bool, str, Optional[int]]:
    """
    業務發起更換驅動指標申請
    強制防呆：當前指標自 start_date 起執行未滿 28 天 (4 週) 則阻擋申請並回傳提示
    回傳值：(成功與否, 提示訊息, 申請單號)
    """
    conn = get_connection()
    try:
        driver_id = data.get("driver_id")
        if not driver_id:
            return False, "缺少欲變更之指標識別碼 (driver_id)", None

        row = conn.execute("SELECT * FROM lead_driver_metric WHERE id = ? AND status = 'active'", (driver_id,)).fetchone()
        if not row:
            return False, "找不到該筆生效中指標紀錄", None

        driver = dict(row)
        today_dt = date.today()
        try:
            start_dt = datetime.strptime(driver["start_date"], "%Y-%m-%d").date()
            days_active = max(0, (today_dt - start_dt).days)
        except Exception:
            days_active = 0

        # 強制防呆：未滿 28 天 (4 週) 阻擋申請
        if days_active < 28:
            rem = 28 - days_active
            return False, f"該指標自 {driver['start_date']} 起僅執行 {days_active} 天，未滿 4 週（28 天）。為避免指標頻繁變更導致行動失焦，尚無法提出更換申請（尚需執行 {rem} 天）。", None

        # 檢查是否有同指標的待審核申請
        pending = conn.execute("""
        SELECT id FROM driver_change_request
        WHERE driver_id = ? AND review_status = '待審核'
        """, (driver_id,)).fetchone()
        if pending:
            return False, "該指標已有一筆變更申請正在審核中，請勿重複提出申請！", pending["id"]

        cursor = conn.cursor()
        cursor.execute("""
        INSERT INTO driver_change_request (
            driver_id, sales_name, product_line, old_driver_name,
            new_driver_name, new_target_val, new_unit, change_reason,
            apply_date, review_status
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, '待審核');
        """, (
            driver_id,
            driver["sales_name"],
            driver["product_line"],
            driver["driver_name"],
            data["new_driver_name"].strip(),
            float(data.get("new_target_val", 0.0)),
            data.get("new_unit", "次").strip(),
            data.get("change_reason", "").strip(),
            str(today_dt)
        ))
        conn.commit()
        req_id = int(cursor.lastrowid or 0)
        return True, "驅動目標變更申請已成功送出，請靜候主管審核！", req_id
    finally:
        conn.close()


def review_driver_change(request_id: int, action: str, reviewer_name: str = "", comment: str = "") -> Tuple[bool, str]:
    """
    主管審核驅動目標變更申請
    action: 'approve' (核准) 或 'reject' (駁回)
    若核准：舊指標轉為 'archived' 並註記 end_date，插入新指標為 'active'（歷史進度完整留存，禁止直接 UPDATE 覆寫）
    若駁回：保留原指標並註記主管反饋意見
    """
    conn = get_connection()
    try:
        row = conn.execute("SELECT * FROM driver_change_request WHERE id = ?", (request_id,)).fetchone()
        if not row:
            return False, "找不到該筆申請紀錄"

        req = dict(row)
        if req["review_status"] != "待審核":
            return False, f"該申請單已處理完成（目前狀態：{req['review_status']}），無法重複審核"

        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        today_str = str(date.today())
        comment_clean = comment.strip()
        if reviewer_name:
            comment_with_reviewer = f"[{reviewer_name}] {comment_clean}" if comment_clean else f"[{reviewer_name}] 審核完成"
        else:
            comment_with_reviewer = comment_clean

        if action == "approve":
            # 1. 封存舊指標 (註記 end_date，不抹除舊資料)
            conn.execute("""
            UPDATE lead_driver_metric
            SET status = 'archived', end_date = ?
            WHERE id = ?;
            """, (today_str, req["driver_id"]))

            # 2. 插入新指標為 active (初始實績為 0.0)
            conn.execute("""
            INSERT INTO lead_driver_metric (
                sales_name, product_line, driver_name, target_val,
                actual_val, unit, period_type, start_date, status
            ) VALUES (?, ?, ?, ?, 0.0, ?, '週', ?, 'active');
            """, (
                req["sales_name"],
                req["product_line"],
                req["new_driver_name"],
                float(req["new_target_val"]),
                req["new_unit"],
                today_str
            ))

            # 3. 更新申請單狀態為「已核准」
            conn.execute("""
            UPDATE driver_change_request
            SET review_status = '已核准', reviewer_comment = ?, reviewed_at = ?
            WHERE id = ?;
            """, (comment_with_reviewer, now_str, request_id))

            conn.commit()
            return True, f"申請單已核准！舊指標已封存，新指標「{req['new_driver_name']}」已正式生效。"

        elif action == "reject":
            # 駁回：原指標維持 active，僅更新申請單狀態
            conn.execute("""
            UPDATE driver_change_request
            SET review_status = '已駁回', reviewer_comment = ?, reviewed_at = ?
            WHERE id = ?;
            """, (comment_with_reviewer, now_str, request_id))
            conn.commit()
            return True, "已駁回該變更申請，原領先指標繼續維持生效。"
        else:
            return False, f"不支援的審核動作：{action}"
    finally:
        conn.close()


def get_driver_change_requests(sales_name: Optional[str] = None, status: Optional[str] = None) -> List[Dict[str, Any]]:
    """
    查詢驅動目標變更申請單清單 (供業務查詢個人申請或主管審核待辦)
    """
    conn = get_connection()
    try:
        sql = "SELECT * FROM driver_change_request WHERE 1=1"
        params: List[Any] = []
        if sales_name:
            sql += " AND sales_name = ?"
            params.append(sales_name)
        if status and status != "ALL":
            sql += " AND review_status = ?"
            params.append(status)
        sql += " ORDER BY id DESC"
        rows = conn.execute(sql, params).fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


# ==============================================================================
# 模組自我驗證入口
# ==============================================================================

if __name__ == "__main__":
    print("[data_access.py] 正在初始化士林電機 FA 資料庫...")
    init_database()
    print("[data_access.py] 資料庫建表與種子資料植入成功！")
