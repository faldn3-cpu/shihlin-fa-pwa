# -*- coding: utf-8 -*-
"""
士林電機 FA 業務戰情與報價系統 - 核心資料存取唯一窗口 (data_access.py)

【架構職責】
1. 集中管理所有資料庫連線、查詢與異動，前端與計算層絕不直接寫 SQL。
2. 100% 採用參數化查詢 (Parameterized Query)，杜絕 SQL 注入攻擊。
3. 嚴格實作三段式單號生成 (Prefix-Sequence-Revision) 與禁止覆寫舊單 (Append-Only)。
4. 落實多人資料隔離：客戶主檔、歷史報價單、手機草稿、個人特殊品皆依業務姓名完全獨立分開。
5. 落實方案 A：共通標準牌價表與業務個人特殊品牌價表雙層架構；當特殊品正式發行於標準牌價表時，自動以公司標準品取代並將個人特殊品封存留紀錄。
6. 預留環境變數切換機制，未來直升公司 SQL Server Express 時，僅需更動本檔連線引擎。
"""

import json
import os
import re
import sqlite3
from datetime import datetime, date, timedelta
from pathlib import Path
from typing import List, Dict, Any, Optional, Tuple

# 資料庫預設本機路徑 (採用 pathlib 跨平台路徑處理)
DEFAULT_DB_FILE = Path(__file__).resolve().parent / "quotation_system.db"
LEGACY_APP_DIR = Path(__file__).resolve().parent / "之前做的應用程式跟網頁" / "_QuotationApp"


def get_connection(db_path: Optional[Path] = None) -> sqlite3.Connection:
    """取得資料庫連線，並開啟欄位字典映射與外鍵約束"""
    path = db_path or DEFAULT_DB_FILE
    conn = sqlite3.connect(str(path))
    conn.row_factory = sqlite3.Row  # 允許透過欄位名稱字典存取
    conn.execute("PRAGMA foreign_keys = ON;")  # 開啟外鍵完整性檢查
    return conn


def _ensure_column(cursor: sqlite3.Cursor, table_name: str, col_name: str, col_def: str) -> None:
    """安全檢查並為既有資料表補齊欄位 (向後相容升級)"""
    cols = [r[1] for r in cursor.execute(f"PRAGMA table_info({table_name})").fetchall()]
    if col_name not in cols:
        cursor.execute(f"ALTER TABLE {table_name} ADD COLUMN {col_name} {col_def};")


def normalize_model_key(text: str) -> str:
    """
    產品型號純化比對鍵：
    1. 去除括號內含中文字之備註（如 (訂購品)、(客製品)）
    2. 去除連字號、斜線與空白，保留英數與小數點並轉小寫
    """
    if not text:
        return ""
    t = str(text).strip()
    t = re.sub(r'[\(（][^\)）]*[\u4e00-\u9fa5]+[^\)）]*[\)）]', '', t)
    t = re.sub(r'[\(（][^\)）]*[\u4e00-\u9fa5]+.*$', '', t)
    t = re.sub(r'[^a-zA-Z0-9.]', '', t)
    return t.lower()


def init_database(db_path: Optional[Path] = None) -> None:
    """
    初始化系統正規化資料表
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

        # 4. 報價單表頭檔 (三段式單號，Append-Only 禁止 UPDATE 覆寫，綁定 sales_name 獨立隔離)
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS quotation_header (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            quote_no TEXT UNIQUE NOT NULL,
            prefix TEXT NOT NULL,
            sequence INTEGER NOT NULL,
            revision INTEGER NOT NULL DEFAULT 0,
            sales_name TEXT NOT NULL DEFAULT '曾維崧',
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
            project_name TEXT NOT NULL DEFAULT '',
            project_type TEXT NOT NULL DEFAULT '一次性專案',
            subtotal REAL NOT NULL DEFAULT 0.0,
            tax_rate REAL NOT NULL DEFAULT 5.0,
            tax_amount REAL NOT NULL DEFAULT 0.0,
            grand_total REAL NOT NULL DEFAULT 0.0,
            status TEXT NOT NULL DEFAULT '追蹤中',
            is_latest INTEGER NOT NULL DEFAULT 1,
            is_won INTEGER NOT NULL DEFAULT 0,
            won_amount REAL NOT NULL DEFAULT 0.0,
            won_date TEXT,
            is_lost INTEGER NOT NULL DEFAULT 0,
            is_ignored INTEGER NOT NULL DEFAULT 0,
            lost_reason TEXT,
            lost_memo TEXT,
            source_type TEXT NOT NULL DEFAULT 'PC_DIRECT',
            is_deleted INTEGER NOT NULL DEFAULT 0,
            deleted_at TEXT,
            created_at TEXT NOT NULL DEFAULT (datetime('now', 'localtime')),
            updated_at TEXT NOT NULL DEFAULT (datetime('now', 'localtime')),
            FOREIGN KEY (deal_id) REFERENCES deal_case(id)
        );
        """)
        _ensure_column(cursor, "quotation_header", "sales_name", "TEXT NOT NULL DEFAULT '曾維崧'")
        _ensure_column(cursor, "quotation_header", "is_latest", "INTEGER NOT NULL DEFAULT 1")
        _ensure_column(cursor, "quotation_header", "is_won", "INTEGER NOT NULL DEFAULT 0")
        _ensure_column(cursor, "quotation_header", "won_amount", "REAL NOT NULL DEFAULT 0.0")
        _ensure_column(cursor, "quotation_header", "won_date", "TEXT")
        _ensure_column(cursor, "quotation_header", "is_lost", "INTEGER NOT NULL DEFAULT 0")
        _ensure_column(cursor, "quotation_header", "is_ignored", "INTEGER NOT NULL DEFAULT 0")
        _ensure_column(cursor, "quotation_header", "lost_reason", "TEXT")
        _ensure_column(cursor, "quotation_header", "lost_memo", "TEXT")
        _ensure_column(cursor, "quotation_header", "source_type", "TEXT NOT NULL DEFAULT 'PC_DIRECT'")

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
            is_won INTEGER NOT NULL DEFAULT 0,
            is_lost INTEGER NOT NULL DEFAULT 0,
            item_status TEXT DEFAULT 'pending',
            lost_reason TEXT,
            FOREIGN KEY (header_id) REFERENCES quotation_header(id) ON DELETE CASCADE
        );
        """)
        _ensure_column(cursor, "quotation_item", "is_won", "INTEGER NOT NULL DEFAULT 0")
        _ensure_column(cursor, "quotation_item", "is_lost", "INTEGER NOT NULL DEFAULT 0")
        _ensure_column(cursor, "quotation_item", "item_status", "TEXT DEFAULT 'pending'")
        _ensure_column(cursor, "quotation_item", "lost_reason", "TEXT")

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

        # 7. 客戶主檔 (綁定 owner_sales，各業務客戶完全獨立分開)
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS customer_master (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT UNIQUE NOT NULL,
            owner_sales TEXT NOT NULL DEFAULT '曾維崧',
            tax_id TEXT,
            default_contact TEXT,
            phone TEXT,
            fax TEXT,
            address TEXT,
            payment_terms TEXT DEFAULT '月結30天',
            status TEXT DEFAULT 'active',
            default_discount REAL DEFAULT 1.0,
            is_favorite INTEGER DEFAULT 0
        );
        """)
        _ensure_column(cursor, "customer_master", "owner_sales", "TEXT NOT NULL DEFAULT '曾維崧'")
        _ensure_column(cursor, "customer_master", "status", "TEXT DEFAULT 'active'")
        _ensure_column(cursor, "customer_master", "default_discount", "REAL DEFAULT 1.0")
        _ensure_column(cursor, "customer_master", "is_favorite", "INTEGER DEFAULT 0")

        # 7b. 客戶聯絡人表 (支援同一客戶多位聯絡窗口)
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS contact_master (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            customer_id INTEGER NOT NULL,
            owner_sales TEXT NOT NULL DEFAULT '曾維崧',
            name TEXT NOT NULL,
            email TEXT DEFAULT '',
            mobile TEXT DEFAULT '',
            FOREIGN KEY (customer_id) REFERENCES customer_master(id) ON DELETE CASCADE
        );
        """)

        # 8. 共通標準牌價表 (全體業務共用之公開標準產品庫)
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS product_master (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            model_no TEXT UNIQUE NOT NULL,
            spec TEXT NOT NULL,
            list_price REAL NOT NULL DEFAULT 0.0,
            dealer_price REAL NOT NULL DEFAULT 0.0,
            stock_qty INTEGER NOT NULL DEFAULT 0,
            sku TEXT DEFAULT '',
            category TEXT DEFAULT '',
            is_order_item INTEGER NOT NULL DEFAULT 0,
            stock_status TEXT DEFAULT 'IN_STOCK'
        );
        """)
        _ensure_column(cursor, "product_master", "sku", "TEXT DEFAULT ''")
        _ensure_column(cursor, "product_master", "category", "TEXT DEFAULT ''")
        _ensure_column(cursor, "product_master", "is_order_item", "INTEGER NOT NULL DEFAULT 0")
        _ensure_column(cursor, "product_master", "stock_status", "TEXT DEFAULT 'IN_STOCK'")

        # 8b. 業務個人特殊品牌價表 (非公開資料，各業務獨立；方案 A：轉標準品時自動封存留紀錄)
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS personal_special_product (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            owner_sales TEXT NOT NULL,
            model_no TEXT NOT NULL,
            spec TEXT NOT NULL DEFAULT '',
            list_price REAL NOT NULL DEFAULT 0.0,
            dealer_price REAL NOT NULL DEFAULT 0.0,
            stock_qty INTEGER NOT NULL DEFAULT 0,
            note TEXT DEFAULT '',
            is_archived INTEGER NOT NULL DEFAULT 0,
            archived_at TEXT,
            archived_standard_price REAL,
            archived_note TEXT DEFAULT '',
            created_at TEXT NOT NULL DEFAULT (datetime('now', 'localtime')),
            UNIQUE(owner_sales, model_no)
        );
        """)

        # 8c. 手機報價草稿資料表 (手機端送出直接入庫，不經雲端硬碟，回電腦桌直接載入)
        cursor.execute("""
        CREATE TABLE IF NOT EXISTS quotation_draft (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            sales_name TEXT NOT NULL DEFAULT '曾維崧',
            customer_id INTEGER,
            customer_name TEXT NOT NULL,
            tax_id TEXT DEFAULT '',
            contact_name TEXT DEFAULT '',
            phone TEXT DEFAULT '',
            fax TEXT DEFAULT '',
            delivery_address TEXT DEFAULT '',
            payment_terms TEXT DEFAULT '月結 30 天電匯',
            project_name TEXT DEFAULT '',
            project_type TEXT DEFAULT '一次性專案',
            deal_id INTEGER,
            crm_links_json TEXT DEFAULT '[]',
            items_json TEXT NOT NULL DEFAULT '[]',
            subtotal REAL NOT NULL DEFAULT 0.0,
            item_count INTEGER NOT NULL DEFAULT 0,
            status TEXT NOT NULL DEFAULT 'pending',
            created_at TEXT NOT NULL DEFAULT (datetime('now', 'localtime')),
            updated_at TEXT NOT NULL DEFAULT (datetime('now', 'localtime'))
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
            metric_type TEXT NOT NULL,
            metric_name TEXT NOT NULL,
            target_val REAL NOT NULL,
            actual_val REAL NOT NULL DEFAULT 0.0,
            unit TEXT DEFAULT '',
            week_idx INTEGER DEFAULT 1,
            lamp_status TEXT DEFAULT '綠燈',
            updated_at TEXT NOT NULL DEFAULT (datetime('now', 'localtime'))
        );
        """)

        # 12. 系統全域設定檔 (捨入位數、捨入方式、稅率、庫存更新時間)
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
            product_line TEXT NOT NULL,
            driver_name TEXT NOT NULL,
            target_val REAL NOT NULL,
            actual_val REAL NOT NULL DEFAULT 0.0,
            unit TEXT NOT NULL,
            period_type TEXT NOT NULL DEFAULT '週',
            start_date TEXT NOT NULL,
            end_date TEXT,
            status TEXT NOT NULL DEFAULT 'active',
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
            review_status TEXT NOT NULL DEFAULT '待審核',
            reviewer_comment TEXT,
            reviewed_at TEXT,
            created_at TEXT NOT NULL DEFAULT (datetime('now', 'localtime')),
            FOREIGN KEY (driver_id) REFERENCES lead_driver_metric(id)
        );
        """)

        # 建立常用核心索引 (加速查詢與業務隔離過濾)
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_quote_prefix_seq ON quotation_header (prefix, sequence);")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_quote_customer ON quotation_header (customer_name);")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_quote_sales ON quotation_header (sales_name, is_deleted);")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_quote_status_del ON quotation_header (status, is_deleted);")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_item_quote_no ON quotation_item (quote_no);")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_cust_owner ON customer_master (owner_sales, name);")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_spec_prod_owner ON personal_special_product (owner_sales, is_archived);")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_draft_sales ON quotation_draft (sales_name, status);")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_daily_sales_date ON daily_report (sales_name, report_date);")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_driver_sales_status ON lead_driver_metric (sales_name, status);")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_driver_req_sales_status ON driver_change_request (sales_name, review_status);")

        # 寫入預設財務與庫存設定 (對齊舊版 _QuotationApp：百位數 / 無條件捨去 / 5% 稅率)
        cursor.execute("INSERT OR IGNORE INTO system_settings (setting_key, setting_val) VALUES ('round_digits', '2');")
        cursor.execute("INSERT OR IGNORE INTO system_settings (setting_key, setting_val) VALUES ('round_mode', '無條件捨去');")
        cursor.execute("INSERT OR IGNORE INTO system_settings (setting_key, setting_val) VALUES ('tax_rate', '5.0');")
        cursor.execute("INSERT OR IGNORE INTO system_settings (setting_key, setting_val) VALUES ('stock_updated_at', '10/15 10:12');")

        # 植入標準展示客戶與產品種子資料 (當資料庫為空時確保基本測試通過)
        cursor.execute("""
        INSERT OR IGNORE INTO customer_master (name, owner_sales, tax_id, default_contact, phone, fax, address, payment_terms)
        VALUES ('赫力電機股份有限公司', '曾維崧', '23456789', '陳經理', '02-2888-1234', '02-2888-5678', '台北市內湖區行愛路100號', '月結30天');
        """)
        cursor.execute("""
        INSERT OR IGNORE INTO customer_master (name, owner_sales, tax_id, default_contact, phone, fax, address, payment_terms)
        VALUES ('贊翔實業有限公司', '曾維崧', '87654321', '陳姿蓉特助', '03-555-8888', '03-555-9999', '新竹市東區光復路一段200號', '票期60天');
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
                ("曾維崧", "士林伺服", "伺服關鍵客群專案洽談", 6.0, 5.0, "場", "月", d35_ago),
                ("曾維崧", "士林變頻器", "大型廠區變頻節能案場勘查", 8.0, 7.0, "場", "月", d35_ago),
                ("曾維崧", "士林PLC", "AX5S 高階 PLC 連線展示", 5.0, 5.0, "場", "月", d35_ago),
                ("曾仁君", "士林伺服", "半導體設備伺服實機試機", 4.0, 3.0, "場", "週", d35_ago),
                ("曾仁君", "士林變頻器", "直賣變頻器升級汰換拜訪", 6.0, 5.0, "次", "週", d35_ago),
                ("曾仁君", "士林PLC", "專案自動化產線 PLC 規格制定", 3.0, 3.0, "案", "週", d35_ago),
                ("張何達", "士林伺服", "經銷商伺服現貨在庫拜訪", 8.0, 8.0, "家", "週", d35_ago),
                ("張何達", "士林變頻器", "盤廠經銷通路變頻專案盤點", 10.0, 9.0, "家", "週", d35_ago),
                ("張何達", "士林PLC", "經銷商推廣 AX 系列技術交流", 6.0, 6.0, "場", "週", d35_ago),
                ("何宛茹", "士林伺服", "直賣伺服節能設備實機測試", 4.0, 3.0, "場", "週", d35_ago),
                ("何宛茹", "士林變頻器", "新竹設備廠變頻器推廣洽談", 6.0, 5.0, "家", "週", d35_ago),
                ("何宛茹", "士林PLC", "包裝機設備 PLC 升級專案", 2.0, 2.0, "案", "週", d35_ago),
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
# 方案 A：特殊品轉公司標準品自動取代與封存機制
# ==============================================================================

def sync_special_to_standard_archive(conn: Optional[sqlite3.Connection] = None) -> int:
    """
    【方案 A 核心防呆機制】
    檢查所有業務手頭上「生效中 (is_archived = 0)」的個人特殊品，
    一旦發現該型號 (model_no) 已經正式存在於「共通標準牌價表 (product_master)」中：
    1. 自動將該筆個人特殊品標記為 is_archived = 1（已轉標準品封存）
    2. 記錄封存時間、轉標準品後之官方經銷價與完整價格對照說明
    3. 確保業務搜尋型號時只會出現 1 筆最新的公司標準品，不會重複打架
    回傳值：本次自動轉封存之筆數
    """
    own_conn = False
    if conn is None:
        conn = get_connection()
        own_conn = True

    try:
        cursor = conn.cursor()
        active_specials = cursor.execute("""
        SELECT id, owner_sales, model_no, dealer_price
        FROM personal_special_product
        WHERE is_archived = 0
        """).fetchall()

        if not active_specials:
            return 0

        archived_count = 0
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        today_str = datetime.now().strftime("%Y/%m/%d")

        for sp in active_specials:
            sp_model = str(sp["model_no"]).strip()
            std_row = cursor.execute("""
            SELECT model_no, dealer_price
            FROM product_master
            WHERE UPPER(TRIM(model_no)) = UPPER(?)
            LIMIT 1
            """, (sp_model,)).fetchone()

            if std_row:
                old_p = int(round(float(sp["dealer_price"] or 0)))
                std_p = int(round(float(std_row["dealer_price"] or 0)))
                note_msg = (
                    f"此型號已於 {today_str} 正式發行於公司共通標準牌價表"
                    f"（原個人特殊價 ${old_p:,} -> 現公司標準價 ${std_p:,}），系統已自動改用標準品並封存此紀錄。"
                )
                cursor.execute("""
                UPDATE personal_special_product
                SET is_archived = 1,
                    archived_at = ?,
                    archived_standard_price = ?,
                    archived_note = ?
                WHERE id = ?
                """, (now_str, float(std_row["dealer_price"] or 0.0), note_msg, sp["id"]))
                archived_count += 1

        if archived_count > 0 and own_conn:
            conn.commit()
        return archived_count
    finally:
        if own_conn:
            conn.close()


def get_personal_special_products(sales_name: str, include_archived: bool = True) -> List[Dict[str, Any]]:
    """取得特定業務員個人的特殊品經銷牌價表（嚴格依業務姓名隔離，他人不可見）"""
    conn = get_connection()
    try:
        sync_special_to_standard_archive(conn)
        conn.commit()
        sales_clean = (sales_name or "曾維崧").strip()
        if include_archived:
            rows = conn.execute("""
            SELECT * FROM personal_special_product
            WHERE owner_sales = ?
            ORDER BY is_archived ASC, id DESC
            """, (sales_clean,)).fetchall()
        else:
            rows = conn.execute("""
            SELECT * FROM personal_special_product
            WHERE owner_sales = ? AND is_archived = 0
            ORDER BY id DESC
            """, (sales_clean,)).fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


def save_personal_special_product(data: Dict[str, Any]) -> Tuple[bool, str, int]:
    """
    新增或更新業務個人的特殊品經銷牌價
    若該型號已經在公司共通標準牌價表中，則依據【方案 A】自動標記為已轉標準品封存並提示業務
    """
    owner_sales = str(data.get("owner_sales") or "曾維崧").strip()
    model_no = str(data.get("model_no") or "").strip()
    if not model_no:
        return False, "產品型號不得為空白", 0

    spec = str(data.get("spec") or model_no).strip()
    dealer_price = float(data.get("dealer_price") or 0.0)
    list_price = float(data.get("list_price") or dealer_price)
    stock_qty = int(data.get("stock_qty") or 0)
    note = str(data.get("note") or "").strip()

    conn = get_connection()
    try:
        cursor = conn.cursor()
        cursor.execute("""
        INSERT INTO personal_special_product (
            owner_sales, model_no, spec, list_price, dealer_price, stock_qty, note, is_archived
        ) VALUES (?, ?, ?, ?, ?, ?, ?, 0)
        ON CONFLICT(owner_sales, model_no) DO UPDATE SET
            spec = excluded.spec,
            list_price = excluded.list_price,
            dealer_price = excluded.dealer_price,
            stock_qty = excluded.stock_qty,
            note = excluded.note,
            is_archived = 0,
            archived_at = NULL,
            archived_standard_price = NULL,
            archived_note = ''
        """, (owner_sales, model_no, spec, list_price, dealer_price, stock_qty, note))

        row = cursor.execute("""
        SELECT id FROM personal_special_product WHERE owner_sales = ? AND model_no = ?
        """, (owner_sales, model_no)).fetchone()
        item_id = int(row["id"]) if row else 0

        # 執行方案 A 自動檢查：若該型號已經在共通標準牌價表中，立即轉為標準品封存
        archived_cnt = sync_special_to_standard_archive(conn)
        conn.commit()

        if archived_cnt > 0:
            return (
                True,
                f"型號「{model_no}」目前已存在於公司共通標準牌價表中！依方案 A 規則，系統已自動以公司標準品取代，並將您的個人特殊品轉入封存紀錄。",
                item_id,
            )
        return True, f"個人特殊品「{model_no}」已成功儲存（僅 {owner_sales} 個人可見）！", item_id
    finally:
        conn.close()


def delete_personal_special_product(item_id: int, sales_name: str) -> bool:
    """刪除業務個人的特殊品項目（防呆驗證 owner_sales 確保只能刪自己的）"""
    conn = get_connection()
    try:
        res = conn.execute("""
        DELETE FROM personal_special_product WHERE id = ? AND owner_sales = ?
        """, (item_id, sales_name.strip()))
        conn.commit()
        return res.rowcount > 0
    finally:
        conn.close()


def promote_special_to_standard(item_id: int, standard_dealer_price: Optional[float] = None) -> Tuple[bool, str]:
    """
    【方案 A 特殊品正式轉為公司標準品發行】
    將指定的特殊品正式寫入共通標準牌價表 (product_master)，
    並自動觸發 sync_special_to_standard_archive 將所有業務手上的同型號特殊品轉為封存留紀錄！
    """
    conn = get_connection()
    try:
        cursor = conn.cursor()
        sp = cursor.execute("SELECT * FROM personal_special_product WHERE id = ?", (item_id,)).fetchone()
        if not sp:
            return False, "找不到該筆個人特殊品紀錄"

        model_no = str(sp["model_no"]).strip()
        spec = str(sp["spec"] or model_no).strip()
        final_std_price = float(standard_dealer_price) if standard_dealer_price is not None else float(sp["dealer_price"] or 0.0)
        list_price = float(sp["list_price"] or final_std_price)
        stock_qty = int(sp["stock_qty"] or 0)

        cursor.execute("""
        INSERT INTO product_master (model_no, spec, list_price, dealer_price, stock_qty, is_order_item)
        VALUES (?, ?, ?, ?, ?, 0)
        ON CONFLICT(model_no) DO UPDATE SET
            spec = excluded.spec,
            list_price = excluded.list_price,
            dealer_price = excluded.dealer_price,
            stock_qty = excluded.stock_qty
        """, (model_no, spec, list_price, final_std_price, stock_qty))

        archived_cnt = sync_special_to_standard_archive(conn)
        conn.commit()
        return (
            True,
            f"型號「{model_no}」已正式升格發行至【公司共通標準牌價表】（標準經銷價 ${int(round(final_std_price)):,}），並自動將 {archived_cnt} 筆個人特殊品轉為封存留紀錄！",
        )
    finally:
        conn.close()


# ==============================================================================
# 客戶與產品基礎主檔存取 (支援業務隔離與雙層牌價合併搜尋)
# ==============================================================================

def get_customers(keyword: str = "", sales_name: Optional[str] = None) -> List[Dict[str, Any]]:
    """
    查詢客戶主檔 (支援關鍵字比對與業務姓名嚴格隔離)
    若有傳入 sales_name，則只回傳該業務名下之客戶與聯絡人清單
    """
    conn = get_connection()
    try:
        sql = "SELECT * FROM customer_master WHERE 1=1"
        params: List[Any] = []

        if sales_name:
            sql += " AND owner_sales = ?"
            params.append(sales_name.strip())

        if keyword:
            like_kw = f"%{keyword.strip()}%"
            sql += " AND (name LIKE ? OR tax_id LIKE ? OR default_contact LIKE ? OR phone LIKE ?)"
            params.extend([like_kw, like_kw, like_kw, like_kw])

        sql += " ORDER BY is_favorite DESC, name ASC"
        rows = conn.execute(sql, params).fetchall()

        # 讀取聯絡人清單並附加於各客戶字典中
        contact_rows = conn.execute("SELECT * FROM contact_master ORDER BY id ASC").fetchall()
        contacts_by_cid: Dict[int, List[Dict[str, Any]]] = {}
        for cr in contact_rows:
            cid = cr["customer_id"]
            contacts_by_cid.setdefault(cid, []).append(dict(cr))

        result = []
        for r in rows:
            d = dict(r)
            c_list = contacts_by_cid.get(d["id"], [])
            if not d.get("default_contact") and c_list:
                d["default_contact"] = c_list[0].get("name", "")
            d["contacts"] = c_list
            result.append(d)
        return result
    finally:
        conn.close()


def save_customer(data: Dict[str, Any]) -> int:
    """新增或更新業務個人所屬之客戶主檔"""
    conn = get_connection()
    try:
        cursor = conn.cursor()
        name = str(data["name"]).strip()
        owner_sales = str(data.get("owner_sales") or "曾維崧").strip()
        tax_id = str(data.get("tax_id") or "").strip()
        default_contact = str(data.get("default_contact") or "").strip()
        phone = str(data.get("phone") or "").strip()
        fax = str(data.get("fax") or "").strip()
        address = str(data.get("address") or "").strip()
        payment_terms = str(data.get("payment_terms") or "月結 30 天電匯").strip()

        cursor.execute("""
        INSERT INTO customer_master (
            name, owner_sales, tax_id, default_contact, phone, fax, address, payment_terms
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        ON CONFLICT(name) DO UPDATE SET
            owner_sales = excluded.owner_sales,
            tax_id = excluded.tax_id,
            default_contact = excluded.default_contact,
            phone = excluded.phone,
            fax = excluded.fax,
            address = excluded.address,
            payment_terms = excluded.payment_terms
        """, (name, owner_sales, tax_id, default_contact, phone, fax, address, payment_terms))
        conn.commit()
        row = cursor.execute("SELECT id FROM customer_master WHERE name = ?", (name,)).fetchone()
        return int(row["id"]) if row else 0
    finally:
        conn.close()


def search_products(query: str = "", limit: int = 80, sales_name: Optional[str] = None) -> List[Dict[str, Any]]:
    """
    搜尋產品牌價庫 (同時涵蓋：1. 該業務未封存之個人特殊品 + 2. 全公司共通標準牌價表)
    搜尋前自動執行方案 A 檢查，確保已轉標準品之型號不會重複出現
    """
    conn = get_connection()
    try:
        sync_special_to_standard_archive(conn)
        conn.commit()

        results: List[Dict[str, Any]] = []
        q_clean = (query or "").strip()
        like_q = f"%{q_clean}%"

        # 1. 先查詢該業務員生效中的「個人特殊品」(若未指定 sales_name 則預設曾維崧)
        target_sales = (sales_name or "曾維崧").strip()
        if q_clean:
            sp_rows = conn.execute("""
            SELECT * FROM personal_special_product
            WHERE owner_sales = ? AND is_archived = 0
              AND (model_no LIKE ? OR spec LIKE ? OR note LIKE ?)
            ORDER BY model_no ASC LIMIT ?
            """, (target_sales, like_q, like_q, like_q, limit)).fetchall()
        else:
            sp_rows = conn.execute("""
            SELECT * FROM personal_special_product
            WHERE owner_sales = ? AND is_archived = 0
            ORDER BY model_no ASC LIMIT ?
            """, (target_sales, min(limit, 30))).fetchall()

        for sp in sp_rows:
            d = dict(sp)
            d["is_special"] = 1
            d["source_label"] = "個人特殊品"
            d["is_order_item"] = 0
            results.append(d)

        # 2. 再查詢全公司「共通標準牌價表」
        rem_limit = max(10, limit - len(results))
        if q_clean:
            std_rows = conn.execute("""
            SELECT * FROM product_master
            WHERE model_no LIKE ? OR spec LIKE ? OR sku LIKE ?
            ORDER BY
                CASE WHEN UPPER(model_no) = UPPER(?) THEN 0
                     WHEN UPPER(model_no) LIKE UPPER(?) THEN 1
                     ELSE 2 END,
                model_no ASC
            LIMIT ?
            """, (like_q, like_q, like_q, q_clean, f"{q_clean}%", rem_limit)).fetchall()
        else:
            std_rows = conn.execute("""
            SELECT * FROM product_master
            ORDER BY stock_qty DESC, model_no ASC LIMIT ?
            """, (rem_limit,)).fetchall()

        for st in std_rows:
            d = dict(st)
            d["is_special"] = 0
            d["source_label"] = "共通標準品"
            results.append(d)

        return results
    finally:
        conn.close()


# ==============================================================================
# 報價單財務捨入與系統設定存取
# ==============================================================================

def get_quotation_finance_settings() -> Dict[str, Any]:
    """讀取報價單財務設定（稅率、捨入位數、捨入方式、庫存更新時間、報價條款）"""
    conn = get_connection()
    try:
        rows = conn.execute("SELECT setting_key, setting_val FROM system_settings").fetchall()
        kv = {r["setting_key"]: r["setting_val"] for r in rows}

        raw_digit = int(kv.get("round_digits", "2"))
        if raw_digit in (0, 1, 2, 3):
            round_factor = 10 ** raw_digit
        else:
            round_factor = raw_digit if raw_digit in (1, 10, 100, 1000) else 100

        digit_label_map = {1: "個位數", 10: "十位數", 100: "百位數", 1000: "千位數"}
        round_mode = kv.get("round_mode", "無條件捨去")
        if round_mode.upper() == "FLOOR":
            round_mode = "無條件捨去"
        elif round_mode.upper() in ("CEIL", "CEILING"):
            round_mode = "無條件進位"
        elif round_mode.upper() == "ROUND":
            round_mode = "四捨五入"

        return {
            "tax_rate": float(kv.get("tax_rate", "5.0")),
            "round_digits": raw_digit,
            "round_factor": round_factor,
            "round_digit_label": digit_label_map.get(round_factor, "百位數"),
            "round_mode": round_mode,
            "round_summary_text": f"捨入：{digit_label_map.get(round_factor, '百位數')} / {round_mode}",
            "stock_updated_at": kv.get("stock_updated_at", "10/15 10:12"),
            "quote_terms": kv.get("quote_terms", ""),
        }
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
# 三段式報價單編號生成、Append-Only 存取與業務完全隔離
# ==============================================================================

def generate_next_quote_no(prefix: str, sequence: Optional[int] = None) -> Tuple[str, int, int]:
    """
    嚴格生成三段式編號 [Prefix]-[Sequence]-[Revision]
    1. 新增報價 (sequence=None)：取該 Prefix 下的最大 sequence + 1，revision 固定為 0
    2. 修改既有單 (給定 sequence)：取該 Prefix+Sequence 下的最大 revision + 1
    回傳值：(完整單號, sequence, revision)
    """
    conn = get_connection()
    try:
        prefix_clean = prefix.strip()
        if sequence is None:
            row = conn.execute("""
            SELECT MAX(sequence) AS max_seq FROM quotation_header WHERE prefix = ?
            """, (prefix_clean,)).fetchone()
            max_seq = row["max_seq"] if (row and row["max_seq"] is not None) else 0
            new_seq = max_seq + 1
            new_rev = 0
        else:
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
    儲存報價單 (鐵律：Append-Only，嚴禁 UPDATE 覆寫舊報價資料，並綁定 sales_name)
    若傳入之單號已存在或要求新版次，自動以 Revision + 1 另存新筆資料
    回傳值：實際存入之唯一報價單號 quote_no
    """
    conn = get_connection()
    try:
        cursor = conn.cursor()
        prefix = header["prefix"].strip()
        req_seq = header.get("sequence")
        sales_name = str(header.get("sales_name") or "曾維崧").strip()

        if req_seq is not None:
            quote_no, seq, rev = generate_next_quote_no(prefix, sequence=req_seq)
        else:
            quote_no, seq, rev = generate_next_quote_no(prefix)

        # 將同 prefix + sequence 舊版標記 is_latest = 0 (不改動舊版內容與金額)
        cursor.execute("""
        UPDATE quotation_header SET is_latest = 0 WHERE prefix = ? AND sequence = ?
        """, (prefix, seq))

        cursor.execute("""
        INSERT INTO quotation_header (
            quote_no, prefix, sequence, revision, sales_name, deal_id, quote_date, valid_date,
            customer_name, tax_id, contact_name, phone, fax, delivery_address,
            payment_terms, project_name, project_type, subtotal, tax_rate,
            tax_amount, grand_total, status, is_latest, is_deleted
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 1, 0);
        """, (
            quote_no, prefix, seq, rev, sales_name,
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
            header.get("project_name", ""),
            header.get("project_type", "一次性專案"),
            float(header.get("subtotal", 0.0)),
            float(header.get("tax_rate", 5.0)),
            float(header.get("tax_amount", 0.0)),
            float(header.get("grand_total", 0.0)),
            header.get("status", "追蹤中")
        ))
        header_id = cursor.lastrowid

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
        i_rows = conn.execute("SELECT * FROM quotation_item WHERE header_id = ? ORDER BY item_index ASC", (header_dict["id"],)).fetchall()
        if not i_rows:
            i_rows = conn.execute("SELECT * FROM quotation_item WHERE quote_no = ? ORDER BY item_index ASC", (quote_no,)).fetchall()
        header_dict["items"] = [dict(r) for r in i_rows]
        return header_dict
    finally:
        conn.close()


def list_quotations(
    keyword: str = "",
    status: str = "ALL",
    include_deleted: bool = False,
    sales_name: Optional[str] = None,
    limit: int = 200
) -> List[Dict[str, Any]]:
    """查詢報價單清單 (支援業務員完全隔離、關鍵字檢索、狀態篩選與垃圾桶過濾)"""
    conn = get_connection()
    try:
        sql = "SELECT * FROM quotation_header WHERE 1=1"
        params: List[Any] = []

        if sales_name:
            sql += " AND sales_name = ?"
            params.append(sales_name.strip())

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

        sql += " ORDER BY quote_date DESC, id DESC LIMIT ?"
        params.append(limit)
        rows = conn.execute(sql, params).fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


def analyze_historical_prices(
    model_keyword: str = "",
    customer_keyword: str = "",
    sales_name: Optional[str] = None,
    limit: int = 100
) -> List[Dict[str, Any]]:
    """查詢歷史單價與折讓紀錄（依業務員 sales_name 完全獨立隔離）"""
    conn = get_connection()
    try:
        sql = """
        SELECT
            i.model_no,
            i.spec,
            i.dealer_price,
            i.discount_pct,
            i.final_unit_price,
            i.quantity,
            i.row_subtotal,
            i.note_lead_time,
            h.quote_no,
            h.quote_date,
            h.customer_name,
            h.project_name,
            h.status,
            h.sales_name
        FROM quotation_item i
        JOIN quotation_header h ON i.header_id = h.id
        WHERE h.is_deleted = 0
          AND i.model_no != '[小計]'
        """
        params: List[Any] = []

        if sales_name:
            sql += " AND h.sales_name = ?"
            params.append(sales_name.strip())

        if model_keyword:
            like_m = f"%{model_keyword.strip()}%"
            sql += " AND (i.model_no LIKE ? OR i.spec LIKE ?)"
            params.extend([like_m, like_m])

        if customer_keyword:
            like_c = f"%{customer_keyword.strip()}%"
            sql += " AND h.customer_name LIKE ?"
            params.append(like_c)

        sql += " ORDER BY h.quote_date DESC, h.id DESC LIMIT ?"
        params.append(limit)
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
# 手機報價草稿匣存取 (手機端送出直接存入資料庫，回到電腦桌直接載入)
# ==============================================================================

def save_quotation_draft(draft_data: Dict[str, Any]) -> int:
    """將手機端或暫存報價草稿直接存入資料庫 quotation_draft 表 (依 sales_name 隔離)"""
    conn = get_connection()
    try:
        cursor = conn.cursor()
        sales_name = str(draft_data.get("sales_name") or "曾維崧").strip()
        customer_name = str(draft_data.get("customer_name") or "未命名客戶").strip()
        items_list = draft_data.get("items", [])
        if isinstance(items_list, str):
            items_json = items_list
            try:
                parsed = json.loads(items_list)
                item_count = len(parsed) if isinstance(parsed, list) else 0
            except Exception:
                item_count = 0
        else:
            items_json = json.dumps(items_list, ensure_ascii=False)
            item_count = len(items_list)

        crm_links = draft_data.get("crm_links", [])
        crm_links_json = crm_links if isinstance(crm_links, str) else json.dumps(crm_links, ensure_ascii=False)

        cursor.execute("""
        INSERT INTO quotation_draft (
            sales_name, customer_id, customer_name, tax_id, contact_name,
            phone, fax, delivery_address, payment_terms, project_name,
            project_type, deal_id, crm_links_json, items_json, subtotal, item_count, status
        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 'pending');
        """, (
            sales_name,
            draft_data.get("customer_id"),
            customer_name,
            str(draft_data.get("tax_id") or ""),
            str(draft_data.get("contact_name") or ""),
            str(draft_data.get("phone") or ""),
            str(draft_data.get("fax") or ""),
            str(draft_data.get("delivery_address") or ""),
            str(draft_data.get("payment_terms") or "月結 30 天電匯"),
            str(draft_data.get("project_name") or ""),
            str(draft_data.get("project_type") or "一次性專案"),
            draft_data.get("deal_id"),
            crm_links_json,
            items_json,
            float(draft_data.get("subtotal") or 0.0),
            item_count,
        ))
        conn.commit()
        return int(cursor.lastrowid or 0)
    finally:
        conn.close()


def get_quotation_drafts(sales_name: Optional[str] = None) -> List[Dict[str, Any]]:
    """讀取該業務員的手機報價草稿匣清單"""
    conn = get_connection()
    try:
        sql = "SELECT * FROM quotation_draft WHERE status = 'pending'"
        params: List[Any] = []
        if sales_name:
            sql += " AND sales_name = ?"
            params.append(sales_name.strip())
        sql += " ORDER BY id DESC"
        rows = conn.execute(sql, params).fetchall()
        result = []
        for r in rows:
            d = dict(r)
            try:
                d["items"] = json.loads(d.get("items_json") or "[]")
            except Exception:
                d["items"] = []
            try:
                d["crm_links"] = json.loads(d.get("crm_links_json") or "[]")
            except Exception:
                d["crm_links"] = []
            result.append(d)
        return result
    finally:
        conn.close()


def delete_quotation_draft(draft_id: int, sales_name: Optional[str] = None) -> bool:
    """刪除單筆手機報價草稿"""
    conn = get_connection()
    try:
        if sales_name:
            res = conn.execute("DELETE FROM quotation_draft WHERE id = ? AND sales_name = ?", (draft_id, sales_name.strip()))
        else:
            res = conn.execute("DELETE FROM quotation_draft WHERE id = ?", (draft_id,))
        conn.commit()
        return res.rowcount > 0
    finally:
        conn.close()


def clear_quotation_drafts(sales_name: str) -> int:
    """清空特定業務員的所有手機報價草稿"""
    conn = get_connection()
    try:
        res = conn.execute("DELETE FROM quotation_draft WHERE sales_name = ?", (sales_name.strip(),))
        conn.commit()
        return res.rowcount
    finally:
        conn.close()


# ==============================================================================
# 舊版 _QuotationApp 真實資料無痛遷移整合引擎
# ==============================================================================

def migrate_from_legacy_quotation_app(
    db_path: Optional[Path] = None,
    legacy_dir: Optional[Path] = None,
    default_owner: str = "曾維崧"
) -> Dict[str, Any]:
    """
    將 `之前做的應用程式跟網頁/_QuotationApp` 內的真實資料安全合併至本系統資料庫：
    1. 53 筆客戶與 53 筆聯絡人 -> 歸屬 default_owner ('曾維崧')
    2. 9,975 筆產品 + L 廠庫存快取 (stock_cache_fg_*.json) -> 寫入共通標準牌價表 product_master
    3. 742 筆歷史報價單與 4,623 筆明細 -> 歸屬 default_owner ('曾維崧')
    4. 系統設定 (百位數 / 無條件捨去 / 庫存更新時間) -> 寫入 system_settings
    5. 100% 保留現有的日報、商機、案件與領先指標資料
    """
    init_database(db_path)
    src_dir = legacy_dir or LEGACY_APP_DIR
    old_db_path = src_dir / "quotation_system.db"
    if not old_db_path.exists():
        return {"status": "skipped", "reason": f"找不到舊資料庫: {old_db_path}"}

    # 1. 讀取 L 廠最新庫存快取與更新時間
    stock_map: Dict[str, Tuple[int, str]] = {}
    stock_updated_label = "10/15 10:12"
    backups_dir = src_dir / "Backups"
    if backups_dir.exists():
        meta_file = backups_dir / "stock_meta.json"
        if meta_file.exists():
            try:
                meta_obj = json.loads(meta_file.read_text(encoding="utf-8"))
                raw_date = str(meta_obj.get("sheet_mod_date") or "")
                # 將 "2026/10/15 AM 10 : 12" 簡化為清爽格式或保留
                m = re.search(r'\d{4}/(\d{1,2}/\d{1,2})\s+(?:AM|PM)?\s*(\d{1,2})\s*:\s*(\d{2})', raw_date)
                if m:
                    stock_updated_label = f"{m.group(1)} {int(m.group(2)):02d}:{m.group(3)}"
                elif raw_date:
                    stock_updated_label = raw_date
            except Exception:
                pass

        cache_files = sorted(backups_dir.glob("stock_cache_fg_*.json"), reverse=True)
        if cache_files:
            try:
                stock_list = json.loads(cache_files[0].read_text(encoding="utf-8"))
                for item in stock_list:
                    spec_str = str(item.get("spec") or "")
                    sku_str = str(item.get("sku") or "")
                    qty_val = int(item.get("avail_qty") or 0)
                    n_key = normalize_model_key(spec_str)
                    if n_key:
                        prev_qty, prev_sku = stock_map.get(n_key, (0, ""))
                        stock_map[n_key] = (prev_qty + qty_val, sku_str or prev_sku)
            except Exception:
                pass

    # 2. 讀取報價單條款 terms.txt
    quote_terms_text = ""
    terms_file = src_dir / "terms.txt"
    if terms_file.exists():
        try:
            quote_terms_text = terms_file.read_text(encoding="utf-8").strip()
        except Exception:
            pass

    conn = get_connection(db_path)
    old_conn = sqlite3.connect(str(old_db_path))
    old_conn.row_factory = sqlite3.Row

    try:
        cursor = conn.cursor()
        old_cur = old_conn.cursor()

        # 更新系統設定
        cursor.execute("INSERT OR REPLACE INTO system_settings (setting_key, setting_val) VALUES ('round_digits', '2');")
        cursor.execute("INSERT OR REPLACE INTO system_settings (setting_key, setting_val) VALUES ('round_mode', '無條件捨去');")
        cursor.execute("INSERT OR REPLACE INTO system_settings (setting_key, setting_val) VALUES ('tax_rate', '5.0');")
        cursor.execute("INSERT OR REPLACE INTO system_settings (setting_key, setting_val) VALUES ('stock_updated_at', ?);", (stock_updated_label,))
        if quote_terms_text:
            cursor.execute("INSERT OR REPLACE INTO system_settings (setting_key, setting_val) VALUES ('quote_terms', ?);", (quote_terms_text,))

        # 3. 匯入客戶與聯絡人 (合併 customer_master、contact_master 與歷史報價單快照中的客戶與窗口)
        old_contacts = old_cur.execute("SELECT * FROM contact_master").fetchall()
        contacts_by_old_cid: Dict[int, List[Dict[str, Any]]] = {}
        for oc in old_contacts:
            contacts_by_old_cid.setdefault(oc["customer_id"], []).append(dict(oc))

        old_customers = old_cur.execute("SELECT * FROM customer_master").fetchall()
        # 依公司名稱歸戶，避免舊庫同公司多筆紀錄互相覆蓋聯絡人
        grouped_customers: Dict[str, Dict[str, Any]] = {}
        for c in old_customers:
            c_name = str(c["name"] or "").strip()
            if not c_name:
                continue
            if c_name not in grouped_customers:
                grouped_customers[c_name] = {
                    "name": c_name,
                    "tax_id": c["tax_id"] or "",
                    "phone": c["phone"] or "",
                    "fax": c["fax"] or "",
                    "address": c["address"] or "",
                    "payment_terms": c["payment_terms"] or "月結 30 天電匯",
                    "status": c["status"] or "active",
                    "default_discount": float(c["default_discount"] if c["default_discount"] is not None else 1.0),
                    "is_favorite": int(c["is_favorite"] or 0),
                    "contacts": [],
                }
            else:
                g = grouped_customers[c_name]
                if not g["tax_id"] and c["tax_id"]:
                    g["tax_id"] = c["tax_id"]
                if not g["phone"] and c["phone"]:
                    g["phone"] = c["phone"]
                if not g["fax"] and c["fax"]:
                    g["fax"] = c["fax"]
                if not g["address"] and c["address"]:
                    g["address"] = c["address"]
                if c["is_favorite"]:
                    g["is_favorite"] = 1

            c_contacts = contacts_by_old_cid.get(c["id"], [])
            if c["contact_name"]:
                c_contacts = [{"name": c["contact_name"], "email": c["contact_email"] or "", "mobile": c["contact_mobile"] or ""}] + c_contacts
            for ct in c_contacts:
                ct_name = str(ct.get("name") or "").strip()
                if ct_name and ct_name.lower() != "nan":
                    existing_names = {x["name"] for x in grouped_customers[c_name]["contacts"]}
                    if ct_name not in existing_names:
                        grouped_customers[c_name]["contacts"].append({
                            "name": ct_name,
                            "email": str(ct.get("email") or "").replace("nan", "").strip(),
                            "mobile": str(ct.get("mobile") or "").replace("nan", "").strip(),
                        })

        # 同時從歷史報價單快照補齊曾報價過之客戶與聯絡窗口
        qh_snapshots = old_cur.execute("""
        SELECT customer_name_snapshot, customer_tax_id_snapshot, contact_name_snapshot,
               customer_phone_snapshot, customer_fax_snapshot, customer_address_snapshot,
               payment_terms_snapshot
        FROM quotation_header
        WHERE customer_name_snapshot IS NOT NULL AND TRIM(customer_name_snapshot) != ''
        ORDER BY id DESC
        """).fetchall()
        for snap in qh_snapshots:
            c_name = str(snap["customer_name_snapshot"] or "").strip()
            if not c_name:
                continue
            if c_name not in grouped_customers:
                grouped_customers[c_name] = {
                    "name": c_name,
                    "tax_id": str(snap["customer_tax_id_snapshot"] or "").strip(),
                    "phone": str(snap["customer_phone_snapshot"] or "").strip(),
                    "fax": str(snap["customer_fax_snapshot"] or "").strip(),
                    "address": str(snap["customer_address_snapshot"] or "").strip(),
                    "payment_terms": str(snap["payment_terms_snapshot"] or "月結 30 天電匯").strip(),
                    "status": "active",
                    "default_discount": 1.0,
                    "is_favorite": 0,
                    "contacts": [],
                }
            else:
                g = grouped_customers[c_name]
                if not g["tax_id"] and snap["customer_tax_id_snapshot"]:
                    g["tax_id"] = str(snap["customer_tax_id_snapshot"]).strip()
                if not g["phone"] and snap["customer_phone_snapshot"]:
                    g["phone"] = str(snap["customer_phone_snapshot"]).strip()
                if not g["fax"] and snap["customer_fax_snapshot"]:
                    g["fax"] = str(snap["customer_fax_snapshot"]).strip()
                if not g["address"] and snap["customer_address_snapshot"]:
                    g["address"] = str(snap["customer_address_snapshot"]).strip()

            snap_contact = str(snap["contact_name_snapshot"] or "").strip()
            if snap_contact and snap_contact.lower() != "nan":
                existing_names = {x["name"] for x in grouped_customers[c_name]["contacts"]}
                if snap_contact not in existing_names:
                    grouped_customers[c_name]["contacts"].append({
                        "name": snap_contact,
                        "email": "",
                        "mobile": "",
                    })

        imported_customers = 0
        imported_contacts = 0

        for c_name, g in grouped_customers.items():
            first_contact = g["contacts"][0]["name"] if g["contacts"] else ""
            cursor.execute("""
            INSERT INTO customer_master (
                name, owner_sales, tax_id, default_contact, phone, fax,
                address, payment_terms, status, default_discount, is_favorite
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(name) DO UPDATE SET
                owner_sales = excluded.owner_sales,
                tax_id = CASE WHEN excluded.tax_id != '' THEN excluded.tax_id ELSE customer_master.tax_id END,
                default_contact = CASE WHEN excluded.default_contact != '' THEN excluded.default_contact ELSE customer_master.default_contact END,
                phone = CASE WHEN excluded.phone != '' THEN excluded.phone ELSE customer_master.phone END,
                fax = CASE WHEN excluded.fax != '' THEN excluded.fax ELSE customer_master.fax END,
                address = CASE WHEN excluded.address != '' THEN excluded.address ELSE customer_master.address END,
                payment_terms = CASE WHEN excluded.payment_terms != '' THEN excluded.payment_terms ELSE customer_master.payment_terms END,
                status = COALESCE(excluded.status, 'active'),
                default_discount = COALESCE(excluded.default_discount, 1.0),
                is_favorite = CASE WHEN excluded.is_favorite > customer_master.is_favorite THEN excluded.is_favorite ELSE customer_master.is_favorite END
            """, (
                c_name,
                default_owner,
                g["tax_id"],
                first_contact,
                g["phone"],
                g["fax"],
                g["address"],
                g["payment_terms"],
                g["status"],
                g["default_discount"],
                g["is_favorite"],
            ))
            imported_customers += 1

            new_c_row = cursor.execute("SELECT id FROM customer_master WHERE name = ?", (c_name,)).fetchone()
            if new_c_row and g["contacts"]:
                new_cid = new_c_row["id"]
                cursor.execute("DELETE FROM contact_master WHERE customer_id = ?", (new_cid,))
                for ct in g["contacts"]:
                    cursor.execute("""
                    INSERT INTO contact_master (customer_id, owner_sales, name, email, mobile)
                    VALUES (?, ?, ?, ?, ?)
                    """, (new_cid, default_owner, ct["name"], ct["email"], ct["mobile"]))
                    imported_contacts += 1

        # 4. 匯入共通標準牌價表 (9,975 筆產品 + 庫存數量與 SKU)
        old_products = old_cur.execute("SELECT * FROM product_master").fetchall()
        imported_products = 0
        for p in old_products:
            m_no = str(p["model_no"]).strip()
            n_key = normalize_model_key(m_no)
            stk_qty, sku_code = stock_map.get(n_key, (int(p["stock_snapshot"] or 0), ""))
            cursor.execute("""
            INSERT INTO product_master (
                model_no, spec, list_price, dealer_price, stock_qty,
                sku, category, is_order_item, stock_status
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            ON CONFLICT(model_no) DO UPDATE SET
                spec = excluded.spec,
                list_price = excluded.list_price,
                dealer_price = excluded.dealer_price,
                stock_qty = excluded.stock_qty,
                sku = CASE WHEN excluded.sku != '' THEN excluded.sku ELSE product_master.sku END,
                category = COALESCE(excluded.category, product_master.category),
                is_order_item = excluded.is_order_item,
                stock_status = excluded.stock_status
            """, (
                m_no,
                p["spec"] or m_no,
                float(p["list_price"] or 0.0),
                float(p["dealer_price"] or 0.0),
                stk_qty,
                sku_code,
                p["category"] or "",
                int(p["is_order_item"] or 0),
                p["stock_status"] or "IN_STOCK",
            ))
            imported_products += 1

        # 5. 匯入 742 筆歷史報價單與 4,623 筆明細 (歸屬 default_owner)
        old_headers = old_cur.execute("SELECT * FROM quotation_header ORDER BY id ASC").fetchall()
        old_items = old_cur.execute("SELECT * FROM quotation_item ORDER BY header_id ASC, row_index ASC, id ASC").fetchall()
        items_by_old_hid: Dict[int, List[Dict[str, Any]]] = {}
        for it in old_items:
            items_by_old_hid.setdefault(it["header_id"], []).append(dict(it))

        imported_quotes = 0
        imported_items = 0

        for h in old_headers:
            q_no = str(h["full_quotation_id"]).strip()
            existing_q = cursor.execute("SELECT id FROM quotation_header WHERE quote_no = ?", (q_no,)).fetchone()
            if existing_q:
                continue

            prefix_str = str(h["prefix"] or q_no).strip()
            try:
                seq_int = int(str(h["seq_no"] or "1").strip())
            except Exception:
                seq_int = 1
            try:
                rev_int = int(str(h["revision"] or "0").strip())
            except Exception:
                rev_int = 0

            q_date = str(h["quotation_date"] or "")[:10] or str(date.today())
            v_date = str(h["valid_until"] or "")[:10] or q_date

            if h["is_won"]:
                status_str = "已成交"
            elif h["is_lost"]:
                status_str = "已失注"
            elif h["is_ignored"]:
                status_str = "不跟進"
            else:
                status_str = "追蹤中"

            cursor.execute("""
            INSERT INTO quotation_header (
                quote_no, prefix, sequence, revision, sales_name, deal_id,
                quote_date, valid_date, customer_name, tax_id, contact_name,
                phone, fax, delivery_address, payment_terms, project_name,
                project_type, subtotal, tax_rate, tax_amount, grand_total,
                status, is_latest, is_won, won_amount, won_date,
                is_lost, is_ignored, lost_reason, lost_memo, source_type,
                is_deleted, deleted_at, created_at
            ) VALUES (?, ?, ?, ?, ?, NULL, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 5.0, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);
            """, (
                q_no,
                prefix_str,
                seq_int,
                rev_int,
                default_owner,
                q_date,
                v_date,
                str(h["customer_name_snapshot"] or "未命名客戶"),
                str(h["customer_tax_id_snapshot"] or ""),
                str(h["contact_name_snapshot"] or ""),
                str(h["customer_phone_snapshot"] or ""),
                str(h["customer_fax_snapshot"] or ""),
                str(h["customer_address_snapshot"] or ""),
                str(h["payment_terms_snapshot"] or ""),
                str(h["project_name"] or ""),
                str(h["project_type"] or "一次性專案"),
                float(h["subtotal_amount"] or 0.0),
                float(h["tax_amount"] or 0.0),
                float(h["total_amount"] or 0.0),
                status_str,
                int(h["is_latest"] if h["is_latest"] is not None else 1),
                int(h["is_won"] or 0),
                float(h["won_amount"] or 0.0),
                str(h["won_date"] or "") if h["won_date"] else None,
                int(h["is_lost"] or 0),
                int(h["is_ignored"] or 0),
                h["lost_reason"],
                h["lost_memo"],
                str(h["source_type"] or "PC_DIRECT"),
                int(h["is_deleted"] or 0),
                str(h["deleted_at"] or "") if h["deleted_at"] else None,
                str(h["created_at"] or datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
            ))
            new_hid = cursor.lastrowid
            imported_quotes += 1

            h_items = items_by_old_hid.get(h["id"], [])
            for idx, itm in enumerate(h_items, start=1):
                cursor.execute("""
                INSERT INTO quotation_item (
                    header_id, quote_no, item_index, model_no, spec,
                    dealer_price, discount_pct, expected_price, final_unit_price,
                    quantity, row_subtotal, note_lead_time, is_won, is_lost,
                    item_status, lost_reason
                ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?);
                """, (
                    new_hid,
                    q_no,
                    int(itm["row_index"]) + 1 if itm.get("row_index") is not None else idx,
                    str(itm["product_model_snapshot"] or ""),
                    str(itm["product_spec_snapshot"] or ""),
                    float(itm["dealer_price"] or 0.0),
                    float(itm["discount_rate"]) if itm.get("discount_rate") is not None else None,
                    float(itm["target_amount"]) if itm.get("target_amount") is not None else None,
                    float(itm["final_unit_price"] or 0.0),
                    float(itm["quantity"] if itm.get("quantity") is not None else 1.0),
                    float(itm["subtotal"] or 0.0),
                    str(itm["remark"] or ""),
                    int(itm.get("is_won") or 0),
                    int(itm.get("is_lost") or 0),
                    str(itm.get("item_status") or "pending"),
                    itm.get("lost_reason"),
                ))
                imported_items += 1

        # 同步執行方案 A 檢查
        sync_special_to_standard_archive(conn)
        conn.commit()

        return {
            "status": "success",
            "imported_customers": imported_customers,
            "imported_contacts": imported_contacts,
            "imported_products": imported_products,
            "imported_quotes": imported_quotes,
            "imported_items": imported_items,
            "stock_updated_at": stock_updated_label,
        }
    finally:
        old_conn.close()
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

        if days_active < 28:
            rem = 28 - days_active
            return False, f"該指標自 {driver['start_date']} 起僅執行 {days_active} 天，未滿 4 週（28 天）。為避免指標頻繁變更導致行動失焦，尚無法提出更換申請（尚需執行 {rem} 天）。", None

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
            conn.execute("""
            UPDATE lead_driver_metric
            SET status = 'archived', end_date = ?
            WHERE id = ?;
            """, (today_str, req["driver_id"]))

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

            conn.execute("""
            UPDATE driver_change_request
            SET review_status = '已核准', reviewer_comment = ?, reviewed_at = ?
            WHERE id = ?;
            """, (comment_with_reviewer, now_str, request_id))

            conn.commit()
            return True, f"申請單已核准！舊指標已封存，新指標「{req['new_driver_name']}」已正式生效。"

        elif action == "reject":
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
    """查詢驅動目標變更申請單清單"""
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


def record_driver_actual(driver_id: int, actual_val: float) -> Tuple[bool, str]:
    """登記前線最新實績數值 (actual_val)"""
    if actual_val < 0:
        return False, "實績數值不得為負數"

    conn = get_connection()
    try:
        row = conn.execute("SELECT id, driver_name, target_val, unit FROM lead_driver_metric WHERE id = ? AND status = 'active'", (driver_id,)).fetchone()
        if not row:
            return False, "找不到該筆生效中之領先指標紀錄"

        conn.execute("""
        UPDATE lead_driver_metric
        SET actual_val = ?
        WHERE id = ?;
        """, (float(actual_val), driver_id))
        conn.commit()
        return True, f"實績已成功登記為 {actual_val} {row['unit']}！"
    finally:
        conn.close()


def update_driver_settings(driver_id: int, driver_name: str, target_val: float, unit: str, period_type: str = "週") -> Tuple[bool, str]:
    """直接設定或編輯領先指標內容 (包含行動名稱、目標數值、單位與週期)"""
    d_name_clean = driver_name.strip()
    if not d_name_clean:
        return False, "指標行動名稱不得為空"
    if target_val <= 0:
        return False, "目標數值必須大於 0"

    unit_clean = unit.strip() or "次"
    p_type_clean = period_type.strip() or "週"

    conn = get_connection()
    try:
        row = conn.execute("SELECT id, product_line FROM lead_driver_metric WHERE id = ? AND status = 'active'", (driver_id,)).fetchone()
        if not row:
            return False, "找不到該筆生效中之領先指標紀錄"

        conn.execute("""
        UPDATE lead_driver_metric
        SET driver_name = ?, target_val = ?, unit = ?, period_type = ?
        WHERE id = ?;
        """, (d_name_clean, float(target_val), unit_clean, p_type_clean, driver_id))
        conn.commit()
        return True, f"「{row['product_line']}」領先指標設定已成功更新！"
    finally:
        conn.close()


if __name__ == "__main__":
    print("[data_access.py] 正在初始化士林電機 FA 資料庫...")
    init_database()
    print("[data_access.py] 資料庫建表與種子資料植入成功！")
