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
from datetime import datetime, date
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

        # 13. 建立常用核心索引 (加速查歷史與過濾)
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_quote_prefix_seq ON quotation_header (prefix, sequence);")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_quote_customer ON quotation_header (customer_name);")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_quote_status_del ON quotation_header (status, is_deleted);")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_item_quote_no ON quotation_item (quote_no);")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_daily_sales_date ON daily_report (sales_name, report_date);")

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
# 模組自我驗證入口
# ==============================================================================

if __name__ == "__main__":
    print("[data_access.py] 正在初始化士林電機 FA 資料庫...")
    init_database()
    print("[data_access.py] 資料庫建表與種子資料植入成功！")
