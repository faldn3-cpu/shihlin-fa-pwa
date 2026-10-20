# -*- coding: utf-8 -*-
"""
士林電機 FA 業務戰情與報價系統 - 資料存取層自動化驗證腳本 (test_data_access.py)
嚴格檢驗：
1. 資料表建立與種子資料
2. 業務日報 (OGSM) 寫入與讀取
3. CRM 商機轉為 Deal 案件
4. 報價單三段式編號生成 (Prefix-Seq-Rev)
5. 報價單 Append-Only 鐵律 (修改產生新版次，舊單不被覆蓋)
6. 垃圾桶軟刪除與一鍵還原
7. 專案售價申請單串聯
"""

import os
import sys
from pathlib import Path

# 將專案目錄加入模組搜尋路徑
project_dir = Path(__file__).resolve().parent
sys.path.insert(0, str(project_dir))

# 設定 Windows 終端輸出編碼為 UTF-8
if sys.platform.startswith("win"):
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass

import data_access


def run_all_tests():
    print("=" * 65)
    print("[TEST] 開始執行 data_access.py 資料存取層全自動化驗證")
    print("=" * 65)

    test_db = project_dir / "test_verification.db"
    if test_db.exists():
        try:
            test_db.unlink()
        except Exception:
            pass

    # 1. 驗證資料庫建立與種子資料
    print("\n[測試 1] 初始化資料庫與檢查種子資料...")
    data_access.init_database(test_db)
    conn = data_access.get_connection(test_db)
    
    # 檢查 13 張表是否全數建立
    tables = conn.execute("SELECT name FROM sqlite_master WHERE type='table';").fetchall()
    table_names = [t["name"] for t in tables]
    required_tables = [
        "daily_report", "crm_opportunity", "deal_case", "quotation_header",
        "quotation_item", "special_price_application", "customer_master",
        "product_master", "manager_task", "industry_strategy",
        "scoreboard_metric", "system_settings"
    ]
    for req in required_tables:
        assert req in table_names, f"缺少必要資料表: {req}"
    print(f"  [OK] 12 大核心業務表全數成功建立 (目前共有 {len(table_names)} 張表格)")

    # 檢查客戶與產品主檔
    customers = conn.execute("SELECT COUNT(*) AS c FROM customer_master;").fetchone()["c"]
    products = conn.execute("SELECT COUNT(*) AS c FROM product_master;").fetchone()["c"]
    assert customers >= 2, "客戶主檔種子資料不足"
    assert products >= 5, "產品庫種子資料不足"
    print(f"  [OK] 主檔種子資料正確植入：客戶 {customers} 家、產品庫 {products} 筆")
    conn.close()

    # 2. 驗證業務日報 (OGSM)
    print("\n[測試 2] 驗證業務日報 (OGSM) 寫入與查詢...")
    # 設定全域測試 DB
    original_db = data_access.DEFAULT_DB_FILE
    data_access.DEFAULT_DB_FILE = test_db

    try:
        report_id = data_access.save_daily_report({
            "report_date": "2026-10-20",
            "sales_name": "曾維崧",
            "customer_name": "贊翔實業有限公司",
            "contact_name": "陳姿蓉特助",
            "visit_type": "實地拜訪",
            "content_summary": "拜訪洽談本季出貨目標，客戶表達有新產線變頻器升級需求。"
        })
        assert report_id > 0, "日報寫入失敗"
        reports = data_access.get_daily_reports(sales_name="曾維崧")
        assert len(reports) == 1, "日報查詢筆數不符合"
        assert reports[0]["customer_name"] == "贊翔實業有限公司"
        print(f"  [OK] 業務日報寫入成功 (編號 ID={report_id})，資料精準回讀！")

        # 3. 驗證商機 (CRM) 轉 Deal 案件
        print("\n[測試 3] 驗證商機 (CRM) 轉成 Deal 專案案件...")
        opp_id = data_access.save_crm_opportunity({
            "daily_report_id": report_id,
            "customer_name": "贊翔實業有限公司",
            "industry_type": "自動化設備產業",
            "product_series": "士林品",
            "est_quarter": "Q4",
            "est_scale": 1500000.0,
            "status": "評估中"
        })
        assert opp_id > 0
        print(f"  [OK] 第二層 CRM 商機建立成功 (ID={opp_id})")

        deal_id = data_access.save_deal_case({
            "opportunity_id": opp_id,
            "customer_name": "贊翔實業有限公司",
            "project_name": "新竹產線變頻器升級案",
            "project_type": "專案特價",
            "comp_brand": "台達",
            "comp_channel": "能麒",
            "est_deal_date": "2026-11-30",
            "est_amount": 1200000.0,
            "stage": "報價中"
        })
        assert deal_id > 0
        deals = data_access.get_deal_cases(stage="報價中")
        assert len(deals) == 1
        assert deals[0]["project_name"] == "新竹產線變頻器升級案"
        print(f"  [OK] 第三層 Deal 專案案件升格成功 (ID={deal_id})")

        # 4. 驗證三段式報價單編號與 Append-Only (修改另存新版，舊單不覆蓋)
        print("\n[測試 4] 驗證報價單 Append-Only 鐵律 (修改產生新版次，禁止覆蓋舊資料)...")
        header_data_v0 = {
            "prefix": "贊翔2610",
            "deal_id": deal_id,
            "quote_date": "2026-10-20",
            "valid_date": "2026-11-19",
            "customer_name": "贊翔實業有限公司",
            "tax_id": "87654321",
            "contact_name": "陳姿蓉特助",
            "project_name": "新竹產線變頻器升級案",
            "project_type": "專案特價",
            "subtotal": 141000.0,
            "tax_rate": 5.0,
            "tax_amount": 7050.0,
            "grand_total": 148050.0,
            "status": "追蹤中"
        }
        items_data_v0 = [
            {
                "model_no": "SE3-023-11K",
                "spec": "3Φ220V 15HP 高機能向量變頻器",
                "dealer_price": 14100.0,
                "discount_pct": 100.0,
                "final_unit_price": 14100.0,
                "quantity": 10.0,
                "row_subtotal": 141000.0,
                "note_lead_time": "現貨庫存"
            }
        ]

        # 第一次儲存：應自動產生 Sequence 001, Revision 00
        quote_no_v0 = data_access.save_quotation(header_data_v0, items_data_v0)
        assert quote_no_v0 == "贊翔2610-001-00", f"單號格式不符合預期: {quote_no_v0}"
        print(f"  [OK] 初版報價單儲存成功：單號為【{quote_no_v0}】")

        # 第二次儲存（業務調整價格另存新版）：指定 sequence=1
        header_data_v1 = header_data_v0.copy()
        header_data_v1["sequence"] = 1  # 針對第 1 號報價單升級版次
        header_data_v1["subtotal"] = 120000.0
        header_data_v1["grand_total"] = 126000.0
        items_data_v1 = [
            {
                "model_no": "SE3-023-11K",
                "spec": "3Φ220V 15HP 高機能向量變頻器",
                "dealer_price": 14100.0,
                "discount_pct": 85.0,
                "final_unit_price": 12000.0,
                "quantity": 10.0,
                "row_subtotal": 120000.0,
                "note_lead_time": "現貨庫存"
            }
        ]
        quote_no_v1 = data_access.save_quotation(header_data_v1, items_data_v1)
        assert quote_no_v1 == "贊翔2610-001-01", f"新版次單號不符合預期: {quote_no_v1}"
        print(f"  [OK] 次版報價單另存成功：單號自動升級為【{quote_no_v1}】")

        # 關鍵驗證：初版 00 必須依然完整存在於資料庫中（禁止覆蓋！）
        old_quote = data_access.get_quotation("贊翔2610-001-00")
        new_quote = data_access.get_quotation("贊翔2610-001-01")
        assert old_quote is not None, "舊版報價單遭到抹除！違反 Append-Only 鐵律！"
        assert new_quote is not None, "新版報價單未成功讀出"
        assert old_quote["grand_total"] == 148050.0, "舊版報價金額被改動！"
        assert new_quote["grand_total"] == 126000.0, "新版報價金額不正確！"
        print("  [OK] 成功通過 Append-Only 鐵律審查：舊版 00 與新版 01 獨立並存、無痕追蹤！")

        # 5. 驗證垃圾桶軟刪除與一鍵復原
        print("\n[測試 5] 驗證垃圾桶機制 (軟刪除與復原)...")
        data_access.soft_delete_quotation("贊翔2610-001-01")
        active_list = data_access.list_quotations(include_deleted=False)
        trash_list = data_access.list_quotations(include_deleted=True)
        assert any(q["quote_no"] == "贊翔2610-001-00" for q in active_list)
        assert not any(q["quote_no"] == "贊翔2610-001-01" for q in active_list)
        assert any(q["quote_no"] == "贊翔2610-001-01" for q in trash_list)
        print("  [OK] 軟刪除成功：單號移入垃圾桶，正常清單已隱藏。")

        data_access.restore_quotation("贊翔2610-001-01")
        active_list_after = data_access.list_quotations(include_deleted=False)
        assert any(q["quote_no"] == "贊翔2610-001-01" for q in active_list_after)
        print("  [OK] 一鍵復原成功：單號順利移回正常報價管理清單。")

        # 6. 驗證專案售價申請單
        print("\n[測試 6] 驗證專案售價申請單...")
        app_id = data_access.save_special_price_application({
            "deal_id": deal_id,
            "quotation_id": new_quote["id"],
            "quote_no": "贊翔2610-001-01",
            "applicant_sales": "曾維崧",
            "apply_reason": "台達報價8,500搶單，為保住年約客戶，建議以85折專案價承做。",
            "comp_price_info": "台達經銷商能麒出價單價約8,500元未稅",
            "target_subtotal": 120000.0,
            "est_margin_pct": 21.5
        })
        assert app_id > 0
        print(f"  [OK] 專案售價申請單成功建立 (ID={app_id})！")

        print("\n" + "=" * 65)
        print("[SUCCESS] 恭喜！data_access.py 全數 6 大單元測試 100% 通過！")
        print("=" * 65)

    finally:
        # 復原路徑設定並安全刪除測試資料庫
        data_access.DEFAULT_DB_FILE = original_db
        if test_db.exists():
            try:
                test_db.unlink()
            except Exception:
                pass


if __name__ == "__main__":
    run_all_tests()
