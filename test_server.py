# -*- coding: utf-8 -*-
"""
Web API、三層漏斗、雙視圖報價、手機草稿匣直連資料庫與【方案 A】特殊品轉標準品整合測試腳本 (test_server.py)
注意：測試全程使用臨時測試用副本資料庫 (test_server_verification.db)，100% 不改動正式 quotation_system.db。
"""

import shutil
import sys
from pathlib import Path

# Windows 輸出編碼設定為 UTF-8
if sys.platform.startswith("win"):
    try:
        reconfig_out = getattr(sys.stdout, "reconfigure", None)
        if callable(reconfig_out):
            reconfig_out(encoding="utf-8")
        reconfig_err = getattr(sys.stderr, "reconfigure", None)
        if callable(reconfig_err):
            reconfig_err(encoding="utf-8")
    except Exception:
        pass

import data_access
import server


def test_api():
    print("=" * 65)
    print("[TEST] 開始驗證 Web API 伺服器、雙視圖報價與多人隔離機制...")
    print("=" * 65)

    project_dir = Path(__file__).resolve().parent
    real_db = data_access.DEFAULT_DB_FILE
    test_db = project_dir / "test_server_verification.db"

    if test_db.exists():
        try:
            test_db.unlink()
        except Exception:
            pass

    if real_db.exists():
        shutil.copy2(real_db, test_db)
    else:
        data_access.init_database(test_db)

    data_access.DEFAULT_DB_FILE = test_db

    try:
        # 1. 測試前端首頁託管、響應式雙視圖結構與業績預測頁面
        index_file = server.WEB_DIR / "index.html"
        assert index_file.exists(), "前端 index.html 不存在"
        content = index_file.read_text(encoding="utf-8")
        assert "士林電機" in content
        assert "第一層：客戶每日日報填寫" in content
        assert "第二層 CRM" in content
        assert "第三層 Deal" in content
        assert "navItemForecast" in content and "panelForecast" in content, "缺少業績預測頁面結構"
        assert "tabQuoteSpecialProducts" in content, "缺少個人特殊品牌價分頁"
        assert "customerName" in content and "itemsContainer" in content, "缺少手機端圖三擬稿介面結構"
        print("  [OK] 前端 index.html 包含三層漏斗、業績預測、個人特殊品牌價與手機圖三擬稿結構")

        # 2. 測試客戶主檔 API 與業務隔離
        customers_tseng = server.api_get_customers(sales_name="曾維崧")
        customers_yeh = server.api_get_customers(sales_name="葉仁豪")
        assert len(customers_tseng) >= 100, f"曾維崧整合後客戶數量異常: {len(customers_tseng)}"
        assert len(customers_yeh) == 0, "葉仁豪不應看到曾維崧的客戶主檔"
        print(f"  [OK] 客戶主檔 API 正常回傳：曾維崧 {len(customers_tseng)} 家、葉仁豪 {len(customers_yeh)} 家 (嚴格隔離)")

        # 3. 測試日報連鎖建立第二層商機與第三層案件
        report_in = server.DailyReportCreate(
            report_date="2026-10-20",
            sales_name="曾維崧",
            customer_id=customers_tseng[0]["id"],
            customer_name=customers_tseng[0]["name"],
            contact_name="陳特助",
            visit_type="實地拜訪",
            content_summary="拜訪洽談新設備案，需要士林 15HP 變頻器支援。",
            create_opportunity=True,
            industry_type="自動化設備產業",
            product_series="士林品",
            est_quarter="Q4",
            est_scale=800000.0,
            create_deal=True,
            project_name="測試產線升級專案",
            comp_brand="台達",
            comp_channel="能麒",
            est_deal_date="2026-12-15",
            est_amount=750000.0
        )

        res_data = server.api_create_daily_report(report_in)
        assert res_data["status"] == "success"
        assert res_data["daily_report_id"] is not None
        assert res_data["opportunity_id"] is not None
        assert res_data["deal_id"] is not None
        print(f"  [OK] 三層連鎖儲存成功：日報 ID={res_data['daily_report_id']} -> 商機 ID={res_data['opportunity_id']} -> 案件 ID={res_data['deal_id']}")

        # 4. 驗證案件清單 API
        deals = server.api_get_deals()
        assert any(d["project_name"] == "測試產線升級專案" for d in deals)
        print("  [OK] Deal 案件看板 API 正常回傳，確認包含新建之「測試產線升級專案」")

        # 5. 驗證三大產品線領先指標 (驅動目標) API
        print("\n[測試 5] 驗證領先指標 API 與三位主管/業務視角...")
        for s_name in ["曾維崧", "曾仁君", "張何達"]:
            driver_res = server.api_get_drivers(sales_name=s_name)
            assert driver_res["status"] == "success"
            assert len(driver_res["drivers"]) == 3, f"{s_name} 指標數量不為 3"
            print(f"  [OK] {s_name} 領先指標 API 回傳正常 (3 條產品線)")

        req_res = server.api_get_driver_requests()
        assert req_res["status"] == "success"
        print(f"  [OK] 變更申請單 API 回傳正常，目前共有 {len(req_res['requests'])} 筆申請待審")

        # 6. 驗證手機報價草稿匣 API (直連資料庫，不經雲端硬碟)
        print("\n[測試 6] 驗證手機報價草稿匣直連資料庫 API 與業務隔離...")
        draft_payload = server.QuotationDraftCreate(
            sales_name="曾維崧",
            customer_name="贊翔實業有限公司",
            tax_id="27866845",
            contact_name="陳小姐",
            phone="04-23358958",
            subtotal=42300.0,
            items=[
                {"model_no": "SE3-023-11K", "spec": "變頻器", "dealer_price": 14100, "qty": 3, "final_price": 14100}
            ]
        )
        draft_create_res = server.api_save_quotation_draft(draft_payload)
        assert draft_create_res["status"] == "success"
        draft_id = draft_create_res["draft_id"]

        draft_list_tseng = server.api_get_quotation_drafts(sales_name="曾維崧")
        draft_list_yeh = server.api_get_quotation_drafts(sales_name="葉仁豪")
        assert any(d["id"] == draft_id for d in draft_list_tseng)
        assert len(draft_list_yeh) == 0, "葉仁豪不應看到曾維崧的手機報價草稿"

        del_res = server.api_delete_quotation_draft(draft_id=draft_id, sales_name="曾維崧")
        assert del_res["status"] == "success"
        print("  [OK] 手機報價草稿送出、電腦端草稿匣讀取與移除 API 驗證通過！")

        # 7. 驗證個人特殊品牌價表與【方案 A】轉標準品自動取代並封存 API
        print("\n[測試 7] 驗證個人特殊品牌價表與【方案 A】API...")
        sp_in = server.SpecialProductCreate(
            owner_sales="曾維崧",
            model_no="API-SP-777",
            spec="測試專用特殊規格控制器",
            dealer_price=18500.0,
            stock_qty=5,
            note="業務專案特規"
        )
        sp_save_res = server.api_save_special_product(sp_in)
        assert sp_save_res["status"] == "success"
        sp_item_id = sp_save_res["id"]

        promo_req = server.SpecialProductPromoteRequest(standard_dealer_price=20000.0)
        promo_api_res = server.api_promote_special_to_standard(sp_item_id, promo_req)
        assert promo_api_res["status"] == "success"

        sp_list_res = server.api_get_special_products(sales_name="曾維崧", include_archived=True)
        archived_item = next(item for item in sp_list_res if item["model_no"] == "API-SP-777")
        assert archived_item["is_archived"] == 1
        assert archived_item["archived_standard_price"] == 20000.0
        print("  [OK] 個人特殊品建立與【方案 A】轉標準品自動取代封存 API 驗證通過！")

        # 8. 驗證歷史報價單、歷史單價分析與財務設定 API
        print("\n[測試 8] 驗證歷史報價單、歷史單價分析與財務設定 API...")
        quotes_tseng = server.api_list_quotations(include_deleted=False, sales_name="曾維崧", limit=1000)
        quotes_yeh = server.api_list_quotations(include_deleted=False, sales_name="葉仁豪", limit=1000)
        assert len(quotes_tseng) >= 700, f"曾維崧歷史報價單筆數異常: {len(quotes_tseng)}"
        assert len(quotes_yeh) == 0, "葉仁豪不應看到曾維崧的歷史報價單"

        analysis_res = server.api_analyze_historical_prices(sales_name="曾維崧", model_keyword="SA3", customer_keyword="")
        assert len(analysis_res) > 0

        settings_res = server.api_get_quotation_settings()
        assert settings_res["round_digit_label"] == "百位數"
        assert settings_res["round_mode"] == "無條件捨去"
        print(f"  [OK] 歷史報價單 ({len(quotes_tseng)} 筆)、歷史單價分析 ({len(analysis_res)} 筆) 與財務捨入設定 API 驗證通過！")

        print("\n" + "=" * 65)
        print("[SUCCESS] Web API 伺服器全數 8 大整合測試 100% 通過！")
        print("=" * 65)

    finally:
        data_access.DEFAULT_DB_FILE = real_db
        if test_db.exists():
            try:
                test_db.unlink()
            except Exception:
                pass


if __name__ == "__main__":
    test_api()
