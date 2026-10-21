# -*- coding: utf-8 -*-
"""
Web API 與三層漏斗整合測試腳本 (test_server.py)
不依賴第三方 testclient/httpx，直接透過 FastAPI 內部處理函式進行精確單元驗證：
1. /api/customers 自動帶入主檔
2. /api/daily-reports 建立日報並同時連鎖建立 CRM 商機與 Deal 案件
3. 前端首頁 index.html 正確託管
"""

import sys
from pathlib import Path

# Windows 輸出編碼
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

import server


def test_api():
    print("=" * 60)
    print("[TEST] 開始驗證 Web API 伺服器與三層漏斗連鎖...")
    print("=" * 60)

    # 1. 測試前端首頁託管
    index_file = server.WEB_DIR / "index.html"
    assert index_file.exists(), "前端 index.html 不存在"
    content = index_file.read_text(encoding="utf-8")
    assert "士林電機" in content
    assert "第一層：客戶每日日報填寫" in content
    assert "第二層 CRM" in content
    assert "第三層 Deal" in content
    print("  [OK] 前端首頁 index.html 存在且包含三層漏斗結構")

    # 2. 測試客戶主檔 API
    customers = server.api_get_customers()
    assert len(customers) >= 2
    assert any(c["name"] == "贊翔實業有限公司" for c in customers)
    print(f"  [OK] 客戶主檔 API 正常回傳 {len(customers)} 筆資料")

    # 3. 測試日報連鎖建立第二層商機與第三層案件
    report_in = server.DailyReportCreate(
        report_date="2026-10-20",
        sales_name="曾維崧",
        customer_id=customers[0]["id"],
        customer_name=customers[0]["name"],
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
    print(f"  [OK] 三層連鎖儲存成功：日報 ID={res_data['daily_report_id']} ➔ 商機 ID={res_data['opportunity_id']} ➔ 案件 ID={res_data['deal_id']}")

    # 4. 驗證案件清單 API
    deals = server.api_get_deals()
    assert any(d["project_name"] == "測試產線升級專案" for d in deals)
    print("  [OK] Deal 案件看板 API 正常回傳，確認包含新建之「測試產線升級專案」")

    print("\n" + "=" * 60)
    print("[SUCCESS] Web API 與三層漏斗功能驗證全數通過！")
    print("=" * 60)


if __name__ == "__main__":
    test_api()
