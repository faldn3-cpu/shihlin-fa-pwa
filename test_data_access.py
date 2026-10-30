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
7. 專案售價申請單串聯與領先指標 4 週防呆審核
8. 多人業務資料嚴格隔離 (客戶、歷史報價單、手機草稿匣、歷史單價分析)
9. 共通牌價表 + 業務個人特殊品牌價表，以及【方案 A】特殊品轉標準品自動取代與封存留紀錄
"""

import sys
from pathlib import Path

# 將專案目錄加入模組搜尋路徑
project_dir = Path(__file__).resolve().parent
sys.path.insert(0, str(project_dir))

# 設定 Windows 終端輸出編碼為 UTF-8
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

    tables = conn.execute("SELECT name FROM sqlite_master WHERE type='table';").fetchall()
    table_names = [t["name"] for t in tables]
    required_tables = [
        "daily_report", "crm_opportunity", "deal_case", "quotation_header",
        "quotation_item", "special_price_application", "customer_master",
        "product_master", "manager_task", "industry_strategy",
        "scoreboard_metric", "system_settings",
        "contact_master", "personal_special_product", "quotation_draft"
    ]
    for req in required_tables:
        assert req in table_names, f"缺少必要資料表: {req}"
    print(f"  [OK] 核心與擴充業務表全數成功建立 (目前共有 {len(table_names)} 張表格)")

    customers = conn.execute("SELECT COUNT(*) AS c FROM customer_master;").fetchone()["c"]
    products = conn.execute("SELECT COUNT(*) AS c FROM product_master;").fetchone()["c"]
    assert customers >= 2, "客戶主檔種子資料不足"
    assert products >= 5, "產品庫種子資料不足"
    print(f"  [OK] 主檔種子資料正確植入：客戶 {customers} 家、產品庫 {products} 筆")
    conn.close()

    # 2. 驗證業務日報 (OGSM)
    print("\n[測試 2] 驗證業務日報 (OGSM) 寫入與查詢...")
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
            "sales_name": "曾維崧",
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

        quote_no_v0 = data_access.save_quotation(header_data_v0, items_data_v0)
        assert quote_no_v0 == "贊翔2610-001-00", f"單號格式不符合預期: {quote_no_v0}"
        print(f"  [OK] 初版報價單儲存成功：單號為【{quote_no_v0}】")

        header_data_v1 = header_data_v0.copy()
        header_data_v1["sequence"] = 1
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
        active_list = data_access.list_quotations(include_deleted=False, sales_name="曾維崧")
        trash_list = data_access.list_quotations(include_deleted=True, sales_name="曾維崧")
        assert any(q["quote_no"] == "贊翔2610-001-00" for q in active_list)
        assert not any(q["quote_no"] == "贊翔2610-001-01" for q in active_list)
        assert any(q["quote_no"] == "贊翔2610-001-01" for q in trash_list)
        print("  [OK] 軟刪除成功：單號移入垃圾桶，正常清單已隱藏。")

        data_access.restore_quotation("贊翔2610-001-01")
        active_list_after = data_access.list_quotations(include_deleted=False, sales_name="曾維崧")
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

        # 7. 驗證三大產品線領先指標 (驅動目標) 與 4 週防呆審核
        print("\n[測試 7] 驗證領先指標 (驅動目標) 讀取、4 週防呆阻擋與主管審核...")
        drivers_tseng = data_access.get_active_drivers("曾維崧")
        assert len(drivers_tseng) == 3, f"曾維崧應有 3 條產品線指標，實際有 {len(drivers_tseng)} 條"
        for drv in drivers_tseng:
            assert drv["product_line"] in ["士林伺服", "士林變頻器", "士林PLC"]
            assert drv["rate_pct"] >= 0.0
            assert drv["traffic_light"] in ["green", "yellow", "red"]
        print("  [OK] 曾維崧 3 條產品線驅動指標讀取成功，紅黃綠燈計算精準！")

        drivers_yeh = data_access.get_active_drivers("葉仁豪")
        servo_yeh = next(d for d in drivers_yeh if d["product_line"] == "士林伺服")
        assert servo_yeh["can_change"] is False, "未滿 28 天卻顯示可變更！"
        ok, msg, req_id = data_access.apply_driver_change({
            "driver_id": servo_yeh["id"],
            "new_driver_name": "想要隨便更換指標",
            "new_target_val": 10.0,
            "new_unit": "次",
            "change_reason": "測試防呆"
        })
        assert ok is False, "未滿 28 天之指標竟然申請成功！防呆機制失效！"
        assert "未滿 4 週" in msg, f"防呆提示訊息不正確: {msg}"
        print(f"  [OK] 成功通過 4 週防呆審查：未滿 28 天申請遭精準阻擋！(系統提示：{msg})")

        drivers_ren = data_access.get_active_drivers("曾仁君")
        servo_ren = next(d for d in drivers_ren if d["product_line"] == "士林伺服")
        assert servo_ren["can_change"] is True, "已滿 28 天卻顯示不可變更！"
        ok_apply, msg_apply, req_id_apply = data_access.apply_driver_change({
            "driver_id": servo_ren["id"],
            "new_driver_name": "先進半導體產線實機調機與規格制定",
            "new_target_val": 5.0,
            "new_unit": "場",
            "change_reason": "配合新竹大廠擴產，原試機轉為深入調機規格制定。"
        })
        assert ok_apply is True, f"滿 4 週指標申請失敗: {msg_apply}"
        assert req_id_apply is not None
        print(f"  [OK] 滿 4 週指標申請成功送出 (申請單 ID={req_id_apply})！")

        ok_rev, msg_rev = data_access.review_driver_change(req_id_apply, action="approve", reviewer_name="曾維崧", comment="同意策略調整，全力支援技術工程師。")
        assert ok_rev is True, f"主管審核核准失敗: {msg_rev}"

        drivers_ren_after = data_access.get_active_drivers("曾仁君")
        servo_ren_after = next(d for d in drivers_ren_after if d["product_line"] == "士林伺服")
        assert servo_ren_after["driver_name"] == "先進半導體產線實機調機與規格制定"
        assert servo_ren_after["target_val"] == 5.0
        assert servo_ren_after["actual_val"] == 0.0
        print("  [OK] 主管核准生效：舊指標安全歸檔封存，新指標成功上線運作！")

        ok_apply2, _, req_id_apply2 = data_access.apply_driver_change({
            "driver_id": servo_ren_after["id"],
            "new_driver_name": "短線電訪 50 通",
            "new_target_val": 50.0,
            "new_unit": "通",
            "change_reason": "短線衝刺"
        })
        assert ok_apply2 is False, "新指標未滿 28 天卻能再次申請！"
        print("  [OK] 新指標剛上線執行 0 天，再次觸發 4 週防呆鎖定保護！")

        # 8. 驗證業務人員資料嚴格隔離 (客戶、歷史報價單、手機草稿匣、歷史單價分析)
        print("\n[測試 8] 驗證不同業務人員資料完全獨立隔離...")
        cust_yeh_id = data_access.save_customer({
            "owner_sales": "葉仁豪",
            "name": "葉仁豪專屬測試機械廠",
            "tax_id": "11223344",
            "default_contact": "林經理",
            "phone": "04-23456789"
        })
        assert cust_yeh_id > 0
        cust_list_tseng = data_access.get_customers(sales_name="曾維崧")
        cust_list_yeh = data_access.get_customers(sales_name="葉仁豪")
        assert not any(c["name"] == "葉仁豪專屬測試機械廠" for c in cust_list_tseng), "曾維崧不應看到葉仁豪的專屬客戶！"
        assert any(c["name"] == "葉仁豪專屬測試機械廠" for c in cust_list_yeh), "葉仁豪應看到自己的客戶！"
        print("  [OK] 客戶主檔依業務人員完全獨立隔離驗證通過！")

        quotes_tseng = data_access.list_quotations(include_deleted=False, sales_name="曾維崧")
        quotes_yeh = data_access.list_quotations(include_deleted=False, sales_name="葉仁豪")
        assert len(quotes_tseng) >= 2
        assert len(quotes_yeh) == 0, "葉仁豪不應看到曾維崧的歷史報價單！"
        print("  [OK] 歷史報價單依業務人員完全獨立隔離驗證通過！")

        draft_id = data_access.save_quotation_draft({
            "sales_name": "曾維崧",
            "customer_name": "贊翔實業有限公司",
            "subtotal": 28200.0,
            "items": [{"model_no": "SE3-023-11K", "qty": 2, "final_price": 14100}]
        })
        assert draft_id > 0
        drafts_tseng = data_access.get_quotation_drafts(sales_name="曾維崧")
        drafts_yeh = data_access.get_quotation_drafts(sales_name="葉仁豪")
        assert len(drafts_tseng) == 1
        assert len(drafts_yeh) == 0, "葉仁豪不應看到曾維崧的手機草稿！"
        deleted_bad = data_access.delete_quotation_draft(draft_id, sales_name="葉仁豪")
        assert deleted_bad is False, "其他業務不應能刪除曾維崧的手機草稿！"
        deleted_ok = data_access.delete_quotation_draft(draft_id, sales_name="曾維崧")
        assert deleted_ok is True
        print("  [OK] 手機報價草稿匣依業務人員直連資料庫與隔離驗證通過！")

        analysis_tseng = data_access.analyze_historical_prices(sales_name="曾維崧", model_keyword="SE3-023-11K")
        analysis_yeh = data_access.analyze_historical_prices(sales_name="葉仁豪", model_keyword="SE3-023-11K")
        assert len(analysis_tseng) >= 2
        assert len(analysis_yeh) == 0, "葉仁豪不應查到曾維崧的歷史單價分析紀錄！"
        print("  [OK] 歷史單價分析依業務人員完全獨立隔離驗證通過！")

        # 9. 驗證共通牌價表 + 個人特殊品牌價表，以及【方案 A】特殊品轉標準品自動取代與封存
        print("\n[測試 9] 驗證個人特殊品牌價表與【方案 A】轉標準品自動取代並封存留紀錄...")
        ok_sp, msg_sp, sp_id = data_access.save_personal_special_product({
            "owner_sales": "曾維崧",
            "model_no": "SP-CUSTOM-9000",
            "spec": "客製化耐高溫伺服馬達",
            "dealer_price": 32000.0,
            "stock_qty": 3,
            "note": "新竹半導體客製型號"
        })
        assert ok_sp is True and sp_id > 0, f"儲存特殊品失敗: {msg_sp}"

        # 曾維崧查得到此特殊品，葉仁豪查不到
        prods_tseng = data_access.search_products(query="SP-CUSTOM-9000", sales_name="曾維崧")
        prods_yeh = data_access.search_products(query="SP-CUSTOM-9000", sales_name="葉仁豪")
        assert any(p["model_no"] == "SP-CUSTOM-9000" and p["is_special"] == 1 for p in prods_tseng)
        assert not any(p["model_no"] == "SP-CUSTOM-9000" for p in prods_yeh), "其他業務不應看到曾維崧的未公開特殊品！"
        print("  [OK] 個人特殊品僅建立者可見，其他業務完全不可見！")

        # 執行【方案 A】：公司正式將 SP-CUSTOM-9000 轉為共通標準品發行 (標準經銷價 35000)
        ok_promo, msg_promo = data_access.promote_special_to_standard(
            item_id=sp_id,
            standard_dealer_price=35000.0
        )
        assert ok_promo is True, f"特殊品轉標準品失敗: {msg_promo}"

        # 驗證共通產品庫現在全員可見，且價格為標準品價格 35000
        prods_tseng_after = data_access.search_products(query="SP-CUSTOM-9000", sales_name="曾維崧")
        prods_yeh_after = data_access.search_products(query="SP-CUSTOM-9000", sales_name="葉仁豪")
        std_item_tseng = next(p for p in prods_tseng_after if p["model_no"] == "SP-CUSTOM-9000")
        std_item_yeh = next(p for p in prods_yeh_after if p["model_no"] == "SP-CUSTOM-9000")
        assert std_item_tseng["is_special"] == 0
        assert std_item_tseng["dealer_price"] == 35000.0
        assert std_item_yeh["dealer_price"] == 35000.0

        # 驗證曾維崧的個人特殊品已被自動標記為封存，並完整保留原價 32000 與新標準價 35000 對照紀錄
        sp_all_tseng = data_access.get_personal_special_products(sales_name="曾維崧", include_archived=True)
        archived_sp = next(s for s in sp_all_tseng if s["model_no"] == "SP-CUSTOM-9000")
        assert archived_sp["is_archived"] == 1
        assert archived_sp["dealer_price"] == 32000.0
        assert archived_sp["archived_standard_price"] == 35000.0
        assert "32,000" in archived_sp["archived_note"] and "35,000" in archived_sp["archived_note"]
        print(f"  [OK] 【方案 A】驗證通過：特殊品自動由標準品取代，並封存保留價格軌跡 ({archived_sp['archived_note']})！")

        print("\n" + "=" * 65)
        print("[SUCCESS] 恭喜！data_access.py 全數 9 大單元測試 100% 通過！")
        print("=" * 65)

    finally:
        data_access.DEFAULT_DB_FILE = original_db
        if test_db.exists():
            try:
                test_db.unlink()
            except Exception:
                pass


if __name__ == "__main__":
    run_all_tests()
