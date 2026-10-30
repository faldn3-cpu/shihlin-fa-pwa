# -*- coding: utf-8 -*-
"""
士林電機 FA 業務戰情與報價系統 - Web API 伺服器 (server.py)

【職責】
1. 作為 PWA 前端網頁與後端資料存取層 (data_access.py) 之間的通訊橋樑。
2. 提供靜態檔案託管服務，讓使用者開啟瀏覽器即可操作系統。
3. 提供乾淨嚴謹的 RESTful API 路由，處理日報、商機、案件、雙層牌價（共通標準品＋個人特殊品方案 A）、手機報價草稿與正式報價單。
"""

import os
from pathlib import Path
from typing import List, Optional, Dict, Any

from fastapi import FastAPI, HTTPException, Query
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from fastapi.responses import FileResponse
from pydantic import BaseModel, Field

import data_access

# 確保資料庫在伺服器啟動時已初始化
data_access.init_database()

app = FastAPI(
    title="士林電機 FA 業務戰情與報價系統",
    description="整合業務日報 (OGSM)、商機 (CRM)、案件 (Deal) 與跨裝置報價單管理系統",
    version="1.0.38"
)

# 開啟跨來源資源共用 (CORS)，方便本地開發與 PWA 運作
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# 前端靜態網頁目錄
WEB_DIR = Path(__file__).resolve().parent / "web"


# ==============================================================================
# Pydantic 資料驗證模型 (Request Models)
# ==============================================================================

class DailyReportCreate(BaseModel):
    report_date: str = Field(..., description="拜訪日期 YYYY-MM-DD")
    sales_name: str = Field(..., description="業務姓名")
    customer_id: Optional[int] = None
    customer_name: str = Field(..., description="客戶名稱")
    contact_name: Optional[str] = ""
    visit_type: str = Field(default="實地拜訪", description="拜訪形式")
    content_summary: str = Field(..., description="洽談重點與行程摘要")
    manager_task_id: Optional[int] = None
    create_opportunity: bool = False
    industry_type: Optional[str] = "自動化設備產業"
    product_series: Optional[str] = "士林品"
    est_quarter: Optional[str] = "Q4"
    est_scale: Optional[float] = 0.0

    create_deal: bool = False
    project_name: Optional[str] = None
    comp_brand: Optional[str] = "台達"
    comp_channel: Optional[str] = "經銷商"
    est_deal_date: Optional[str] = None
    est_amount: Optional[float] = 0.0


class OpportunityCreate(BaseModel):
    daily_report_id: Optional[int] = None
    customer_id: Optional[int] = None
    customer_name: str
    industry_type: str = "自動化設備產業"
    product_series: str = "士林品"
    est_quarter: str = "Q4"
    est_scale: float = 0.0
    status: str = "挖掘中"


class DealCaseCreate(BaseModel):
    opportunity_id: Optional[int] = None
    customer_id: Optional[int] = None
    customer_name: str
    project_name: str
    project_type: str = "一次性專案"
    comp_brand: str = "台達"
    comp_channel: str = "經銷商"
    est_deal_date: str
    est_amount: float = 0.0
    stage: str = "初訪開拓"
    win_loss_reason: Optional[str] = ""


class CustomerCreate(BaseModel):
    name: str
    owner_sales: str = "曾維崧"
    tax_id: Optional[str] = ""
    default_contact: Optional[str] = ""
    phone: Optional[str] = ""
    fax: Optional[str] = ""
    address: Optional[str] = ""
    payment_terms: Optional[str] = "月結 30 天電匯"


class SpecialProductCreate(BaseModel):
    owner_sales: str = "曾維崧"
    model_no: str
    spec: Optional[str] = ""
    list_price: float = 0.0
    dealer_price: float = 0.0
    stock_qty: int = 0
    note: Optional[str] = ""


class SpecialProductPromoteRequest(BaseModel):
    standard_dealer_price: Optional[float] = None


class QuotationItemCreate(BaseModel):
    model_no: str
    spec: Optional[str] = ""
    dealer_price: float = 0.0
    discount_pct: Optional[float] = None
    expected_price: Optional[float] = None
    final_unit_price: float = 0.0
    quantity: float = 1.0
    row_subtotal: float = 0.0
    note_lead_time: Optional[str] = ""


class QuotationCreate(BaseModel):
    prefix: str
    sequence: Optional[int] = None
    sales_name: str = "曾維崧"
    deal_id: Optional[int] = None
    quote_date: str
    valid_date: str
    customer_name: str
    tax_id: Optional[str] = ""
    contact_name: Optional[str] = ""
    phone: Optional[str] = ""
    fax: Optional[str] = ""
    delivery_address: Optional[str] = ""
    payment_terms: Optional[str] = ""
    project_name: Optional[str] = ""
    project_type: str = "一次性專案"
    subtotal: float = 0.0
    tax_rate: float = 5.0
    tax_amount: float = 0.0
    grand_total: float = 0.0
    status: str = "追蹤中"
    items: List[QuotationItemCreate] = []


class QuotationDraftCreate(BaseModel):
    sales_name: str = "曾維崧"
    customer_id: Optional[int] = None
    customer_name: str
    tax_id: Optional[str] = ""
    contact_name: Optional[str] = ""
    phone: Optional[str] = ""
    fax: Optional[str] = ""
    delivery_address: Optional[str] = ""
    payment_terms: Optional[str] = "月結 30 天電匯"
    project_name: Optional[str] = ""
    project_type: Optional[str] = "一次性專案"
    deal_id: Optional[int] = None
    crm_links: List[Dict[str, Any]] = []
    items: List[Dict[str, Any]] = []
    subtotal: float = 0.0


class DriverChangeApplyRequest(BaseModel):
    driver_id: int
    new_driver_name: str
    new_target_val: float
    new_unit: str = "次"
    change_reason: str


class DriverReviewRequest(BaseModel):
    request_id: int
    action: str
    reviewer_name: Optional[str] = ""
    comment: Optional[str] = ""


class DriverReportActualRequest(BaseModel):
    driver_id: int
    actual_val: float


class DriverUpdateSettingsRequest(BaseModel):
    driver_id: int
    driver_name: str
    target_val: float
    unit: str = "次"
    period_type: str = "週"


# ==============================================================================
# API 路由實作
# ==============================================================================

@app.get("/api/customers")
def api_get_customers(keyword: str = "", sales_name: Optional[str] = None):
    """查詢客戶主檔清單（支援依 sales_name 完全獨立隔離）"""
    return data_access.get_customers(keyword=keyword, sales_name=sales_name)


@app.post("/api/customers")
def api_save_customer(cust: CustomerCreate):
    """新增或更新業務個人客戶主檔"""
    cid = data_access.save_customer(cust.model_dump())
    return {"status": "success", "id": cid, "message": "客戶資料已成功儲存"}


@app.get("/api/products")
def api_search_products(q: str = "", limit: int = 80, sales_name: Optional[str] = None):
    """搜尋產品庫（合併業務個人未封存特殊品 ＋ 全公司共通標準品，自動執行方案 A 防呆）"""
    return data_access.search_products(query=q, limit=limit, sales_name=sales_name)


@app.get("/api/quotation-settings")
def api_get_quotation_settings():
    """取得報價單財務捨入設定、稅率與最新庫存更新時間"""
    return data_access.get_quotation_finance_settings()


@app.get("/api/special-products")
def api_get_special_products(sales_name: str = "曾維崧", include_archived: bool = True):
    """取得特定業務員個人的特殊品牌價清單（含方案 A 已轉標準品之封存紀錄）"""
    return data_access.get_personal_special_products(sales_name=sales_name, include_archived=include_archived)


@app.post("/api/special-products")
def api_save_special_product(sp_data: SpecialProductCreate):
    """新增或更新業務個人特殊品（若已存在於標準品中，依方案 A 自動轉封存）"""
    ok, msg, item_id = data_access.save_personal_special_product(sp_data.model_dump())
    if not ok:
        raise HTTPException(status_code=400, detail=msg)
    return {"status": "success", "id": item_id, "message": msg}


@app.delete("/api/special-products/{item_id}")
def api_delete_special_product(item_id: int, sales_name: str = "曾維崧"):
    """刪除業務個人的特殊品項目"""
    ok = data_access.delete_personal_special_product(item_id, sales_name=sales_name)
    if not ok:
        raise HTTPException(status_code=404, detail="找不到指定的個人特殊品或無權限刪除")
    return {"status": "success", "message": "個人特殊品已刪除"}


@app.post("/api/special-products/{item_id}/promote")
def api_promote_special_to_standard(item_id: int, req: Optional[SpecialProductPromoteRequest] = None):
    """【方案 A】將個人特殊品正式發行至公司共通標準牌價表，並自動封存個人特殊品留紀錄"""
    std_price = req.standard_dealer_price if req else None
    ok, msg = data_access.promote_special_to_standard(item_id, standard_dealer_price=std_price)
    if not ok:
        raise HTTPException(status_code=400, detail=msg)
    return {"status": "success", "message": msg}


@app.get("/api/daily-reports")
def api_get_daily_reports(start_date: Optional[str] = None, end_date: Optional[str] = None, sales_name: Optional[str] = None):
    """取得業務日報清單"""
    return data_access.get_daily_reports(start_date=start_date, end_date=end_date, sales_name=sales_name)


@app.post("/api/daily-reports")
def api_create_daily_report(report: DailyReportCreate):
    """建立業務日報 (支援三層自動連鎖：可同時產生第二層商機或第三層案件)"""
    report_dict = report.model_dump()
    report_id = data_access.save_daily_report(report_dict)

    opp_id = None
    deal_id = None

    if report.create_opportunity:
        opp_id = data_access.save_crm_opportunity({
            "daily_report_id": report_id,
            "customer_id": report.customer_id,
            "customer_name": report.customer_name,
            "industry_type": report.industry_type,
            "product_series": report.product_series,
            "est_quarter": report.est_quarter,
            "est_scale": report.est_scale,
            "status": "挖掘中"
        })

    if report.create_deal and report.project_name:
        deal_id = data_access.save_deal_case({
            "opportunity_id": opp_id,
            "customer_id": report.customer_id,
            "customer_name": report.customer_name,
            "project_name": report.project_name,
            "project_type": "一次性專案",
            "comp_brand": report.comp_brand,
            "comp_channel": report.comp_channel,
            "est_deal_date": report.est_deal_date or report.report_date,
            "est_amount": report.est_amount,
            "stage": "初訪開拓"
        })

    return {
        "status": "success",
        "daily_report_id": report_id,
        "opportunity_id": opp_id,
        "deal_id": deal_id,
        "message": "日報與連動案件儲存成功！"
    }


@app.get("/api/opportunities")
def api_get_opportunities():
    """取得 CRM 商機清單"""
    conn = data_access.get_connection()
    try:
        rows = conn.execute("SELECT * FROM crm_opportunity ORDER BY id DESC").fetchall()
        return [dict(r) for r in rows]
    finally:
        conn.close()


@app.post("/api/opportunities")
def api_create_opportunity(opp: OpportunityCreate):
    """獨立新增 CRM 商機"""
    opp_id = data_access.save_crm_opportunity(opp.model_dump())
    return {"status": "success", "id": opp_id}


@app.get("/api/deals")
def api_get_deals(stage: Optional[str] = None):
    """取得 Deal 專案案件清單"""
    return data_access.get_deal_cases(stage=stage)


@app.post("/api/deals")
def api_create_deal(deal: DealCaseCreate):
    """獨立新增 Deal 專案案件"""
    deal_id = data_access.save_deal_case(deal.model_dump())
    return {"status": "success", "id": deal_id}


@app.get("/api/quotations")
def api_list_quotations(
    keyword: str = "",
    status: str = "ALL",
    include_deleted: bool = False,
    sales_name: Optional[str] = None,
    limit: int = 200
):
    """查詢報價單清單 (支援依業務員獨立隔離、關鍵字搜尋與垃圾桶過濾)"""
    return data_access.list_quotations(
        keyword=keyword,
        status=status,
        include_deleted=include_deleted,
        sales_name=sales_name,
        limit=limit
    )


@app.get("/api/quotations/{quote_no}")
def api_get_quotation(quote_no: str):
    """讀取單筆報價單完整內容"""
    quote = data_access.get_quotation(quote_no)
    if not quote:
        raise HTTPException(status_code=404, detail="找不到指定的報價單")
    return quote


@app.post("/api/quotations")
def api_save_quotation(quote_data: QuotationCreate):
    """儲存報價單 (Append-Only，修改自動另存新版次，並綁定開單業務)"""
    header_data = quote_data.model_dump()
    items_data = header_data.pop("items", [])
    assigned_quote_no = data_access.save_quotation(header_data, items_data)
    return {
        "status": "success",
        "quote_no": assigned_quote_no,
        "message": f"報價單儲存成功：單號【{assigned_quote_no}】"
    }


@app.delete("/api/quotations/{quote_no}")
def api_delete_quotation(quote_no: str):
    """軟刪除報價單 (移至垃圾桶，保留 30 天防呆緩衝期)"""
    ok = data_access.soft_delete_quotation(quote_no)
    if not ok:
        raise HTTPException(status_code=404, detail="找不到指定的報價單或已在垃圾桶中")
    return {"status": "success", "message": f"報價單【{quote_no}】已成功移至垃圾桶"}


@app.post("/api/quotations/{quote_no}/restore")
def api_restore_quotation(quote_no: str):
    """自垃圾桶一鍵復原報價單"""
    ok = data_access.restore_quotation(quote_no)
    if not ok:
        raise HTTPException(status_code=404, detail="找不到欲復原的報價單")
    return {"status": "success", "message": f"報價單【{quote_no}】已成功自垃圾桶復原"}


@app.get("/api/price-analysis")
def api_analyze_historical_prices(
    model_keyword: str = "",
    customer_keyword: str = "",
    sales_name: Optional[str] = None,
    limit: int = 100
):
    """歷史報價單價與折讓區間查詢（依業務員 sales_name 完全獨立隔離）"""
    return data_access.analyze_historical_prices(
        model_keyword=model_keyword,
        customer_keyword=customer_keyword,
        sales_name=sales_name,
        limit=limit
    )


@app.get("/api/quotation-drafts")
def api_get_quotation_drafts(sales_name: Optional[str] = None):
    """取得手機報價草稿匣清單（直接自資料庫讀取，依業務員獨立隔離）"""
    return data_access.get_quotation_drafts(sales_name=sales_name)


@app.post("/api/quotation-drafts")
def api_save_quotation_draft(draft: QuotationDraftCreate):
    """手機端送出報價草稿（直接寫入伺服器資料庫，免經雲端硬碟）"""
    draft_id = data_access.save_quotation_draft(draft.model_dump())
    return {
        "status": "success",
        "draft_id": draft_id,
        "message": "手機報價草稿已直接存入系統草稿匣！回到電腦桌即可載入轉正式報價單。"
    }


@app.delete("/api/quotation-drafts/{draft_id}")
def api_delete_quotation_draft(draft_id: int, sales_name: Optional[str] = None):
    """刪除單筆手機報價草稿"""
    ok = data_access.delete_quotation_draft(draft_id, sales_name=sales_name)
    if not ok:
        raise HTTPException(status_code=404, detail="找不到指定的手機草稿")
    return {"status": "success", "message": "手機草稿已刪除"}


@app.delete("/api/quotation-drafts")
def api_clear_quotation_drafts(sales_name: str = "曾維崧"):
    """清空該業務員的所有手機報價草稿"""
    cnt = data_access.clear_quotation_drafts(sales_name=sales_name)
    return {"status": "success", "deleted_count": cnt, "message": "已清空您的手機報價草稿匣"}


# ==============================================================================
# 產品線領先指標 (驅動目標) 與變更審核 API
# ==============================================================================

@app.get("/api/scoreboard/drivers")
def api_get_drivers(sales_name: str = Query("曾維崧", description="業務人員姓名")):
    """取得業務員三大產品線當前生效中的領先指標 (驅動目標) 與達成率燈號"""
    drivers = data_access.get_active_drivers(sales_name)
    return {
        "status": "success",
        "sales_name": sales_name,
        "drivers": drivers
    }


@app.post("/api/scoreboard/drivers/apply-change")
def api_apply_driver_change(req_data: DriverChangeApplyRequest):
    """業務發起更換驅動指標申請 (內建未滿 4 週防呆阻擋)"""
    ok, msg, req_id = data_access.apply_driver_change(req_data.model_dump())
    if not ok:
        raise HTTPException(status_code=400, detail=msg)
    return {
        "status": "success",
        "message": msg,
        "request_id": req_id
    }


@app.post("/api/scoreboard/drivers/review")
def api_review_driver_change(rev_data: DriverReviewRequest):
    """主管審核驅動指標變更申請 (核准或駁回)"""
    ok, msg = data_access.review_driver_change(
        rev_data.request_id,
        action=rev_data.action,
        reviewer_name=rev_data.reviewer_name or "",
        comment=rev_data.comment or ""
    )
    if not ok:
        raise HTTPException(status_code=400, detail=msg)
    return {
        "status": "success",
        "message": msg
    }


@app.get("/api/scoreboard/drivers/requests")
def api_get_driver_requests(
    sales_name: Optional[str] = None,
    status: Optional[str] = None
):
    """查詢變更申請單清單 (供業務查看申請狀態或主管審核待辦)"""
    requests = data_access.get_driver_change_requests(sales_name=sales_name, status=status)
    return {
        "status": "success",
        "requests": requests
    }


@app.post("/api/scoreboard/drivers/report-actual")
def api_report_driver_actual(req_data: DriverReportActualRequest):
    """業務登記前線最新實績數值 (actual_val)"""
    ok, msg = data_access.record_driver_actual(req_data.driver_id, req_data.actual_val)
    if not ok:
        raise HTTPException(status_code=400, detail=msg)
    return {
        "status": "success",
        "message": msg
    }


@app.post("/api/scoreboard/drivers/update-settings")
def api_update_driver_settings(req_data: DriverUpdateSettingsRequest):
    """直接設定或編輯領先指標內容 (包含行動名稱、目標數值、單位與週期)"""
    ok, msg = data_access.update_driver_settings(
        req_data.driver_id,
        driver_name=req_data.driver_name,
        target_val=req_data.target_val,
        unit=req_data.unit,
        period_type=req_data.period_type
    )
    if not ok:
        raise HTTPException(status_code=400, detail=msg)
    return {
        "status": "success",
        "message": msg
    }


# ==============================================================================
# 前端靜態網頁託管
# ==============================================================================

@app.get("/")
def serve_index():
    """提供首頁入口"""
    index_file = WEB_DIR / "index.html"
    if index_file.exists():
        return FileResponse(str(index_file))
    return {"message": "士林電機 FA 系統 Web API 運行中，前端檔案準備中..."}

if WEB_DIR.exists():
    app.mount("/static", StaticFiles(directory=str(WEB_DIR)), name="static")
    app.mount("/web", StaticFiles(directory=str(WEB_DIR), html=True), name="web")
    if (WEB_DIR / "css").exists():
        app.mount("/css", StaticFiles(directory=str(WEB_DIR / "css")), name="css")
    if (WEB_DIR / "js").exists():
        app.mount("/js", StaticFiles(directory=str(WEB_DIR / "js")), name="js")

@app.get("/manifest.json")
def serve_manifest():
    return FileResponse(str(WEB_DIR / "manifest.json"))

@app.get("/sw.js")
def serve_sw():
    return FileResponse(str(WEB_DIR / "sw.js"))

@app.get("/favicon-v2.png")
def serve_favicon():
    return FileResponse(str(WEB_DIR / "favicon-v2.png"))

@app.get("/apple-touch-icon-v2.png")
def serve_apple_icon():
    return FileResponse(str(WEB_DIR / "apple-touch-icon-v2.png"))

@app.get("/icon-192x192-v2.png")
def serve_icon_192():
    return FileResponse(str(WEB_DIR / "icon-192x192-v2.png"))

@app.get("/icon-512x512-v2.png")
def serve_icon_512():
    return FileResponse(str(WEB_DIR / "icon-512x512-v2.png"))


if __name__ == "__main__":
    import uvicorn
    print("[server.py] 正在啟動士林電機 FA Web API (http://127.0.0.1:8000)...")
    uvicorn.run("server:app", host="127.0.0.1", port=8000, reload=True)
