# -*- coding: utf-8 -*-
"""
士林電機 FA 業務戰情與報價系統 - Web API 伺服器 (server.py)

【職責】
1. 作為 PWA 前端網頁與後端資料存取層 (data_access.py) 之間的通訊橋樑。
2. 提供靜態檔案託管服務，讓使用者開啟瀏覽器即可操作系統。
3. 提供乾淨嚴謹的 RESTful API 路由，處理日報、商機、案件與報價單資料。
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
    description="整合業務日報 (OGSM)、商機 (CRM)、案件 (Deal) 與報價單管理系統",
    version="1.0.0"
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
    # 擴充子卡片選填欄位：如果勾選產生商機或案件，自動連鎖建立
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
    project_name: str
    project_type: str = "一次性專案"
    subtotal: float = 0.0
    tax_rate: float = 5.0
    tax_amount: float = 0.0
    grand_total: float = 0.0
    status: str = "追蹤中"
    items: List[QuotationItemCreate] = []


class DriverChangeApplyRequest(BaseModel):
    driver_id: int
    new_driver_name: str
    new_target_val: float
    new_unit: str = "次"
    change_reason: str


class DriverReviewRequest(BaseModel):
    request_id: int
    action: str  # 'approve' or 'reject'
    reviewer_name: Optional[str] = ""
    comment: Optional[str] = ""


# ==============================================================================
# API 路由實作
# ==============================================================================

@app.get("/api/customers")
def api_get_customers(keyword: str = ""):
    """查詢客戶主檔清單"""
    return data_access.get_customers(keyword=keyword)


@app.get("/api/products")
def api_search_products(q: str = ""):
    """搜尋士林產品庫"""
    return data_access.search_products(query=q)


@app.get("/api/daily-reports")
def api_get_daily_reports(start_date: Optional[str] = None, end_date: Optional[str] = None, sales_name: Optional[str] = None):
    """取得業務日報清單"""
    return data_access.get_daily_reports(start_date=start_date, end_date=end_date, sales_name=sales_name)


@app.post("/api/daily-reports")
def api_create_daily_report(report: DailyReportCreate):
    """
    建立業務日報 (支援三層自動連鎖：可同時產生第二層商機或第三層案件)
    """
    report_dict = report.model_dump()
    report_id = data_access.save_daily_report(report_dict)

    opp_id = None
    deal_id = None

    # 連鎖 A：若勾選建立 CRM 商機
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

    # 連鎖 B：若勾選建立正式 Deal 專案
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
def api_list_quotations(keyword: str = "", status: str = "ALL", include_deleted: bool = False):
    """查詢報價單清單 (含搜尋與垃圾桶過濾)"""
    return data_access.list_quotations(keyword=keyword, status=status, include_deleted=include_deleted)


@app.get("/api/quotations/{quote_no}")
def api_get_quotation(quote_no: str):
    """讀取單筆報價單完整內容"""
    quote = data_access.get_quotation(quote_no)
    if not quote:
        raise HTTPException(status_code=404, detail="找不到指定的報價單")
    return quote


@app.post("/api/quotations")
def api_save_quotation(quote_data: QuotationCreate):
    """儲存報價單 (Append-Only，修改自動另存新版次)"""
    header_data = quote_data.model_dump()
    items_data = header_data.pop("items", [])
    assigned_quote_no = data_access.save_quotation(header_data, items_data)
    return {
        "status": "success",
        "quote_no": assigned_quote_no,
        "message": f"報價單儲存成功：單號【{assigned_quote_no}】"
    }


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

# 掛載靜態資源目錄 (同時相容 /web、/css、/js、/static 與 PWA 資源)
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
