/**
 * 士林電機 FA 業務戰情與報價系統 - 前端主程式 (app.js)
 * 嚴格遵循：
 * 1. 100% 採用 Office / Windows Fluent 洗鍊線條樣式
 * 2. 徹底剔除彩色 Emoji 表情符號
 * 3. 拜訪單一紀錄規範：選取客戶自動帶入統編、電話、地址與聯絡人
 * 4. 三層漏斗：日報 ➔ 商機 ➔ 案件 連鎖建立
 */

document.addEventListener("DOMContentLoaded", () => {
  // 全域狀態快取
  let customersCache = [];

  // DOM 元素快取
  const selectCustomer = document.getElementById("selectCustomer");
  const inputTaxId = document.getElementById("inputTaxId");
  const inputPhone = document.getElementById("inputPhone");
  const inputAddress = document.getElementById("inputAddress");
  const inputContactName = document.getElementById("inputContactName");
  const autofillHint = document.getElementById("customerAutofillHint");

  const checkCreateOpp = document.getElementById("checkCreateOpp");
  const bodyOppToggle = document.getElementById("bodyOppToggle");
  const checkCreateDeal = document.getElementById("checkCreateDeal");
  const bodyDealToggle = document.getElementById("bodyDealToggle");

  const dailyReportForm = document.getElementById("dailyReportForm");
  const dailyReportsTableBody = document.getElementById("dailyReportsTableBody");
  const dealsTableBody = document.getElementById("dealsTableBody");
  const btnRefreshReports = document.getElementById("btnRefreshReports");
  const btnResetForm = document.getElementById("btnResetForm");

  // 分頁導覽切換
  const tabCases = document.getElementById("navTabCases");
  const tabQuotes = document.getElementById("navTabQuotes");
  const tabSpecials = document.getElementById("navTabSpecials");
  const tabScoreboard = document.getElementById("navTabScoreboard");

  const panelCases = document.getElementById("panelCases");
  const panelQuotes = document.getElementById("panelQuotes");
  const panelSpecials = document.getElementById("panelSpecials");
  const panelScoreboard = document.getElementById("panelScoreboard");

  // 初始化預設今日日期
  const todayStr = new Date().toISOString().split("T")[0];
  document.getElementById("inputReportDate").value = todayStr;
  document.getElementById("headerClock").innerText = todayStr;

  // ============================================================================
  // 分頁切換邏輯
  // ============================================================================
  function switchTab(activeBtn, activePanel) {
    [tabCases, tabQuotes, tabSpecials, tabScoreboard].forEach(b => b.classList.remove("active"));
    [panelCases, panelQuotes, panelSpecials, panelScoreboard].forEach(p => p.classList.add("hidden"));

    activeBtn.classList.add("active");
    activePanel.classList.remove("hidden");
  }

  tabCases.addEventListener("click", () => switchTab(tabCases, panelCases));
  tabQuotes.addEventListener("click", () => switchTab(tabQuotes, panelQuotes));
  tabSpecials.addEventListener("click", () => switchTab(tabSpecials, panelSpecials));
  tabScoreboard.addEventListener("click", () => switchTab(tabScoreboard, panelScoreboard));

  // ============================================================================
  // 子卡片手風琴展開 / 收合
  // ============================================================================
  checkCreateOpp.addEventListener("change", (e) => {
    if (e.target.checked) {
      bodyOppToggle.classList.add("expanded");
    } else {
      bodyOppToggle.classList.remove("expanded");
    }
  });

  checkCreateDeal.addEventListener("change", (e) => {
    if (e.target.checked) {
      bodyDealToggle.classList.add("expanded");
      // 若建立案件，自動預填成案日期為 30 天後
      const futureDate = new Date();
      futureDate.setDate(futureDate.getDate() + 30);
      document.getElementById("inputEstDealDate").value = futureDate.toISOString().split("T")[0];
    } else {
      bodyDealToggle.classList.remove("expanded");
    }
  });

  // ============================================================================
  // 客戶主檔載入與「自動帶入」規範 (免重複打字)
  // ============================================================================
  async function loadCustomers() {
    try {
      const resp = await fetch("/api/customers");
      if (!resp.ok) throw new Error("讀取客戶失敗");
      customersCache = await resp.json();

      selectCustomer.innerHTML = '<option value="">-- 請選擇拜訪客戶 (或打字過濾) --</option>';
      customersCache.forEach(cust => {
        const opt = document.createElement("option");
        opt.value = cust.id;
        opt.textContent = `${cust.name} (${cust.tax_id || "無統編"})`;
        selectCustomer.appendChild(opt);
      });
    } catch (err) {
      console.error("載入客戶主檔異常:", err);
    }
  }

  selectCustomer.addEventListener("change", () => {
    const custId = parseInt(selectCustomer.value, 10);
    const selected = customersCache.find(c => c.id === custId);

    if (selected) {
      inputTaxId.value = selected.tax_id || "";
      inputPhone.value = selected.phone || "";
      inputAddress.value = selected.address || "";
      if (!inputContactName.value) {
        inputContactName.value = selected.default_contact || "";
      }
      autofillHint.classList.remove("hidden");
    } else {
      inputTaxId.value = "";
      inputPhone.value = "";
      inputAddress.value = "";
      autofillHint.classList.add("hidden");
    }
  });

  // ============================================================================
  // 日報清單載入
  // ============================================================================
  async function loadDailyReports() {
    try {
      const resp = await fetch("/api/daily-reports");
      if (!resp.ok) throw new Error("讀取日報失敗");
      const list = await resp.json();

      if (list.length === 0) {
        dailyReportsTableBody.innerHTML = '<tr><td colspan="6" style="text-align:center; color:var(--text-tertiary); padding:16px;">目前尚無日報紀錄，請於上方填寫。</td></tr>';
        return;
      }

      dailyReportsTableBody.innerHTML = list.map(r => `
        <tr>
          <td><span style="font-family:monospace; font-weight:600;">${r.report_date}</span></td>
          <td>${escapeHtml(r.sales_name)}</td>
          <td><strong>${escapeHtml(r.customer_name)}</strong></td>
          <td>${escapeHtml(r.contact_name || "-")}</td>
          <td><span class="status-pill">${escapeHtml(r.visit_type)}</span></td>
          <td>${escapeHtml(r.content_summary)}</td>
        </tr>
      `).join("");
    } catch (err) {
      dailyReportsTableBody.innerHTML = `<tr><td colspan="6" style="text-align:center; color:var(--color-danger); padding:16px;">載入日報失敗：${escapeHtml(err.message)}</td></tr>`;
    }
  }

  // ============================================================================
  // Deal 專案案件看板載入
  // ============================================================================
  async function loadDeals() {
    try {
      const resp = await fetch("/api/deals");
      if (!resp.ok) throw new Error("讀取案件失敗");
      const deals = await resp.json();

      if (deals.length === 0) {
        dealsTableBody.innerHTML = '<tr><td colspan="8" style="text-align:center; color:var(--text-tertiary); padding:16px;">目前尚無進行中 Deal 案件。</td></tr>';
        return;
      }

      dealsTableBody.innerHTML = deals.map(d => `
        <tr>
          <td style="font-family:monospace;">#${d.id}</td>
          <td><strong>${escapeHtml(d.customer_name)}</strong></td>
          <td>${escapeHtml(d.project_name)}</td>
          <td><span class="status-pill warning">${escapeHtml(d.comp_brand || "無")}</span></td>
          <td style="font-family:monospace; font-weight:bold; color:var(--color-accent);">$${Number(d.est_amount || 0).toLocaleString()}</td>
          <td style="font-family:monospace;">${d.est_deal_date || "-"}</td>
          <td><span class="status-pill success">${escapeHtml(d.stage)}</span></td>
          <td style="text-align:center;">
            <button class="fluent-btn" style="height:26px; padding:0 8px; font-size:11px;" onclick="window.alert('即將帶入案件【${escapeHtml(d.project_name)}】進入報價單（里程碑 3 連動啟用！）')">
              發起報價
            </button>
          </td>
        </tr>
      `).join("");
    } catch (err) {
      dealsTableBody.innerHTML = `<tr><td colspan="8" style="text-align:center; color:var(--color-danger); padding:16px;">載入案件失敗：${escapeHtml(err.message)}</td></tr>`;
    }
  }

  // ============================================================================
  // 日報表單提交 (連動商機與案件)
  // ============================================================================
  dailyReportForm.addEventListener("submit", async (e) => {
    e.preventDefault();

    const selectedCust = customersCache.find(c => c.id === parseInt(selectCustomer.value, 10));
    const customerName = selectedCust ? selectedCust.name : selectCustomer.options[selectCustomer.selectedIndex].text;

    const payload = {
      report_date: document.getElementById("inputReportDate").value,
      sales_name: document.getElementById("inputSalesName").value.trim(),
      customer_id: selectedCust ? selectedCust.id : null,
      customer_name: customerName,
      contact_name: inputContactName.value.trim(),
      visit_type: document.getElementById("selectVisitType").value,
      content_summary: document.getElementById("textContentSummary").value.trim(),
      
      create_opportunity: checkCreateOpp.checked,
      industry_type: document.getElementById("inputIndustry").value,
      product_series: document.getElementById("inputProductSeries").value,
      est_quarter: document.getElementById("inputEstQuarter").value,
      est_scale: parseFloat(document.getElementById("inputEstScale").value) || 0,

      create_deal: checkCreateDeal.checked,
      project_name: document.getElementById("inputProjectName").value.trim(),
      comp_brand: document.getElementById("inputCompBrand").value,
      comp_channel: document.getElementById("inputCompChannel").value.trim(),
      est_deal_date: document.getElementById("inputEstDealDate").value,
      est_amount: parseFloat(document.getElementById("inputEstDealAmount").value) || 0
    };

    if (payload.create_deal && !payload.project_name) {
      alert("請填寫專案案場名稱！");
      return;
    }

    try {
      const resp = await fetch("/api/daily-reports", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload)
      });
      const res = await resp.json();

      if (!resp.ok || res.status !== "success") {
        throw new Error(res.detail || res.message || "儲存失敗");
      }

      alert("日報與連動案件儲存成功！");
      
      // 清空洽談內容與子卡片
      document.getElementById("textContentSummary").value = "";
      checkCreateOpp.checked = false;
      bodyOppToggle.classList.remove("expanded");
      checkCreateDeal.checked = false;
      bodyDealToggle.classList.remove("expanded");

      // 刷新清單
      loadDailyReports();
      loadDeals();
    } catch (err) {
      alert(`儲存失敗: ${err.message}`);
    }
  });

  btnResetForm.addEventListener("click", () => {
    dailyReportForm.reset();
    document.getElementById("inputReportDate").value = todayStr;
    document.getElementById("inputSalesName").value = "曾維崧";
    inputTaxId.value = "";
    inputPhone.value = "";
    inputAddress.value = "";
    autofillHint.classList.add("hidden");
    bodyOppToggle.classList.remove("expanded");
    bodyDealToggle.classList.remove("expanded");
  });

  btnRefreshReports.addEventListener("click", () => {
    loadDailyReports();
    loadDeals();
  });

  // 工具函式：XSS 安全防護
  function escapeHtml(str) {
    if (!str) return "";
    return String(str)
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#039;");
  }

  // 初始載入
  loadCustomers();
  loadDailyReports();
  loadDeals();
});
