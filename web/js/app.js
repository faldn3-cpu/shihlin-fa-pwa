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

  // 初始化預設今日日期與時鐘
  const todayStr = new Date().toISOString().split("T")[0];
  const inputReportDate = document.getElementById("inputReportDate");
  const headerClock = document.getElementById("headerClock");
  if (inputReportDate) inputReportDate.value = todayStr;
  if (headerClock) headerClock.innerText = todayStr;

  // 分頁導覽切換 (首頁純粹戰情計分板 / 日報案件 / 報價 / 特價)
  const tabHome = document.getElementById("navTabHome");
  const tabCases = document.getElementById("navTabCases");
  const tabQuotes = document.getElementById("navTabQuotes");
  const tabSpecials = document.getElementById("navTabSpecials");

  const panelHome = document.getElementById("panelHome");
  const panelCases = document.getElementById("panelCases");
  const panelQuotes = document.getElementById("panelQuotes");
  const panelSpecials = document.getElementById("panelSpecials");

  function switchTab(activeBtn, activePanel) {
    [tabHome, tabCases, tabQuotes, tabSpecials].forEach(b => b && b.classList.remove("active"));
    [panelHome, panelCases, panelQuotes, panelSpecials].forEach(p => p && p.classList.add("hidden"));

    if (activeBtn) activeBtn.classList.add("active");
    if (activePanel) activePanel.classList.remove("hidden");
  }

  if (tabHome) tabHome.addEventListener("click", () => switchTab(tabHome, panelHome));
  if (tabCases) tabCases.addEventListener("click", () => switchTab(tabCases, panelCases));
  if (tabQuotes) tabQuotes.addEventListener("click", () => switchTab(tabQuotes, panelQuotes));
  if (tabSpecials) tabSpecials.addEventListener("click", () => switchTab(tabSpecials, panelSpecials));

  // ============================================================================
  // 外觀顯示模式切換 (裝置跟隨系統 / 淺色 / 深色) - 參照微軟 Office / Windows 11
  // ============================================================================
  const selectThemeMode = document.getElementById("selectThemeMode");
  const systemDarkMedia = window.matchMedia("(prefers-color-scheme: dark)");

  function applyTheme(themeMode) {
    if (themeMode === "dark") {
      document.documentElement.setAttribute("data-theme", "dark");
    } else if (themeMode === "light") {
      document.documentElement.removeAttribute("data-theme");
    } else {
      // 裝置 (跟隨系統設定)
      if (systemDarkMedia.matches) {
        document.documentElement.setAttribute("data-theme", "dark");
      } else {
        document.documentElement.removeAttribute("data-theme");
      }
    }
  }

  function initTheme() {
    const savedTheme = localStorage.getItem("shihlin_pwa_theme_mode") || "auto";
    if (selectThemeMode) {
      selectThemeMode.value = savedTheme;
      selectThemeMode.addEventListener("change", (e) => {
        const mode = e.target.value;
        localStorage.setItem("shihlin_pwa_theme_mode", mode);
        applyTheme(mode);
      });
    }

    // 監聽作業系統深淺色即時變更
    systemDarkMedia.addEventListener("change", () => {
      const currentMode = localStorage.getItem("shihlin_pwa_theme_mode") || "auto";
      if (currentMode === "auto") {
        applyTheme("auto");
      }
    });

    // 立即套用
    applyTheme(savedTheme);
  }

  initTheme();

  // ============================================================================
  // 行動功能快捷看板 (Action Hub) 點擊事件
  // ============================================================================
  const btnQuickReport = document.getElementById("btnQuickReport");
  const btnQuickDeals = document.getElementById("btnQuickDeals");
  const btnQuickQuotes = document.getElementById("btnQuickQuotes");
  const btnQuickSpecials = document.getElementById("btnQuickSpecials");

  if (btnQuickReport) {
    btnQuickReport.addEventListener("click", () => {
      switchTab(tabCases, panelCases);
      const form = document.getElementById("dailyReportForm");
      if (form) {
        form.scrollIntoView({ behavior: "smooth", block: "start" });
        selectCustomer.focus();
      }
    });
  }

  if (btnQuickDeals) {
    btnQuickDeals.addEventListener("click", () => {
      switchTab(tabCases, panelCases);
      const dealsTable = document.getElementById("dealsTableBody");
      if (dealsTable) {
        dealsTable.scrollIntoView({ behavior: "smooth", block: "center" });
      }
    });
  }

  if (btnQuickQuotes) {
    btnQuickQuotes.addEventListener("click", () => {
      switchTab(tabQuotes, panelQuotes);
    });
  }

  if (btnQuickSpecials) {
    btnQuickSpecials.addEventListener("click", () => {
      switchTab(tabSpecials, panelSpecials);
    });
  }

  // ============================================================================
  // 主管戰情目標設定與動態計算 (支援自訂文字、數字，以及金額/百分比/次數/件數 4 大類型)
  // ============================================================================
  const modalTargetSettings = document.getElementById("modalTargetSettings");
  const btnOpenTargetModal = document.getElementById("btnOpenTargetModal");
  const btnCloseTargetModal = document.getElementById("btnCloseTargetModal");
  const btnCancelTargetModal = document.getElementById("btnCancelTargetModal");
  const btnResetTargetDefaults = document.getElementById("btnResetTargetDefaults");
  const formTargetSettings = document.getElementById("formTargetSettings");

  // 4 大指標預設範本
  const defaultTargets = {
    m1: {
      name: "年度 / 季度營收成果目標",
      type: "amount", // amount (金額-萬元), percent (百分比-%), frequency (次數-次), count (件數-件)
      target: 2250,
      actual: 1852,
      desc: "季度目標對標"
    },
    m2: {
      name: "客戶拜訪量 (行動層)",
      type: "frequency",
      target: 25,
      actual: 28,
      desc: "維持高頻客情拜訪"
    },
    m3: {
      name: "新商機挖掘 (CRM 層)",
      type: "count",
      target: 12,
      actual: 9,
      desc: "建議加強自動化設備案源"
    },
    m4: {
      name: "進行中 Deal / 標案",
      type: "count",
      target: 8,
      actual: 6,
      desc: "預估本月成案推進"
    }
  };

  function formatMoneyNum(num) {
    return new Intl.NumberFormat("zh-TW").format(num);
  }

  function getSavedTargets() {
    try {
      const saved = localStorage.getItem("shihlin_scoreboard_targets_v2");
      if (saved) {
        const parsed = JSON.parse(saved);
        if (parsed.m1 && parsed.m2 && parsed.m3 && parsed.m4) return parsed;
      }
    } catch (e) {
      console.warn("無法讀取目標設定:", e);
    }
    return defaultTargets;
  }

  // 格式化單一指標資料與號誌
  function computeMetricData(m) {
    const target = Number(m.target) || 1;
    const actual = Number(m.actual) || 0;
    const rawRate = (actual / target) * 100;
    const rate = Math.round(rawRate);
    const rateFormatted = rawRate.toFixed(1);
    const diff = target - actual;

    let unit = "";
    let actualText = "";
    let targetText = "";
    let subText = "";

    switch (m.type) {
      case "amount": // 金額 (萬元)
        unit = "萬元";
        actualText = `${formatMoneyNum(actual)} <span class="unit-text">${unit}</span>`;
        targetText = `/ 目標 ${formatMoneyNum(target)} ${unit}`;
        if (diff > 0) {
          subText = `季度目標：${formatMoneyNum(target)} ${unit} ｜ 差距：${formatMoneyNum(diff)} ${unit}`;
        } else {
          subText = `季度目標：${formatMoneyNum(target)} ${unit} ｜ 超標：${formatMoneyNum(Math.abs(diff))} ${unit}`;
        }
        break;

      case "percent": // 百分比 (%)
        unit = "%";
        actualText = `${actual} <span class="unit-text">${unit}</span>`;
        targetText = `/ 目標 ${target} ${unit}`;
        if (diff > 0) {
          subText = `目標：${target}% ｜ 尚差：${diff}%`;
        } else {
          subText = `目標：${target}% ｜ 超標：${Math.abs(diff)}%`;
        }
        break;

      case "frequency": // 次數 (次)
        unit = "次";
        actualText = `${formatMoneyNum(actual)}`;
        targetText = `/ 目標 ${formatMoneyNum(target)} ${unit}`;
        subText = `達成率 ${rate}% ｜ ${m.desc || "實地拜訪"}`;
        break;

      case "count": // 件數 (件)
      default:
        unit = "件";
        actualText = `${formatMoneyNum(actual)}`;
        targetText = `/ 目標 ${formatMoneyNum(target)} ${unit}`;
        subText = `達成率 ${rate}% ｜ ${m.desc || "推進案件"}`;
        break;
    }

    let badgeClass = "lamp-badge green";
    let badgeText = `${rate}% 達標`;
    if (rate >= 100) {
      badgeClass = "lamp-badge green";
      badgeText = "良好";
    } else if (rate >= 75) {
      badgeClass = "lamp-badge yellow";
      badgeText = "預警";
    } else {
      badgeClass = "lamp-badge red";
      badgeText = "落後";
    }

    return {
      rate,
      rateFormatted,
      actualText,
      targetText,
      subText,
      badgeClass,
      badgeText,
      progressWidth: Math.min(Math.max(rate, 0), 100) + "%"
    };
  }

  function renderScoreboardUI(targets) {
    const data1 = computeMetricData(targets.m1);
    const data2 = computeMetricData(targets.m2);
    const data3 = computeMetricData(targets.m3);
    const data4 = computeMetricData(targets.m4);

    // 指標 1 (主成果)
    const elM1Name = document.getElementById("displayMetric1Name");
    const elM1Sub = document.getElementById("displayMetric1Sub");
    const elM1Actual = document.getElementById("displayActualMetric1");
    const elM1Target = document.getElementById("displayTargetMetric1");
    const elM1Bar = document.getElementById("barMetric1Progress");
    const elM1Badge = document.getElementById("badgeMetric1Rate");

    if (elM1Name) elM1Name.innerText = targets.m1.name || "成果目標";
    if (elM1Sub) elM1Sub.innerText = data1.subText;
    if (elM1Actual) elM1Actual.innerHTML = data1.actualText;
    if (elM1Target) elM1Target.innerText = data1.targetText;
    if (elM1Bar) elM1Bar.style.width = data1.progressWidth;
    if (elM1Badge) {
      elM1Badge.className = data1.badgeClass;
      elM1Badge.innerText = `${data1.rateFormatted}% 達標`;
    }

    // 指標 2
    const elM2Name = document.getElementById("displayMetric2Name");
    const elM2Actual = document.getElementById("displayActualMetric2");
    const elM2Target = document.getElementById("displayTargetMetric2");
    const elM2Bar = document.getElementById("barMetric2Progress");
    const elM2Badge = document.getElementById("badgeMetric2Rate");
    const elM2Desc = document.getElementById("descMetric2Target");

    if (elM2Name) elM2Name.innerText = targets.m2.name;
    if (elM2Actual) elM2Actual.innerHTML = data2.actualText;
    if (elM2Target) elM2Target.innerText = data2.targetText;
    if (elM2Bar) elM2Bar.style.width = data2.progressWidth;
    if (elM2Badge) {
      elM2Badge.className = data2.badgeClass;
      elM2Badge.innerText = data2.badgeText;
    }
    if (elM2Desc) elM2Desc.innerText = data2.subText;

    // 指標 3
    const elM3Name = document.getElementById("displayMetric3Name");
    const elM3Actual = document.getElementById("displayActualMetric3");
    const elM3Target = document.getElementById("displayTargetMetric3");
    const elM3Bar = document.getElementById("barMetric3Progress");
    const elM3Badge = document.getElementById("badgeMetric3Rate");
    const elM3Desc = document.getElementById("descMetric3Target");

    if (elM3Name) elM3Name.innerText = targets.m3.name;
    if (elM3Actual) elM3Actual.innerHTML = data3.actualText;
    if (elM3Target) elM3Target.innerText = data3.targetText;
    if (elM3Bar) elM3Bar.style.width = data3.progressWidth;
    if (elM3Badge) {
      elM3Badge.className = data3.badgeClass;
      elM3Badge.innerText = data3.badgeText;
    }
    if (elM3Desc) elM3Desc.innerText = data3.subText;

    // 指標 4
    const elM4Name = document.getElementById("displayMetric4Name");
    const elM4Actual = document.getElementById("displayActualMetric4");
    const elM4Target = document.getElementById("displayTargetMetric4");
    const elM4Bar = document.getElementById("barMetric4Progress");
    const elM4Badge = document.getElementById("badgeMetric4Rate");
    const elM4Desc = document.getElementById("descMetric4Target");

    if (elM4Name) elM4Name.innerText = targets.m4.name;
    if (elM4Actual) elM4Actual.innerHTML = data4.actualText;
    if (elM4Target) elM4Target.innerText = data4.targetText;
    if (elM4Bar) elM4Bar.style.width = data4.progressWidth;
    if (elM4Badge) {
      elM4Badge.className = data4.badgeClass;
      elM4Badge.innerText = data4.badgeText;
    }
    if (elM4Desc) elM4Desc.innerText = data4.subText;
  }

  // 將設定值灌入彈窗欄位
  function populateModalFields(targets) {
    const setVal = (id, val) => {
      const el = document.getElementById(id);
      if (el) el.value = val;
    };

    setVal("inputM1Name", targets.m1.name);
    setVal("inputM1Type", targets.m1.type);
    setVal("inputM1Target", targets.m1.target);
    setVal("inputM1Actual", targets.m1.actual);
    setVal("inputM1Desc", targets.m1.desc);

    setVal("inputM2Name", targets.m2.name);
    setVal("inputM2Type", targets.m2.type);
    setVal("inputM2Target", targets.m2.target);
    setVal("inputM2Actual", targets.m2.actual);
    setVal("inputM2Desc", targets.m2.desc);

    setVal("inputM3Name", targets.m3.name);
    setVal("inputM3Type", targets.m3.type);
    setVal("inputM3Target", targets.m3.target);
    setVal("inputM3Actual", targets.m3.actual);
    setVal("inputM3Desc", targets.m3.desc);

    setVal("inputM4Name", targets.m4.name);
    setVal("inputM4Type", targets.m4.type);
    setVal("inputM4Target", targets.m4.target);
    setVal("inputM4Actual", targets.m4.actual);
    setVal("inputM4Desc", targets.m4.desc);
  }

  // 從彈窗讀取使用者輸入的值
  function readModalFields() {
    const getVal = (id) => {
      const el = document.getElementById(id);
      return el ? el.value.trim() : "";
    };
    const getNum = (id, fallback) => {
      const val = Number(getVal(id));
      return isNaN(val) ? fallback : val;
    };

    return {
      m1: {
        name: getVal("inputM1Name") || defaultTargets.m1.name,
        type: getVal("inputM1Type") || "amount",
        target: getNum("inputM1Target", defaultTargets.m1.target),
        actual: getNum("inputM1Actual", defaultTargets.m1.actual),
        desc: getVal("inputM1Desc")
      },
      m2: {
        name: getVal("inputM2Name") || defaultTargets.m2.name,
        type: getVal("inputM2Type") || "frequency",
        target: getNum("inputM2Target", defaultTargets.m2.target),
        actual: getNum("inputM2Actual", defaultTargets.m2.actual),
        desc: getVal("inputM2Desc")
      },
      m3: {
        name: getVal("inputM3Name") || defaultTargets.m3.name,
        type: getVal("inputM3Type") || "count",
        target: getNum("inputM3Target", defaultTargets.m3.target),
        actual: getNum("inputM3Actual", defaultTargets.m3.actual),
        desc: getVal("inputM3Desc")
      },
      m4: {
        name: getVal("inputM4Name") || defaultTargets.m4.name,
        type: getVal("inputM4Type") || "count",
        target: getNum("inputM4Target", defaultTargets.m4.target),
        actual: getNum("inputM4Actual", defaultTargets.m4.actual),
        desc: getVal("inputM4Desc")
      }
    };
  }

  // 初始化計分板
  const initialTargets = getSavedTargets();
  renderScoreboardUI(initialTargets);

  // 主管設定 Modal 開關事件
  if (btnOpenTargetModal) {
    btnOpenTargetModal.addEventListener("click", () => {
      const targets = getSavedTargets();
      populateModalFields(targets);
      if (modalTargetSettings) modalTargetSettings.classList.remove("hidden");
    });
  }

  function closeTargetModal() {
    if (modalTargetSettings) modalTargetSettings.classList.add("hidden");
  }

  if (btnCloseTargetModal) btnCloseTargetModal.addEventListener("click", closeTargetModal);
  if (btnCancelTargetModal) btnCancelTargetModal.addEventListener("click", closeTargetModal);

  // 恢復預設值
  if (btnResetTargetDefaults) {
    btnResetTargetDefaults.addEventListener("click", () => {
      populateModalFields(defaultTargets);
    });
  }

  if (formTargetSettings) {
    formTargetSettings.addEventListener("submit", (e) => {
      e.preventDefault();
      const newTargets = readModalFields();
      localStorage.setItem("shihlin_scoreboard_targets_v2", JSON.stringify(newTargets));
      renderScoreboardUI(newTargets);
      closeTargetModal();
    });
  }

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
