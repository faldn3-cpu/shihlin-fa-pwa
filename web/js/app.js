/**
 * 士林電機 FA 業務戰情與報價系統 - 前端主程式 (app.js)
 * 嚴格遵循：
 * 1. 100% 採用 Office / Windows Fluent 洗鍊線條樣式
 * 2. 徹底剔除彩色 Emoji 表情符號，採用洗鍊純文字狀態號誌 [達標] / [預警] / [落後]
 * 3. 業績卡更名：士林伺服 (LEB)、士林變頻器 (LNH)、士林PLC (LPB+LPN+LPD 加總)
 * 4. 曾維崧為主管，免計個人目標，直接看全區總戰況
 * 5. 目標維護支援 Excel 貼上自動解析 + 在線 Excel 試算表格子直接點選微調
 * 6. 拜訪單一紀錄規範：選取客戶自動帶入統編、電話、地址與聯絡人
 * 7. 三層漏斗：日報 ➔ 商機 ➔ 案件 連鎖建立
 */

document.addEventListener("DOMContentLoaded", () => {
  // ============================================================================
  // 全域變數與快取
  // ============================================================================
  let customersCache = [];

  // ============================================================================
  // 1. 頂部即時時鐘 (MM-DD HH:mm) - 確保單行顯示
  // ============================================================================
  const headerClock = document.getElementById("headerClock");
  function updateHeaderClock() {
    const now = new Date();
    const mm = String(now.getMonth() + 1).padStart(2, "0");
    const dd = String(now.getDate()).padStart(2, "0");
    const hh = String(now.getHours()).padStart(2, "0");
    const min = String(now.getMinutes()).padStart(2, "0");
    if (headerClock) {
      headerClock.innerText = `${mm}-${dd} ${hh}:${min}`;
    }
  }
  updateHeaderClock();
  setInterval(updateHeaderClock, 15000); // 每 15 秒更新一次分鐘

  // ============================================================================
  // 2. 四方格應用程式滑出抽屜導覽 (App Launcher Drawer)
  // ============================================================================
  const btnAppLauncher = document.getElementById("btnAppLauncher");
  const drawerOverlay = document.getElementById("drawerOverlay");
  const btnCloseDrawer = document.getElementById("btnCloseDrawer");

  const navItemHome = document.getElementById("navItemHome");
  const navItemCases = document.getElementById("navItemCases");
  const navItemQuotes = document.getElementById("navItemQuotes");
  const navItemSpecials = document.getElementById("navItemSpecials");
  const navItemSettings = document.getElementById("navItemSettings");
  const navItemAdminTargets = document.getElementById("navItemAdminTargets");

  const panelHome = document.getElementById("panelHome");
  const panelCases = document.getElementById("panelCases");
  const panelQuotes = document.getElementById("panelQuotes");
  const panelSpecials = document.getElementById("panelSpecials");
  const panelSettings = document.getElementById("panelSettings");
  const panelAdminTargets = document.getElementById("panelAdminTargets");

  const navItems = [
    { btn: navItemHome, panel: panelHome },
    { btn: navItemCases, panel: panelCases },
    { btn: navItemQuotes, panel: panelQuotes },
    { btn: navItemSpecials, panel: panelSpecials },
    { btn: navItemSettings, panel: panelSettings },
    { btn: navItemAdminTargets, panel: panelAdminTargets }
  ];

  function openDrawer() {
    if (drawerOverlay) drawerOverlay.classList.remove("hidden");
  }

  function closeDrawer() {
    if (drawerOverlay) drawerOverlay.classList.add("hidden");
  }

  if (btnAppLauncher) btnAppLauncher.addEventListener("click", openDrawer);
  if (btnCloseDrawer) btnCloseDrawer.addEventListener("click", closeDrawer);
  if (drawerOverlay) {
    drawerOverlay.addEventListener("click", (e) => {
      if (e.target === drawerOverlay) closeDrawer();
    });
  }

  function switchPage(activeBtn, activePanel) {
    navItems.forEach(item => {
      if (item.btn) item.btn.classList.remove("active");
      if (item.panel) item.panel.classList.add("hidden");
    });

    if (activeBtn) activeBtn.classList.add("active");
    if (activePanel) {
      activePanel.classList.remove("hidden");
      window.scrollTo({ top: 0, behavior: "smooth" });
    }
    closeDrawer();
  }

  navItems.forEach(item => {
    if (item.btn && item.panel) {
      item.btn.addEventListener("click", () => switchPage(item.btn, item.panel));
    }
  });

  // ============================================================================
  // 3. 行動功能快捷看板 (Action Hub) 點擊事件
  // ============================================================================
  const btnQuickReport = document.getElementById("btnQuickReport");
  const btnQuickDeals = document.getElementById("btnQuickDeals");
  const btnQuickQuotes = document.getElementById("btnQuickQuotes");
  const btnQuickSpecials = document.getElementById("btnQuickSpecials");

  if (btnQuickReport) {
    btnQuickReport.addEventListener("click", () => {
      switchPage(navItemCases, panelCases);
      const form = document.getElementById("dailyReportForm");
      if (form) {
        form.scrollIntoView({ behavior: "smooth", block: "start" });
        const custSelect = document.getElementById("selectCustomer");
        if (custSelect) custSelect.focus();
      }
    });
  }

  if (btnQuickDeals) {
    btnQuickDeals.addEventListener("click", () => {
      switchPage(navItemCases, panelCases);
      const dealsTable = document.getElementById("dealsTableBody");
      if (dealsTable) dealsTable.scrollIntoView({ behavior: "smooth", block: "center" });
    });
  }

  if (btnQuickQuotes) {
    btnQuickQuotes.addEventListener("click", () => {
      switchPage(navItemQuotes, panelQuotes);
    });
  }

  if (btnQuickSpecials) {
    btnQuickSpecials.addEventListener("click", () => {
      switchPage(navItemSpecials, panelSpecials);
    });
  }

  // ============================================================================
  // 4. 設定頁面控制 (外觀顯示 / 字型 / 字級 / 業務身份切換)
  // ============================================================================
  const selectThemeMode = document.getElementById("selectThemeMode");
  const selectSystemFont = document.getElementById("selectSystemFont");
  const selectFontSize = document.getElementById("selectFontSize");
  const selectActiveSales = document.getElementById("selectActiveSales");

  const headerUserAvatar = document.getElementById("headerUserAvatar");
  const currentSalesDisplay = document.getElementById("currentSalesDisplay");
  const displaySalesNameTarget = document.getElementById("displaySalesNameTarget");
  const inputSalesName = document.getElementById("inputSalesName");

  const systemDarkMedia = window.matchMedia("(prefers-color-scheme: dark)");

  function applyTheme(themeMode) {
    if (themeMode === "dark") {
      document.documentElement.setAttribute("data-theme", "dark");
    } else if (themeMode === "light") {
      document.documentElement.removeAttribute("data-theme");
    } else {
      // 裝置 (跟隨作業系統)
      if (systemDarkMedia.matches) {
        document.documentElement.setAttribute("data-theme", "dark");
      } else {
        document.documentElement.removeAttribute("data-theme");
      }
    }
  }

  function applyFont(fontMode) {
    if (fontMode === "system") {
      document.documentElement.style.setProperty(
        "--app-font-family",
        "-apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif"
      );
    } else {
      document.documentElement.style.setProperty(
        "--app-font-family",
        "'Microsoft JhengHei', '微軟正黑體', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif"
      );
    }
  }

  function applyFontSize(size) {
    document.documentElement.style.setProperty("--app-font-size", size);
  }

  function updateActiveSalesUI(salesName) {
    const isManager = salesName === "曾維崧";
    if (headerUserAvatar) {
      headerUserAvatar.innerText = salesName.charAt(0) || "曾";
    }
    if (currentSalesDisplay) {
      currentSalesDisplay.innerText = isManager ? `${salesName} (主管)` : salesName;
    }
    if (displaySalesNameTarget) {
      displaySalesNameTarget.innerText = salesName;
    }
    if (inputSalesName) {
      inputSalesName.value = salesName;
    }
    renderHomeDashboard();
  }

  function initSettings() {
    // 1. 外觀主題
    const savedTheme = localStorage.getItem("shihlin_pwa_theme_mode") || "auto";
    if (selectThemeMode) {
      selectThemeMode.value = savedTheme;
      selectThemeMode.addEventListener("change", (e) => {
        const mode = e.target.value;
        localStorage.setItem("shihlin_pwa_theme_mode", mode);
        applyTheme(mode);
      });
    }
    systemDarkMedia.addEventListener("change", () => {
      const mode = localStorage.getItem("shihlin_pwa_theme_mode") || "auto";
      if (mode === "auto") applyTheme("auto");
    });
    applyTheme(savedTheme);

    // 2. 系統字型
    const savedFont = localStorage.getItem("shihlin_pwa_font_mode") || "jhenghei";
    if (selectSystemFont) {
      selectSystemFont.value = savedFont;
      selectSystemFont.addEventListener("change", (e) => {
        const font = e.target.value;
        localStorage.setItem("shihlin_pwa_font_mode", font);
        applyFont(font);
      });
    }
    applyFont(savedFont);

    // 3. 字級大小
    const savedSize = localStorage.getItem("shihlin_pwa_font_size") || "12pt";
    if (selectFontSize) {
      selectFontSize.value = savedSize;
      selectFontSize.addEventListener("change", (e) => {
        const sz = e.target.value;
        localStorage.setItem("shihlin_pwa_font_size", sz);
        applyFontSize(sz);
      });
    }
    applyFontSize(savedSize);

    // 4. 業務身份切換 (預設曾維崧主管)
    const savedSales = localStorage.getItem("shihlin_pwa_active_sales") || "曾維崧";
    if (selectActiveSales) {
      selectActiveSales.value = savedSales;
      selectActiveSales.addEventListener("change", (e) => {
        const sName = e.target.value;
        localStorage.setItem("shihlin_pwa_active_sales", sName);
        updateActiveSalesUI(sName);
      });
    }
    updateActiveSalesUI(savedSales);
  }

  // ============================================================================
  // 5. 業務目標數據庫與試算模型 (對標 Excel 實況與曾維崧主管模式)
  // ============================================================================
  const defaultTargetData = {
    monthTitle: "10 月營收挑戰目標",
    regionTotal: {
      target: 9300,
      actual: 10370
    },
    products: {
      LEB: { name: "士林伺服", target: 516, actual: 196 },
      LNH: { name: "士林變頻器", target: 516, actual: 243 },
      PLC: { name: "士林PLC", target: 210, actual: 246 }
    },
    salesTargets: {
      "曾維崧": { isManager: true, target: 0, actual: 0, leb: 0, lnh: 0, plc: 0 },
      "何宛茹": { target: 550, actual: 238, leb: 28.21, lnh: 4.86, plc: 0.30 },
      "邱文輝": { target: 1400, actual: 1388, leb: 3.65, lnh: 9.83, plc: 143.75 },
      "張何達": { target: 900, actual: 1451, leb: 32.85, lnh: 49.62, plc: 1.35 },
      "曾仁君": { target: 810, actual: 802, leb: 28.20, lnh: 4.90, plc: 12.00 },
      "溫達仁": { target: 1200, actual: 1147, leb: 42.10, lnh: 35.00, plc: 20.30 },
      "楊家豪": { target: 700, actual: 721, leb: 18.50, lnh: 22.40, plc: 15.00 },
      "莊富丞": { target: 2400, actual: 3858, leb: 65.00, lnh: 78.50, plc: 35.20 },
      "葉仁豪": { target: 700, actual: 473, leb: 15.00, lnh: 18.20, plc: 8.50 },
      "張書偉": { target: 600, actual: 278, leb: 9.40, lnh: 4.08, plc: 3.26 },
      "黃柏翰": { target: 40, actual: 12, leb: 1.50, lnh: 2.10, plc: 0.80 }
    }
  };

  function getTargetDatabase() {
    try {
      const saved = localStorage.getItem("shihlin_target_database_v2");
      if (saved) {
        const parsed = JSON.parse(saved);
        if (parsed.regionTotal && parsed.salesTargets && parsed.products) {
          return parsed;
        }
      }
    } catch (e) {
      console.warn("無法讀取自訂目標數據:", e);
    }
    return defaultTargetData;
  }

  function saveTargetDatabase(db) {
    localStorage.setItem("shihlin_target_database_v2", JSON.stringify(db));
  }

  function formatMoney(num) {
    return new Intl.NumberFormat("zh-TW").format(num);
  }

  function getStatusBadge(rate) {
    if (rate >= 100) {
      return { className: "status-badge success", text: `${rate.toFixed(1)}% 達標` };
    } else if (rate >= 70) {
      return { className: "status-badge warning", text: `${rate.toFixed(1)}% 預警` };
    } else {
      return { className: "status-badge danger", text: `${rate.toFixed(1)}% 落後` };
    }
  }

  // 渲染首頁目標與戰況看板 (含主管檢視模式與一般業務模式)
  function renderHomeDashboard() {
    const db = getTargetDatabase();
    const currentSales = localStorage.getItem("shihlin_pwa_active_sales") || "曾維崧";
    const isManager = currentSales === "曾維崧";
    const salesInfo = db.salesTargets[currentSales] || {
      target: 800, actual: 700, leb: 20.0, lnh: 15.0, plc: 10.0
    };

    // 1. 全區核心主目標
    const elMetric1Name = document.getElementById("displayMetric1Name");
    const elMetric1Sub = document.getElementById("displayMetric1Sub");
    const elBadgeMetric1 = document.getElementById("badgeMetric1Rate");
    const elDisplayActualMetric1 = document.getElementById("displayActualMetric1");
    const elDisplayTargetMetric1 = document.getElementById("displayTargetMetric1");
    const elBarMetric1Progress = document.getElementById("barMetric1Progress");

    const totalTarget = db.regionTotal.target || 9300;
    const totalActual = db.regionTotal.actual || 10370;
    const totalRate = (totalActual / totalTarget) * 100;
    const totalDiff = totalActual - totalTarget;

    if (elMetric1Name) elMetric1Name.innerText = db.monthTitle || "10 月營收挑戰目標";
    if (elMetric1Sub) {
      if (totalDiff >= 0) {
        elMetric1Sub.innerText = `全區合計：${formatMoney(totalActual)} 萬元 / 目標 ${formatMoney(totalTarget)} 萬元 (超標 ${formatMoney(totalDiff)} 萬元)`;
      } else {
        elMetric1Sub.innerText = `全區合計：${formatMoney(totalActual)} 萬元 / 目標 ${formatMoney(totalTarget)} 萬元 (尚差 ${formatMoney(Math.abs(totalDiff))} 萬元)`;
      }
    }
    if (elBadgeMetric1) {
      const badge = getStatusBadge(totalRate);
      elBadgeMetric1.className = badge.className;
      elBadgeMetric1.innerText = badge.text;
    }
    if (elDisplayActualMetric1) {
      elDisplayActualMetric1.innerHTML = `${formatMoney(totalActual)} <span class="unit-text">萬元</span>`;
    }
    if (elDisplayTargetMetric1) {
      elDisplayTargetMetric1.innerText = `全區目標 ${formatMoney(totalTarget)} 萬元`;
    }
    if (elBarMetric1Progress) {
      elBarMetric1Progress.style.width = `${Math.min(Math.max(totalRate, 0), 100)}%`;
    }

    // 2. 個人進度摘要列 (主管模式 vs 一般業務個人進度)
    const elPersonalTargetLabel = document.getElementById("personalTargetLabel");
    const elDisplayPersonalRev = document.getElementById("displayPersonalRev");

    if (isManager) {
      // 曾維崧為主管，沒有個人目標，直接看全區戰況
      if (elPersonalTargetLabel) {
        elPersonalTargetLabel.innerHTML = `主管檢視模式 (<span id="displaySalesNameTarget">曾維崧</span>)：`;
      }
      if (elDisplayPersonalRev) {
        elDisplayPersonalRev.innerText = `免計個人目標 ｜ 掌舵全區總目標 ${formatMoney(totalTarget)} 萬，實績 ${formatMoney(totalActual)} 萬 (${totalRate.toFixed(1)}% 達標)`;
      }
    } else {
      // 一般業務人員
      const pTarget = salesInfo.target || 1;
      const pActual = salesInfo.actual || 0;
      const pRate = ((pActual / pTarget) * 100).toFixed(1);
      const pDiff = pTarget - pActual;

      if (elPersonalTargetLabel) {
        elPersonalTargetLabel.innerHTML = `個人進度 (<span id="displaySalesNameTarget">${currentSales}</span>)：`;
      }
      if (elDisplayPersonalRev) {
        if (pDiff <= 0) {
          elDisplayPersonalRev.innerText = `${formatMoney(pActual)} 萬元 / 目標 ${formatMoney(pTarget)} 萬元 (${pRate}%，超標 ${formatMoney(Math.abs(pDiff))} 萬元)`;
        } else {
          elDisplayPersonalRev.innerText = `${formatMoney(pActual)} 萬元 / 目標 ${formatMoney(pTarget)} 萬元 (${pRate}%，尚差 ${formatMoney(pDiff)} 萬元)`;
        }
      }
    }

    // 3. 自製品三大卡片 (士林伺服 LEB、士林變頻器 LNH、士林PLC)
    const prodLEB = db.products.LEB || { target: 516, actual: 196 };
    const prodLNH = db.products.LNH || { target: 516, actual: 243 };
    const prodPLC = db.products.PLC || { target: 210, actual: 246 };

    renderProductCardUI({
      cardTitle: "士林伺服",
      badgeId: "badgeProdLEB",
      labelPersonalId: "labelPersonalLEB",
      personalValId: "displayPersonalLEB",
      teamValId: "displayTeamLEB",
      barId: "barProdLEB",
      descId: "descProdLEB",
      isManager: isManager,
      personalVal: salesInfo.leb || 0,
      prodData: prodLEB
    });

    renderProductCardUI({
      cardTitle: "士林變頻器",
      badgeId: "badgeProdLNH",
      labelPersonalId: "labelPersonalLNH",
      personalValId: "displayPersonalLNH",
      teamValId: "displayTeamLNH",
      barId: "barProdLNH",
      descId: "descProdLNH",
      isManager: isManager,
      personalVal: salesInfo.lnh || 0,
      prodData: prodLNH
    });

    renderProductCardUI({
      cardTitle: "士林PLC",
      badgeId: "badgeProdPLC",
      labelPersonalId: "labelPersonalPLC",
      personalValId: "displayPersonalPLC",
      teamValId: "displayTeamPLC",
      barId: "barProdPLC",
      descId: "descProdPLC",
      isManager: isManager,
      personalVal: salesInfo.plc || 0,
      prodData: prodPLC
    });
  }

  function renderProductCardUI(cfg) {
    const badgeEl = document.getElementById(cfg.badgeId);
    const labelEl = document.getElementById(cfg.labelPersonalId);
    const personalEl = document.getElementById(cfg.personalValId);
    const teamEl = document.getElementById(cfg.teamValId);
    const barEl = document.getElementById(cfg.barId);
    const descEl = document.getElementById(cfg.descId);

    const target = cfg.prodData.target || 1;
    const actual = cfg.prodData.actual || 0;
    const rate = (actual / target) * 100;
    const diff = target - actual;

    if (badgeEl) {
      const b = getStatusBadge(rate);
      badgeEl.className = b.className;
      badgeEl.innerText = b.text;
    }

    if (cfg.isManager) {
      // 主管模式：個人實績欄位直接標註「全區累計」
      if (labelEl) labelEl.innerText = "全區實績";
      if (personalEl) personalEl.innerHTML = `${formatMoney(actual)} <span class="unit-text">萬元</span>`;
    } else {
      // 一般業務：個人實績欄位
      if (labelEl) labelEl.innerText = "個人實績";
      if (personalEl) personalEl.innerHTML = `${cfg.personalVal.toFixed(2)} <span class="unit-text">萬元</span>`;
    }

    if (teamEl) {
      teamEl.innerText = `${formatMoney(actual)} / ${formatMoney(target)} 萬元`;
    }
    if (barEl) {
      barEl.style.width = `${Math.min(Math.max(rate, 0), 100)}%`;
    }
    if (descEl) {
      if (diff > 0) {
        descEl.innerText = `全區目標 ${formatMoney(target)} 萬 ｜ 差距 ${formatMoney(diff)} 萬`;
      } else {
        descEl.innerText = `全區目標 ${formatMoney(target)} 萬 ｜ 超標 ${formatMoney(Math.abs(diff))} 萬`;
      }
    }
  }

  // ============================================================================
  // 6. 左右雙 Excel 接收表格 (左：挑戰目標總表 / 右：產品出貨明細表)
  // 支援直接按 Ctrl+V 完整還原 Excel 欄位與列數
  // ============================================================================
  const containerTableTargets = document.getElementById("containerTableTargets");
  const promptTargets = document.getElementById("promptTargets");
  const tablePastedTargets = document.getElementById("tablePastedTargets");
  const badgeTargetCount = document.getElementById("badgeTargetCount");
  const btnPasteTargetDemo = document.getElementById("btnPasteTargetDemo");
  const btnClearTargetTable = document.getElementById("btnClearTargetTable");

  const containerTableShipments = document.getElementById("containerTableShipments");
  const promptShipments = document.getElementById("promptShipments");
  const tablePastedShipments = document.getElementById("tablePastedShipments");
  const badgeShipmentCount = document.getElementById("badgeShipmentCount");
  const btnPasteShipmentDemo = document.getElementById("btnPasteShipmentDemo");
  const btnClearShipmentTable = document.getElementById("btnClearShipmentTable");

  const btnLoadSampleExcel = document.getElementById("btnLoadSampleExcel");
  const btnResetAdminDefaults = document.getElementById("btnResetAdminDefaults");
  const btnConfirmTablesReady = document.getElementById("btnConfirmTablesReady");

  // 將剪貼簿 TSV 文字渲染為原生 Excel 格式表格
  function renderPastedTable(tsvText, tableEl, promptEl, badgeEl) {
    if (!tsvText || !tsvText.trim() || !tableEl) return;
    const lines = tsvText.trim().split(/\r?\n/);
    if (lines.length === 0) return;

    let html = "";
    let maxCols = 0;

    lines.forEach((line, rIdx) => {
      const cols = line.split("\t");
      if (cols.length > maxCols) maxCols = cols.length;

      html += "<tr>";
      cols.forEach(col => {
        const val = col.trim();
        const tag = rIdx === 0 ? "th" : "td";
        // 判斷是否為數字靠右排
        const isNum = !isNaN(parseFloat(val.replace(/,/g, "").replace(/%/g, ""))) && !val.includes("月") && !val.includes("區");
        const alignStyle = isNum ? "text-align:right;" : "text-align:left;";
        html += `<${tag} style="${alignStyle}">${escapeHtml(val || "-")}</${tag}>`;
      });
      html += "</tr>";
    });

    tableEl.innerHTML = html;
    tableEl.classList.remove("hidden");
    if (promptEl) promptEl.classList.add("hidden");
    if (badgeEl) {
      badgeEl.innerText = `${lines.length} 列 × ${maxCols} 欄 (已接收)`;
      badgeEl.classList.add("active");
    }
  }

  function clearTable(tableEl, promptEl, badgeEl) {
    if (tableEl) {
      tableEl.innerHTML = "";
      tableEl.classList.add("hidden");
    }
    if (promptEl) promptEl.classList.remove("hidden");
    if (badgeEl) {
      badgeEl.innerText = "待貼入";
      badgeEl.classList.remove("active");
    }
  }

  // 監聽左邊表格容器 paste 事件
  if (containerTableTargets) {
    containerTableTargets.addEventListener("paste", (e) => {
      e.preventDefault();
      const clipboardData = e.clipboardData || window.clipboardData;
      const text = clipboardData.getData("text");
      if (text) {
        renderPastedTable(text, tablePastedTargets, promptTargets, badgeTargetCount);
        localStorage.setItem("shihlin_pasted_targets_raw", text);
      }
    });
  }

  // 監聽右邊表格容器 paste 事件
  if (containerTableShipments) {
    containerTableShipments.addEventListener("paste", (e) => {
      e.preventDefault();
      const clipboardData = e.clipboardData || window.clipboardData;
      const text = clipboardData.getData("text");
      if (text) {
        renderPastedTable(text, tablePastedShipments, promptShipments, badgeShipmentCount);
        localStorage.setItem("shihlin_pasted_shipments_raw", text);
      }
    });
  }

  // 10 月真實目標總表範例
  const sampleTargetsExcelTSV = `北區\t10\t月挑戰目標\t9,300\t萬
項次\t營業員\t年計月目標\t10月挑戰目標\t10月堆疊業績\t10/21累積目標\t10/21實績\t堆疊目標差異(100%)\t挑戰目標差異\t挑戰目標達成率\t前一個月實績
經銷\t張何達\t758\t900\t900\t610\t1,451\t551\t551\t161%\t957
\t葉仁豪\t664\t700\t705\t477\t473\t-231\t-227\t68%\t692
\t邱文輝\t1,450\t1,400\t1,400\t948\t1,388\t-12\t-12\t99%\t1,696
直賣\t曾仁君\t477\t810\t810\t549\t802\t-8\t-8\t99%\t1,754
\t溫達仁\t834\t1,200\t1,200\t813\t1,147\t-53\t-53\t96%\t863
\t楊家豪\t658\t700\t720\t488\t721\t1\t21\t103%\t1,218
\t莊富丞\t1,206\t2,400\t3,650\t2,473\t3,858\t208\t1,458\t161%\t2,942
\t何宛茹\t401\t550\t398\t270\t238\t-160\t-312\t43%\t371
\t張書偉\t612\t600\t629\t426\t278\t-351\t-322\t46%\t513
\t黃柏翰\t\t40\t40\t\t12\t\t\t\t20
\t合計\t7,060\t9,300\t10,452\t7,053\t10,370\t-54\t1,070\t112%\t11,026
北區\t自製品\t挑戰目標\t1,242\t萬
1\tLEB\t468\t516\t382\t259\t196\t-63\t-320\t38%\t283
2\tLNH\t567\t516\t320\t217\t243\t26\t-273\t47%\t691
3\tLPN+LPB+LP\t259\t210\t220\t149\t246\t97\t36\t117%\t194
\t合計\t1,294\t1,242\t922\t625\t685\t60\t-557\t55%\t1,168`;

  // 10 月真實出貨明細樞紐表範例
  const sampleShipmentsExcelTSV = `出貨金額(萬元)\t\t月份\t10
營業員\t產品別\t10\t總計
何宛茹-台北\tLOG (太陽能變流器)\t52.30\t52.30
何宛茹-台北 合計\t\t52.30\t52.30
何宛茹-桃園\tBHP (電湧保護器)\t8.94\t8.94
\tLCA (三菱大型PLC)\t25.84\t25.84
\tLCH (台製品端子臺/線材)\t.47\t.47
\tLEA (三菱伺服)\t116.59\t116.59
\tLEB (士林伺服)\t28.21\t28.21
\tLER (減速機)\t37.20\t37.20
\tLFB (PATLITE警示燈/蜂鳴器)\t2.55\t2.55
\tLNH (士林變頻器)\t4.86\t4.86
\tLNM (三菱變頻器)\t25.89\t25.89
\tLPB (士林小型PLC)\t.30\t.30
\tLPC (三菱小型PLC)\t2.49\t2.49
\tLSA (松下感測器)\t15.13\t15.13
何宛茹-桃園 合計\t\t268.47\t268.47
邱文輝-台北\tLSA (松下感測器)\t5.03\t5.03
邱文輝-台北 合計\t\t5.03\t5.03
邱文輝-桃園\tLCA (三菱大型PLC)\t196.40\t196.40
\tLCG (三菱人機)\t12.33\t12.33
\tLEA (三菱伺服)\t283.96\t283.96
\tLEB (士林伺服)\t3.65\t3.65
\tLNH (士林變頻器)\t9.83\t9.83
\tLPB (士林小型PLC)\t11.55\t11.55
\tLPC (三菱小型PLC)\t656.16\t656.16
\tLPN (小型PLC_技提)\t132.20\t132.20
\tLSA (松下感測器)\t22.02\t22.02
邱文輝-桃園 合計\t\t1,400.39\t1,400.39
張何達-台北\tLCE (士林人機)\t.36\t.36
\tLPC (三菱小型PLC)\t.21\t.21
\tLPN (小型PLC_技提)\t.50\t.50
張何達-台北 合計\t\t1.07\t1.07
張何達-桃園\tLEA (三菱伺服)\t453.05\t453.05
\tLEB (士林伺服)\t32.85\t32.85
\tLNH (士林變頻器)\t49.62\t49.62
\tLNM (三菱變頻器)\t502.11\t502.11
\tLPB (士林小型PLC)\t.21\t.21
\tLPC (三菱小型PLC)\t400.56\t400.56
\tLPN (小型PLC_技提)\t.64\t.64
張何達-桃園 合計\t\t1,573.17\t1,573.17
張書偉-新竹\tLEA (三菱伺服)\t151.83\t151.83
\tLEB (士林伺服)\t9.40\t9.40
\tLNH (士林變頻器)\t4.08\t4.08
\tLPB (士林小型PLC)\t1.11\t1.11
\tLPC (三菱小型PLC)\t56.30\t56.30
\tLPN (小型PLC_技提)\t2.15\t2.15`;

  // 左表按鈕
  if (btnPasteTargetDemo) {
    btnPasteTargetDemo.addEventListener("click", () => {
      renderPastedTable(sampleTargetsExcelTSV, tablePastedTargets, promptTargets, badgeTargetCount);
      localStorage.setItem("shihlin_pasted_targets_raw", sampleTargetsExcelTSV);
    });
  }
  if (btnClearTargetTable) {
    btnClearTargetTable.addEventListener("click", () => {
      clearTable(tablePastedTargets, promptTargets, badgeTargetCount);
      localStorage.removeItem("shihlin_pasted_targets_raw");
    });
  }

  // 右表按鈕
  if (btnPasteShipmentDemo) {
    btnPasteShipmentDemo.addEventListener("click", () => {
      renderPastedTable(sampleShipmentsExcelTSV, tablePastedShipments, promptShipments, badgeShipmentCount);
      localStorage.setItem("shihlin_pasted_shipments_raw", sampleShipmentsExcelTSV);
    });
  }
  if (btnClearShipmentTable) {
    btnClearShipmentTable.addEventListener("click", () => {
      clearTable(tablePastedShipments, promptShipments, badgeShipmentCount);
      localStorage.removeItem("shihlin_pasted_shipments_raw");
    });
  }

  // 頂部全部載入/清空按鈕
  if (btnLoadSampleExcel) {
    btnLoadSampleExcel.addEventListener("click", () => {
      renderPastedTable(sampleTargetsExcelTSV, tablePastedTargets, promptTargets, badgeTargetCount);
      renderPastedTable(sampleShipmentsExcelTSV, tablePastedShipments, promptShipments, badgeShipmentCount);
      localStorage.setItem("shihlin_pasted_targets_raw", sampleTargetsExcelTSV);
      localStorage.setItem("shihlin_pasted_shipments_raw", sampleShipmentsExcelTSV);
      alert("已成功將兩份 10 月 Excel 範例資料分別載入左右兩張表格！");
    });
  }

  if (btnResetAdminDefaults) {
    btnResetAdminDefaults.addEventListener("click", () => {
      if (confirm("確定要清空左右兩張表格嗎？")) {
        clearTable(tablePastedTargets, promptTargets, badgeTargetCount);
        clearTable(tablePastedShipments, promptShipments, badgeShipmentCount);
        localStorage.removeItem("shihlin_pasted_targets_raw");
        localStorage.removeItem("shihlin_pasted_shipments_raw");
      }
    });
  }

  if (btnConfirmTablesReady) {
    btnConfirmTablesReady.addEventListener("click", () => {
      alert("左右兩張表格已順利接收 Excel 資料！\n您可以檢視每一行、每一欄是否與 Excel 原始資料完全對齊。");
    });
  }

  // 載入先前記住的已貼入表格
  const savedTargetsRaw = localStorage.getItem("shihlin_pasted_targets_raw");
  if (savedTargetsRaw) {
    renderPastedTable(savedTargetsRaw, tablePastedTargets, promptTargets, badgeTargetCount);
  }
  const savedShipmentsRaw = localStorage.getItem("shihlin_pasted_shipments_raw");
  if (savedShipmentsRaw) {
    renderPastedTable(savedShipmentsRaw, tablePastedShipments, promptShipments, badgeShipmentCount);
  }

  // ============================================================================
  // 8. 客戶主檔載入與單一輸入自動帶入規範 (免重複打字)
  // ============================================================================
  const selectCustomer = document.getElementById("selectCustomer");
  const inputTaxId = document.getElementById("inputTaxId");
  const inputPhone = document.getElementById("inputPhone");
  const inputAddress = document.getElementById("inputAddress");
  const inputContactName = document.getElementById("inputContactName");
  const autofillHint = document.getElementById("customerAutofillHint");

  async function loadCustomers() {
    try {
      const resp = await fetch("/api/customers");
      if (!resp.ok) throw new Error("讀取客戶失敗");
      customersCache = await resp.json();

      if (selectCustomer) {
        selectCustomer.innerHTML = '<option value="">-- 請選擇拜訪客戶 (或打字過濾) --</option>';
        customersCache.forEach(cust => {
          const opt = document.createElement("option");
          opt.value = cust.id;
          opt.textContent = `${cust.name} (${cust.tax_id || "無統編"})`;
          selectCustomer.appendChild(opt);
        });
      }
    } catch (err) {
      console.warn("載入客戶主檔異常:", err);
    }
  }

  if (selectCustomer) {
    selectCustomer.addEventListener("change", () => {
      const custId = parseInt(selectCustomer.value, 10);
      const selected = customersCache.find(c => c.id === custId);

      if (selected) {
        if (inputTaxId) inputTaxId.value = selected.tax_id || "";
        if (inputPhone) inputPhone.value = selected.phone || "";
        if (inputAddress) inputAddress.value = selected.address || "";
        if (inputContactName && !inputContactName.value) {
          inputContactName.value = selected.default_contact || "";
        }
        if (autofillHint) autofillHint.classList.remove("hidden");
      } else {
        if (inputTaxId) inputTaxId.value = "";
        if (inputPhone) inputPhone.value = "";
        if (inputAddress) inputAddress.value = "";
        if (autofillHint) autofillHint.classList.add("hidden");
      }
    });
  }

  // ============================================================================
  // 9. 日報清單與 Deal 案件看板載入
  // ============================================================================
  const dailyReportsTableBody = document.getElementById("dailyReportsTableBody");
  const dealsTableBody = document.getElementById("dealsTableBody");
  const btnRefreshReports = document.getElementById("btnRefreshReports");
  const btnResetForm = document.getElementById("btnResetForm");
  const dailyReportForm = document.getElementById("dailyReportForm");

  const checkCreateOpp = document.getElementById("checkCreateOpp");
  const bodyOppToggle = document.getElementById("bodyOppToggle");
  const checkCreateDeal = document.getElementById("checkCreateDeal");
  const bodyDealToggle = document.getElementById("bodyDealToggle");

  if (checkCreateOpp && bodyOppToggle) {
    checkCreateOpp.addEventListener("change", (e) => {
      if (e.target.checked) bodyOppToggle.classList.add("expanded");
      else bodyOppToggle.classList.remove("expanded");
    });
  }

  if (checkCreateDeal && bodyDealToggle) {
    checkCreateDeal.addEventListener("change", (e) => {
      if (e.target.checked) {
        bodyDealToggle.classList.add("expanded");
        const futureDate = new Date();
        futureDate.setDate(futureDate.getDate() + 30);
        const estDealDateEl = document.getElementById("inputEstDealDate");
        if (estDealDateEl) estDealDateEl.value = futureDate.toISOString().split("T")[0];
      } else {
        bodyDealToggle.classList.remove("expanded");
      }
    });
  }

  async function loadDailyReports() {
    if (!dailyReportsTableBody) return;
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

  async function loadDeals() {
    if (!dealsTableBody) return;
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

  if (dailyReportForm) {
    dailyReportForm.addEventListener("submit", async (e) => {
      e.preventDefault();

      const selectedCust = customersCache.find(c => c.id === parseInt(selectCustomer.value, 10));
      const customerName = selectedCust ? selectedCust.name : (selectCustomer.options[selectCustomer.selectedIndex]?.text || "");

      const payload = {
        report_date: document.getElementById("inputReportDate")?.value,
        sales_name: document.getElementById("inputSalesName")?.value.trim(),
        customer_id: selectedCust ? selectedCust.id : null,
        customer_name: customerName,
        contact_name: inputContactName ? inputContactName.value.trim() : "",
        visit_type: document.getElementById("selectVisitType")?.value,
        content_summary: document.getElementById("textContentSummary")?.value.trim(),
        
        create_opportunity: checkCreateOpp ? checkCreateOpp.checked : false,
        industry_type: document.getElementById("inputIndustry")?.value,
        product_series: document.getElementById("inputProductSeries")?.value,
        est_quarter: document.getElementById("inputEstQuarter")?.value,
        est_scale: parseFloat(document.getElementById("inputEstScale")?.value) || 0,

        create_deal: checkCreateDeal ? checkCreateDeal.checked : false,
        project_name: document.getElementById("inputProjectName")?.value.trim(),
        comp_brand: document.getElementById("inputCompBrand")?.value,
        comp_channel: document.getElementById("inputCompChannel")?.value.trim(),
        est_deal_date: document.getElementById("inputEstDealDate")?.value,
        est_amount: parseFloat(document.getElementById("inputEstDealAmount")?.value) || 0
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
        
        const summaryEl = document.getElementById("textContentSummary");
        if (summaryEl) summaryEl.value = "";
        if (checkCreateOpp) checkCreateOpp.checked = false;
        if (bodyOppToggle) bodyOppToggle.classList.remove("expanded");
        if (checkCreateDeal) checkCreateDeal.checked = false;
        if (bodyDealToggle) bodyDealToggle.classList.remove("expanded");

        loadDailyReports();
        loadDeals();
      } catch (err) {
        alert(`儲存失敗: ${err.message}`);
      }
    });
  }

  if (btnResetForm && dailyReportForm) {
    btnResetForm.addEventListener("click", () => {
      dailyReportForm.reset();
      const todayStr = new Date().toISOString().split("T")[0];
      const reportDateEl = document.getElementById("inputReportDate");
      if (reportDateEl) reportDateEl.value = todayStr;
      const curSales = localStorage.getItem("shihlin_pwa_active_sales") || "曾維崧";
      if (inputSalesName) inputSalesName.value = curSales;
      if (inputTaxId) inputTaxId.value = "";
      if (inputPhone) inputPhone.value = "";
      if (inputAddress) inputAddress.value = "";
      if (autofillHint) autofillHint.classList.add("hidden");
      if (bodyOppToggle) bodyOppToggle.classList.remove("expanded");
      if (bodyDealToggle) bodyDealToggle.classList.remove("expanded");
    });
  }

  if (btnRefreshReports) {
    btnRefreshReports.addEventListener("click", () => {
      loadDailyReports();
      loadDeals();
    });
  }

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

  // 初始載入設定與數據
  initSettings();
  renderAdminExcelGrid();
  renderHomeDashboard();

  // 初始載入客戶主檔與後端日報
  loadCustomers();
  loadDailyReports();
  loadDeals();
});
