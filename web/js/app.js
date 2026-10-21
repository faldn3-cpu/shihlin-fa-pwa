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
  // ============================================================================
  // 7. 管理員目標資料維護（微軟原生 Excel 雙試算表格子）
  // 支援 RFC 4180 試算表標準（精準處理雙引號內 Alt+Enter 換行，還原 Synology 表格）
  // 支援右鍵快顯選單（插入/刪除列）、DEL 鍵快速清空與水平捲軸自適應歸零
  // ============================================================================
  const containerTableTargets = document.getElementById("containerTableTargets");
  const tablePastedTargets = document.getElementById("tablePastedTargets");
  const badgeTargetCount = document.getElementById("badgeTargetCount");
  const btnPasteTargetDemo = document.getElementById("btnPasteTargetDemo");
  const btnClearTargetTable = document.getElementById("btnClearTargetTable");

  const containerTableShipments = document.getElementById("containerTableShipments");
  const tablePastedShipments = document.getElementById("tablePastedShipments");
  const badgeShipmentCount = document.getElementById("badgeShipmentCount");
  const btnPasteShipmentDemo = document.getElementById("btnPasteShipmentDemo");
  const btnClearShipmentTable = document.getElementById("btnClearShipmentTable");

  const btnLoadSampleExcel = document.getElementById("btnLoadSampleExcel");
  const btnResetAdminDefaults = document.getElementById("btnResetAdminDefaults");
  const btnConfirmTablesReady = document.getElementById("btnConfirmTablesReady");

  // 右鍵快顯功能選單元件
  const excelContextMenu = document.getElementById("excelContextMenu");
  const menuInsertRowAbove = document.getElementById("menuInsertRowAbove");
  const menuInsertRowBelow = document.getElementById("menuInsertRowBelow");
  const menuDeleteRow = document.getElementById("menuDeleteRow");
  const menuInsertColLeft = document.getElementById("menuInsertColLeft");
  const menuInsertColRight = document.getElementById("menuInsertColRight");
  const menuDeleteCol = document.getElementById("menuDeleteCol");
  const menuAutoFitCol = document.getElementById("menuAutoFitCol");
  const menuAutoFitAllCols = document.getElementById("menuAutoFitAllCols");
  const menuClearCell = document.getElementById("menuClearCell");

  let activeContextTable = null;
  let activeContextRow = -1;
  let activeContextCol = -1;
  let activeContextCell = null;
  let activeSelectedCol = -1;

  // 滑鼠多格框選狀態
  let isRangeSelecting = false;
  let selectionStart = null; // { table, r, c }
  let activeRange = null;    // { table, minR, maxR, minC, maxC }

  // 產生 Excel 欄標籤 A, B, C... Z, AA, AB...
  function getExcelColName(idx) {
    let name = "";
    let i = idx;
    while (i >= 0) {
      name = String.fromCharCode((i % 26) + 65) + name;
      i = Math.floor(i / 26) - 1;
    }
    return name;
  }

  // 100% 格式對照：解析 Excel / Synology Office 剪貼簿產生的 HTML 表格樣式 (含顏色、字型大小、底色、粗體)
  function parseExcelHTML(htmlText) {
    if (!htmlText || !htmlText.includes("<table")) return null;
    try {
      const parser = new DOMParser();
      const doc = parser.parseFromString(htmlText, "text/html");
      const table = doc.querySelector("table");
      if (!table) return null;

      // 提取 <style> 中的 class 樣式規則
      const classStyles = {};
      const styleTags = doc.querySelectorAll("style");
      styleTags.forEach(styleTag => {
        const cssText = styleTag.textContent || "";
        const ruleRegex = /\.([a-zA-Z0-9_-]+)\s*\{([^}]+)\}/g;
        let match;
        while ((match = ruleRegex.exec(cssText)) !== null) {
          const clsName = match[1];
          const declarations = match[2].split(";");
          if (!classStyles[clsName]) classStyles[clsName] = {};
          declarations.forEach(dec => {
            const parts = dec.split(":");
            if (parts.length >= 2) {
              const prop = parts[0].trim().toLowerCase();
              const val = parts.slice(1).join(":").trim();
              classStyles[clsName][prop] = val;
            }
          });
        }
      });

      const rows = [];
      const trList = table.querySelectorAll("tr");
      trList.forEach(tr => {
        const rowData = [];
        const cellList = tr.querySelectorAll("td, th");
        cellList.forEach(cell => {
          const text = (cell.innerText || "").trim();
          let bg = cell.style.backgroundColor || cell.getAttribute("bgcolor") || "";
          let color = cell.style.color || "";
          let bold = cell.style.fontWeight === "bold" || parseInt(cell.style.fontWeight, 10) >= 600 || !!cell.querySelector("b, strong");
          let fontSize = cell.style.fontSize || "";
          let align = cell.style.textAlign || cell.getAttribute("align") || "";

          // 讀取 class 樣式
          cell.classList.forEach(cls => {
            if (classStyles[cls]) {
              const s = classStyles[cls];
              if (!bg && (s["background-color"] || s["background"])) bg = s["background-color"] || s["background"];
              if (!color && s["color"]) color = s["color"];
              if (!bold && (s["font-weight"] === "bold" || parseInt(s["font-weight"], 10) >= 600)) bold = true;
              if (!fontSize && s["font-size"]) fontSize = s["font-size"];
              if (!align && s["text-align"]) align = s["text-align"];
            }
          });

          // 讀取子標籤 (<font>, <span>)
          const fontEl = cell.querySelector("font");
          if (fontEl && fontEl.color && !color) color = fontEl.color;
          const spanEl = cell.querySelector("span");
          if (spanEl) {
            if (spanEl.style.color && !color) color = spanEl.style.color;
            if (spanEl.style.backgroundColor && !bg) bg = spanEl.style.backgroundColor;
            if (spanEl.style.fontSize && !fontSize) fontSize = spanEl.style.fontSize;
          }

          if (bg || color || bold || fontSize || align) {
            rowData.push({
              val: text,
              bg: bg,
              color: color,
              bold: bold ? "bold" : "",
              fontSize: fontSize,
              align: align
            });
          } else {
            rowData.push(text);
          }
        });
        if (rowData.length > 0) {
          rows.push(rowData);
        }
      });

      // 移除尾端空白列
      while (rows.length > 0 && rows[rows.length - 1].every(c => (typeof c === "object" ? !c.val : !c))) {
        rows.pop();
      }

      return rows.length > 0 ? rows : null;
    } catch (err) {
      console.warn("[Excel Paste] HTML 解析失敗，將回退為 TSV 純文字：", err);
      return null;
    }
  }

  // 支援 RFC 4180 試算表規範：能正確辨識雙引號內的 Alt+Enter 換行，防止儲存格內換行碎裂跑版
  function parseExcelTSV(text) {
    if (!text || !text.trim()) return [];
    const rows = [];
    let currentRow = [];
    let currentCell = "";
    let inQuotes = false;
    const len = text.length;

    for (let i = 0; i < len; i++) {
      const char = text[i];
      const nextChar = i + 1 < len ? text[i + 1] : "";

      if (char === '"') {
        if (inQuotes && nextChar === '"') {
          currentCell += '"';
          i++; // 跳過連續雙引號轉義
        } else {
          inQuotes = !inQuotes;
        }
      } else if (char === '\t' && !inQuotes) {
        currentRow.push(currentCell.trim());
        currentCell = "";
      } else if ((char === '\r' || char === '\n') && !inQuotes) {
        if (char === '\r' && nextChar === '\n') {
          i++; // 跳過 \r\n
        }
        currentRow.push(currentCell.trim());
        rows.push(currentRow);
        currentRow = [];
        currentCell = "";
      } else {
        currentCell += char;
      }
    }

    if (currentCell !== "" || currentRow.length > 0) {
      currentRow.push(currentCell.trim());
      rows.push(currentRow);
    }

    while (rows.length > 0 && rows[rows.length - 1].every(cell => !cell)) {
      rows.pop();
    }

    return rows;
  }

  // 清除多格連續框選
  function clearRangeSelection() {
    if (activeRange && activeRange.table) {
      const prevCells = activeRange.table.querySelectorAll("td.cell-selected-range");
      prevCells.forEach(c => c.classList.remove("cell-selected-range"));
    }
    activeRange = null;
  }

  // 更新多格連續框選高亮區域
  function updateRangeSelection(tableEl, r1, c1, r2, c2) {
    clearRangeSelection();
    activeRange = {
      table: tableEl,
      minR: Math.min(r1, r2),
      maxR: Math.max(r1, r2),
      minC: Math.min(c1, c2),
      maxC: Math.max(c1, c2)
    };
    for (let r = activeRange.minR; r <= activeRange.maxR; r++) {
      for (let c = activeRange.minC; c <= activeRange.maxC; c++) {
        const cell = tableEl.querySelector(`td.grid-cell[data-row="${r}"][data-col="${c}"]`);
        if (cell) cell.classList.add("cell-selected-range");
      }
    }
  }

  // 清除整列與整欄選取
  function clearRowColSelection(tableEl) {
    if (!tableEl) return;
    tableEl.querySelectorAll("tr.row-selected").forEach(tr => tr.classList.remove("row-selected"));
    tableEl.querySelectorAll("th.col-selected").forEach(th => th.classList.remove("col-selected"));
    tableEl.querySelectorAll("td.col-selected").forEach(td => td.classList.remove("col-selected"));
    activeSelectedCol = -1;
  }

  // 渲染微軟原生 Excel 試算表格子 (含 100% 樣式支援、A/B/C 欄標籤與 1/2/3 列號，支援拖曳框選、調整欄寬與排序)
  function renderExcelGrid(tableEl, inputData, minRows, minCols, badgeEl, preserveScroll = false) {
    if (!tableEl) return;

    let dataMatrix = [];
    if (typeof inputData === "string") {
      if (inputData.trim().startsWith("[")) {
        try {
          dataMatrix = JSON.parse(inputData);
        } catch (e) {
          dataMatrix = parseExcelHTML(inputData) || parseExcelTSV(inputData);
        }
      } else {
        dataMatrix = parseExcelHTML(inputData) || parseExcelTSV(inputData);
      }
    } else if (Array.isArray(inputData)) {
      dataMatrix = inputData;
    }

    tableEl._gridMatrix = dataMatrix;
    tableEl._minRows = minRows || 20;
    tableEl._minCols = minCols || 10;
    tableEl._badgeEl = badgeEl;

    let maxCols = tableEl._minCols;
    dataMatrix.forEach(row => {
      if (row.length > maxCols) maxCols = row.length;
    });

    const totalRows = Math.max(tableEl._minRows, dataMatrix.length);
    const totalCols = maxCols;

    // 建立 thead (頂部 A, B, C... 欄標籤，含拖曳欄寬與雙擊自動適配把手)
    let theadHtml = `<thead><tr><th class="corner-cell" title="點擊列號或儲存格可按右鍵操作選單，按 DEL 清空"></th>`;
    for (let c = 0; c < totalCols; c++) {
      theadHtml += `
        <th class="col-header" data-col-index="${c}">
          <div class="col-header-content">
            <span class="col-title">${getExcelColName(c)}</span>
          </div>
          <div class="col-resizer" data-col-index="${c}" title="拖曳調整欄寬，或雙擊邊界自動最佳化欄寬"></div>
        </th>`;
    }
    theadHtml += "</tr></thead>";

    // 建立 tbody (左側 1, 2, 3... 列號與真實儲存格，完整還原字型、大小、底色與顏色，自動列高撐開)
    let tbodyHtml = "<tbody>";
    for (let r = 0; r < totalRows; r++) {
      tbodyHtml += `<tr data-row-index="${r}"><th class="row-header" title="點擊選取第 ${r + 1} 列 (可點右鍵選單)">${r + 1}</th>`;
      const rowData = dataMatrix[r] || [];
      for (let c = 0; c < totalCols; c++) {
        const cellData = rowData[c];
        let val = "";
        const styleParts = [];

        if (cellData && typeof cellData === "object") {
          val = cellData.val !== undefined ? cellData.val : "";
          if (cellData.bg && cellData.bg !== "transparent" && cellData.bg !== "inherit") {
            styleParts.push(`background-color:${cellData.bg} !important`);
          }
          if (cellData.color && cellData.color !== "inherit") {
            styleParts.push(`color:${cellData.color} !important`);
          }
          if (cellData.bold) {
            styleParts.push(`font-weight:${cellData.bold}`);
          }
          if (cellData.fontSize) {
            styleParts.push(`font-size:${cellData.fontSize}`);
          }
          if (cellData.align) {
            styleParts.push(`text-align:${cellData.align}`);
          }
        } else {
          val = cellData !== undefined ? cellData : "";
        }

        // 若來源未特別指定對齊，依據數值與字串自動靠右/靠左
        if (!styleParts.some(p => p.startsWith("text-align"))) {
          const isNum = val && !val.includes("\n") && !isNaN(parseFloat(val.replace(/,/g, "").replace(/%/g, ""))) && !val.includes("月") && !val.includes("區");
          styleParts.push(isNum ? "text-align:right;" : "text-align:left;");
        }

        const styleAttr = styleParts.length > 0 ? `style="${styleParts.join(";")}"` : "";
        tbodyHtml += `<td class="grid-cell" contenteditable="true" ${styleAttr} data-row="${r}" data-col="${c}">${escapeHtml(val)}</td>`;
      }
      tbodyHtml += "</tr>";
    }
    tbodyHtml += "</tbody>";

    tableEl.innerHTML = theadHtml + tbodyHtml;

    if (!preserveScroll && tableEl.parentElement) {
      tableEl.parentElement.scrollLeft = 0;
    }

    // 綁定儲存格事件、拖曳選取與欄寬調整
    bindExcelGridEvents(tableEl);

    // 若有資料，自動為所有欄位計算最佳欄寬 (Auto-fit Columns)
    if (dataMatrix.length > 0) {
      setTimeout(() => autoFitAllColWidths(tableEl), 10);
    }

    if (badgeEl) {
      if (dataMatrix.length > 0) {
        badgeEl.innerText = `${dataMatrix.length} 列 × ${maxCols} 欄 (已填入)`;
        badgeEl.classList.add("active");
      } else {
        badgeEl.innerText = `空白試算表 (${totalRows} 列 × ${totalCols} 欄)`;
        badgeEl.classList.remove("active");
      }
    }
  }

  // 選取特定列高亮
  function selectGridRow(tableEl, rowIdx) {
    clearRowColSelection(tableEl);
    clearRangeSelection();
    tableEl._selectedRow = rowIdx;
    activeContextTable = tableEl;
    activeContextRow = rowIdx;
    activeContextCol = -1;
    const tr = tableEl.querySelector(`tbody tr[data-row-index="${rowIdx}"]`);
    if (tr) tr.classList.add("row-selected");
  }

  // 選取特定欄高亮
  function selectGridCol(tableEl, colIdx) {
    clearRowColSelection(tableEl);
    clearRangeSelection();
    activeSelectedCol = colIdx;
    activeContextTable = tableEl;
    activeContextCol = colIdx;
    activeContextRow = -1;
    const th = tableEl.querySelector(`th.col-header[data-col-index="${colIdx}"]`);
    if (th) th.classList.add("col-selected");
    tableEl.querySelectorAll(`td.grid-cell[data-col="${colIdx}"]`).forEach(td => td.classList.add("col-selected"));
  }

  // 自動計算並調整單一欄位的最佳寬度 (Auto-fit Column Width)
  function autoFitColWidth(tableEl, colIdx) {
    if (!tableEl) return;
    const th = tableEl.querySelector(`th.col-header[data-col-index="${colIdx}"]`);
    if (!th) return;

    let maxTextLen = 3; // 至少能容納欄標籤 A, B, C
    const cells = tableEl.querySelectorAll(`td.grid-cell[data-col="${colIdx}"]`);
    cells.forEach(td => {
      const text = td.innerText || "";
      const lines = text.split("\n");
      lines.forEach(line => {
        let len = 0;
        for (let i = 0; i < line.length; i++) {
          len += line.charCodeAt(i) > 255 ? 2 : 1; // 中文字或全形佔 2 個字元寬度
        }
        if (len > maxTextLen) maxTextLen = len;
      });
    });

    const optimalWidth = Math.max(55, Math.min(420, Math.round(maxTextLen * 8.2 + 22)));
    th.style.width = `${optimalWidth}px`;
    th.style.minWidth = `${optimalWidth}px`;
    th.style.maxWidth = `${optimalWidth}px`;
  }

  // 自動調整整張表格的所有欄寬 (Auto-fit All Column Widths)
  function autoFitAllColWidths(tableEl) {
    if (!tableEl) return;
    const colHeaders = tableEl.querySelectorAll("th.col-header");
    colHeaders.forEach(th => {
      const colIdx = parseInt(th.dataset.colIndex, 10);
      if (!isNaN(colIdx)) {
        autoFitColWidth(tableEl, colIdx);
      }
    });
  }

  // 開啟微軟風格右鍵快顯功能選單 (支援列、欄與自動適配寬度)
  function openContextMenu(e, tableEl, rowIdx, colIdx, cellEl) {
    if (!excelContextMenu) return;
    activeContextTable = tableEl;
    activeContextRow = rowIdx;
    activeContextCol = colIdx;
    activeContextCell = cellEl;

    let x = e.clientX;
    let y = e.clientY;
    const menuWidth = 220;
    const menuHeight = 280;

    if (x + menuWidth > window.innerWidth) x = window.innerWidth - menuWidth - 10;
    if (y + menuHeight > window.innerHeight) y = window.innerHeight - menuHeight - 10;

    excelContextMenu.style.left = `${x}px`;
    excelContextMenu.style.top = `${y}px`;
    excelContextMenu.classList.remove("hidden");
  }

  // 綁定儲存格與列號互動事件 (含拖曳框選、欄寬調整、雙擊自動適配欄寬、Tab/Enter 導航與右鍵選單)
  function bindExcelGridEvents(tableEl) {
    // 1. 列號點選與右鍵
    const rowHeaders = tableEl.querySelectorAll("th.row-header");
    rowHeaders.forEach(th => {
      th.addEventListener("click", () => {
        const rowIdx = parseInt(th.parentElement.dataset.rowIndex, 10);
        selectGridRow(tableEl, rowIdx);
      });
      th.addEventListener("contextmenu", (e) => {
        e.preventDefault();
        const rowIdx = parseInt(th.parentElement.dataset.rowIndex, 10);
        selectGridRow(tableEl, rowIdx);
        openContextMenu(e, tableEl, rowIdx, -1, null);
      });
    });

    // 2. 欄標籤點選、拖曳欄寬與雙擊自動適配
    const colHeaders = tableEl.querySelectorAll("th.col-header");
    colHeaders.forEach(th => {
      const colIdx = parseInt(th.dataset.colIndex, 10);
      th.addEventListener("click", (e) => {
        if (e.target.classList.contains("col-resizer")) return;
        selectGridCol(tableEl, colIdx);
      });

      th.addEventListener("contextmenu", (e) => {
        e.preventDefault();
        selectGridCol(tableEl, colIdx);
        openContextMenu(e, tableEl, -1, colIdx, null);
      });

      // 欄寬拖曳與雙擊把手
      const resizer = th.querySelector(".col-resizer");
      if (resizer) {
        // 雙擊邊界：自動最佳化該欄欄寬 (就像 Excel 點兩下邊線)
        resizer.addEventListener("dblclick", (e) => {
          e.preventDefault();
          e.stopPropagation();
          autoFitColWidth(tableEl, colIdx);
        });

        // 按住拖曳：手動調整欄寬
        resizer.addEventListener("mousedown", (e) => {
          e.preventDefault();
          e.stopPropagation();
          const startX = e.pageX;
          const startWidth = th.offsetWidth;
          resizer.classList.add("resizing");

          function onMouseMove(ev) {
            const diff = ev.pageX - startX;
            const newWidth = Math.max(45, startWidth + diff);
            th.style.width = `${newWidth}px`;
            th.style.minWidth = `${newWidth}px`;
            th.style.maxWidth = `${newWidth}px`;
          }

          function onMouseUp() {
            resizer.classList.remove("resizing");
            document.removeEventListener("mousemove", onMouseMove);
            document.removeEventListener("mouseup", onMouseUp);
          }

          document.addEventListener("mousemove", onMouseMove);
          document.addEventListener("mouseup", onMouseUp);
        });
      }
    });

    // 3. 儲存格事件：多格拖曳框選、鍵盤導航 (Tab, Enter) 與編輯監聽
    const cells = tableEl.querySelectorAll("td.grid-cell");
    cells.forEach(td => {
      const r = parseInt(td.dataset.row, 10);
      const c = parseInt(td.dataset.col, 10);

      // 滑鼠按下：開始多格框選
      td.addEventListener("mousedown", (e) => {
        if (e.button === 0) { // 滑鼠左鍵
          isRangeSelecting = true;
          clearRowColSelection(tableEl);
          selectionStart = { table: tableEl, r: r, c: c };
          updateRangeSelection(tableEl, r, c, r, c);
          activeContextTable = tableEl;
          activeContextRow = r;
          activeContextCol = c;
          activeContextCell = td;
        }
      });

      // 滑鼠移入：擴展框選區域
      td.addEventListener("mouseenter", () => {
        if (isRangeSelecting && selectionStart && selectionStart.table === tableEl) {
          updateRangeSelection(tableEl, selectionStart.r, selectionStart.c, r, c);
        }
      });

      // 儲存格焦點
      td.addEventListener("focus", () => {
        activeContextTable = tableEl;
        activeContextRow = r;
        activeContextCol = c;
        activeContextCell = td;
      });

      // 右鍵選單
      td.addEventListener("contextmenu", (e) => {
        e.preventDefault();
        activeContextTable = tableEl;
        activeContextRow = r;
        activeContextCol = c;
        activeContextCell = td;
        td.focus();
        openContextMenu(e, tableEl, r, c, td);
      });

      // 儲存格內容變更即時寫回矩陣
      td.addEventListener("blur", () => {
        if (tableEl._gridMatrix) {
          while (tableEl._gridMatrix.length <= r) {
            tableEl._gridMatrix.push([]);
          }
          const existing = tableEl._gridMatrix[r][c];
          if (existing && typeof existing === "object") {
            existing.val = td.innerText;
          } else {
            tableEl._gridMatrix[r][c] = td.innerText;
          }
        }
      });

      // 鍵盤導航：Tab 換欄、Enter 換列
      td.addEventListener("keydown", (e) => {
        if (e.key === "Tab") {
          e.preventDefault();
          const nextCol = e.shiftKey ? c - 1 : c + 1;
          const targetCell = tableEl.querySelector(`td.grid-cell[data-row="${r}"][data-col="${nextCol}"]`);
          if (targetCell) targetCell.focus();
        } else if (e.key === "Enter" && !e.shiftKey && !e.altKey) {
          e.preventDefault();
          const nextRow = r + 1;
          const targetCell = tableEl.querySelector(`td.grid-cell[data-row="${nextRow}"][data-col="${c}"]`);
          if (targetCell) targetCell.focus();
        }
      });
    });
  }

  // 全域放開滑鼠時結束框選
  document.addEventListener("mouseup", () => {
    isRangeSelecting = false;
  });

  // 點擊選單外任一處或按 Esc 自動隱藏右鍵選單
  document.addEventListener("click", (e) => {
    if (excelContextMenu && !excelContextMenu.contains(e.target)) {
      excelContextMenu.classList.add("hidden");
    }
  });

  document.addEventListener("keydown", (e) => {
    if (e.key === "Escape" && excelContextMenu) {
      excelContextMenu.classList.add("hidden");
    }
  });

  // 右鍵選單：在上方插入列
  if (menuInsertRowAbove) {
    menuInsertRowAbove.addEventListener("click", () => {
      if (!activeContextTable) return;
      const matrix = activeContextTable._gridMatrix || [];
      const numCols = matrix[0] ? matrix[0].length : (activeContextTable._minCols || 10);
      const insertIdx = activeContextRow >= 0 ? activeContextRow : 0;
      matrix.splice(insertIdx, 0, new Array(numCols).fill(""));
      renderExcelGrid(activeContextTable, matrix, activeContextTable._minRows, activeContextTable._minCols, activeContextTable._badgeEl, true);
      selectGridRow(activeContextTable, insertIdx);
      if (excelContextMenu) excelContextMenu.classList.add("hidden");
    });
  }

  // 右鍵選單：在下方插入列
  if (menuInsertRowBelow) {
    menuInsertRowBelow.addEventListener("click", () => {
      if (!activeContextTable) return;
      const matrix = activeContextTable._gridMatrix || [];
      const numCols = matrix[0] ? matrix[0].length : (activeContextTable._minCols || 10);
      const insertIdx = (activeContextRow >= 0 ? activeContextRow : matrix.length - 1) + 1;
      matrix.splice(insertIdx, 0, new Array(numCols).fill(""));
      renderExcelGrid(activeContextTable, matrix, activeContextTable._minRows, activeContextTable._minCols, activeContextTable._badgeEl, true);
      selectGridRow(activeContextTable, insertIdx);
      if (excelContextMenu) excelContextMenu.classList.add("hidden");
    });
  }

  // 右鍵選單：刪除此列
  if (menuDeleteRow) {
    menuDeleteRow.addEventListener("click", () => {
      if (!activeContextTable || !activeContextTable._gridMatrix) return;
      const matrix = activeContextTable._gridMatrix;
      if (activeContextRow >= 0 && activeContextRow < matrix.length) {
        matrix.splice(activeContextRow, 1);
        renderExcelGrid(activeContextTable, matrix, activeContextTable._minRows, activeContextTable._minCols, activeContextTable._badgeEl, true);
        if (matrix.length > 0) {
          const nextIdx = Math.min(activeContextRow, matrix.length - 1);
          selectGridRow(activeContextTable, nextIdx);
        }
      }
      if (excelContextMenu) excelContextMenu.classList.add("hidden");
    });
  }

  // 右鍵選單：在左方插入一欄
  if (menuInsertColLeft) {
    menuInsertColLeft.addEventListener("click", () => {
      if (!activeContextTable) return;
      const matrix = activeContextTable._gridMatrix || [];
      const insertIdx = activeContextCol >= 0 ? activeContextCol : 0;
      matrix.forEach(row => row.splice(insertIdx, 0, ""));
      activeContextTable._minCols = (activeContextTable._minCols || 10) + 1;
      renderExcelGrid(activeContextTable, matrix, activeContextTable._minRows, activeContextTable._minCols, activeContextTable._badgeEl, true);
      selectGridCol(activeContextTable, insertIdx);
      if (excelContextMenu) excelContextMenu.classList.add("hidden");
    });
  }

  // 右鍵選單：在右方插入一欄
  if (menuInsertColRight) {
    menuInsertColRight.addEventListener("click", () => {
      if (!activeContextTable) return;
      const matrix = activeContextTable._gridMatrix || [];
      const numCols = matrix[0] ? matrix[0].length : (activeContextTable._minCols || 10);
      const insertIdx = (activeContextCol >= 0 ? activeContextCol : numCols - 1) + 1;
      matrix.forEach(row => row.splice(insertIdx, 0, ""));
      activeContextTable._minCols = (activeContextTable._minCols || 10) + 1;
      renderExcelGrid(activeContextTable, matrix, activeContextTable._minRows, activeContextTable._minCols, activeContextTable._badgeEl, true);
      selectGridCol(activeContextTable, insertIdx);
      if (excelContextMenu) excelContextMenu.classList.add("hidden");
    });
  }

  // 右鍵選單：刪除此欄
  if (menuDeleteCol) {
    menuDeleteCol.addEventListener("click", () => {
      if (!activeContextTable || !activeContextTable._gridMatrix) return;
      const matrix = activeContextTable._gridMatrix;
      if (activeContextCol >= 0) {
        matrix.forEach(row => {
          if (row.length > activeContextCol) row.splice(activeContextCol, 1);
        });
        if (activeContextTable._minCols > 3) activeContextTable._minCols--;
        renderExcelGrid(activeContextTable, matrix, activeContextTable._minRows, activeContextTable._minCols, activeContextTable._badgeEl, true);
        selectGridCol(activeContextTable, Math.max(0, activeContextCol - 1));
      }
      if (excelContextMenu) excelContextMenu.classList.add("hidden");
    });
  }

  // 右鍵選單：自動調整此欄寬度
  if (menuAutoFitCol) {
    menuAutoFitCol.addEventListener("click", () => {
      if (activeContextTable && activeContextCol >= 0) {
        autoFitColWidth(activeContextTable, activeContextCol);
      }
      if (excelContextMenu) excelContextMenu.classList.add("hidden");
    });
  }

  // 右鍵選單：自動調整全部欄寬
  if (menuAutoFitAllCols) {
    menuAutoFitAllCols.addEventListener("click", () => {
      if (activeContextTable) {
        autoFitAllColWidths(activeContextTable);
      }
      if (excelContextMenu) excelContextMenu.classList.add("hidden");
    });
  }

  // 右鍵選單：清空內容 (或鍵盤 DEL)
  if (menuClearCell) {
    menuClearCell.addEventListener("click", () => {
      clearSelectedCellOrRowOrRange();
      if (excelContextMenu) excelContextMenu.classList.add("hidden");
    });
  }

  // 清空選取單元格、整列、整欄或框選區域共用函式
  function clearSelectedCellOrRowOrRange() {
    // 1. 若有多格連續框選區域
    if (activeRange && activeRange.table) {
      const tbl = activeRange.table;
      for (let r = activeRange.minR; r <= activeRange.maxR; r++) {
        for (let c = activeRange.minC; c <= activeRange.maxC; c++) {
          if (tbl._gridMatrix && tbl._gridMatrix[r]) {
            const item = tbl._gridMatrix[r][c];
            if (item && typeof item === "object") item.val = "";
            else tbl._gridMatrix[r][c] = "";
          }
          const cell = tbl.querySelector(`td.grid-cell[data-row="${r}"][data-col="${c}"]`);
          if (cell) cell.innerText = "";
        }
      }
      return;
    }

    // 2. 若選取了整欄
    if (activeSelectedCol >= 0 && activeContextTable) {
      const col = activeSelectedCol;
      const matrix = activeContextTable._gridMatrix || [];
      matrix.forEach(row => {
        if (row[col] !== undefined) {
          if (row[col] && typeof row[col] === "object") row[col].val = "";
          else row[col] = "";
        }
      });
      activeContextTable.querySelectorAll(`td.grid-cell[data-col="${col}"]`).forEach(td => td.innerText = "");
      return;
    }

    // 3. 若選取了單一儲存格
    if (activeContextCell) {
      activeContextCell.innerText = "";
      const r = parseInt(activeContextCell.dataset.row, 10);
      const c = parseInt(activeContextCell.dataset.col, 10);
      if (activeContextTable && activeContextTable._gridMatrix && activeContextTable._gridMatrix[r]) {
        const item = activeContextTable._gridMatrix[r][c];
        if (item && typeof item === "object") item.val = "";
        else activeContextTable._gridMatrix[r][c] = "";
      }
      return;
    }

    // 4. 若選取了整列
    if (activeContextTable && activeContextTable._gridMatrix && activeContextTable._selectedRow !== undefined) {
      const r = activeContextTable._selectedRow;
      if (activeContextTable._gridMatrix[r]) {
        activeContextTable._gridMatrix[r] = new Array(activeContextTable._gridMatrix[r].length).fill("");
        renderExcelGrid(activeContextTable, activeContextTable._gridMatrix, activeContextTable._minRows, activeContextTable._minCols, activeContextTable._badgeEl, true);
        selectGridRow(activeContextTable, r);
      }
    }
  }

  // 監聽鍵盤 DEL / Delete 鍵快速清空選取內容 (支援單格、多格框選、整欄、整列)
  document.addEventListener("keydown", (e) => {
    if (e.key === "Delete" || e.key === "Del") {
      clearSelectedCellOrRowOrRange();
    }
  });

  // 監聽左邊表格容器 paste 事件 (優先解析 HTML 還原 100% 格式，退回 TSV)
  if (containerTableTargets) {
    containerTableTargets.addEventListener("paste", (e) => {
      e.preventDefault();
      const clipboardData = e.clipboardData || window.clipboardData;
      const htmlText = clipboardData.getData("text/html");
      const plainText = clipboardData.getData("text");
      const matrix = parseExcelHTML(htmlText) || parseExcelTSV(plainText);
      if (matrix && matrix.length > 0) {
        renderExcelGrid(tablePastedTargets, matrix, 20, 11, badgeTargetCount, false);
        localStorage.setItem("shihlin_pasted_targets_raw", JSON.stringify(matrix));
      }
    });
  }

  // 監聽右邊表格容器 paste 事件 (優先解析 HTML 還原 100% 格式，退回 TSV)
  if (containerTableShipments) {
    containerTableShipments.addEventListener("paste", (e) => {
      e.preventDefault();
      const clipboardData = e.clipboardData || window.clipboardData;
      const htmlText = clipboardData.getData("text/html");
      const plainText = clipboardData.getData("text");
      const matrix = parseExcelHTML(htmlText) || parseExcelTSV(plainText);
      if (matrix && matrix.length > 0) {
        renderExcelGrid(tablePastedShipments, matrix, 25, 5, badgeShipmentCount, false);
        localStorage.setItem("shihlin_pasted_shipments_raw", JSON.stringify(matrix));
      }
    });
  }

  // 10 月真實目標總表範例 (包含 Alt+Enter 換行之真實格式)
  const sampleTargetsExcelTSV = `北區\t10\t月挑戰目標\t9,300\t萬
項次\t營業員\t年計月目標\t"10月\n挑戰目標"\t"10月\n堆疊業績"\t"10/21\n累積目標"\t"10/21\n實績"\t"堆疊目標\n差異(100%)"\t"挑戰目標\n差異"\t"挑戰目標\n達成率"\t"前一個月\n實績"
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
      renderExcelGrid(tablePastedTargets, sampleTargetsExcelTSV, 20, 11, badgeTargetCount);
      localStorage.setItem("shihlin_pasted_targets_raw", sampleTargetsExcelTSV);
    });
  }
  if (btnClearTargetTable) {
    btnClearTargetTable.addEventListener("click", () => {
      renderExcelGrid(tablePastedTargets, "", 20, 11, badgeTargetCount);
      localStorage.removeItem("shihlin_pasted_targets_raw");
    });
  }

  // 右表按鈕
  if (btnPasteShipmentDemo) {
    btnPasteShipmentDemo.addEventListener("click", () => {
      renderExcelGrid(tablePastedShipments, sampleShipmentsExcelTSV, 25, 5, badgeShipmentCount);
      localStorage.setItem("shihlin_pasted_shipments_raw", sampleShipmentsExcelTSV);
    });
  }
  if (btnClearShipmentTable) {
    btnClearShipmentTable.addEventListener("click", () => {
      renderExcelGrid(tablePastedShipments, "", 25, 5, badgeShipmentCount);
      localStorage.removeItem("shihlin_pasted_shipments_raw");
    });
  }

  // 頂部全部載入/清空按鈕
  if (btnLoadSampleExcel) {
    btnLoadSampleExcel.addEventListener("click", () => {
      renderExcelGrid(tablePastedTargets, sampleTargetsExcelTSV, 20, 11, badgeTargetCount);
      renderExcelGrid(tablePastedShipments, sampleShipmentsExcelTSV, 25, 5, badgeShipmentCount);
      localStorage.setItem("shihlin_pasted_targets_raw", sampleTargetsExcelTSV);
      localStorage.setItem("shihlin_pasted_shipments_raw", sampleShipmentsExcelTSV);
      alert("已成功將兩份 10 月 Excel 範例資料分別載入左右兩張試算表格子！");
    });
  }

  if (btnResetAdminDefaults) {
    btnResetAdminDefaults.addEventListener("click", () => {
      if (confirm("確定要重設左右兩張表格為初始空白 Excel 格子嗎？")) {
        renderExcelGrid(tablePastedTargets, "", 20, 11, badgeTargetCount);
        renderExcelGrid(tablePastedShipments, "", 25, 5, badgeShipmentCount);
        localStorage.removeItem("shihlin_pasted_targets_raw");
        localStorage.removeItem("shihlin_pasted_shipments_raw");
      }
    });
  }

  if (btnConfirmTablesReady) {
    btnConfirmTablesReady.addEventListener("click", () => {
      alert("左右兩張表格已順利呈現 Excel 格子！\n您可以檢視每一行、每一欄是否與 Excel 原始資料完全對齊，或點選儲存格直接編輯。");
    });
  }

  // 初始載入左右雙試算表格子 (若有先前儲存的資料則顯示資料，否則直接呈現空白 Excel 格子)
  const savedTargetsRaw = localStorage.getItem("shihlin_pasted_targets_raw");
  renderExcelGrid(tablePastedTargets, savedTargetsRaw || "", 20, 11, badgeTargetCount);

  const savedShipmentsRaw = localStorage.getItem("shihlin_pasted_shipments_raw");
  renderExcelGrid(tablePastedShipments, savedShipmentsRaw || "", 25, 5, badgeShipmentCount);

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
