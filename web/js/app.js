/**
 * 士林電機 FA 業務戰情與報價系統 - 前端主程式 (app.js)
 * 嚴格遵循：
 * 1. 100% 採用 Office / Windows Fluent 洗鍊線條樣式
 * 2. 徹底剔除彩色 Emoji 表情符號，採用洗鍊純文字狀態號誌 [達標] / [預警] / [落後]
 * 3. 頂部單行精簡化、即時時鐘 (MM-DD HH:mm)、4方格滑出式抽屜選單
 * 4. 戰情計分板升級：雙軌追蹤 (個人進度 + 全區目標)，支援 Excel 貼上自動解析
 * 5. 拜訪單一紀錄規範：選取客戶自動帶入統編、電話、地址與聯絡人
 * 6. 三層漏斗：日報 ➔ 商機 ➔ 案件 連鎖建立
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
  // 4. 設定頁面控制 (外觀顯示 / 字型 / 字級 / 業務切換)
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
    if (headerUserAvatar) {
      headerUserAvatar.innerText = salesName.charAt(0) || "曾";
    }
    if (currentSalesDisplay) {
      currentSalesDisplay.innerText = salesName;
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

    // 4. 業務身份切換
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

  initSettings();

  // ============================================================================
  // 5. 業務目標數據庫與試算模型 (對標附圖二與附圖三之真實業務數據)
  // ============================================================================
  const defaultTargetData = {
    monthTitle: "10 月營收挑戰目標",
    regionTotal: {
      target: 9300,
      actual: 10370
    },
    products: {
      LEB: { name: "LEB (士林伺服)", target: 516, actual: 196 },
      LNH: { name: "LNH (士林變頻器)", target: 516, actual: 243 },
      PLC: { name: "士林 PLC 系列", target: 210, actual: 246 }
    },
    salesTargets: {
      "曾維崧": { target: 810, actual: 802, leb: 28.2, lnh: 4.9, plc: 12.0 },
      "曾仁君": { target: 810, actual: 802, leb: 28.2, lnh: 4.9, plc: 12.0 },
      "張何達": { target: 900, actual: 1451, leb: 32.85, lnh: 49.62, plc: 0.64 },
      "葉仁豪": { target: 700, actual: 473, leb: 15.0, lnh: 18.2, plc: 8.5 },
      "邱文輝": { target: 1400, actual: 1388, leb: 3.65, lnh: 9.83, plc: 143.75 },
      "溫達仁": { target: 1200, actual: 1147, leb: 42.1, lnh: 35.0, plc: 20.3 },
      "楊家豪": { target: 700, actual: 721, leb: 18.5, lnh: 22.4, plc: 15.0 },
      "莊富丞": { target: 2400, actual: 3858, leb: 65.0, lnh: 78.5, plc: 35.2 },
      "何宛茹": { target: 550, actual: 238, leb: 28.21, lnh: 4.86, plc: 2.49 },
      "張書偉": { target: 600, actual: 278, leb: 10.2, lnh: 12.0, plc: 5.6 },
      "黃柏翰": { target: 40, actual: 12, leb: 1.5, lnh: 2.1, plc: 0.8 }
    }
  };

  function getTargetDatabase() {
    try {
      const saved = localStorage.getItem("shihlin_target_database_v1");
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

  // 渲染首頁目標與戰況看板
  function renderHomeDashboard() {
    const db = getTargetDatabase();
    const currentSales = localStorage.getItem("shihlin_pwa_active_sales") || "曾維崧";
    const salesInfo = db.salesTargets[currentSales] || {
      target: 800, actual: 700, leb: 20.0, lnh: 15.0, plc: 10.0
    };

    // 1. 主目標 (全區總成果)
    const elMetric1Name = document.getElementById("displayMetric1Name");
    const elMetric1Sub = document.getElementById("displayMetric1Sub");
    const elBadgeMetric1 = document.getElementById("badgeMetric1Rate");
    const elDisplayPersonalRev = document.getElementById("displayPersonalRev");
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

    // 個人目標試算
    const pTarget = salesInfo.target || 1;
    const pActual = salesInfo.actual || 0;
    const pRate = ((pActual / pTarget) * 100).toFixed(1);
    const pDiff = pTarget - pActual;
    if (elDisplayPersonalRev) {
      if (pDiff <= 0) {
        elDisplayPersonalRev.innerText = `${formatMoney(pActual)} 萬元 / 目標 ${formatMoney(pTarget)} 萬元 (${pRate}%，超標 ${formatMoney(Math.abs(pDiff))} 萬元)`;
      } else {
        elDisplayPersonalRev.innerText = `${formatMoney(pActual)} 萬元 / 目標 ${formatMoney(pTarget)} 萬元 (${pRate}%，尚差 ${formatMoney(pDiff)} 萬元)`;
      }
    }

    // 2. 自製品目標 1: LEB (士林伺服)
    renderSubProductUI({
      badgeId: "badgeProdLEB",
      personalValId: "displayPersonalLEB",
      teamValId: "displayTeamLEB",
      barId: "barProdLEB",
      descId: "descProdLEB",
      personalVal: salesInfo.leb || 0,
      prodData: db.products.LEB || { target: 516, actual: 196 }
    });

    // 3. 自製品目標 2: LNH (士林變頻器)
    renderSubProductUI({
      badgeId: "badgeProdLNH",
      personalValId: "displayPersonalLNH",
      teamValId: "displayTeamLNH",
      barId: "barProdLNH",
      descId: "descProdLNH",
      personalVal: salesInfo.lnh || 0,
      prodData: db.products.LNH || { target: 516, actual: 243 }
    });

    // 4. 自製品目標 3: 士林 PLC 系列
    renderSubProductUI({
      badgeId: "badgeProdPLC",
      personalValId: "displayPersonalPLC",
      teamValId: "displayTeamPLC",
      barId: "barProdPLC",
      descId: "descProdPLC",
      personalVal: salesInfo.plc || 0,
      prodData: db.products.PLC || { target: 210, actual: 246 }
    });
  }

  function renderSubProductUI(cfg) {
    const badgeEl = document.getElementById(cfg.badgeId);
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
    if (personalEl) {
      personalEl.innerHTML = `${cfg.personalVal.toFixed(1)} <span class="unit-text">萬元</span>`;
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

  renderHomeDashboard();

  // ============================================================================
  // 6. 管理員目標資料維護 (貼上 Excel 自動解析模組)
  // ============================================================================
  const btnLoadSampleExcel = document.getElementById("btnLoadSampleExcel");
  const textExcelTargets = document.getElementById("textExcelTargets");
  const textExcelShipments = document.getElementById("textExcelShipments");
  const btnResetAdminDefaults = document.getElementById("btnResetAdminDefaults");
  const formAdminTargets = document.getElementById("formAdminTargets");

  const sampleTargetsExcelTSV = `營業員\t挑戰目標\t實績\t達成率
北區合計\t9300\t10370\t111.5%
自製品LEB\t516\t196\t38.0%
自製品LNH\t516\t243\t47.1%
自製品PLC\t210\t246\t117.1%
張何達\t900\t1451\t161.2%
葉仁豪\t700\t473\t67.6%
邱文輝\t1400\t1388\t99.1%
曾仁君\t810\t802\t99.0%
曾維崧\t810\t802\t99.0%
溫達仁\t1200\t1147\t95.6%
楊家豪\t700\t721\t103.0%
莊富丞\t2400\t3858\t160.8%
何宛茹\t550\t238\t43.3%
張書偉\t600\t278\t46.3%
黃柏翰\t40\t12\t30.0%`;

  const sampleShipmentsExcelTSV = `營業員\tLEB(伺服)\tLNH(變頻器)\tPLC系列
何宛茹\t28.21\t4.86\t2.49
邱文輝\t3.65\t9.83\t143.75
張何達\t32.85\t49.62\t0.64
曾仁君\t28.20\t4.90\t12.00
曾維崧\t28.20\t4.90\t12.00
溫達仁\t42.10\t35.00\t20.30
楊家豪\t18.50\t22.40\t15.00
莊富丞\t65.00\t78.50\t35.20
葉仁豪\t15.00\t18.20\t8.50
張書偉\t10.20\t12.00\t5.60
黃柏翰\t1.50\t2.10\t0.80`;

  if (btnLoadSampleExcel) {
    btnLoadSampleExcel.addEventListener("click", () => {
      if (textExcelTargets) textExcelTargets.value = sampleTargetsExcelTSV;
      if (textExcelShipments) textExcelShipments.value = sampleShipmentsExcelTSV;
      alert("已填入 10 月 Excel 範例資料，點擊「解析並發布更新」即可生效！");
    });
  }

  if (btnResetAdminDefaults) {
    btnResetAdminDefaults.addEventListener("click", () => {
      if (confirm("確定要恢復系統初始預設目標數據嗎？")) {
        localStorage.removeItem("shihlin_target_database_v1");
        if (textExcelTargets) textExcelTargets.value = "";
        if (textExcelShipments) textExcelShipments.value = "";
        renderHomeDashboard();
        alert("已重設為系統初始數據！");
      }
    });
  }

  function parseExcelTSV(tsvText) {
    const lines = tsvText.trim().split(/\r?\n/);
    const rows = [];
    for (const line of lines) {
      if (!line.trim()) continue;
      // 支援 Tab 或逗號分隔
      const cols = line.includes("\t")
        ? line.split("\t").map(c => c.trim())
        : line.split(",").map(c => c.trim());
      rows.push(cols);
    }
    return rows;
  }

  if (formAdminTargets) {
    formAdminTargets.addEventListener("submit", (e) => {
      e.preventDefault();
      const targetsRaw = textExcelTargets ? textExcelTargets.value.trim() : "";
      const shipmentsRaw = textExcelShipments ? textExcelShipments.value.trim() : "";

      if (!targetsRaw && !shipmentsRaw) {
        alert("請先貼上目標總表或產品出貨明細表！");
        return;
      }

      const db = getTargetDatabase();

      // 1. 解析目標總表
      if (targetsRaw) {
        const rows = parseExcelTSV(targetsRaw);
        for (const row of rows) {
          if (row.length < 3) continue;
          const name = row[0];
          const targetVal = parseFloat(row[1].replace(/,/g, "")) || 0;
          const actualVal = parseFloat(row[2].replace(/,/g, "")) || 0;

          if (name.includes("合計") || name.includes("全區") || name.includes("北區")) {
            db.regionTotal.target = targetVal;
            db.regionTotal.actual = actualVal;
          } else if (name.toUpperCase().includes("LEB")) {
            db.products.LEB.target = targetVal;
            db.products.LEB.actual = actualVal;
          } else if (name.toUpperCase().includes("LNH")) {
            db.products.LNH.target = targetVal;
            db.products.LNH.actual = actualVal;
          } else if (name.toUpperCase().includes("PLC")) {
            db.products.PLC.target = targetVal;
            db.products.PLC.actual = actualVal;
          } else {
            // 業務人員個人目標
            if (!db.salesTargets[name]) {
              db.salesTargets[name] = { target: targetVal, actual: actualVal, leb: 0, lnh: 0, plc: 0 };
            } else {
              db.salesTargets[name].target = targetVal;
              db.salesTargets[name].actual = actualVal;
            }
          }
        }
      }

      // 2. 解析產品出貨明細
      if (shipmentsRaw) {
        const rows = parseExcelTSV(shipmentsRaw);
        for (const row of rows) {
          if (row.length < 4) continue;
          const name = row[0];
          if (name === "營業員" || name.includes("姓名")) continue;

          const leb = parseFloat(row[1].replace(/,/g, "")) || 0;
          const lnh = parseFloat(row[2].replace(/,/g, "")) || 0;
          const plc = parseFloat(row[3].replace(/,/g, "")) || 0;

          if (!db.salesTargets[name]) {
            db.salesTargets[name] = { target: 800, actual: 700, leb, lnh, plc };
          } else {
            db.salesTargets[name].leb = leb;
            db.salesTargets[name].lnh = lnh;
            db.salesTargets[name].plc = plc;
          }
        }
      }

      // 儲存至本地快取
      localStorage.setItem("shihlin_target_database_v1", JSON.stringify(db));
      renderHomeDashboard();
      alert("目標資料解析成功！首頁數據已即刻更新完畢。");
      switchPage(navItemHome, panelHome);
    });
  }

  // ============================================================================
  // 7. 客戶主檔載入與單一輸入自動帶入規範 (免重複打字)
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
  // 8. 日報清單與 Deal 案件看板載入
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

  // 初始化預設今日日期
  const todayStr = new Date().toISOString().split("T")[0];
  const inputReportDate = document.getElementById("inputReportDate");
  if (inputReportDate) inputReportDate.value = todayStr;

  // 初始載入既有數據
  loadCustomers();
  loadDailyReports();
  loadDeals();
});
