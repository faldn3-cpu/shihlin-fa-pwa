/**
 * 士林電機 FA 業務戰情與報價系統 - 前端主程式 (app.js)
 * 嚴格遵循：
 * 1. 100% 採用 Office / Windows Fluent 洗鍊線條樣式
 * 2. 徹底剔除彩色 Emoji 表情符號，採用洗鍊純文字狀態號誌 [達標] / [預警] / [落後]
 * 3. 業績卡更名：士林伺服 (LEB)、士林變頻器 (LNH)、士林PLC (LPB+LPN+LPD 加總)
 * 4. 曾維崧為主管，免計個人目標，直接看北區總戰況
 * 5. 目標維護支援 Excel 貼上自動解析 + 在線 Excel 試算表格子直接點選微調
 * 6. 拜訪單一紀錄規範：選取客戶自動帶入統編、電話、地址與聯絡人
 * 7. 三層漏斗：日報 ➔ 商機 ➔ 案件 連鎖建立
 */

document.addEventListener("DOMContentLoaded", () => {
  // ============================================================================
  // 全域變數與快取
  // ============================================================================
  let customersCache = [];
  let dealsCache = [];

  // ============================================================================
  // 1. 頂部即時時鐘 (格式：2026/10/21 下午 04:32) - 支援繁體中文 12 小時制
  // ============================================================================
  const headerClock = document.getElementById("headerClock");
  function updateHeaderClock() {
    const now = new Date();
    const yyyy = now.getFullYear();
    const mm = String(now.getMonth() + 1).padStart(2, "0");
    const dd = String(now.getDate()).padStart(2, "0");
    const hours = now.getHours();
    const min = String(now.getMinutes()).padStart(2, "0");
    const period = hours >= 12 ? "下午" : "上午";
    const h12 = hours === 0 ? 12 : (hours > 12 ? hours - 12 : hours);
    const hh = String(h12).padStart(2, "0");
    if (headerClock) {
      headerClock.innerText = `${yyyy}/${mm}/${dd} ${period} ${hh}:${min}`;
    }
  }
  updateHeaderClock();
  setInterval(updateHeaderClock, 10000); // 每 10 秒即時更新

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
  const navItemForecast = document.getElementById("navItemForecast");
  const navItemSettings = document.getElementById("navItemSettings");
  const navItemAdminTargets = document.getElementById("navItemAdminTargets");
  const navItemDriverMgmt = document.getElementById("navItemDriverMgmt");

  const panelHome = document.getElementById("panelHome");
  const panelCases = document.getElementById("panelCases");
  const panelQuotes = document.getElementById("panelQuotes");
  const panelSpecials = document.getElementById("panelSpecials");
  const panelForecast = document.getElementById("panelForecast");
  const panelSettings = document.getElementById("panelSettings");
  const panelAdminTargets = document.getElementById("panelAdminTargets");
  const panelDriverMgmt = document.getElementById("panelDriverMgmt");

  const navItems = [
    { btn: navItemHome, panel: panelHome },
    { btn: navItemCases, panel: panelCases },
    { btn: navItemQuotes, panel: panelQuotes },
    { btn: navItemSpecials, panel: panelSpecials },
    { btn: navItemForecast, panel: panelForecast },
    { btn: navItemAdminTargets, panel: panelAdminTargets },
    { btn: navItemDriverMgmt, panel: panelDriverMgmt },
    { btn: navItemSettings, panel: panelSettings }
  ];

  const DEFAULT_MANAGERS = ["曾維崧", "曾仁君", "張何達"];

  function getManagerSalesList() {
    const saved = localStorage.getItem("shihlin_pwa_manager_sales");
    if (saved) {
      try {
        const list = JSON.parse(saved);
        if (Array.isArray(list) && list.length > 0) return list;
      } catch (e) {
        console.warn("解析管理者名單失敗:", e);
      }
    }
    return [...DEFAULT_MANAGERS];
  }

  function isSalesManager(name) {
    const list = getManagerSalesList();
    return list.includes(name);
  }

  function renderManagerCheckboxes() {
    const container = document.getElementById("containerManagerCheckboxes");
    if (!container) return;
    const allSales = getOrderedSalespersons();
    const currentManagers = getManagerSalesList();

    container.innerHTML = allSales.map(s => {
      const isChecked = currentManagers.includes(s.name) ? "checked" : "";
      return `
        <label class="manager-checkbox-item">
          <input type="checkbox" class="chk-manager-sales" value="${escapeHtml(s.name)}" ${isChecked} />
          <span>${escapeHtml(s.name)}</span>
        </label>
      `;
    }).join("");
  }

  function updateManagerUIVisibility(activeSalesName) {
    const isMgr = isSalesManager(activeSalesName);
    
    // 1. 控制功能選單中「業績數據資料庫」的顯示權限
    const navAdmin = document.getElementById("navItemAdminTargets");
    if (navAdmin) {
      navAdmin.style.display = isMgr ? "" : "none";
    }

    // 2. 控制領先指標審核中心面板顯示
    const panelReview = document.getElementById("panelDriverReview");
    if (panelReview) {
      panelReview.style.display = isMgr ? "block" : "none";
    }
  }

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

  // ============================================================================
  // 清除快取並強制重整 (Clear Cache & Hard Reload)
  // ============================================================================
  async function clearAppCacheAndReload(triggerBtn) {
    if (triggerBtn) {
      triggerBtn.disabled = true;
      triggerBtn.style.opacity = "0.7";
      const span = triggerBtn.querySelector("span");
      if (span) span.textContent = "清除中...";
    }

    try {
      // 1. 註銷現有所有 Service Worker
      if ("serviceWorker" in navigator) {
        const registrations = await navigator.serviceWorker.getRegistrations();
        for (const reg of registrations) {
          await reg.unregister();
        }
      }

      // 2. 徹底清空 Cache Storage 中所有快取檔案
      if ("caches" in window) {
        const cacheKeys = await caches.keys();
        await Promise.all(cacheKeys.map(key => caches.delete(key)));
      }
    } catch (err) {
      console.warn("[Cache] 清理快取時發生非預期錯誤:", err);
    } finally {
      // 3. 帶上時間戳記強制刷新，繞過瀏覽器本地快取，向伺服器取得最新版本
      const timestamp = Date.now();
      const newUrl = window.location.origin + window.location.pathname + "?v=" + timestamp;
      window.location.replace(newUrl);
    }
  }

  const btnClearCacheDrawer = document.getElementById("btnClearCacheDrawer");
  if (btnClearCacheDrawer) {
    btnClearCacheDrawer.addEventListener("click", (e) => {
      e.stopPropagation();
      clearAppCacheAndReload(btnClearCacheDrawer);
    });
  }

  const btnClearCacheSettings = document.getElementById("btnClearCacheSettings");
  if (btnClearCacheSettings) {
    btnClearCacheSettings.addEventListener("click", (e) => {
      e.stopPropagation();
      clearAppCacheAndReload(btnClearCacheSettings);
    });
  }

  function setSidebarMode(isQuotationMode) {
    const mainList = document.getElementById("mainSidebarNavList");
    const quoteList = document.getElementById("quoteSidebarNavList");
    const drawerTitle = document.getElementById("drawerMainTitle");
    const appDrawer = document.getElementById("appDrawer");
    if (isQuotationMode) {
      if (mainList) mainList.classList.add("hidden");
      if (quoteList) quoteList.classList.remove("hidden");
      if (drawerTitle) drawerTitle.textContent = "報價管理系統";
      // 標記目前處於報價單模式，讓報價單專用的版面規則只在報價單內生效
      document.body.classList.add("quote-mode");
      if (appDrawer) appDrawer.classList.remove("collapsed");
    } else {
      if (mainList) mainList.classList.remove("hidden");
      if (quoteList) quoteList.classList.add("hidden");
      if (drawerTitle) drawerTitle.textContent = "功能選單";
      // 離開報價單模式，移除報價單專用版面標記，其他頁面恢復原樣
      document.body.classList.remove("quote-mode");
    }
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

      if (activePanel === panelDriverMgmt) {
        setSidebarMode(false);
        const curSales = localStorage.getItem("shihlin_pwa_active_sales") || "曾維崧";
        currentViewingDriverSales = currentViewingDriverSales || curSales;
        renderDriverManagementPage(currentViewingDriverSales);
      } else if (activePanel === panelQuotes) {
        setSidebarMode(true);
        // 進入報價單時若視窗小於 1280px，直接進入網頁版報價單模式 (只顯示報價擬稿)
        forceQuoteWebModeIfNarrow();
        initQuotationModule();
        if (typeof updateQuoteDraftsBadge === "function") {
          updateQuoteDraftsBadge();
        }
      } else {
        setSidebarMode(false);
      }
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

  // 6 大核心商務主題色盤 (含深淺雙模高對比自適應引擎)
  const ACCENT_PALETTES = {
    blue: {
      name: "經典科技藍",
      light: { accent: "#0067C0", btnBg: "#0067C0", btnText: "#FFFFFF", lightBg: "rgba(0, 103, 192, 0.10)" },
      dark:  { accent: "#2899F5", btnBg: "#0078D4", btnText: "#FFFFFF", lightBg: "rgba(40, 153, 245, 0.16)" }
    },
    green: {
      name: "翡翠商務綠",
      light: { accent: "#107C41", btnBg: "#107C41", btnText: "#FFFFFF", lightBg: "rgba(16, 124, 65, 0.10)" },
      dark:  { accent: "#2ECC71", btnBg: "#107C41", btnText: "#FFFFFF", lightBg: "rgba(46, 204, 113, 0.16)" }
    },
    orange: {
      name: "活力朝氣橘",
      light: { accent: "#D83B01", btnBg: "#D83B01", btnText: "#FFFFFF", lightBg: "rgba(216, 59, 1, 0.10)" },
      dark:  { accent: "#FF8C00", btnBg: "#D83B01", btnText: "#FFFFFF", lightBg: "rgba(255, 140, 0, 0.16)" }
    },
    purple: {
      name: "專業高階紫",
      light: { accent: "#5C2D91", btnBg: "#5C2D91", btnText: "#FFFFFF", lightBg: "rgba(92, 45, 145, 0.10)" },
      dark:  { accent: "#A779E9", btnBg: "#7B38A8", btnText: "#FFFFFF", lightBg: "rgba(167, 121, 233, 0.16)" }
    },
    black: {
      name: "沉穩商務黑",
      light: { accent: "#2B2B2B", btnBg: "#2B2B2B", btnText: "#FFFFFF", lightBg: "rgba(0, 0, 0, 0.08)" },
      // 深色模式下反轉為高對比鉑金銀白按鈕與黑字，確保最高易讀性
      dark:  { accent: "#F0F2F5", btnBg: "#F0F2F5", btnText: "#1A1A1A", lightBg: "rgba(255, 255, 255, 0.12)" }
    },
    cyan: {
      name: "深海科技青",
      light: { accent: "#008272", btnBg: "#008272", btnText: "#FFFFFF", lightBg: "rgba(0, 130, 114, 0.10)" },
      dark:  { accent: "#00B7C3", btnBg: "#008272", btnText: "#FFFFFF", lightBg: "rgba(0, 183, 195, 0.16)" }
    }
  };

  const HEX_TO_KEY = {
    "#0067C0": "blue",
    "#0078D4": "blue",
    "#107C41": "green",
    "#D83B01": "orange",
    "#5C2D91": "purple",
    "#2B2B2B": "black",
    "#1A1A1A": "black",
    "#000000": "black",
    "#008272": "cyan"
  };

  function applyTheme(themeMode) {
    if (themeMode === "dark") {
      document.documentElement.setAttribute("data-theme", "dark");
    } else if (themeMode === "light") {
      document.documentElement.removeAttribute("data-theme");
    } else {
      if (systemDarkMedia.matches) {
        document.documentElement.setAttribute("data-theme", "dark");
      } else {
        document.documentElement.removeAttribute("data-theme");
      }
    }

    // 更新膠囊分段控制器外觀
    const segThemeMode = document.getElementById("segThemeMode");
    if (segThemeMode) {
      segThemeMode.querySelectorAll(".seg-btn").forEach(btn => {
        btn.classList.toggle("active", btn.dataset.val === themeMode);
      });
    }

    // 當切換深淺模式時，同步重新計算當前色系之對比配色（特別是黑色系反轉為鉑金銀白按鈕）
    const currentAccent = localStorage.getItem("shihlin_pwa_theme_accent") || "blue";
    applyThemeAccent(currentAccent);
  }

  function applyThemeAccent(colorKey) {
    if (!colorKey) colorKey = "blue";
    const key = HEX_TO_KEY[colorKey] || (ACCENT_PALETTES[colorKey] ? colorKey : "blue");
    const palette = ACCENT_PALETTES[key] || ACCENT_PALETTES.blue;

    const isDark = document.documentElement.getAttribute("data-theme") === "dark";
    const themeColors = isDark ? palette.dark : palette.light;

    document.documentElement.style.setProperty("--color-accent", themeColors.accent);
    document.documentElement.style.setProperty("--color-accent-hover", themeColors.accent);
    document.documentElement.style.setProperty("--color-accent-light", themeColors.lightBg);
    document.documentElement.style.setProperty("--btn-primary-bg", themeColors.btnBg);
    document.documentElement.style.setProperty("--btn-primary-text", themeColors.btnText);

    // 同步色票按鈕選取狀態
    const swatchesContainer = document.getElementById("swatchesThemeAccent");
    if (swatchesContainer) {
      swatchesContainer.querySelectorAll(".swatch-btn").forEach(btn => {
        btn.classList.toggle("active", btn.dataset.color === key);
      });
    }
  }

  function applyFont(fontMode) {
    // 微軟 Office 繁體中文常用 7 大字型選單
    const fontMap = {
      jhenghei: "'Microsoft JhengHei', '微軟正黑體', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif",
      ming: "'PMingLiU', '新細明體', 'MingLiU', serif",
      kai: "'DFKai-SB', '標楷體', 'BiauKai', 'KaiTi', serif",
      yahei: "'Microsoft YaHei', '微軟雅黑', sans-serif",
      arial_jhenghei: "Arial, 'Microsoft JhengHei', '微軟正黑體', sans-serif",
      calibri_jhenghei: "Calibri, 'Microsoft JhengHei', '微軟正黑體', sans-serif",
      times_ming: "'Times New Roman', 'PMingLiU', '新細明體', serif"
    };
    const val = fontMap[fontMode] || fontMap.jhenghei;
    document.documentElement.style.setProperty("--app-font-family", val);
  }

  function applyFontSize(size) {
    if (!size) size = "12pt";
    document.documentElement.style.setProperty("--app-font-size", size);
    // 移除行內 style.fontSize，交由 CSS 中的 html { font-size: var(--app-font-size) } 與行動端 @media 智慧縮放
    document.documentElement.style.removeProperty("font-size");
  }

  function applyDensity(mode) {
    document.body.setAttribute("data-density", mode || "normal");
  }

  function applyBorderRadius(rad) {
    document.documentElement.style.setProperty("--border-radius", rad || "6px");
  }

  function applyTabularNums(mode) {
    document.body.setAttribute("data-tabular", mode || "tabular");
    const segTabularNums = document.getElementById("segTabularNums");
    if (segTabularNums) {
      segTabularNums.querySelectorAll(".seg-btn").forEach(btn => {
        btn.classList.toggle("active", btn.dataset.val === mode);
      });
    }
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
    if (selectActiveSales && selectActiveSales.value !== salesName) {
      selectActiveSales.value = salesName;
    }
    const forecastBadge = document.getElementById("forecastSalesBadge");
    if (forecastBadge) forecastBadge.innerText = `檢視身份：${salesName}`;
    const spOwnerBadge = document.getElementById("specialProductOwnerBadge");
    if (spOwnerBadge) spOwnerBadge.innerText = `所屬業務：${salesName}`;
    renderHomeDashboard(salesName);
  }

  // 全系統標準業務人員預設排序名單 (Single Source of Truth)
  const DEFAULT_SALESPERSONS = [
    { id: "曾維崧", name: "曾維崧", role: "北區總管", group: "主管" },
    { id: "曾仁君", name: "曾仁君", role: "直賣主管", group: "主管" },
    { id: "何宛茹", name: "何宛茹", role: "直賣", group: "直賣" },
    { id: "張何達", name: "張何達", role: "經銷主管", group: "主管" },
    { id: "葉仁豪", name: "葉仁豪", role: "經銷", group: "經銷" },
    { id: "邱文輝", name: "邱文輝", role: "經銷", group: "經銷" },
    { id: "溫達仁", name: "溫達仁", role: "直賣", group: "直賣" },
    { id: "楊家豪", name: "楊家豪", role: "直賣", group: "直賣" },
    { id: "莊富丞", name: "莊富丞", role: "直賣", group: "直賣" },
    { id: "張書偉", name: "張書偉", role: "直賣", group: "直賣" },
    { id: "黃柏翰", name: "黃柏翰", role: "直賣", group: "直賣" }
  ];

  // 取得當前有效業務人員排序清單 (優先自 localStorage 讀取，並自動清洗歷史快取中的全區字樣)
  function getOrderedSalespersons() {
    const saved = localStorage.getItem("shihlin_salespersons_order") || localStorage.getItem("shihlin_salespersons_default_order");
    const currentRegion = getCurrentRegion();
    if (saved) {
      try {
        const parsed = JSON.parse(saved);
        if (Array.isArray(parsed) && parsed.length > 0) {
          return parsed.map(s => ({
            ...s,
            role: (s.role || "").replace(/全區/g, currentRegion) || (s.name === "曾維崧" ? `${currentRegion}總管` : (s.group || "直賣"))
          }));
        }
      } catch (e) {
        console.warn("讀取業務人員排序失敗，使用預設值：", e);
      }
    }
    return [...DEFAULT_SALESPERSONS];
  }

  // 儲存業務人員排序清單並即刻驅動全系統下拉選單與排序表
  function saveSalespersonsOrder(list) {
    localStorage.setItem("shihlin_salespersons_order", JSON.stringify(list));
    renderActiveSalesSelectOptions();
    renderSalespersonOrderTable();
    if (typeof renderAdminImpersonateOptions === "function") {
      renderAdminImpersonateOptions();
    }
  }

  // 依據排序設定表動態渲染「目前業務身份」下拉選單
  function renderActiveSalesSelectOptions() {
    const selectActiveSales = document.getElementById("selectActiveSales");
    if (!selectActiveSales) return;
    const currentVal = selectActiveSales.value || localStorage.getItem("shihlin_pwa_active_sales") || "曾維崧";
    const list = getOrderedSalespersons();
    selectActiveSales.innerHTML = "";
    list.forEach(sales => {
      const opt = document.createElement("option");
      opt.value = sales.name;
      opt.innerText = sales.name;
      selectActiveSales.appendChild(opt);
    });
    const exists = list.some(s => s.name === currentVal);
    selectActiveSales.value = exists ? currentVal : (list[0] && list[0].name) || "曾維崧";
  }

  // 記憶體編輯工作草稿 (未按「保存」前不覆寫本地永久預設)
  let currentSalespersonsDraft = null;
  let pendingDeleteSalesName = null;

  function promptDeleteSalesperson(name) {
    if (!currentSalespersonsDraft) initSalespersonsDraft();
    if (currentSalespersonsDraft.length <= 1) {
      alert("系統名冊至少需保留一名業務人員，無法刪除！");
      return;
    }
    pendingDeleteSalesName = name;
    const modal = document.getElementById("modalConfirmDeleteSalesperson");
    const spanName = document.getElementById("spanDeleteSalesName");
    if (spanName) spanName.innerText = name;
    if (modal) modal.classList.remove("hidden");
  }

  function initSalespersonsDraft() {
    const list = getOrderedSalespersons();
    const mgrs = getManagerSalesList();
    currentSalespersonsDraft = list.map(s => {
      let grp = s.group || "";
      if (grp === "主管") {
        if (s.role && s.role.includes("直賣")) grp = "直賣";
        else if (s.role && s.role.includes("經銷")) grp = "經銷";
        else grp = "";
      } else if (!grp && s.role) {
        if (s.role.includes("直賣")) grp = "直賣";
        else if (s.role.includes("經銷")) grp = "經銷";
      }
      if (s.name === "曾維崧") grp = "";
      return {
        id: s.id || s.name,
        name: s.name,
        group: grp,
        isManager: mgrs.includes(s.name)
      };
    });
  }

  // 渲染業務人員與權限設定總表 (二合一整合表：通路下拉、純Checkbox、上移下移、人員增減刪除)
  function renderSalespersonOrderTable() {
    const tbody = document.getElementById("tbodySalesOrder");
    const containerMobile = document.getElementById("containerSalesOrderMobile");
    if (!tbody && !containerMobile) return;

    if (!currentSalespersonsDraft) {
      initSalespersonsDraft();
    }

    // -------------------------------------------------------------
    // 方案 A：渲染電腦端標準寬表格 (tbodySalesOrder)
    // -------------------------------------------------------------
    if (tbody) {
      tbody.innerHTML = "";
      currentSalespersonsDraft.forEach((sales, idx) => {
        const tr = document.createElement("tr");

        // 1. 序號
        const tdIndex = document.createElement("td");
        tdIndex.style.textAlign = "center";
        tdIndex.style.fontWeight = "600";
        tdIndex.innerText = (idx + 1).toString();

        // 2. 姓名
        const tdName = document.createElement("td");
        tdName.style.fontWeight = "600";
        tdName.innerText = sales.name;

        // 3. 通路 (直賣 / 經銷 / 空白[無文字])
        const tdGroup = document.createElement("td");
        tdGroup.style.textAlign = "center";
        const selGroup = document.createElement("select");
        selGroup.className = "sales-group-select";
        selGroup.innerHTML = `
          <option value="直賣" ${sales.group === "直賣" ? "selected" : ""}>直賣</option>
          <option value="經銷" ${sales.group === "經銷" ? "selected" : ""}>經銷</option>
          <option value="" ${(!sales.group || sales.group === "空白") ? "selected" : ""}></option>
        `;
        selGroup.onchange = (e) => {
          sales.group = e.target.value;
          renderSalespersonOrderTable();
        };
        tdGroup.appendChild(selGroup);

        // 4. 數據管理權限 (純 Checkbox 置中)
        const tdMgr = document.createElement("td");
        tdMgr.style.textAlign = "center";
        const chkMgr = document.createElement("input");
        chkMgr.type = "checkbox";
        chkMgr.className = "chk-manager-sales";
        chkMgr.checked = !!sales.isManager;
        chkMgr.style.width = "16px";
        chkMgr.style.height = "16px";
        chkMgr.style.cursor = "pointer";
        chkMgr.onchange = (e) => {
          sales.isManager = e.target.checked;
          renderSalespersonOrderTable();
        };
        tdMgr.appendChild(chkMgr);

        // 5. 順序調整 (上移 / 下移)
        const tdOrder = document.createElement("td");
        tdOrder.style.textAlign = "center";

        const btnUp = document.createElement("button");
        btnUp.type = "button";
        btnUp.className = "btn-order-move";
        btnUp.innerText = "上移 ⬆";
        btnUp.disabled = (idx === 0);
        btnUp.onclick = () => {
          if (idx > 0) {
            const temp = currentSalespersonsDraft[idx - 1];
            currentSalespersonsDraft[idx - 1] = currentSalespersonsDraft[idx];
            currentSalespersonsDraft[idx] = temp;
            renderSalespersonOrderTable();
          }
        };

        const btnDown = document.createElement("button");
        btnDown.type = "button";
        btnDown.className = "btn-order-move";
        btnDown.innerText = "下移 ⬇";
        btnDown.disabled = (idx === currentSalespersonsDraft.length - 1);
        btnDown.onclick = () => {
          if (idx < currentSalespersonsDraft.length - 1) {
            const temp = currentSalespersonsDraft[idx + 1];
            currentSalespersonsDraft[idx + 1] = currentSalespersonsDraft[idx];
            currentSalespersonsDraft[idx] = temp;
            renderSalespersonOrderTable();
          }
        };

        tdOrder.appendChild(btnUp);
        tdOrder.appendChild(btnDown);

        // 6. 人員增減 (刪除按鈕附帶警示確認)
        const tdActions = document.createElement("td");
        tdActions.style.textAlign = "center";
        const btnDel = document.createElement("button");
        btnDel.type = "button";
        btnDel.className = "btn-order-delete";
        btnDel.innerHTML = `<svg class="icon-svg" style="width:13px; height:13px;" viewBox="0 0 24 24"><path d="M6 19c0 1.1.9 2 2 2h8c1.1 0 2-.9 2-2V7H6v12zM19 4h-3.5l-1-1h-5l-1 1H5v2h14V4z"/></svg> 刪除`;
        btnDel.onclick = () => {
          promptDeleteSalesperson(sales.name);
        };
        tdActions.appendChild(btnDel);

        tr.appendChild(tdIndex);
        tr.appendChild(tdName);
        tr.appendChild(tdGroup);
        tr.appendChild(tdMgr);
        tr.appendChild(tdOrder);
        tr.appendChild(tdActions);

        tbody.appendChild(tr);
      });
    }

    // -------------------------------------------------------------
    // 方案 B：渲染手機端專用觸控卡片清單 (containerSalesOrderMobile)
    // -------------------------------------------------------------
    if (containerMobile) {
      containerMobile.innerHTML = "";
      currentSalespersonsDraft.forEach((sales, idx) => {
        const card = document.createElement("div");
        card.className = "sales-mobile-card";

        // 第一列：序號標章 + 姓名 + 通路下拉 + 數據管理核取方塊
        const row1 = document.createElement("div");
        row1.className = "sales-mobile-card-row1";

        const leftGroup = document.createElement("div");
        leftGroup.style.display = "flex";
        leftGroup.style.alignItems = "center";
        leftGroup.style.gap = "8px";

        const badge = document.createElement("span");
        badge.className = "sales-order-badge";
        badge.innerText = `#${idx + 1}`;

        const nameSpan = document.createElement("span");
        nameSpan.className = "sales-mobile-name";
        nameSpan.innerText = sales.name;

        leftGroup.appendChild(badge);
        leftGroup.appendChild(nameSpan);

        const rightGroup = document.createElement("div");
        rightGroup.style.display = "flex";
        rightGroup.style.alignItems = "center";
        rightGroup.style.gap = "10px";

        // 手機端通路選擇
        const selGroupM = document.createElement("select");
        selGroupM.className = "sales-group-select mobile-select";
        selGroupM.innerHTML = `
          <option value="直賣" ${sales.group === "直賣" ? "selected" : ""}>直賣</option>
          <option value="經銷" ${sales.group === "經銷" ? "selected" : ""}>經銷</option>
          <option value="" ${(!sales.group || sales.group === "空白") ? "selected" : ""}></option>
        `;
        selGroupM.onchange = (e) => {
          sales.group = e.target.value;
          renderSalespersonOrderTable();
        };

        // 手機端數據管理權限
        const labelMgrM = document.createElement("label");
        labelMgrM.style.display = "flex";
        labelMgrM.style.alignItems = "center";
        labelMgrM.style.gap = "4px";
        labelMgrM.style.margin = "0";
        labelMgrM.style.fontSize = "0.82rem";
        labelMgrM.style.cursor = "pointer";

        const chkMgrM = document.createElement("input");
        chkMgrM.type = "checkbox";
        chkMgrM.className = "chk-manager-sales";
        chkMgrM.checked = !!sales.isManager;
        chkMgrM.style.width = "16px";
        chkMgrM.style.height = "16px";
        chkMgrM.onchange = (e) => {
          sales.isManager = e.target.checked;
          renderSalespersonOrderTable();
        };

        const mgrText = document.createElement("span");
        mgrText.innerText = "管理";

        labelMgrM.appendChild(chkMgrM);
        labelMgrM.appendChild(mgrText);

        rightGroup.appendChild(selGroupM);
        rightGroup.appendChild(labelMgrM);

        row1.appendChild(leftGroup);
        row1.appendChild(rightGroup);

        // 第二列：上移 / 下移 按鈕 + 刪除按鈕
        const row2 = document.createElement("div");
        row2.className = "sales-mobile-card-row2";

        const moveBox = document.createElement("div");
        moveBox.style.display = "flex";
        moveBox.style.gap = "6px";

        const btnUpM = document.createElement("button");
        btnUpM.type = "button";
        btnUpM.className = "btn-order-move mobile-move";
        btnUpM.innerText = "上移 ⬆";
        btnUpM.disabled = (idx === 0);
        btnUpM.onclick = () => {
          if (idx > 0) {
            const temp = currentSalespersonsDraft[idx - 1];
            currentSalespersonsDraft[idx - 1] = currentSalespersonsDraft[idx];
            currentSalespersonsDraft[idx] = temp;
            renderSalespersonOrderTable();
          }
        };

        const btnDownM = document.createElement("button");
        btnDownM.type = "button";
        btnDownM.className = "btn-order-move mobile-move";
        btnDownM.innerText = "下移 ⬇";
        btnDownM.disabled = (idx === currentSalespersonsDraft.length - 1);
        btnDownM.onclick = () => {
          if (idx < currentSalespersonsDraft.length - 1) {
            const temp = currentSalespersonsDraft[idx + 1];
            currentSalespersonsDraft[idx + 1] = currentSalespersonsDraft[idx];
            currentSalespersonsDraft[idx] = temp;
            renderSalespersonOrderTable();
          }
        };

        moveBox.appendChild(btnUpM);
        moveBox.appendChild(btnDownM);

        const btnDelM = document.createElement("button");
        btnDelM.type = "button";
        btnDelM.className = "btn-order-delete mobile-del";
        btnDelM.innerHTML = `<svg class="icon-svg" style="width:13px; height:13px;" viewBox="0 0 24 24"><path d="M6 19c0 1.1.9 2 2 2h8c1.1 0 2-.9 2-2V7H6v12zM19 4h-3.5l-1-1h-5l-1 1H5v2h14V4z"/></svg> 刪除`;
        btnDelM.onclick = () => {
          promptDeleteSalesperson(sales.name);
        };

        row2.appendChild(moveBox);
        row2.appendChild(btnDelM);

        card.appendChild(row1);
        card.appendChild(row2);

        containerMobile.appendChild(card);
      });
    }
  }

  function initSettings() {
    // 1. 外觀主題模式 (Segmented Control 膠囊分段控制器)
    const segThemeMode = document.getElementById("segThemeMode");
    const savedTheme = localStorage.getItem("shihlin_pwa_theme_mode") || "auto";
    if (segThemeMode) {
      segThemeMode.querySelectorAll(".seg-btn").forEach(btn => {
        btn.addEventListener("click", () => {
          const mode = btn.dataset.val;
          localStorage.setItem("shihlin_pwa_theme_mode", mode);
          applyTheme(mode);
        });
      });
    }
    systemDarkMedia.addEventListener("change", () => {
      const mode = localStorage.getItem("shihlin_pwa_theme_mode") || "auto";
      if (mode === "auto") applyTheme("auto");
    });
    applyTheme(savedTheme);

    // 2. 主題色系 (彩色圓點色票按鈕 Color Swatches)
    const swatchesThemeAccent = document.getElementById("swatchesThemeAccent");
    const savedAccent = localStorage.getItem("shihlin_pwa_theme_accent") || "blue";
    if (swatchesThemeAccent) {
      swatchesThemeAccent.querySelectorAll(".swatch-btn").forEach(btn => {
        btn.addEventListener("click", () => {
          const colKey = btn.dataset.color;
          localStorage.setItem("shihlin_pwa_theme_accent", colKey);
          applyThemeAccent(colKey);
        });
      });
    }
    applyThemeAccent(savedAccent);

    // 3. 系統字型 (微軟 Office 繁體常用字型)
    const selectSystemFont = document.getElementById("selectSystemFont");
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

    // 4. 字級大小 (8pt ~ 24pt 純淨文字選單)
    const selectFontSize = document.getElementById("selectFontSize");
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

    // 5. 數字排版 (Segmented Control 膠囊分段控制器)
    const segTabularNums = document.getElementById("segTabularNums");
    const savedTabular = localStorage.getItem("shihlin_pwa_tabular_nums") || "tabular";
    if (segTabularNums) {
      segTabularNums.querySelectorAll(".seg-btn").forEach(btn => {
        btn.addEventListener("click", () => {
          const tab = btn.dataset.val;
          localStorage.setItem("shihlin_pwa_tabular_nums", tab);
          applyTabularNums(tab);
        });
      });
    }
    applyTabularNums(savedTabular);

    // 6. 所屬營運區域設定 (北區 / 中區 / 南區，切換即刻生效並自動儲存為預設)
    const selectActiveRegion = document.getElementById("selectActiveRegion");
    const currentReg = getCurrentRegion();
    if (selectActiveRegion) {
      selectActiveRegion.value = currentReg;
      selectActiveRegion.addEventListener("change", (e) => {
        const rVal = e.target.value;
        localStorage.setItem("shihlin_pwa_current_region", rVal);
        localStorage.setItem("shihlin_pwa_default_region", rVal);
        renderActiveSalesSelectOptions();
        renderHomeDashboard();
      });
    }

    // 7. 動態初始化業務人員排序表與目前業務下拉選單 (切換即刻生效並自動儲存)
    renderActiveSalesSelectOptions();
    renderSalespersonOrderTable();

    const savedSales = localStorage.getItem("shihlin_pwa_active_sales") || "曾維崧";
    const selectActiveSales = document.getElementById("selectActiveSales");
    if (selectActiveSales) {
      selectActiveSales.value = savedSales;
      selectActiveSales.addEventListener("change", (e) => {
        const sName = e.target.value;
        switchSimulatedSalesperson(sName, true);
      });
    }
    updateActiveSalesUI(savedSales);
    updateManagerUIVisibility(savedSales);

    // 11. 新增業務人員對話框 (彈跳 Modal，完全取代 window.prompt)
    const modalAddSalesperson = document.getElementById("modalAddSalesperson");
    const btnAddSalesperson = document.getElementById("btnAddSalesperson");
    const btnCloseAddSalesModal = document.getElementById("btnCloseAddSalesModal");
    const btnCancelAddSales = document.getElementById("btnCancelAddSales");
    const btnConfirmAddSales = document.getElementById("btnConfirmAddSales");
    const inputNewSalesName = document.getElementById("inputNewSalesName");
    const selectNewSalesGroup = document.getElementById("selectNewSalesGroup");
    const chkNewSalesIsManager = document.getElementById("chkNewSalesIsManager");

    if (btnAddSalesperson && modalAddSalesperson) {
      btnAddSalesperson.addEventListener("click", () => {
        if (!currentSalespersonsDraft) initSalespersonsDraft();
        if (inputNewSalesName) inputNewSalesName.value = "";
        if (selectNewSalesGroup) selectNewSalesGroup.value = "直賣";
        if (chkNewSalesIsManager) chkNewSalesIsManager.checked = false;
        modalAddSalesperson.classList.remove("hidden");
        setTimeout(() => inputNewSalesName && inputNewSalesName.focus(), 80);
      });
    }

    const hideAddModal = () => {
      if (modalAddSalesperson) modalAddSalesperson.classList.add("hidden");
    };
    if (btnCloseAddSalesModal) btnCloseAddSalesModal.addEventListener("click", hideAddModal);
    if (btnCancelAddSales) btnCancelAddSales.addEventListener("click", hideAddModal);

    // 刪除業務人員確認對話框事件監聽 (取代容易受瀏覽器封鎖之 window.confirm)
    const modalConfirmDeleteSalesperson = document.getElementById("modalConfirmDeleteSalesperson");
    const btnCloseDeleteSalesModal = document.getElementById("btnCloseDeleteSalesModal");
    const btnCancelDeleteSales = document.getElementById("btnCancelDeleteSales");
    const btnConfirmDeleteSales = document.getElementById("btnConfirmDeleteSales");

    const hideDeleteModal = () => {
      pendingDeleteSalesName = null;
      if (modalConfirmDeleteSalesperson) modalConfirmDeleteSalesperson.classList.add("hidden");
    };

    if (btnCloseDeleteSalesModal) btnCloseDeleteSalesModal.addEventListener("click", hideDeleteModal);
    if (btnCancelDeleteSales) btnCancelDeleteSales.addEventListener("click", hideDeleteModal);

    if (btnConfirmDeleteSales) {
      btnConfirmDeleteSales.addEventListener("click", () => {
        if (!pendingDeleteSalesName) {
          hideDeleteModal();
          return;
        }
        const targetName = pendingDeleteSalesName;
        hideDeleteModal();

        if (!currentSalespersonsDraft) initSalespersonsDraft();
        const targetIdx = currentSalespersonsDraft.findIndex(s => s.name === targetName);
        if (targetIdx !== -1) {
          currentSalespersonsDraft.splice(targetIdx, 1);
          renderSalespersonOrderTable();
          alert(`已將業務人員【${targetName}】自編輯名冊移除，請點擊右上角「保存」按鈕確認生效！`);
        }
      });
    }

    if (btnConfirmAddSales) {
      btnConfirmAddSales.addEventListener("click", () => {
        const nameVal = inputNewSalesName ? inputNewSalesName.value.trim() : "";
        if (!nameVal) {
          alert("請輸入業務人員姓名！");
          if (inputNewSalesName) inputNewSalesName.focus();
          return;
        }
        if (!currentSalespersonsDraft) initSalespersonsDraft();
        if (currentSalespersonsDraft.some(s => s.name === nameVal)) {
          alert(`業務人員【${nameVal}】已在名冊中，無法重複新增！`);
          return;
        }

        const grpVal = selectNewSalesGroup ? selectNewSalesGroup.value : "直賣";
        const isMgrVal = chkNewSalesIsManager ? chkNewSalesIsManager.checked : false;

        currentSalespersonsDraft.push({
          id: nameVal,
          name: nameVal,
          group: grpVal,
          isManager: isMgrVal
        });

        hideAddModal();
        renderSalespersonOrderTable();
        alert(`已將【${nameVal}】加入名冊編輯清單，請點擊右上角「保存」按鈕確認生效！`);
      });
    }

    // 12. 保存名冊與權限設定按鈕 (修改保存前彈出微軟 Fluent 警示確認對話盒)
    const btnSaveSalespersonSettings = document.getElementById("btnSaveSalespersonSettings");
    const modalConfirmSaveSalespersonSettings = document.getElementById("modalConfirmSaveSalespersonSettings");
    const btnCloseSaveSalesModal = document.getElementById("btnCloseSaveSalesModal");
    const btnCancelSaveSalesSettings = document.getElementById("btnCancelSaveSalesSettings");
    const btnConfirmSaveSalesSettings = document.getElementById("btnConfirmSaveSalesSettings");
    const spanSaveSummaryTotal = document.getElementById("spanSaveSummaryTotal");
    const spanSaveSummaryManagers = document.getElementById("spanSaveSummaryManagers");

    const hideSaveModal = () => {
      if (modalConfirmSaveSalespersonSettings) modalConfirmSaveSalespersonSettings.classList.add("hidden");
    };

    if (btnCloseSaveSalesModal) btnCloseSaveSalesModal.addEventListener("click", hideSaveModal);
    if (btnCancelSaveSalesSettings) btnCancelSaveSalesSettings.addEventListener("click", hideSaveModal);

    // 點擊「保存」按鈕 -> 執行前置檢查並彈出警示對話盒
    if (btnSaveSalespersonSettings) {
      btnSaveSalespersonSettings.addEventListener("click", () => {
        if (!currentSalespersonsDraft || currentSalespersonsDraft.length === 0) {
          alert("名冊人員不能為空！");
          return;
        }

        const mgrList = currentSalespersonsDraft.filter(s => s.isManager).map(s => s.name);
        if (mgrList.length === 0) {
          alert("請至少勾選一位具備數據管理權限之人員！");
          return;
        }

        // 動態注入當前名冊變更摘要
        if (spanSaveSummaryTotal) {
          spanSaveSummaryTotal.innerText = `${currentSalespersonsDraft.length} 位業務人員`;
        }
        if (spanSaveSummaryManagers) {
          spanSaveSummaryManagers.innerText = `${mgrList.length} 位（${mgrList.join("、")}）`;
        }

        // 彈出警示對話盒
        if (modalConfirmSaveSalespersonSettings) {
          modalConfirmSaveSalespersonSettings.classList.remove("hidden");
        }
      });
    }

    // 警示對話盒內點擊「確認保存」 -> 正式寫入系統並同步生效
    if (btnConfirmSaveSalesSettings) {
      btnConfirmSaveSalesSettings.addEventListener("click", () => {
        hideSaveModal();

        if (!currentSalespersonsDraft || currentSalespersonsDraft.length === 0) return;
        const mgrList = currentSalespersonsDraft.filter(s => s.isManager).map(s => s.name);
        if (mgrList.length === 0) return;

        const curRegion = getCurrentRegion();
        const cleanList = currentSalespersonsDraft.map(s => ({
          id: s.id || s.name,
          name: s.name,
          group: s.group || "",
          role: s.name === "曾維崧" ? `${curRegion}總管` : (s.group || "直賣")
        }));

        localStorage.setItem("shihlin_salespersons_default_order", JSON.stringify(cleanList));
        localStorage.setItem("shihlin_salespersons_order", JSON.stringify(cleanList));
        localStorage.setItem("shihlin_pwa_manager_sales", JSON.stringify(mgrList));

        // 同步更新「目前業務身份」選單 (若被刪除者為當前登入者，回退至第1位)
        const curActive = localStorage.getItem("shihlin_pwa_active_sales") || "曾維崧";
        const stillExists = cleanList.some(s => s.name === curActive);
        const finalActive = stillExists ? curActive : cleanList[0].name;
        localStorage.setItem("shihlin_pwa_active_sales", finalActive);

        renderActiveSalesSelectOptions();
        updateActiveSalesUI(finalActive);
        updateManagerUIVisibility(finalActive);
        renderManagerReviewPanel(finalActive);
        renderHomeDashboard(finalActive);

        alert("[保存成功] 業務人員名冊與數據管理權限已同步更新生效！");
      });
    }
  }

  function getCurrentRegion() {
    return localStorage.getItem("shihlin_pwa_current_region") || localStorage.getItem("shihlin_pwa_default_region") || "北區";
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

  // 戰情看板檢視視角狀態
  let currentScoreboardScope = null; // "region" (北區/營運區目標) 或 "personal" (個人目標)
  let lastParsedMetrics = null;
  let isDriverViewActive = false; // 是否折疊成果目標並僅檢視領先指標

  // 渲染首頁目標與戰況看板 (統一連動「業績數據資料庫」雙表格真實解析數據)
  function renderHomeDashboard(specificSales) {
    const activeSales = specificSales || localStorage.getItem("shihlin_pwa_active_sales") || "曾維崧";
    const currentGrid = (typeof tablePastedTargets !== "undefined" && tablePastedTargets && tablePastedTargets._gridMatrix && tablePastedTargets._gridMatrix.length > 0)
      ? tablePastedTargets._gridMatrix
      : getSampleTargetsGrid();
    const currentShipments = (typeof tablePastedShipments !== "undefined" && tablePastedShipments && tablePastedShipments._gridMatrix && tablePastedShipments._gridMatrix.length > 0)
      ? tablePastedShipments._gridMatrix
      : getSampleShipmentsGrid();
    const metrics = parseTargetsGridToMetrics(currentGrid, currentShipments, activeSales);
    lastParsedMetrics = metrics;

    // 依主管身分自動給定預設：主管預設營運區目標 ("region")，其他人預設個人目標 ("personal")
    if (!currentScoreboardScope || specificSales) {
      currentScoreboardScope = isSalesManager(activeSales) ? "region" : "personal";
    }
    updateScopeSwitchUI(currentScoreboardScope);
    applyScoreboardMetrics(metrics, currentScoreboardScope);
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

    const curRegion = getCurrentRegion();

    if (cfg.isManager) {
      // 主管模式：個人實績欄位直接標註「營運區實績」
      if (labelEl) labelEl.innerText = `${curRegion}實績`;
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
        descEl.innerText = `${curRegion}目標 ${formatMoney(target)} 萬 ｜ 差距 ${formatMoney(diff)} 萬`;
      } else {
        descEl.innerText = `${curRegion}目標 ${formatMoney(target)} 萬 ｜ 超標 ${formatMoney(Math.abs(diff))} 萬`;
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

  // 雙模表格色彩淨化與變數映射防呆機制 (Contrast Guard)
  function normalizeGridColor(color) {
    if (!color) return "";
    const c = color.trim().toLowerCase();
    // 1. 淨化寫死的黑色與深灰底字 (讓其自然繼承系統主題色 var(--text-primary))
    if (c === "#000" || c === "#000000" || c === "black" || c === "windowtext" ||
        c === "rgb(0, 0, 0)" || c === "rgb(0,0,0)" || c === "rgba(0, 0, 0, 1)" ||
        c === "#111827" || c === "#1f2937" || c === "#201f1e" || c === "#1f1f1f" ||
        c === "#333" || c === "#333333" || c === "#222" || c === "#2b2b2b") {
      return "";
    }
    // 2. 映射深藍/純藍數字 -> 雙模變數 var(--grid-num-blue)
    if (c === "#0000ff" || c === "#0052cc" || c === "#1d4ed8" || c === "#0000cd" || c === "blue" || c === "#2563eb" || c === "#1e40af") {
      return "var(--grid-num-blue)";
    }
    // 3. 映射紅色數字 -> 雙模變數 var(--grid-num-red)
    if (c === "#d93025" || c === "#b91c1c" || c === "#ff0000" || c === "red" || c === "#dc2626" || c === "#ef4444" || c === "#991b1b") {
      return "var(--grid-num-red)";
    }
    // 4. 映射綠色數字 -> 雙模變數 var(--grid-num-green)
    if (c === "#0d8050" || c === "#15803d" || c === "#008000" || c === "green" || c === "#16a34a" || c === "#107c41" || c === "#166534") {
      return "var(--grid-num-green)";
    }
    // 5. 映射橘色數字 -> 雙模變數 var(--grid-num-orange)
    if (c === "#c2410c" || c === "#ea580c" || c === "#f97316" || c === "orange" || c === "#9a3412") {
      return "var(--grid-num-orange)";
    }
    return color;
  }

  function normalizeGridBg(bg) {
    if (!bg || bg === "transparent" || bg === "inherit") return "";
    const b = bg.trim().toLowerCase();
    // 淺藍底色 -> 雙模變數 var(--grid-bg-blue) 或 var(--grid-bg-total)
    if (b === "#dbeafe" || b === "#bfdbfe" || b === "#e0f2fe" || b === "#bae6fd" || b === "#e6f4ea") {
      return b === "#e0f2fe" ? "var(--grid-bg-total)" : "var(--grid-bg-blue)";
    }
    // 淺紅底色 -> 雙模變數 var(--grid-bg-red)
    if (b === "#fee2e2" || b === "#fecaca" || b === "#ffebee") {
      return "var(--grid-bg-red)";
    }
    // 淺綠底色 -> 雙模變數 var(--grid-bg-green)
    if (b === "#dcfce7" || b === "#bbf7d0" || b === "#e8f5e9") {
      return "var(--grid-bg-green)";
    }
    // 淺橘底色 -> 雙模變數 var(--grid-bg-orange)
    if (b === "#fed7aa" || b === "#ffedd5" || b === "#fff3e0") {
      return "var(--grid-bg-orange)";
    }
    // 白色/純淺灰背景 (Excel 預設 white/rgb(255,255,255)) 在深色模式下淨化，避免黑底泛白塊
    if (b === "#fff" || b === "#ffffff" || b === "white" || b === "rgb(255, 255, 255)" || b === "rgb(255,255,255)") {
      return "";
    }
    return bg;
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

      // 核心升級：建立 2D 矩陣追蹤 rowspan 與 colspan，徹底防止合併儲存格造成後續欄位左移錯位
      const grid = [];
      const trList = table.querySelectorAll("tr");

      trList.forEach((tr, r) => {
        if (!grid[r]) grid[r] = [];
        let colIdx = 0;

        const cellList = tr.querySelectorAll("td, th");
        cellList.forEach(cell => {
          // 跳過已被上方列的垂直跨列 (rowspan) 或左側跨欄 (colspan) 佔據的座標
          while (grid[r][colIdx] !== undefined) {
            colIdx++;
          }

          const rowspan = parseInt(cell.getAttribute("rowspan") || "1", 10);
          const colspan = parseInt(cell.getAttribute("colspan") || "1", 10);

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

          // 執行貼上色彩淨化與變數轉換
          const sanitizedColor = normalizeGridColor(color);
          const sanitizedBg = normalizeGridBg(bg);

          const cellObj = (sanitizedBg || sanitizedColor || bold || fontSize || align)
            ? { val: text, bg: sanitizedBg, color: sanitizedColor, bold: bold ? "bold" : "", fontSize: fontSize, align: align }
            : text;

          // 填入 2D 矩陣
          for (let dr = 0; dr < rowspan; dr++) {
            const targetR = r + dr;
            if (!grid[targetR]) grid[targetR] = [];
            for (let dc = 0; dc < colspan; dc++) {
              const targetC = colIdx + dc;
              if (dr === 0 && dc === 0) {
                grid[targetR][targetC] = cellObj;
              } else {
                // 合併儲存格覆蓋之附屬座標，放置空字串佔位，防止同一列後續資料往左錯位
                grid[targetR][targetC] = "";
              }
            }
          }

          colIdx += colspan;
        });
      });

      // 移除尾端全空白列
      while (grid.length > 0 && grid[grid.length - 1].every(c => (typeof c === "object" ? !c.val : !c))) {
        grid.pop();
      }

      return grid.length > 0 ? grid : null;
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
          const normBg = normalizeGridBg(cellData.bg);
          const normColor = normalizeGridColor(cellData.color);

          if (normBg) {
            styleParts.push(`background-color:${normBg} !important`);
          }
          if (normColor) {
            styleParts.push(`color:${normColor} !important`);
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

  // ============================================================================
  // 10 月真實目標總表範例 (100% 對照圖二：含顏色、粗體、自製品三大項與合計)
  // ============================================================================
  function getSampleTargetsGrid() {
    const blue = (val) => ({ val, color: "var(--grid-num-blue)", bold: "bold" });
    const red = (val) => ({ val, color: "var(--grid-num-red)", bold: "bold" });
    const green = (val) => ({ val, color: "var(--grid-num-green)", bold: "bold" });
    const negRed = (val) => ({ val, color: "var(--grid-num-red)" });
    const hdrBlue = (val) => ({ val, color: "var(--grid-num-blue)", bg: "var(--grid-bg-blue)", bold: "bold" });
    const hdrRed = (val) => ({ val, color: "var(--grid-num-red)", bg: "var(--grid-bg-red)", bold: "bold" });
    const hdrGreen = (val) => ({ val, color: "var(--grid-num-green)", bg: "var(--grid-bg-green)", bold: "bold" });
    const hdrActBlue = (val) => ({ val, color: "var(--grid-num-blue)", bg: "var(--grid-bg-total)", bold: "bold" });
    const hdrOrange = (val) => ({ val, bg: "var(--grid-bg-orange)" });

    return [
      ["北區", "10", "月挑戰目標", "9,300", "萬", "", "", "", "", "", ""],
      ["項次", "營業員", hdrBlue("年計月目標"), hdrRed("10月\n挑戰目標"), hdrGreen("10月\n堆疊業績"), "10/21\n累積目標", hdrActBlue("10/21\n實績"), "堆疊目標\n差異(100%)", "挑戰目標\n差異", "挑戰目標\n達成率", hdrOrange("前一個月\n實績")],
      ["經銷", "張何達", blue("758"), red("900"), green("900"), "610", blue("1,451"), "551", "551", "161%", "957"],
      ["", "葉仁豪", blue("664"), red("700"), green("705"), "477", blue("473"), negRed("-231"), negRed("-227"), "68%", "692"],
      ["", "邱文輝", blue("1,450"), red("1,400"), green("1,400"), "948", blue("1,388"), negRed("-12"), negRed("-12"), "99%", "1,696"],
      ["直賣", "曾仁君", blue("477"), red("810"), green("810"), "549", blue("802"), negRed("-8"), negRed("-8"), "99%", "1,754"],
      ["", "溫達仁", blue("834"), red("1,200"), green("1,200"), "813", blue("1,147"), negRed("-53"), negRed("-53"), "96%", "863"],
      ["", "楊家豪", blue("658"), red("700"), green("720"), "488", blue("721"), "1", "21", "103%", "1,218"],
      ["", "莊富丞", blue("1,206"), red("2,400"), green("3,650"), "2,473", blue("3,858"), "208", "1,458", "161%", "2,942"],
      ["", "何宛茹", blue("401"), red("550"), green("398"), "270", blue("238"), negRed("-160"), negRed("-312"), "43%", "371"],
      ["", "張書偉", blue("612"), red("600"), green("629"), "426", blue("278"), negRed("-351"), negRed("-322"), "46%", "513"],
      ["", "黃柏翰", "", red("40"), green("40"), "", blue("12"), "", "", "", "20"],
      ["", "合計", blue("7,060"), red("9,300"), green("10,452"), "7,053", blue("10,370"), negRed("-54"), "1,070", "112%", "11,026"],
      ["北區", "自製品", "挑戰目標", "1,242", "萬", "", "", "", "", "", hdrOrange("")],
      ["1", "LEB", blue("468"), red("516"), green("382"), "259", blue("196"), negRed("-63"), negRed("-320"), "38%", "283"],
      ["2", "LNH", blue("567"), red("516"), green("320"), "217", blue("243"), "26", negRed("-273"), "47%", "691"],
      ["3", "LPN+LPB+LP", blue("259"), red("210"), green("220"), "149", blue("246"), "97", "36", "117%", "194"],
      ["", "合計", blue("1,294"), red("1,242"), green("922"), "625", blue("685"), "60", negRed("-557"), "55%", "1,168"]
    ];
  }

  // ============================================================================
  // 10 月真實出貨明細樞紐表範例 (100% 對照圖三：出貨金額與產品別)
  // ============================================================================
  function getSampleShipmentsGrid() {
    return [
      ["出貨金額(萬元)", "", "月份", "10", ""],
      ["營業員", "產品別", "10", "總計", ""],
      ["何宛茹-桃園", "LCA (三菱大型PLC)", "25.84", "25.84", ""],
      ["", "LCH (台製品端子臺/線材)", ".47", ".47", ""],
      ["", "LEA (三菱伺服)", "116.59", "116.59", ""],
      ["", "LEB (士林伺服)", "28.21", "28.21", ""],
      ["", "LER (減速機)", "37.20", "37.20", ""],
      ["", "LFB (PATLITE警示燈/蜂鳴器)", "2.55", "2.55", ""],
      ["", "LNH (士林變頻器)", "4.86", "4.86", ""],
      ["", "LNM (三菱變頻器)", "25.89", "25.89", ""],
      ["", "LPB (士林小型PLC)", ".30", ".30", ""],
      ["", "LPC (三菱小型PLC)", "2.49", "2.49", ""],
      ["", "LSA (松下感測器)", "15.13", "15.13", ""],
      ["何宛茹-桃園 合計", "", "259.53", "259.53", ""],
      ["周柏翰-桃園", "LPC (三菱小型PLC)", ".23", ".23", ""],
      ["周柏翰-桃園 合計", "", ".23", ".23", ""],
      ["邱文輝-台北", "LSA (松下感測器)", "5.03", "5.03", ""],
      ["邱文輝-台北 合計", "", "5.03", "5.03", ""],
      ["邱文輝-桃園", "LCA (三菱大型PLC)", "196.40", "196.40", ""],
      ["", "LCG (三菱人機)", "12.33", "12.33", ""],
      ["", "LEA (三菱伺服)", "283.96", "283.96", ""],
      ["", "LEB (士林伺服)", "3.65", "3.65", ""],
      ["", "LNH (士林變頻器)", "9.83", "9.83", ""],
      ["", "LNK (SD-INV/變頻器配件)", "1.93", "1.93", ""],
      ["", "LNM (三菱變頻器)", "70.36", "70.36", ""],
      ["", "LPB (士林小型PLC)", "11.55", "11.55", ""],
      ["", "LPC (三菱小型PLC)", "656.16", "656.16", ""],
      ["曾仁君-新竹", "LEB (士林伺服)", "65.87", "65.87", ""],
      ["", "LNH (士林變頻器)", "55.22", "55.22", ""],
      ["", "LPB (士林小型PLC)", ".06", ".06", ""],
      ["", "LPN (小型PLC_技提)", "8.97", "8.97", ""],
      ["", "LPC (三菱小型PLC)", "219.06", "219.06", ""],
      ["曾仁君-新竹 合計", "", "349.18", "349.18", ""]
    ];
  }

  // 輔助函式：提取儲存格內純文字
  function getCellStr(cell) {
    if (cell === null || cell === undefined) return "";
    if (typeof cell === "object") return (cell.val || "").toString().trim();
    return cell.toString().trim();
  }

  // 取得達成率紅綠燈狀態與漸進式多階層漸層色彩 (>=100% 翡翠綠, 80%~99.9% 金黃暖橙, 50%~79.9% 珊瑚橙紅, <50% 寶石艷紅)
  function getTrafficLight(rate) {
    const num = parseFloat(rate) || 0;
    if (num >= 100) {
      return { type: "success", barClass: "bar-success", text: "達標" };
    } else if (num >= 80) {
      return { type: "warning", barClass: "bar-warning", text: "預警" };
    } else if (num >= 50) {
      return { type: "alert", barClass: "bar-alert", text: "落後" };
    } else {
      return { type: "danger", barClass: "bar-danger", text: "嚴重落後" };
    }
  }

  // 計算特定業務員在出貨明細表中的重點自製品實績 (LEB、LNH、PLC)
  // 鐵律：嚴格排除三菱產品 (如三菱大型PLC LCA、三菱小型PLC LPC、三菱伺服 LEA、三菱變頻器 LNM)
  // 鐵律：自動跳過包含「合計」或「總計」的加總列，防止重複計算
  function calculateSalespersonProductShipments(shipmentsMatrix, salesName) {
    const result = { leb: 0, lnh: 0, plc: 0 };
    if (!shipmentsMatrix || !Array.isArray(shipmentsMatrix) || shipmentsMatrix.length === 0 || !salesName) {
      return result;
    }

    let currentSales = "";
    for (let r = 0; r < shipmentsMatrix.length; r++) {
      const row = shipmentsMatrix[r];
      const col0 = getCellStr(row[0]);
      const col1 = getCellStr(row[1]);
      const colAmtStr = getCellStr(row[2]) || getCellStr(row[3]);

      // 跳過表頭列、包含「合計」或「總計」的加總列，避免重複累計
      if (col0.includes("合計") || col0.includes("總計") || col1.includes("合計") || col1.includes("總計")) {
        continue;
      }

      if (col0 && !col0.includes("出貨金額") && !col0.includes("營業員")) {
        currentSales = col0;
      }

      // 當前列屬於該業務員 (比對業務員姓名，例如「曾仁君」)
      if (currentSales.includes(salesName)) {
        const amt = parseFloat(colAmtStr.replace(/,/g, "")) || 0;

        // 1. 士林伺服：僅認列 LEB (排除三菱 LEA)
        if (col1.includes("LEB") || (col1.includes("士林伺服") && !col1.includes("三菱"))) {
          result.leb += amt;
        }
        // 2. 士林變頻器：僅認列 LNH (排除三菱 LNM)
        else if (col1.includes("LNH") || (col1.includes("士林變頻器") && !col1.includes("三菱"))) {
          result.lnh += amt;
        }
        // 3. 士林PLC：僅認列 LPB (士林小型PLC)、LPN (小型PLC_技提)、LPD 等士林機種，嚴格排除 LPC (三菱小型PLC) 與 LCA (三菱大型PLC)
        else if (
          (col1.includes("LPB") || col1.includes("LPN") || col1.includes("LPD") || col1.includes("士林小型PLC") || col1.includes("小型PLC_技提") || (col1.includes("PLC") && col1.includes("士林"))) &&
          !col1.includes("LPC") && !col1.includes("LCA") && !col1.includes("三菱")
        ) {
          result.plc += amt;
        }
      }
    }
    return result;
  }

  // 解析目標表格矩陣為首頁戰情總覽指標 (精準定位：年計月目標、挑戰目標、堆疊目標)
  function parseTargetsGridToMetrics(matrix, shipmentsMatrix, activeSalesName) {
    if (!matrix || !Array.isArray(matrix) || matrix.length === 0) return null;

    const targetSales = activeSalesName || localStorage.getItem("shihlin_pwa_active_sales") || "曾維崧";
    const isManager = (targetSales === "曾維崧");

    const metrics = {
      isManager: isManager,
      salesName: targetSales,
      month: "10",
      annualTarget: "7,060",
      totalTarget: "9,300",
      stackTarget: "10,452",
      totalActual: "10,370",
      totalDiff: "1,070",
      totalRate: "112%",
      totalRateNum: 112,
      diffText: "超標 1,070 萬元",

      salesTarget: "810",
      salesActual: "802",
      salesDiff: "-8",
      salesRate: "99.0%",
      salesRateNum: 99.0,
      salesDiffText: "尚差 8 萬元",

      personalLeb: "0.00",
      personalLnh: "0.00",
      personalPlc: "0.00",

      lebTarget: "516",
      lebActual: "196",
      lebDiff: "-320",
      lebRate: "38%",
      lebRateNum: 38,
      lebDiffText: "差距 320 萬",

      lnhTarget: "516",
      lnhActual: "243",
      lnhDiff: "-273",
      lnhRate: "47%",
      lnhRateNum: 47,
      lnhDiffText: "差距 273 萬",

      plcTarget: "210",
      plcActual: "246",
      plcDiff: "36",
      plcRate: "117%",
      plcRateNum: 117,
      plcDiffText: "超標 36 萬"
    };

    try {
      // 1. 偵測月份與總月目標 (Row 0: 北區 | 10 | 月挑戰目標 | 9,300 | 萬)
      for (let r = 0; r < Math.min(3, matrix.length); r++) {
        const row = matrix[r];
        for (let c = 0; c < row.length; c++) {
          const txt = getCellStr(row[c]);
          if (txt.includes("月挑戰目標") || txt.includes("挑戰目標")) {
            const nextVal = getCellStr(row[c + 1]);
            if (nextVal && !isNaN(parseFloat(nextVal.replace(/,/g, "")))) {
              metrics.totalTarget = nextVal;
            }
          }
          if (/^\d{1,2}$/.test(txt) && c <= 2) {
            metrics.month = txt;
          }
        }
      }

      // 2. 尋找標題列 (鎖定包含 "營業員" 或同時包含 "年計" 與 "實績" 的欄位名稱列)
      let headerRowIdx = -1;
      for (let r = 0; r < Math.min(6, matrix.length); r++) {
        const rowStr = matrix[r].map(getCellStr).join(" ");
        if (rowStr.includes("營業員") || (rowStr.includes("年計") && (rowStr.includes("挑戰") || rowStr.includes("實績")))) {
          headerRowIdx = r;
          break;
        }
      }

      // 3. 尋找業務員總計列 (第一組「合計」列)
      let totalRowIdx = -1;
      for (let r = (headerRowIdx >= 0 ? headerRowIdx + 1 : 0); r < matrix.length; r++) {
        const row = matrix[r];
        if (row.some(cell => getCellStr(cell) === "合計")) {
          totalRowIdx = r;
          break;
        }
      }

      // 4. 定位標題欄位索引 (年計月目標、挑戰目標、堆疊目標、實績、達成率)
      let colAnnualTarget = -1;    // 年計月目標
      let colChallengeTarget = -1; // 挑戰目標 (忽略月份)
      let colStackTarget = -1;     // 堆疊目標 (忽略月份)
      let colActual = -1;          // 實績
      let colDiff = -1;            // 挑戰目標差異
      let colRate = -1;            // 達成率

      if (headerRowIdx >= 0) {
        const headerRow = matrix[headerRowIdx];
        headerRow.forEach((c, idx) => {
          const txt = getCellStr(c);
          if (txt.includes("年計")) {
            colAnnualTarget = idx;
          }
          if (txt.includes("挑戰目標") && !txt.includes("差異") && !txt.includes("達成率")) {
            colChallengeTarget = idx;
          }
          if ((txt.includes("堆疊") || txt.includes("堆疊目標") || txt.includes("堆疊業績")) && !txt.includes("差異")) {
            colStackTarget = idx;
          }
          if (txt.includes("實績") && !txt.includes("前一個月")) {
            colActual = idx;
          }
          if (txt.includes("達成率")) {
            colRate = idx;
          }
          if (txt.includes("挑戰目標差異") || (txt.includes("差異") && !txt.includes("堆疊"))) {
            colDiff = idx;
          }
        });
      }

      // 5. 從合計列依欄位精準提取數字 (北區/營運區數據)
      if (totalRowIdx >= 0) {
        const totRow = matrix[totalRowIdx];
        if (colAnnualTarget >= 0 && getCellStr(totRow[colAnnualTarget])) {
          metrics.annualTarget = getCellStr(totRow[colAnnualTarget]);
        }
        if (colChallengeTarget >= 0 && getCellStr(totRow[colChallengeTarget])) {
          metrics.totalTarget = getCellStr(totRow[colChallengeTarget]);
        }
        if (colStackTarget >= 0 && getCellStr(totRow[colStackTarget])) {
          metrics.stackTarget = getCellStr(totRow[colStackTarget]);
        }
        if (colActual >= 0 && getCellStr(totRow[colActual])) {
          metrics.totalActual = getCellStr(totRow[colActual]);
        }
        if (colDiff >= 0 && getCellStr(totRow[colDiff])) {
          metrics.totalDiff = getCellStr(totRow[colDiff]);
        }
        if (colRate >= 0 && getCellStr(totRow[colRate])) {
          metrics.totalRate = getCellStr(totRow[colRate]);
          metrics.totalRateNum = parseFloat(metrics.totalRate.replace(/%/g, "")) || 100;
        }

        const dNum = parseFloat((metrics.totalDiff || "0").replace(/,/g, ""));
        if (!isNaN(dNum)) {
          metrics.diffText = dNum >= 0 ? `超標 ${metrics.totalDiff} 萬元` : `差距 ${Math.abs(dNum)} 萬元`;
        }
      }

      // 6. 自製品三大項目 (LEB, LNH, PLC) - 北區/營運區戰況
      for (let r = (totalRowIdx > 0 ? totalRowIdx : 0); r < matrix.length; r++) {
        const row = matrix[r];
        const rowStr = row.map(getCellStr).join(" ");

        if (rowStr.includes("LEB")) {
          const target = (colChallengeTarget >= 0 && getCellStr(row[colChallengeTarget])) || "516";
          const actual = (colActual >= 0 && getCellStr(row[colActual])) || "196";
          const diff = (colDiff >= 0 && getCellStr(row[colDiff])) || "-320";
          const rate = (colRate >= 0 && getCellStr(row[colRate])) || "38%";
          metrics.lebTarget = target;
          metrics.lebActual = actual;
          metrics.lebRate = rate.includes("%") ? rate : `${rate}%`;
          metrics.lebRateNum = parseFloat(rate.replace(/%/g, "")) || 38;
          const dNum = parseFloat(diff.replace(/,/g, ""));
          metrics.lebDiffText = dNum >= 0 ? `超標 ${diff} 萬` : `差距 ${Math.abs(dNum)} 萬`;
        } else if (rowStr.includes("LNH")) {
          const target = (colChallengeTarget >= 0 && getCellStr(row[colChallengeTarget])) || "516";
          const actual = (colActual >= 0 && getCellStr(row[colActual])) || "243";
          const diff = (colDiff >= 0 && getCellStr(row[colDiff])) || "-273";
          const rate = (colRate >= 0 && getCellStr(row[colRate])) || "47%";
          metrics.lnhTarget = target;
          metrics.lnhActual = actual;
          metrics.lnhRate = rate.includes("%") ? rate : `${rate}%`;
          metrics.lnhRateNum = parseFloat(rate.replace(/%/g, "")) || 47;
          const dNum = parseFloat(diff.replace(/,/g, ""));
          metrics.lnhDiffText = dNum >= 0 ? `超標 ${diff} 萬` : `差距 ${Math.abs(dNum)} 萬`;
        } else if (rowStr.includes("LPN") || rowStr.includes("PLC")) {
          const target = (colChallengeTarget >= 0 && getCellStr(row[colChallengeTarget])) || "210";
          const actual = (colActual >= 0 && getCellStr(row[colActual])) || "246";
          const diff = (colDiff >= 0 && getCellStr(row[colDiff])) || "36";
          const rate = (colRate >= 0 && getCellStr(row[colRate])) || "117%";
          metrics.plcTarget = target;
          metrics.plcActual = actual;
          metrics.plcRate = rate.includes("%") ? rate : `${rate}%`;
          metrics.plcRateNum = parseFloat(rate.replace(/%/g, "")) || 117;
          const dNum = parseFloat(diff.replace(/,/g, ""));
          metrics.plcDiffText = dNum >= 0 ? `超標 ${diff} 萬` : `差距 ${Math.abs(dNum)} 萬`;
        }
      }

      // 7. 依據選定的業務員 (targetSales) 動態提取個人業績或主管北區業績
      if (isManager) {
        // 主管模式 (曾維崧)：卡片 1 直接以「營運區營收戰報」為主視角
        metrics.salesName = "曾維崧";
        metrics.salesTarget = metrics.totalTarget;
        metrics.salesActual = metrics.totalActual;
        metrics.salesDiff = metrics.totalDiff;
        metrics.salesRate = metrics.totalRate;
        metrics.salesRateNum = metrics.totalRateNum;
        metrics.salesDiffText = metrics.diffText;
        // 自製品三大卡片直接顯示營運區實績
        metrics.personalLeb = metrics.lebActual;
        metrics.personalLnh = metrics.lnhActual;
        metrics.personalPlc = metrics.plcActual;
      } else {
        // 一般業務員模式 (如葉仁豪、曾仁君、邱文輝等)：在總表中搜尋該業務員
        let foundSales = false;
        for (let r = (headerRowIdx >= 0 ? headerRowIdx + 1 : 0); r < (totalRowIdx > 0 ? totalRowIdx : matrix.length); r++) {
          const row = matrix[r];
          const rowStr = row.map(getCellStr).join(" ");
          if (rowStr.includes(targetSales)) {
            metrics.salesName = targetSales;
            foundSales = true;
            if (colChallengeTarget >= 0 && getCellStr(row[colChallengeTarget])) {
              metrics.salesTarget = getCellStr(row[colChallengeTarget]);
            }
            if (colActual >= 0 && getCellStr(row[colActual])) {
              metrics.salesActual = getCellStr(row[colActual]);
            }
            if (colRate >= 0 && getCellStr(row[colRate])) {
              metrics.salesRate = getCellStr(row[colRate]);
              metrics.salesRateNum = parseFloat(metrics.salesRate.replace(/%/g, "")) || 0;
            }
            if (colDiff >= 0 && getCellStr(row[colDiff])) {
              metrics.salesDiff = getCellStr(row[colDiff]);
            }

            const sT = parseFloat((metrics.salesTarget || "0").replace(/,/g, ""));
            const sA = parseFloat((metrics.salesActual || "0").replace(/,/g, ""));
            if (sT > 0) {
              if (!metrics.salesRate || metrics.salesRate === "0%") {
                const rVal = ((sA / sT) * 100).toFixed(1);
                metrics.salesRate = `${rVal}%`;
                metrics.salesRateNum = parseFloat(rVal) || 0;
              }
              const diff = metrics.salesDiff ? parseFloat(metrics.salesDiff.replace(/,/g, "")) : (sA - sT);
              metrics.salesDiffText = diff >= 0 ? `超標 ${Math.round(diff)} 萬元` : `尚差 ${Math.abs(Math.round(diff))} 萬元`;
            }
            break;
          }
        }

        if (!foundSales) {
          metrics.salesName = targetSales;
          metrics.salesTarget = "0";
          metrics.salesActual = "0";
          metrics.salesRate = "0%";
          metrics.salesRateNum = 0;
          metrics.salesDiffText = "尚無個人目標數據";
        }

        // 依業務員真實計算右表自製品個人出貨明細 (若右表無該員紀錄則真實回傳 0.00 萬)
        if (shipmentsMatrix) {
          const personalShipments = calculateSalespersonProductShipments(shipmentsMatrix, targetSales);
          metrics.personalLeb = personalShipments.leb.toFixed(2);
          metrics.personalLnh = personalShipments.lnh.toFixed(2);
          metrics.personalPlc = personalShipments.plc.toFixed(2);
        }
      }
    } catch (err) {
      console.warn("[Scoreboard Sync] 解析指標時發生輕微錯誤，使用安全預設值：", err);
    }

    return metrics;
  }

  // 將解析好的指標即時套用至首頁四大核心大數字戰情卡片與紅綠燈 (支援北區/營運區目標 vs 個人目標切換)
  function applyScoreboardMetrics(metrics, scope) {
    if (!metrics) return;
    const viewScope = scope || currentScoreboardScope || "region";

    const isRegion = (viewScope === "region");

    // 1. 頂部月度挑戰目標、實績與完成率串流 (標籤與數值依營運區/個人切換)
    const labelMainTarget = document.getElementById("labelMainTarget");
    const labelMainActual = document.getElementById("labelMainActual");
    const flowTargetVal = document.getElementById("flowTargetVal");
    const flowActualVal = document.getElementById("flowActualVal");
    const flowRateVal = document.getElementById("flowRateVal");

    const curRegion = getCurrentRegion();

    if (isRegion) {
      if (labelMainTarget) labelMainTarget.innerText = `本月${curRegion}目標：`;
      if (labelMainActual) labelMainActual.innerText = `${curRegion}實績：`;
      if (flowTargetVal) flowTargetVal.innerText = `${metrics.totalTarget || "9,300"}萬`;
      if (flowActualVal) flowActualVal.innerText = `${metrics.totalActual || "10,370"}萬`;
      if (flowRateVal) {
        const rateNum = parseFloat((metrics.totalRate || "112%").replace(/%/g, "")) || 0;
        const light = getTrafficLight(rateNum);
        flowRateVal.className = `flow-val traffic-light-badge ${light.type}`;
        flowRateVal.innerText = metrics.totalRate || "112%";
      }
    } else {
      if (labelMainTarget) labelMainTarget.innerText = "本月個人目標：";
      if (labelMainActual) labelMainActual.innerText = "個人實績：";
      if (flowTargetVal) flowTargetVal.innerText = `${metrics.salesTarget || "810"}萬`;
      if (flowActualVal) flowActualVal.innerText = `${metrics.salesActual || "802"}萬`;
      if (flowRateVal) {
        const rateNum = parseFloat((metrics.salesRate || "99.0%").replace(/%/g, "")) || 0;
        const light = getTrafficLight(rateNum);
        flowRateVal.className = `flow-val traffic-light-badge ${light.type}`;
        flowRateVal.innerText = metrics.salesRate || "99.0%";
      }
    }

    // 2. 卡片 1：總業績 (營運區總業績 / 個人總業績)
    const kpiCardTitlePersonal = document.getElementById("kpiCardTitlePersonal");
    const kpiValPersonal = document.getElementById("kpiValPersonal");
    const kpiTrafficPersonal = document.getElementById("kpiTrafficPersonal");
    const kpiDetailPersonal = document.getElementById("kpiDetailPersonal");
    const kpiBarPersonal = document.getElementById("kpiBarPersonal");

    if (isRegion) {
      if (kpiCardTitlePersonal) kpiCardTitlePersonal.innerText = `${curRegion}總業績`;
      if (kpiValPersonal) kpiValPersonal.innerText = metrics.totalActual || "10,370";
      const rateNum = metrics.totalRateNum || 112;
      const light = getTrafficLight(rateNum);
      if (kpiTrafficPersonal) {
        kpiTrafficPersonal.className = `traffic-light-badge ${light.type}`;
        kpiTrafficPersonal.innerText = metrics.totalRate || "112%";
      }
      if (kpiDetailPersonal) {
        const tVal = parseFloat((metrics.totalTarget || "9300").replace(/,/g, ""));
        const aVal = parseFloat((metrics.totalActual || "10370").replace(/,/g, ""));
        const diff = aVal - tVal;
        const dLabel = diff >= 0 ? "超標" : "差距";
        kpiDetailPersonal.innerHTML = `<span>目標：${metrics.totalTarget || "9,300"} 萬</span> <span>｜</span> <span>實績：${metrics.totalActual || "10,370"} 萬</span> <span>｜</span> <span>${dLabel}：${Math.abs(Math.round(diff))} 萬</span>`;
      }
      if (kpiBarPersonal) {
        kpiBarPersonal.className = `fluent-progress-bar ${light.barClass}`;
        kpiBarPersonal.style.width = `${Math.min(100, Math.max(0, rateNum))}%`;
      }
    } else {
      if (kpiCardTitlePersonal) kpiCardTitlePersonal.innerText = "個人總業績";
      if (kpiValPersonal) kpiValPersonal.innerText = metrics.salesActual || "802";
      const ratePersonalNum = metrics.salesRateNum !== undefined ? metrics.salesRateNum : (parseFloat((metrics.salesRate || "0").replace(/%/g, "")) || 0);
      const lightPersonal = getTrafficLight(ratePersonalNum);
      if (kpiTrafficPersonal) {
        kpiTrafficPersonal.className = `traffic-light-badge ${lightPersonal.type}`;
        kpiTrafficPersonal.innerText = metrics.salesRate || "99.0%";
      }
      if (kpiDetailPersonal) {
        const sT = parseFloat((metrics.salesTarget || "0").replace(/,/g, ""));
        const sA = parseFloat((metrics.salesActual || "0").replace(/,/g, ""));
        const diffVal = metrics.salesDiff ? parseFloat(metrics.salesDiff.replace(/,/g, "")) : (sA - sT);
        const diffLabel = diffVal >= 0 ? "超標" : "差距";
        kpiDetailPersonal.innerHTML = `<span>目標：${metrics.salesTarget || "810"} 萬</span> <span>｜</span> <span>實績：${metrics.salesActual || "802"} 萬</span> <span>｜</span> <span>${diffLabel}：${Math.abs(Math.round(diffVal))} 萬</span>`;
      }
      if (kpiBarPersonal) {
        kpiBarPersonal.className = `fluent-progress-bar ${lightPersonal.barClass}`;
        kpiBarPersonal.style.width = `${Math.min(100, Math.max(0, ratePersonalNum))}%`;
      }
    }

    // 3. 卡片 2：士林伺服
    const kpiValLEB = document.getElementById("kpiValLEB");
    const kpiTrafficLEB = document.getElementById("kpiTrafficLEB");
    const kpiDetailLEB = document.getElementById("kpiDetailLEB");
    const kpiBarLEB = document.getElementById("kpiBarLEB");

    if (isRegion) {
      if (kpiValLEB) kpiValLEB.innerText = metrics.lebActual || "196";
      const rateLEBNum = metrics.lebRateNum || 38;
      const lightLEB = getTrafficLight(rateLEBNum);
      if (kpiTrafficLEB) {
        kpiTrafficLEB.style.display = isDriverViewActive ? "none" : "";
        kpiTrafficLEB.className = `traffic-light-badge ${lightLEB.type}`;
        kpiTrafficLEB.innerText = metrics.lebRate || "38%";
      }
      if (kpiDetailLEB) {
        const diffNum = parseFloat((metrics.lebDiff || "-320").replace(/,/g, ""));
        const diffLabel = diffNum >= 0 ? "超標" : "差距";
        kpiDetailLEB.innerHTML = `<span>${curRegion}目標：${metrics.lebTarget || "516"} 萬</span> <span>｜</span> <span>${curRegion}實績：${metrics.lebActual || "196"} 萬</span> <span>｜</span> <span>${diffLabel}：${Math.abs(Math.round(diffNum))} 萬</span>`;
      }
      if (kpiBarLEB) {
        kpiBarLEB.className = `fluent-progress-bar ${lightLEB.barClass}`;
        kpiBarLEB.style.width = `${Math.min(100, Math.max(0, rateLEBNum))}%`;
      }
    } else {
      if (kpiValLEB) kpiValLEB.innerText = metrics.personalLeb || "0.00";
      const pVal = parseFloat(metrics.personalLeb || "0");
      if (kpiTrafficLEB) {
        // 個人視角下隱藏右上標籤，保持畫面乾淨不干擾
        kpiTrafficLEB.style.display = "none";
      }
      if (kpiDetailLEB) {
        kpiDetailLEB.innerHTML = `<span>個人出貨：<strong>${metrics.personalLeb || "0.00"}</strong> 萬</span> <span>｜</span> <span>${curRegion}挑戰目標：${metrics.lebTarget || "516"} 萬</span>`;
      }
      if (kpiBarLEB) {
        kpiBarLEB.className = `fluent-progress-bar bar-success`;
        kpiBarLEB.style.width = pVal > 0 ? `${Math.min(100, Math.round((pVal / 100) * 100))}%` : "0%";
      }
    }

    // 4. 卡片 3：士林變頻器
    const kpiValLNH = document.getElementById("kpiValLNH");
    const kpiTrafficLNH = document.getElementById("kpiTrafficLNH");
    const kpiDetailLNH = document.getElementById("kpiDetailLNH");
    const kpiBarLNH = document.getElementById("kpiBarLNH");

    if (isRegion) {
      if (kpiValLNH) kpiValLNH.innerText = metrics.lnhActual || "243";
      const rateLNHNum = metrics.lnhRateNum || 47;
      const lightLNH = getTrafficLight(rateLNHNum);
      if (kpiTrafficLNH) {
        kpiTrafficLNH.style.display = isDriverViewActive ? "none" : "";
        kpiTrafficLNH.className = `traffic-light-badge ${lightLNH.type}`;
        kpiTrafficLNH.innerText = metrics.lnhRate || "47%";
      }
      if (kpiDetailLNH) {
        const diffNum = parseFloat((metrics.lnhDiff || "-273").replace(/,/g, ""));
        const diffLabel = diffNum >= 0 ? "超標" : "差距";
        kpiDetailLNH.innerHTML = `<span>${curRegion}目標：${metrics.lnhTarget || "516"} 萬</span> <span>｜</span> <span>${curRegion}實績：${metrics.lnhActual || "243"} 萬</span> <span>｜</span> <span>${diffLabel}：${Math.abs(diffNum)} 萬</span>`;
      }
      if (kpiBarLNH) {
        kpiBarLNH.className = `fluent-progress-bar ${lightLNH.barClass}`;
        kpiBarLNH.style.width = `${Math.min(100, Math.max(0, rateLNHNum))}%`;
      }
    } else {
      if (kpiValLNH) kpiValLNH.innerText = metrics.personalLnh || "0.00";
      const pVal = parseFloat(metrics.personalLnh || "0");
      if (kpiTrafficLNH) {
        // 個人視角下隱藏右上標籤，保持畫面乾淨不干擾
        kpiTrafficLNH.style.display = "none";
      }
      if (kpiDetailLNH) {
        kpiDetailLNH.innerHTML = `<span>個人出貨：<strong>${metrics.personalLnh || "0.00"}</strong> 萬</span> <span>｜</span> <span>${curRegion}挑戰目標：${metrics.lnhTarget || "516"} 萬</span>`;
      }
      if (kpiBarLNH) {
        kpiBarLNH.className = `fluent-progress-bar bar-success`;
        kpiBarLNH.style.width = pVal > 0 ? `${Math.min(100, Math.round((pVal / 100) * 100))}%` : "0%";
      }
    }

    // 5. 卡片 4：士林PLC
    const kpiValPLC = document.getElementById("kpiValPLC");
    const kpiTrafficPLC = document.getElementById("kpiTrafficPLC");
    const kpiDetailPLC = document.getElementById("kpiDetailPLC");
    const kpiBarPLC = document.getElementById("kpiBarPLC");

    if (isRegion) {
      if (kpiValPLC) kpiValPLC.innerText = metrics.plcActual || "246";
      const ratePLCNum = metrics.plcRateNum || 117;
      const lightPLC = getTrafficLight(ratePLCNum);
      if (kpiTrafficPLC) {
        kpiTrafficPLC.style.display = isDriverViewActive ? "none" : "";
        kpiTrafficPLC.className = `traffic-light-badge ${lightPLC.type}`;
        kpiTrafficPLC.innerText = metrics.plcRate || "117%";
      }
      if (kpiDetailPLC) {
        const diffNum = parseFloat((metrics.plcDiff || "36").replace(/,/g, ""));
        const diffLabel = diffNum >= 0 ? "超標" : "差距";
        kpiDetailPLC.innerHTML = `<span>${curRegion}目標：${metrics.plcTarget || "210"} 萬</span> <span>｜</span> <span>${curRegion}實績：${metrics.plcActual || "246"} 萬</span> <span>｜</span> <span>${diffLabel}：${Math.abs(diffNum)} 萬</span>`;
      }
      if (kpiBarPLC) {
        kpiBarPLC.className = `fluent-progress-bar ${lightPLC.barClass}`;
        kpiBarPLC.style.width = `${Math.min(100, Math.max(0, ratePLCNum))}%`;
      }
    } else {
      if (kpiValPLC) kpiValPLC.innerText = metrics.personalPlc || "0.00";
      const pVal = parseFloat(metrics.personalPlc || "0");
      if (kpiTrafficPLC) {
        // 個人視角下隱藏右上標籤，保持畫面乾淨不干擾
        kpiTrafficPLC.style.display = "none";
      }
      if (kpiDetailPLC) {
        kpiDetailPLC.innerHTML = `<span>個人出貨：<strong>${metrics.personalPlc || "0.00"}</strong> 萬</span> <span>｜</span> <span>${curRegion}挑戰目標：${metrics.plcTarget || "210"} 萬</span>`;
      }
      if (kpiBarPLC) {
        kpiBarPLC.className = `fluent-progress-bar bar-success`;
        kpiBarPLC.style.width = pVal > 0 ? `${Math.min(100, Math.round((pVal / 100) * 100))}%` : "0%";
      }
    }

    // 6. 頂部導覽列業務員名稱與頭像即時連動
    const currentSalesDisplay = document.getElementById("currentSalesDisplay");
    const headerUserAvatar = document.getElementById("headerUserAvatar");
    if (currentSalesDisplay && metrics.salesName) {
      currentSalesDisplay.innerText = metrics.salesName;
    }
    if (headerUserAvatar && metrics.salesName) {
      headerUserAvatar.innerText = metrics.salesName.charAt(0) || "曾";
    }

    // 7. 同步更新系統設定下拉選單
    const selectActiveSales = document.getElementById("selectActiveSales");
    if (selectActiveSales && metrics.salesName && selectActiveSales.value !== metrics.salesName) {
      selectActiveSales.value = metrics.salesName;
    }

    // 8. 載入並渲染三大產品線領先指標 (驅動目標) 與審核面板
    loadAndRenderDriverMetrics(metrics.salesName);
    updateGlobalDriverView(isDriverViewActive);
  }

  // 左表載入範例與清空
  if (btnPasteTargetDemo) {
    btnPasteTargetDemo.addEventListener("click", () => {
      const targetsGrid = getSampleTargetsGrid();
      renderExcelGrid(tablePastedTargets, targetsGrid, 20, 11, badgeTargetCount);
      localStorage.setItem("shihlin_pasted_targets_raw", JSON.stringify(targetsGrid));
    });
  }
  if (btnClearTargetTable) {
    btnClearTargetTable.addEventListener("click", () => {
      renderExcelGrid(tablePastedTargets, [], 20, 11, badgeTargetCount);
      localStorage.removeItem("shihlin_pasted_targets_raw");
      clearRangeSelection();
      clearRowColSelection(tablePastedTargets);
    });
  }

  // 右表載入範例與清空
  if (btnPasteShipmentDemo) {
    btnPasteShipmentDemo.addEventListener("click", () => {
      const shipmentsGrid = getSampleShipmentsGrid();
      renderExcelGrid(tablePastedShipments, shipmentsGrid, 28, 5, badgeShipmentCount);
      localStorage.setItem("shihlin_pasted_shipments_raw", JSON.stringify(shipmentsGrid));
    });
  }
  if (btnClearShipmentTable) {
    btnClearShipmentTable.addEventListener("click", () => {
      renderExcelGrid(tablePastedShipments, [], 25, 5, badgeShipmentCount);
      localStorage.removeItem("shihlin_pasted_shipments_raw");
      clearRangeSelection();
      clearRowColSelection(tablePastedShipments);
    });
  }

  // 頂部「載入範例」按鈕
  if (btnLoadSampleExcel) {
    btnLoadSampleExcel.addEventListener("click", () => {
      const targetsGrid = getSampleTargetsGrid();
      const shipmentsGrid = getSampleShipmentsGrid();
      renderExcelGrid(tablePastedTargets, targetsGrid, 20, 11, badgeTargetCount);
      renderExcelGrid(tablePastedShipments, shipmentsGrid, 28, 5, badgeShipmentCount);
      localStorage.setItem("shihlin_pasted_targets_raw", JSON.stringify(targetsGrid));
      localStorage.setItem("shihlin_pasted_shipments_raw", JSON.stringify(shipmentsGrid));
      alert("已成功載入最新 10 月範例數據！\n可點擊最下方「資料確認」同步更新首頁戰情看板。");
    });
  }

  // 頂部「全部清空」按鈕：即時重設兩張表格、清空狀態標籤與暫存
  if (btnResetAdminDefaults) {
    btnResetAdminDefaults.addEventListener("click", () => {
      renderExcelGrid(tablePastedTargets, [], 20, 11, badgeTargetCount);
      renderExcelGrid(tablePastedShipments, [], 25, 5, badgeShipmentCount);
      localStorage.removeItem("shihlin_pasted_targets_raw");
      localStorage.removeItem("shihlin_pasted_shipments_raw");
      clearRangeSelection();
      clearRowColSelection(tablePastedTargets);
      clearRowColSelection(tablePastedShipments);
      alert("已成功全部清空！兩張表格已重設為乾淨的空白試算表格子。");
    });
  }

  // 最下方「資料確認」按鈕：解析表格數據並同步連動首頁戰情總覽
  if (btnConfirmTablesReady) {
    btnConfirmTablesReady.addEventListener("click", () => {
      const activeSales = localStorage.getItem("shihlin_pwa_active_sales") || "曾維崧";
      const currentGrid = (tablePastedTargets && tablePastedTargets._gridMatrix) || [];
      const currentShipments = (tablePastedShipments && tablePastedShipments._gridMatrix) || [];
      const metrics = parseTargetsGridToMetrics(currentGrid, currentShipments, activeSales);
      if (metrics) {
        applyScoreboardMetrics(metrics);
        alert("資料確認完成！\n首頁戰情總覽之四大核心大數字與紅綠燈狀態已成功同步連動！");
      } else {
        alert("資料已確認！");
      }
    });
  }

  // 初始載入左右雙試算表格子
  const savedTargetsRaw = localStorage.getItem("shihlin_pasted_targets_raw");
  renderExcelGrid(tablePastedTargets, savedTargetsRaw || "", 20, 11, badgeTargetCount);

  const savedShipmentsRaw = localStorage.getItem("shihlin_pasted_shipments_raw");
  renderExcelGrid(tablePastedShipments, savedShipmentsRaw || "", 25, 5, badgeShipmentCount);

  // 初始載入戰情看板指標 (直接依據當前儲存表格與業務身份動態渲染，不殘留過期快取)
  renderHomeDashboard();

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
      const activeSales = localStorage.getItem("shihlin_pwa_active_sales") || "曾維崧";
      const resp = await fetch(`/api/customers?sales_name=${encodeURIComponent(activeSales)}`);
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
      dealsCache = Array.isArray(deals) ? deals : [];

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

  // ============================================================================
  // 6. 產品線領先指標 (驅動目標) 與變更審核管理
  // ============================================================================

  // (DEFAULT_MANAGERS & manager helpers moved to top of file to avoid TDZ)

  function getManagerTitle(name) {
    if (name === "曾維崧") return `FA ${getCurrentRegion()}主管`;
    if (name === "曾仁君") return "直賣業務主管";
    if (name === "張何達") return "經銷業務主管";
    return "業務主管";
  }

  // 預設種子資料快取 (相容於離線或純靜態 GitHub Pages 託管環境)
  const LOCAL_SEED_DRIVERS = {
    "曾維崧": [
      { id: 1, sales_name: "曾維崧", product_line: "士林伺服", driver_name: "伺服關鍵客群專案洽談", target_val: 6, actual_val: 5, unit: "場", period_type: "月", start_date: "2026-09-17", days_active: 35, weeks_active: 5.0, can_change: true, days_remaining: 0, rate_pct: 83.3, traffic_light: "yellow", light_text: "預警", bar_class: "bar-warning" },
      { id: 2, sales_name: "曾維崧", product_line: "士林變頻器", driver_name: "大型廠區變頻節能案場勘查", target_val: 8, actual_val: 7, unit: "場", period_type: "月", start_date: "2026-09-17", days_active: 35, weeks_active: 5.0, can_change: true, days_remaining: 0, rate_pct: 87.5, traffic_light: "yellow", light_text: "預警", bar_class: "bar-warning" },
      { id: 3, sales_name: "曾維崧", product_line: "士林PLC", driver_name: "AX5S 高階 PLC 連線展示", target_val: 5, actual_val: 5, unit: "場", period_type: "月", start_date: "2026-09-17", days_active: 35, weeks_active: 5.0, can_change: true, days_remaining: 0, rate_pct: 100.0, traffic_light: "green", light_text: "達標", bar_class: "bar-success" }
    ],
    "曾仁君": [
      { id: 4, sales_name: "曾仁君", product_line: "士林伺服", driver_name: "半導體設備伺服實機試機", target_val: 4, actual_val: 3, unit: "場", period_type: "週", start_date: "2026-09-17", days_active: 35, weeks_active: 5.0, can_change: true, days_remaining: 0, rate_pct: 75.0, traffic_light: "yellow", light_text: "預警", bar_class: "bar-warning" },
      { id: 5, sales_name: "曾仁君", product_line: "士林變頻器", driver_name: "直賣變頻器升級汰換拜訪", target_val: 6, actual_val: 5, unit: "次", period_type: "週", start_date: "2026-09-17", days_active: 35, weeks_active: 5.0, can_change: true, days_remaining: 0, rate_pct: 83.3, traffic_light: "yellow", light_text: "預警", bar_class: "bar-warning" },
      { id: 6, sales_name: "曾仁君", product_line: "士林PLC", driver_name: "專案自動化產線 PLC 規格制定", target_val: 3, actual_val: 3, unit: "案", period_type: "週", start_date: "2026-09-17", days_active: 35, weeks_active: 5.0, can_change: true, days_remaining: 0, rate_pct: 100.0, traffic_light: "green", light_text: "達標", bar_class: "bar-success" }
    ],
    "張何達": [
      { id: 7, sales_name: "張何達", product_line: "士林伺服", driver_name: "經銷商伺服現貨在庫拜訪", target_val: 8, actual_val: 8, unit: "家", period_type: "週", start_date: "2026-09-17", days_active: 35, weeks_active: 5.0, can_change: true, days_remaining: 0, rate_pct: 100.0, traffic_light: "green", light_text: "達標", bar_class: "bar-success" },
      { id: 8, sales_name: "張何達", product_line: "士林變頻器", driver_name: "盤廠經銷通路變頻專案盤點", target_val: 10, actual_val: 9, unit: "家", period_type: "週", start_date: "2026-09-17", days_active: 35, weeks_active: 5.0, can_change: true, days_remaining: 0, rate_pct: 90.0, traffic_light: "green", light_text: "達標", bar_class: "bar-success" },
      { id: 9, sales_name: "張何達", product_line: "士林PLC", driver_name: "經銷商推廣 AX 系列技術交流", target_val: 6, actual_val: 6, unit: "場", period_type: "週", start_date: "2026-09-17", days_active: 35, weeks_active: 5.0, can_change: true, days_remaining: 0, rate_pct: 100.0, traffic_light: "green", light_text: "達標", bar_class: "bar-success" }
    ],
    "葉仁豪": [
      { id: 10, sales_name: "葉仁豪", product_line: "士林伺服", driver_name: "伺服新樣品送樣驗證", target_val: 4, actual_val: 1, unit: "組", period_type: "週", start_date: "2026-10-08", days_active: 14, weeks_active: 2.0, can_change: false, days_remaining: 14, rate_pct: 25.0, traffic_light: "red", light_text: "嚴重落後", bar_class: "bar-danger" },
      { id: 11, sales_name: "葉仁豪", product_line: "士林變頻器", driver_name: "變頻器舊換新推廣電訪", target_val: 12, actual_val: 10, unit: "通", period_type: "週", start_date: "2026-10-08", days_active: 14, weeks_active: 2.0, can_change: false, days_remaining: 14, rate_pct: 83.3, traffic_light: "yellow", light_text: "預警", bar_class: "bar-warning" },
      { id: 12, sales_name: "葉仁豪", product_line: "士林PLC", driver_name: "PLC 搭配變頻配套展示", target_val: 4, actual_val: 3, unit: "次", period_type: "週", start_date: "2026-10-08", days_active: 14, weeks_active: 2.0, can_change: false, days_remaining: 14, rate_pct: 75.0, traffic_light: "yellow", light_text: "預警", bar_class: "bar-warning" }
    ],
    "何宛茹": [
      { id: 13, sales_name: "何宛茹", product_line: "士林伺服", driver_name: "直賣伺服節能設備實機測試", target_val: 4, actual_val: 3, unit: "場", period_type: "週", start_date: "2026-09-20", days_active: 32, weeks_active: 4.6, can_change: true, days_remaining: 0, rate_pct: 75.0, traffic_light: "yellow", light_text: "預警", bar_class: "bar-warning" },
      { id: 14, sales_name: "何宛茹", product_line: "士林變頻器", driver_name: "新竹設備廠變頻器推廣洽談", target_val: 6, actual_val: 5, unit: "家", period_type: "週", start_date: "2026-09-20", days_active: 32, weeks_active: 4.6, can_change: true, days_remaining: 0, rate_pct: 83.3, traffic_light: "yellow", light_text: "預警", bar_class: "bar-warning" },
      { id: 15, sales_name: "何宛茹", product_line: "士林PLC", driver_name: "包裝機設備 PLC 升級專案", target_val: 2, actual_val: 2, unit: "案", period_type: "週", start_date: "2026-09-20", days_active: 32, weeks_active: 4.6, can_change: true, days_remaining: 0, rate_pct: 100.0, traffic_light: "green", light_text: "達標", bar_class: "bar-success" }
    ]
  };

  const LOCAL_SEED_REQUESTS = [
    {
      id: 1,
      driver_id: 14,
      sales_name: "何宛茹",
      product_line: "士林變頻器",
      old_driver_name: "新竹設備廠變頻器推廣洽談",
      new_driver_name: "半導體耗電產線大馬力變頻替換",
      new_target_val: 5.0,
      new_unit: "場",
      change_reason: "原推廣洽談轉化率佳，配合節能政策聚焦大馬力機型實地能耗檢測。",
      apply_date: "2026-10-22",
      review_status: "待審核"
    }
  ];

  function getLocalDriverMetrics(salesName) {
    const key = "shihlin_pwa_driver_metrics";
    let store = null;
    try {
      const raw = localStorage.getItem(key);
      if (raw) store = JSON.parse(raw);
    } catch (e) {
      console.warn("讀取本地驅動指標失敗：", e);
    }
    if (!store) {
      store = { ...LOCAL_SEED_DRIVERS };
      localStorage.setItem(key, JSON.stringify(store));
    }
    if (!store[salesName]) {
      store[salesName] = [
        { id: Date.now(), sales_name: salesName, product_line: "士林伺服", driver_name: "伺服客群實地拜訪與技術交流", target_val: 4, actual_val: 3, unit: "場", period_type: "週", start_date: "2026-09-20", days_active: 32, weeks_active: 4.6, can_change: true, days_remaining: 0, rate_pct: 75.0, traffic_light: "yellow", light_text: "預警", bar_class: "bar-warning" },
        { id: Date.now() + 1, sales_name: salesName, product_line: "士林變頻器", driver_name: "變頻器節能案場勘查與推廣", target_val: 6, actual_val: 5, unit: "次", period_type: "週", start_date: "2026-09-20", days_active: 32, weeks_active: 4.6, can_change: true, days_remaining: 0, rate_pct: 83.3, traffic_light: "yellow", light_text: "預警", bar_class: "bar-warning" },
        { id: Date.now() + 2, sales_name: salesName, product_line: "士林PLC", driver_name: "PLC 與自動化連線架構制定", target_val: 3, actual_val: 2, unit: "案", period_type: "週", start_date: "2026-09-20", days_active: 32, weeks_active: 4.6, can_change: true, days_remaining: 0, rate_pct: 66.7, traffic_light: "red", light_text: "低標", bar_class: "bar-alert" }
      ];
      localStorage.setItem(key, JSON.stringify(store));
    }
    return store[salesName];
  }

  function getLocalDriverRequests() {
    const key = "shihlin_pwa_driver_requests";
    try {
      const raw = localStorage.getItem(key);
      if (raw) return JSON.parse(raw);
    } catch (e) {
      console.warn("讀取本地申請單失敗：", e);
    }
    const reqs = [...LOCAL_SEED_REQUESTS];
    localStorage.setItem(key, JSON.stringify(reqs));
    return reqs;
  }

  function saveLocalDriverRequests(reqs) {
    localStorage.setItem("shihlin_pwa_driver_requests", JSON.stringify(reqs));
  }

  function saveLocalDriverMetrics(salesName, list) {
    const key = "shihlin_pwa_driver_metrics";
    let store = {};
    try {
      const raw = localStorage.getItem(key);
      if (raw) store = JSON.parse(raw);
    } catch (e) {}
    store[salesName] = list;
    localStorage.setItem(key, JSON.stringify(store));
  }

  let currentViewingDriverSales = null;

  // 業務人員快速切換膠囊按鈕列 (主管巡視 10 位同仁 / 一般業務鎖定個人)
  function renderDriverSalesFilterBar(viewingSales, loggedInSales, isManager) {
    const container = document.getElementById("driverSalesFilterBar");
    const hint = document.getElementById("hintDriverFilterRole");
    if (!container) return;

    const allSales = getOrderedSalespersons();

    if (hint) {
      if (isManager) {
        hint.innerText = "主管模式：點選任一同仁姓名，即可立即巡視其專屬三大產品線指標";
        hint.style.color = "var(--color-accent)";
      } else {
        hint.innerText = "一般業務模式：已鎖定個人專屬指標檢視";
        hint.style.color = "var(--text-tertiary)";
      }
    }

    container.innerHTML = allSales.map(s => {
      const isSelected = (s.name === viewingSales);
      const isSelf = (s.name === loggedInSales);
      let classes = "driver-sales-chip";
      let title = "";

      if (isSelected) classes += " active";
      if (!isManager && !isSelf) {
        classes += " disabled-other";
        title = "一般業務僅可檢視個人指標";
      }

      return `
        <button type="button" class="${classes}" data-sales="${escapeHtml(s.name)}" title="${title}">
          <span>${escapeHtml(s.name)}</span>
          ${isSelf ? '<span style="font-size:0.7rem; opacity:0.85;">(我)</span>' : ''}
        </button>
      `;
    }).join("");

    container.querySelectorAll(".driver-sales-chip").forEach(btn => {
      btn.onclick = () => {
        const target = btn.getAttribute("data-sales");
        if (!isManager && target !== loggedInSales) return;
        renderDriverManagementPage(target);
      };
    });
  }

  // 渲染領先指標管理與審核專頁
  async function renderDriverManagementPage(targetSalesName) {
    const curSales = localStorage.getItem("shihlin_pwa_active_sales") || "曾維崧";
    const isMgr = isSalesManager(curSales);
    const viewingSales = targetSalesName || curSales;
    currentViewingDriverSales = viewingSales;

    // 1. 渲染切換膠囊列
    renderDriverSalesFilterBar(viewingSales, curSales, isMgr);

    // 2. 更新頂部徽章與卡片標題
    const badgeViewing = document.getElementById("badgeCurrentDriverViewingSales");
    if (badgeViewing) badgeViewing.innerText = viewingSales;

    const labelTitle = document.getElementById("labelDriverCardsTitle");
    if (labelTitle) {
      labelTitle.innerText = `${viewingSales} 的三大產品線領先指標 (驅動目標)`;
    }

    // 3. 向後端請求該業務員的領先指標
    let drivers = [];
    try {
      const res = await fetch(`/api/scoreboard/drivers?sales_name=${encodeURIComponent(viewingSales)}`);
      if (res.ok) {
        const data = await res.json();
        if (data.status === "success" && Array.isArray(data.drivers)) {
          drivers = data.drivers;
        }
      }
    } catch (err) {
      console.log("[Drivers] 後端連線備援，切換至本地快取：", err.message);
    }

    if (!drivers || drivers.length === 0) {
      drivers = getLocalDriverMetrics(viewingSales);
    }

    // 4. 渲染 3 張產品線管理卡片
    renderSettingsDriverManagement(viewingSales, drivers);

    // 5. 渲染主管審核專區
    renderManagerReviewPanel(curSales);
    updateManagerUIVisibility(curSales);
  }

  // 載入並渲染驅動目標 (首頁戰情看板專用)
  async function loadAndRenderDriverMetrics(salesName) {
    const sName = salesName || localStorage.getItem("shihlin_pwa_active_sales") || "曾維崧";
    let drivers = [];

    // 優先向後端 API 請求，若離線或 404 則無縫由本地儲存備援
    try {
      const res = await fetch(`/api/scoreboard/drivers?sales_name=${encodeURIComponent(sName)}`);
      if (res.ok) {
        const data = await res.json();
        if (data.status === "success" && Array.isArray(data.drivers)) {
          drivers = data.drivers;
        }
      }
    } catch (err) {
      console.log("[Drivers] 後端連線備援，切換至本地快取：", err.message);
    }

    if (!drivers || drivers.length === 0) {
      drivers = getLocalDriverMetrics(sName);
    }

    // 渲染三大產品線折疊抽屜 (首頁)
    renderDriverDrawer("士林伺服", "drawerDriverLEB", drivers);
    renderDriverDrawer("士林變頻器", "drawerDriverLNH", drivers);
    renderDriverDrawer("士林PLC", "drawerDriverPLC", drivers);

    // 同步渲染管理頁面
    const viewingSales = currentViewingDriverSales || sName;
    renderDriverManagementPage(viewingSales);
  }

  // 渲染單一抽屜內容 (首頁三大產品線領先指標卡片)
  function renderDriverDrawer(productLine, drawerId, drivers) {
    const drawerEl = document.getElementById(drawerId);
    if (!drawerEl) return;

    const d = drivers.find(item => item.product_line === productLine);
    if (!d) {
      drawerEl.innerHTML = `<div style="font-size:0.75rem; color:var(--text-tertiary); padding:4px 0;">暫無此產品線驅動指標</div>`;
      return;
    }

    const lightClass = d.traffic_light === "green" ? "success" : (d.traffic_light === "yellow" ? "warning" : (d.rate_pct >= 50 ? "alert" : "danger"));
    const barGradientClass = d.bar_class || (d.traffic_light === "green" ? "bar-success" : (d.traffic_light === "yellow" ? "bar-warning" : (d.rate_pct >= 50 ? "bar-alert" : "bar-danger")));

    drawerEl.innerHTML = `
      <div class="driver-item-card">
        <div class="driver-header-row">
          <div class="driver-title-group">
            <svg class="icon-svg" style="width:14px; height:14px; color:var(--color-accent); flex-shrink:0;" viewBox="0 0 24 24">
              <path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm-2 15l-5-5 1.41-1.41L10 14.17l7.59-7.59L19 8l-9 9z"/>
            </svg>
            <span class="driver-name" title="${escapeHtml(d.driver_name)}">${escapeHtml(d.driver_name)}</span>
          </div>
          <span class="traffic-light-badge ${lightClass}" style="font-size:0.72rem; padding:1px 6px;">
            ${d.rate_pct}% (${escapeHtml(d.light_text || '進度')})
          </span>
        </div>
        <div class="driver-metrics-row">
          <span>目標：<strong>${d.target_val}</strong> ${escapeHtml(d.unit)}</span>
          <span>實績：<strong>${d.actual_val}</strong> ${escapeHtml(d.unit)}</span>
          <span>週期：${escapeHtml(d.period_type || '週')}</span>
        </div>
        <div class="driver-progress-bg">
          <div class="driver-progress-bar ${barGradientClass}" style="width: ${Math.min(100, Math.max(0, d.rate_pct))}%;"></div>
        </div>
        <div class="driver-footer-row">
          <span>始於：${escapeHtml(d.start_date || '')}</span>
          <span>${d.can_change ? '<span style="color:#059669; font-weight:600;">已執行滿 4 週 (可微調)</span>' : `<span style="color:#dc2626; font-weight:600;">已執行 ${d.days_active} 天 (尚需 ${d.days_remaining} 天滿 4 週)</span>`}</span>
        </div>
        <div class="driver-card-action-bar">
          <button type="button" class="btn btn-primary btn-sm btn-report-driver-actual"
            data-id="${d.id}" data-product="${escapeHtml(d.product_line)}" data-name="${escapeHtml(d.driver_name)}"
            data-target="${d.target_val}" data-actual="${d.actual_val}" data-unit="${escapeHtml(d.unit)}">
            <svg class="icon-svg" style="width:13px; height:13px;" viewBox="0 0 24 24"><path d="M19 3H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zm-2 10h-4v4h-2v-4H7v-2h4V7h2v4h4v2z"/></svg>
            <span>回報實績</span>
          </button>
          <button type="button" class="btn btn-secondary btn-sm btn-edit-driver-settings"
            data-id="${d.id}" data-product="${escapeHtml(d.product_line)}" data-name="${escapeHtml(d.driver_name)}"
            data-target="${d.target_val}" data-unit="${escapeHtml(d.unit)}" data-period="${escapeHtml(d.period_type || '週')}">
            <svg class="icon-svg" style="width:13px; height:13px;" viewBox="0 0 24 24"><path d="M3 17.25V21h3.75L17.81 9.94l-3.75-3.75L3 17.25zM20.71 7.04c.39-.39.39-1.02 0-1.41l-2.34-2.34c-.39-.39-1.02-.39-1.41 0l-1.83 1.83 3.75 3.75 1.83-1.83z"/></svg>
            <span>設定指標</span>
          </button>
          ${d.can_change ? `
            <button type="button" class="btn btn-secondary btn-sm btn-driver-apply-home btn-change-span"
              data-id="${d.id}" data-product="${escapeHtml(d.product_line)}" data-name="${escapeHtml(d.driver_name)}">
              <svg class="icon-svg" style="width:13px; height:13px;" viewBox="0 0 24 24"><path d="M12 4V1L8 5l4 4V6c3.31 0 6 2.69 6 6 0 1.01-.25 1.97-.7 2.8l1.46 1.46C19.54 15.03 20 13.57 20 12c0-4.42-3.58-8-8-8zm0 14c-3.31 0-6-2.69-6-6 0-1.01.25-1.97.7-2.8L5.24 7.74C4.46 8.97 4 10.43 4 12c0 4.42 3.58 8 8 8v3l4-4-4-4v3z"/></svg>
              <span>申請調整 (滿 4 週)</span>
            </button>
          ` : `
            <button type="button" class="btn btn-secondary btn-sm btn-change-span" disabled style="opacity:0.55; cursor:not-allowed;" title="執行未滿 4 週，尚需 ${d.days_remaining} 天">
              <svg class="icon-svg" style="width:13px; height:13px;" viewBox="0 0 24 24"><path d="M18 8h-1V6c0-2.76-2.24-5-5-5S7 3.24 7 6v2H6c-1.1 0-2 .9-2 2v10c0 1.1.9 2 2 2h12c1.1 0 2-.9 2-2V10c0-1.1-.9-2-2-2zm-6 9c-1.1 0-2-.9-2-2s.9-2 2-2 2 .9 2 2-.9 2-2 2zm3.1-9H8.9V6c0-1.71 1.39-3.1 3.1-3.1 1.71 0 3.1 1.39 3.1 3.1v2z"/></svg>
              <span>未滿 4 週 (尚需 ${d.days_remaining} 天)</span>
            </button>
          `}
        </div>
      </div>
    `;

    // 綁定首頁卡片內部之按鈕
    drawerEl.querySelectorAll(".btn-report-driver-actual").forEach(btn => {
      btn.onclick = () => {
        openDriverReportActualModal(
          btn.getAttribute("data-id"),
          btn.getAttribute("data-product"),
          btn.getAttribute("data-name"),
          btn.getAttribute("data-target"),
          btn.getAttribute("data-actual"),
          btn.getAttribute("data-unit")
        );
      };
    });

    drawerEl.querySelectorAll(".btn-edit-driver-settings").forEach(btn => {
      btn.onclick = () => {
        openDriverEditSettingsModal(
          btn.getAttribute("data-id"),
          btn.getAttribute("data-product"),
          btn.getAttribute("data-name"),
          btn.getAttribute("data-target"),
          btn.getAttribute("data-unit"),
          btn.getAttribute("data-period")
        );
      };
    });

    drawerEl.querySelectorAll(".btn-driver-apply-home").forEach(btn => {
      btn.onclick = () => {
        const dId = btn.getAttribute("data-id");
        const pLine = btn.getAttribute("data-product");
        const oName = btn.getAttribute("data-name");
        openDriverApplyModal(dId, pLine, oName);
      };
    });
  }

  // 1. 綁定附圖一紅框處「單一立體切換按鈕」 (區域目標 ⮂ 個人目標)
  function setupScoreboardScopeSwitch() {
    const btnToggle = document.getElementById("btnToggleScope");
    if (!btnToggle) return;

    btnToggle.onclick = () => {
      // 點一下在 region 與 personal 之間切換
      currentScoreboardScope = (currentScoreboardScope === "region") ? "personal" : "region";
      updateScopeSwitchUI(currentScoreboardScope);
      if (lastParsedMetrics) applyScoreboardMetrics(lastParsedMetrics, currentScoreboardScope);
    };
  }

  function updateScopeSwitchUI(scope) {
    const btnToggle = document.getElementById("btnToggleScope");
    const label = document.getElementById("txtScopeBtnLabel");
    const currentRegion = getCurrentRegion();
    if (!btnToggle) return;

    if (scope === "personal") {
      btnToggle.className = "scope-toggle-btn scope-mode-personal";
      btnToggle.setAttribute("data-scope", "personal");
      if (label) label.innerText = "個人目標";
    } else {
      btnToggle.className = "scope-toggle-btn scope-mode-region";
      btnToggle.setAttribute("data-scope", "region");
      if (label) label.innerText = `${currentRegion}目標`;
    }
  }

  // 2. 綁定單一全局展開按鈕：點擊後成果目標完全折疊，只保留領先指標
  function setupGlobalDriverToggle() {
    const btn = document.getElementById("btnGlobalToggleDriver");
    if (!btn) return;
    btn.onclick = () => {
      isDriverViewActive = !isDriverViewActive;
      updateGlobalDriverView(isDriverViewActive);
    };
  }

  function updateGlobalDriverView(isActive) {
    const btn = document.getElementById("btnGlobalToggleDriver");
    const hint = document.getElementById("txtCurrentViewMode");
    const cardPersonal = document.getElementById("cardHeroPersonal");

    if (btn) {
      if (isActive) {
        btn.classList.add("expanded");
        btn.setAttribute("aria-expanded", "true");
        if (hint) hint.innerText = "目前檢視：領先指標";
        if (cardPersonal) cardPersonal.classList.add("collapsed");
        document.querySelectorAll(".kpi-outcome-body").forEach(el => el.classList.add("collapsed"));
        document.querySelectorAll(".kpi-driver-body").forEach(el => el.classList.add("expanded"));
      } else {
        btn.classList.remove("expanded");
        btn.setAttribute("aria-expanded", "false");
        if (hint) hint.innerText = "目前檢視：業績狀況";
        if (cardPersonal) cardPersonal.classList.remove("collapsed");
        document.querySelectorAll(".kpi-outcome-body").forEach(el => el.classList.remove("collapsed"));
        document.querySelectorAll(".kpi-driver-body").forEach(el => el.classList.remove("expanded"));
      }

      // 控制伺服、變頻器、PLC 三大產品右上角標籤：
      // 在個人視角或領先指標模式下全面隱藏；僅在北區目標且業績狀況模式時顯示
      const kpiTrafficLEB = document.getElementById("kpiTrafficLEB");
      const kpiTrafficLNH = document.getElementById("kpiTrafficLNH");
      const kpiTrafficPLC = document.getElementById("kpiTrafficPLC");
      const showTraffic = (!isActive && currentScoreboardScope === "region");
      [kpiTrafficLEB, kpiTrafficLNH, kpiTrafficPLC].forEach(el => {
        if (el) el.style.display = showTraffic ? "" : "none";
      });
    }
  }

  // 渲染設定頁個人驅動指標管理卡片
  function renderSettingsDriverManagement(salesName, drivers) {
    const gridEl = document.getElementById("gridSalesDrivers");
    if (!gridEl) return;

    if (!drivers || drivers.length === 0) {
      gridEl.innerHTML = `<div style="color:var(--text-secondary); font-size:0.85rem;">目前無指標資料</div>`;
      return;
    }

    gridEl.innerHTML = drivers.map(d => {
      const lightClass = d.traffic_light === "green" ? "success" : (d.traffic_light === "yellow" ? "warning" : (d.rate_pct >= 50 ? "alert" : "danger"));
      return `
        <div class="driver-mgmt-card">
          <div class="driver-mgmt-card-header">
            <span class="driver-mgmt-product">${escapeHtml(d.product_line)}</span>
            <span class="traffic-light-badge ${lightClass}" style="font-size:0.75rem;">
              達成率 ${d.rate_pct}% (${escapeHtml(d.light_text || '')})
            </span>
          </div>
          <div style="font-size:0.9rem; font-weight:600; color:var(--text-primary);">
            ${escapeHtml(d.driver_name)}
          </div>
          <div style="font-size:0.82rem; color:var(--text-secondary); display:flex; justify-content:space-between;">
            <span>目標：${d.target_val} ${escapeHtml(d.unit)} (${escapeHtml(d.period_type || '週')})</span>
            <span>實績：${d.actual_val} ${escapeHtml(d.unit)}</span>
          </div>
          <div class="driver-progress-bg" style="margin:2px 0;">
            <div class="driver-progress-bar ${d.bar_class || 'bar-warning'}" style="width: ${Math.min(100, Math.max(0, d.rate_pct))}%;"></div>
          </div>
          <div style="font-size:0.78rem; color:var(--text-tertiary); display:flex; justify-content:space-between; align-items:center; margin-top:2px;">
            <span>始於：${escapeHtml(d.start_date)}</span>
            ${d.can_change ? `<span style="color:#059669; font-weight:600;">已執行 ${d.days_active} 天 (滿 4 週)</span>` : `<span style="color:#dc2626; font-weight:600;">已執行 ${d.days_active} 天 (差 ${d.days_remaining} 天)</span>`}
          </div>
          <div class="driver-card-action-bar">
            <button type="button" class="btn btn-primary btn-sm btn-report-driver-actual"
              data-id="${d.id}" data-product="${escapeHtml(d.product_line)}" data-name="${escapeHtml(d.driver_name)}"
              data-target="${d.target_val}" data-actual="${d.actual_val}" data-unit="${escapeHtml(d.unit)}">
              <svg class="icon-svg" style="width:13px; height:13px;" viewBox="0 0 24 24"><path d="M19 3H5c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zm-2 10h-4v4h-2v-4H7v-2h4V7h2v4h4v2z"/></svg>
              <span>回報實績</span>
            </button>
            <button type="button" class="btn btn-secondary btn-sm btn-edit-driver-settings"
              data-id="${d.id}" data-product="${escapeHtml(d.product_line)}" data-name="${escapeHtml(d.driver_name)}"
              data-target="${d.target_val}" data-unit="${escapeHtml(d.unit)}" data-period="${escapeHtml(d.period_type || '週')}">
              <svg class="icon-svg" style="width:13px; height:13px;" viewBox="0 0 24 24"><path d="M3 17.25V21h3.75L17.81 9.94l-3.75-3.75L3 17.25zM20.71 7.04c.39-.39.39-1.02 0-1.41l-2.34-2.34c-.39-.39-1.02-.39-1.41 0l-1.83 1.83 3.75 3.75 1.83-1.83z"/></svg>
              <span>設定指標</span>
            </button>
            ${d.can_change ? `
              <button type="button" class="btn btn-secondary btn-sm btn-open-driver-modal btn-change-span"
                data-id="${d.id}" data-product="${escapeHtml(d.product_line)}" data-name="${escapeHtml(d.driver_name)}">
                <svg class="icon-svg" style="width:13px; height:13px;" viewBox="0 0 24 24"><path d="M12 4V1L8 5l4 4V6c3.31 0 6 2.69 6 6 0 1.01-.25 1.97-.7 2.8l1.46 1.46C19.54 15.03 20 13.57 20 12c0-4.42-3.58-8-8-8zm0 14c-3.31 0-6-2.69-6-6 0-1.01.25-1.97.7-2.8L5.24 7.74C4.46 8.97 4 10.43 4 12c0 4.42 3.58 8 8 8v3l4-4-4-4v3z"/></svg>
                <span>申請調整 (滿 4 週)</span>
              </button>
            ` : `
              <button type="button" class="btn btn-secondary btn-sm btn-change-span" disabled style="opacity:0.55; cursor:not-allowed;" title="執行未滿 4 週，尚需 ${d.days_remaining} 天">
                <svg class="icon-svg" style="width:13px; height:13px;" viewBox="0 0 24 24"><path d="M18 8h-1V6c0-2.76-2.24-5-5-5S7 3.24 7 6v2H6c-1.1 0-2 .9-2 2v10c0 1.1.9 2 2 2h12c1.1 0 2-.9 2-2V10c0-1.1-.9-2-2-2zm-6 9c-1.1 0-2-.9-2-2s.9-2 2-2 2 .9 2 2-.9 2-2 2zm3.1-9H8.9V6c0-1.71 1.39-3.1 3.1-3.1 1.71 0 3.1 1.39 3.1 3.1v2z"/></svg>
                <span>未滿 4 週 (尚需 ${d.days_remaining} 天)</span>
              </button>
            `}
          </div>
        </div>
      `;
    }).join("");

    // 綁定管理頁面卡片按鈕
    gridEl.querySelectorAll(".btn-report-driver-actual").forEach(btn => {
      btn.onclick = () => {
        openDriverReportActualModal(
          btn.getAttribute("data-id"),
          btn.getAttribute("data-product"),
          btn.getAttribute("data-name"),
          btn.getAttribute("data-target"),
          btn.getAttribute("data-actual"),
          btn.getAttribute("data-unit")
        );
      };
    });

    gridEl.querySelectorAll(".btn-edit-driver-settings").forEach(btn => {
      btn.onclick = () => {
        openDriverEditSettingsModal(
          btn.getAttribute("data-id"),
          btn.getAttribute("data-product"),
          btn.getAttribute("data-name"),
          btn.getAttribute("data-target"),
          btn.getAttribute("data-unit"),
          btn.getAttribute("data-period")
        );
      };
    });

    gridEl.querySelectorAll(".btn-open-driver-modal").forEach(btn => {
      btn.onclick = () => {
        const dId = btn.getAttribute("data-id");
        const pLine = btn.getAttribute("data-product");
        const oName = btn.getAttribute("data-name");
        openDriverApplyModal(dId, pLine, oName);
      };
    });
  }

  // 渲染主管審核專區 (曾維崧、曾仁君、張何達)
  async function renderManagerReviewPanel(salesName) {
    const panel = document.getElementById("panelDriverReview");
    const listEl = document.getElementById("listDriverReviewRequests");
    const badgeRole = document.getElementById("badgeManagerRole");
    if (!panel || !listEl) return;

    if (!isSalesManager(salesName)) {
      panel.style.display = "none";
      return;
    }

    panel.style.display = "block";
    if (badgeRole) {
      badgeRole.innerText = `${salesName} (${getManagerTitle(salesName)})`;
    }

    let requests = [];
    try {
      const res = await fetch("/api/scoreboard/drivers/requests");
      if (res.ok) {
        const data = await res.json();
        if (data.status === "success" && Array.isArray(data.requests)) {
          requests = data.requests;
        }
      }
    } catch (e) {
      console.log("[Drivers Review] 後端連線備援，切換至本地申請清單");
    }

    if (!requests || requests.length === 0) {
      requests = getLocalDriverRequests();
    }

    // 篩選待審核之申請單
    const pendingReqs = requests.filter(r => r.review_status === "待審核");

    if (pendingReqs.length === 0) {
      listEl.innerHTML = `<div style="font-size:0.85rem; color:var(--text-tertiary); padding:10px 0; text-align:center;">目前無待審核之領先指標申請案件。</div>`;
      return;
    }

    listEl.innerHTML = pendingReqs.map(req => `
      <div class="driver-review-item" data-req-id="${req.id}">
        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:6px;">
          <span style="font-weight:700; color:var(--text-primary); font-size:0.95rem;">
            ${escapeHtml(req.sales_name)} ➔ ${escapeHtml(req.product_line)}
          </span>
          <span class="traffic-light-badge warning" style="font-size:0.75rem;">${escapeHtml(req.review_status)}</span>
        </div>
        <div style="font-size:0.84rem; color:var(--text-secondary); margin-bottom:4px;">
          <strong>原指標：</strong>${escapeHtml(req.old_driver_name)}
        </div>
        <div style="font-size:0.84rem; color:var(--text-primary); margin-bottom:4px;">
          <strong>申請新指標：</strong><span style="color:var(--color-accent); font-weight:600;">${escapeHtml(req.new_driver_name)}</span> (目標：${req.new_target_val} ${escapeHtml(req.new_unit || '次')})
        </div>
        <div style="font-size:0.82rem; color:var(--text-secondary); margin-bottom:8px; background:var(--bg-subtle, rgba(0,0,0,0.03)); padding:6px 8px; border-radius:4px;">
          <strong>變更理由：</strong>${escapeHtml(req.change_reason)}
        </div>
        <div style="display:flex; gap:8px; align-items:center; margin-top:8px;">
          <input type="text" id="reviewComment_${req.id}" class="form-control" placeholder="主管簽核指引 / 反饋意見 (選填)" style="flex:1; font-size:0.82rem; padding:4px 8px;" />
          <button type="button" class="btn btn-primary btn-sm btn-approve-driver" data-id="${req.id}" style="padding:4px 12px; font-size:0.82rem;">核准生效</button>
          <button type="button" class="btn btn-secondary btn-sm btn-reject-driver" data-id="${req.id}" style="padding:4px 12px; font-size:0.82rem;">駁回申請</button>
        </div>
      </div>
    `).join("");

    // 綁定核准與駁回按鈕
    listEl.querySelectorAll(".btn-approve-driver").forEach(btn => {
      btn.onclick = () => {
        const reqId = btn.getAttribute("data-id");
        const commentInput = document.getElementById(`reviewComment_${reqId}`);
        const comment = commentInput ? commentInput.value : "";
        handleReviewAction(reqId, "approve", salesName, comment);
      };
    });

    listEl.querySelectorAll(".btn-reject-driver").forEach(btn => {
      btn.onclick = () => {
        const reqId = btn.getAttribute("data-id");
        const commentInput = document.getElementById(`reviewComment_${reqId}`);
        const comment = commentInput ? commentInput.value : "";
        handleReviewAction(reqId, "reject", salesName, comment);
      };
    });
  }

  // 處理主管審核動作
  async function handleReviewAction(requestId, action, reviewerName, comment) {
    const actionText = action === "approve" ? "核准" : "駁回";

    let success = false;
    let message = "";

    try {
      const res = await fetch("/api/scoreboard/drivers/review", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          request_id: parseInt(requestId),
          action: action,
          reviewer_name: reviewerName,
          comment: comment
        })
      });
      if (res.ok) {
        const data = await res.json();
        success = true;
        message = data.message;
      } else {
        const err = await res.json();
        message = err.detail || "審核失敗";
      }
    } catch (e) {
      // 離線備援處理
      const localReqs = getLocalDriverRequests();
      const r = localReqs.find(item => item.id == requestId);
      if (r) {
        r.review_status = action === "approve" ? "已核准" : "已駁回";
        r.reviewer_comment = `[${reviewerName}] ${comment || '審核完成'}`;
        saveLocalDriverRequests(localReqs);

        if (action === "approve") {
          const metrics = getLocalDriverMetrics(r.sales_name);
          const m = metrics.find(item => item.product_line === r.product_line);
          if (m) {
            m.driver_name = r.new_driver_name;
            m.target_val = r.new_target_val;
            m.actual_val = 0;
            m.unit = r.new_unit;
            m.start_date = new Date().toISOString().split("T")[0];
            m.days_active = 0;
            m.weeks_active = 0;
            m.can_change = false;
            m.days_remaining = 28;
            m.rate_pct = 0;
            m.traffic_light = "red";
            m.light_text = "嚴重落後";
            m.bar_class = "bar-danger";
            saveLocalDriverMetrics(r.sales_name, metrics);
          }
        }
        success = true;
        message = action === "approve" ? "申請已核准！新指標正式生效。" : "已駁回申請。";
      }
    }

    if (success) {
      alert(`[審核成功] ${message}`);
      const activeSales = localStorage.getItem("shihlin_pwa_active_sales") || "曾維崧";
      loadAndRenderDriverMetrics(activeSales);
      renderManagerReviewPanel(activeSales);
      if (currentViewingDriverSales) {
        renderDriverManagementPage(currentViewingDriverSales);
      }
    } else {
      alert(`[審核失敗] ${message}`);
    }
  }

  // 實績回報彈窗控制
  function openDriverReportActualModal(driverId, productLine, driverName, targetVal, currentActual, unit) {
    const modal = document.getElementById("modalDriverReportActual");
    const dIdEl = document.getElementById("reportDriverId");
    const pLineEl = document.getElementById("reportProductLine");
    const dNameEl = document.getElementById("reportDriverName");
    const tValEl = document.getElementById("reportTargetVal");
    const aValEl = document.getElementById("reportActualVal");
    const untEl1 = document.getElementById("reportUnitText1");
    const untEl2 = document.getElementById("reportUnitText2");

    if (!modal) return;
    if (dIdEl) dIdEl.value = driverId;
    if (pLineEl) pLineEl.value = productLine;
    if (dNameEl) dNameEl.value = driverName;
    if (tValEl) tValEl.value = targetVal;
    if (aValEl) aValEl.value = currentActual;
    if (untEl1) untEl1.innerText = unit;
    if (untEl2) untEl2.innerText = unit;

    modal.style.display = "flex";
  }

  function closeDriverReportActualModal() {
    const modal = document.getElementById("modalDriverReportActual");
    if (modal) modal.style.display = "none";
  }

  // 指標設定編輯彈窗控制
  function openDriverEditSettingsModal(driverId, productLine, driverName, targetVal, unit, periodType) {
    const modal = document.getElementById("modalDriverEditSettings");
    const dIdEl = document.getElementById("editDriverId");
    const pLineEl = document.getElementById("editProductLine");
    const dNameEl = document.getElementById("editDriverName");
    const tValEl = document.getElementById("editTargetVal");
    const untEl = document.getElementById("editUnit");
    const pTypeEl = document.getElementById("editPeriodType");

    if (!modal) return;
    if (dIdEl) dIdEl.value = driverId;
    if (pLineEl) pLineEl.value = productLine;
    if (dNameEl) dNameEl.value = driverName;
    if (tValEl) tValEl.value = targetVal;
    if (untEl) untEl.value = unit || "場";
    if (pTypeEl) pTypeEl.value = periodType || "週";

    modal.style.display = "flex";
  }

  function closeDriverEditSettingsModal() {
    const modal = document.getElementById("modalDriverEditSettings");
    if (modal) modal.style.display = "none";
  }

  // 申請彈窗控制
  function openDriverApplyModal(driverId, productLine, currentName) {
    const modal = document.getElementById("modalDriverChangeApply");
    const dIdEl = document.getElementById("applyDriverId");
    const pLineEl = document.getElementById("applyProductLine");
    const oNameEl = document.getElementById("applyOldDriverName");
    const nNameEl = document.getElementById("applyNewDriverName");
    const nValEl = document.getElementById("applyNewTargetVal");
    const nReasonEl = document.getElementById("applyChangeReason");

    if (!modal) return;
    if (dIdEl) dIdEl.value = driverId;
    if (pLineEl) pLineEl.value = productLine;
    if (oNameEl) oNameEl.value = currentName;
    if (nNameEl) nNameEl.value = "";
    if (nValEl) nValEl.value = "";
    if (nReasonEl) nReasonEl.value = "";

    modal.style.display = "flex";
  }

  function closeDriverApplyModal() {
    const modal = document.getElementById("modalDriverChangeApply");
    if (modal) modal.style.display = "none";
  }

  // 本地快取實績更新
  function updateLocalDriverActual(driverId, actualVal) {
    const key = "shihlin_pwa_driver_metrics";
    let store = {};
    try {
      const raw = localStorage.getItem(key);
      if (raw) store = JSON.parse(raw);
    } catch (e) {}

    Object.keys(store).forEach(salesName => {
      const list = store[salesName];
      if (Array.isArray(list)) {
        const d = list.find(item => item.id == driverId);
        if (d) {
          d.actual_val = parseFloat(actualVal);
          d.rate_pct = d.target_val > 0 ? Math.round((d.actual_val / d.target_val) * 1000) / 10 : 0;
          if (d.rate_pct >= 90) {
            d.traffic_light = "green";
            d.light_text = "達標";
            d.bar_class = "bar-success";
          } else if (d.rate_pct >= 75) {
            d.traffic_light = "yellow";
            d.light_text = "預警";
            d.bar_class = "bar-warning";
          } else if (d.rate_pct >= 50) {
            d.traffic_light = "red";
            d.light_text = "低標";
            d.bar_class = "bar-alert";
          } else {
            d.traffic_light = "red";
            d.light_text = "嚴重落後";
            d.bar_class = "bar-danger";
          }
        }
      }
    });
    localStorage.setItem(key, JSON.stringify(store));
  }

  // 本地快取設定更新
  function updateLocalDriverSettings(driverId, driverName, targetVal, unit, periodType) {
    const key = "shihlin_pwa_driver_metrics";
    let store = {};
    try {
      const raw = localStorage.getItem(key);
      if (raw) store = JSON.parse(raw);
    } catch (e) {}

    Object.keys(store).forEach(salesName => {
      const list = store[salesName];
      if (Array.isArray(list)) {
        const d = list.find(item => item.id == driverId);
        if (d) {
          d.driver_name = driverName;
          d.target_val = parseFloat(targetVal);
          d.unit = unit;
          d.period_type = periodType;
          d.rate_pct = d.target_val > 0 ? Math.round((d.actual_val / d.target_val) * 1000) / 10 : 0;
          if (d.rate_pct >= 90) {
            d.traffic_light = "green";
            d.light_text = "達標";
            d.bar_class = "bar-success";
          } else if (d.rate_pct >= 75) {
            d.traffic_light = "yellow";
            d.light_text = "預警";
            d.bar_class = "bar-warning";
          } else if (d.rate_pct >= 50) {
            d.traffic_light = "red";
            d.light_text = "低標";
            d.bar_class = "bar-alert";
          } else {
            d.traffic_light = "red";
            d.light_text = "嚴重落後";
            d.bar_class = "bar-danger";
          }
        }
      }
    });
    localStorage.setItem(key, JSON.stringify(store));
  }

  // 設定 Modal 事件
  function setupDriverModalEvents() {
    const btnClose = document.getElementById("btnCloseDriverApplyModal");
    const btnCancel = document.getElementById("btnCancelDriverApply");
    const modal = document.getElementById("modalDriverChangeApply");
    const form = document.getElementById("formDriverChangeApply");
    const btnRefresh = document.getElementById("btnRefreshDriverRequests");

    if (btnClose) btnClose.addEventListener("click", closeDriverApplyModal);
    if (btnCancel) btnCancel.addEventListener("click", closeDriverApplyModal);
    if (modal) {
      modal.addEventListener("click", (e) => {
        if (e.target === modal) closeDriverApplyModal();
      });
    }

    // 1. 回報實績 Modal 事件
    const modalReport = document.getElementById("modalDriverReportActual");
    const btnCloseReport = document.getElementById("btnCloseDriverReportModal");
    const btnCancelReport = document.getElementById("btnCancelDriverReport");
    const formReport = document.getElementById("formDriverReportActual");
    const btnQuickAddOne = document.getElementById("btnQuickAddOne");

    if (btnCloseReport) btnCloseReport.addEventListener("click", closeDriverReportActualModal);
    if (btnCancelReport) btnCancelReport.addEventListener("click", closeDriverReportActualModal);
    if (modalReport) {
      modalReport.addEventListener("click", (e) => {
        if (e.target === modalReport) closeDriverReportActualModal();
      });
    }

    if (btnQuickAddOne) {
      btnQuickAddOne.addEventListener("click", () => {
        const inp = document.getElementById("reportActualVal");
        if (inp) {
          const val = parseFloat(inp.value) || 0;
          inp.value = (val + 1).toString();
        }
      });
    }

    if (formReport) {
      formReport.addEventListener("submit", async (e) => {
        e.preventDefault();
        const driverId = parseInt(document.getElementById("reportDriverId").value);
        const actualVal = parseFloat(document.getElementById("reportActualVal").value);
        const activeSales = localStorage.getItem("shihlin_pwa_active_sales") || "曾維崧";

        if (isNaN(actualVal) || actualVal < 0) {
          alert("實績數值不得為負數！");
          return;
        }

        let success = false;
        let message = "";
        try {
          const res = await fetch("/api/scoreboard/drivers/report-actual", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
              driver_id: driverId,
              actual_val: actualVal
            })
          });
          if (res.ok) {
            const data = await res.json();
            success = true;
            message = data.message;
          } else {
            const err = await res.json();
            message = err.detail || "實績回報失敗";
          }
        } catch (err) {
          // 離線備援
          updateLocalDriverActual(driverId, actualVal);
          success = true;
          message = "實績已登記至本地快取！";
        }

        if (success) {
          alert(message);
          closeDriverReportActualModal();
          loadAndRenderDriverMetrics(activeSales);
          if (currentViewingDriverSales) {
            renderDriverManagementPage(currentViewingDriverSales);
          }
        } else {
          alert(message);
        }
      });
    }

    // 2. 設定指標 Modal 事件
    const modalEdit = document.getElementById("modalDriverEditSettings");
    const btnCloseEdit = document.getElementById("btnCloseDriverEditModal");
    const btnCancelEdit = document.getElementById("btnCancelDriverEdit");
    const formEdit = document.getElementById("formDriverEditSettings");

    if (btnCloseEdit) btnCloseEdit.addEventListener("click", closeDriverEditSettingsModal);
    if (btnCancelEdit) btnCancelEdit.addEventListener("click", closeDriverEditSettingsModal);
    if (modalEdit) {
      modalEdit.addEventListener("click", (e) => {
        if (e.target === modalEdit) closeDriverEditSettingsModal();
      });
    }

    if (formEdit) {
      formEdit.addEventListener("submit", async (e) => {
        e.preventDefault();
        const driverId = parseInt(document.getElementById("editDriverId").value);
        const driverName = document.getElementById("editDriverName").value.trim();
        const targetVal = parseFloat(document.getElementById("editTargetVal").value);
        const unit = document.getElementById("editUnit").value;
        const periodType = document.getElementById("editPeriodType").value;
        const activeSales = localStorage.getItem("shihlin_pwa_active_sales") || "曾維崧";

        if (!driverName || isNaN(targetVal) || targetVal <= 0) {
          alert("請輸入有效的指標名稱與目標數值（須大於 0）！");
          return;
        }

        let success = false;
        let message = "";
        try {
          const res = await fetch("/api/scoreboard/drivers/update-settings", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
              driver_id: driverId,
              driver_name: driverName,
              target_val: targetVal,
              unit: unit,
              period_type: periodType
            })
          });
          if (res.ok) {
            const data = await res.json();
            success = true;
            message = data.message;
          } else {
            const err = await res.json();
            message = err.detail || "設定更新失敗";
          }
        } catch (err) {
          // 離線備援
          updateLocalDriverSettings(driverId, driverName, targetVal, unit, periodType);
          success = true;
          message = "指標設定已更新至本地快取！";
        }

        if (success) {
          alert(message);
          closeDriverEditSettingsModal();
          loadAndRenderDriverMetrics(activeSales);
          if (currentViewingDriverSales) {
            renderDriverManagementPage(currentViewingDriverSales);
          }
        } else {
          alert(message);
        }
      });
    }

    if (btnRefresh) {
      btnRefresh.addEventListener("click", () => {
        const activeSales = localStorage.getItem("shihlin_pwa_active_sales") || "曾維崧";
        renderManagerReviewPanel(activeSales);
        if (currentViewingDriverSales) {
          renderDriverManagementPage(currentViewingDriverSales);
        }
      });
    }

    if (form) {
      form.addEventListener("submit", async (e) => {
        e.preventDefault();
        const driverId = parseInt(document.getElementById("applyDriverId").value);
        const pLine = document.getElementById("applyProductLine").value;
        const oName = document.getElementById("applyOldDriverName").value;
        const newName = document.getElementById("applyNewDriverName").value.trim();
        const newVal = parseFloat(document.getElementById("applyNewTargetVal").value);
        const newUnit = document.getElementById("applyNewUnit").value;
        const reason = document.getElementById("applyChangeReason").value.trim();
        const activeSales = localStorage.getItem("shihlin_pwa_active_sales") || "曾維崧";

        if (!newName || !newVal || !reason) {
          alert("請完整填寫所有必填欄位！");
          return;
        }

        const btnSubmit = document.getElementById("btnSubmitDriverApply");
        if (btnSubmit) {
          btnSubmit.disabled = true;
          btnSubmit.innerText = "送出申請中...";
        }

        let success = false;
        let message = "";

        try {
          const res = await fetch("/api/scoreboard/drivers/apply-change", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
              driver_id: driverId,
              new_driver_name: newName,
              new_target_val: newVal,
              new_unit: newUnit,
              change_reason: reason
            })
          });
          if (res.ok) {
            const data = await res.json();
            success = true;
            message = data.message;
          } else {
            const err = await res.json();
            message = err.detail || "申請失敗";
          }
        } catch (err) {
          // 離線備援
          const reqs = getLocalDriverRequests();
          const newReq = {
            id: Date.now(),
            driver_id: driverId,
            sales_name: activeSales,
            product_line: pLine,
            old_driver_name: oName,
            new_driver_name: newName,
            new_target_val: newVal,
            new_unit: newUnit,
            change_reason: reason,
            apply_date: new Date().toISOString().split("T")[0],
            review_status: "待審核"
          };
          reqs.unshift(newReq);
          saveLocalDriverRequests(reqs);
          success = true;
          message = "更換申請已成功送出，請靜候主管審核！";
        } finally {
          if (btnSubmit) {
            btnSubmit.disabled = false;
            btnSubmit.innerText = "送出更換申請 (送主管審核)";
          }
        }

        if (success) {
          alert(`[申請成功] ${message}`);
          closeDriverApplyModal();
          loadAndRenderDriverMetrics(activeSales);
          renderManagerReviewPanel(activeSales);
          if (currentViewingDriverSales) {
            renderDriverManagementPage(currentViewingDriverSales);
          }
        } else {
          alert(`[無法申請] ${message}`);
        }
      });
    }
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

  // ============================================================================
  // 13. 報價單管理模組 (電腦端圖二 Excel 試算表 + 手機端圖三卡片式擬稿雙視角整合版)
  // ============================================================================
  let isQuotationModuleInitialized = false;
  let quoteItemCount = 0;
  let currentEditingItemId = null;
  let currentSelectedCustomer = null;
  let quoteProductsCache = [];
  let itemSkuMap = {};
  let quoteAttachedImages = [];
  let isTrashViewActive = false;
  let currentActiveQuoteTab = 'create';
  let pendingDeleteQuoteNo = null;
  let currentBaseQuoteNo = null;
  let currentLoadedDraftId = null;

  // 財務捨入與庫存同步設定 (自 /api/quotation-settings 讀取，預設百位數無條件捨去)
  let quoteFinanceSettings = {
    round_digits: 2,
    round_factor: 100,
    round_mode: "無條件捨去",
    round_unit_label: "百位數",
    round_summary_text: "捨入：百位數 / 無條件捨去",
    tax_rate: 5.0,
    stock_updated_at: "10/15 10:12",
    default_terms: ""
  };

  let PRESET_DELIVERY_OPTIONS = [
    "現貨供應",
    "訂單確認後 1~2 週",
    "訂單確認後 2~3 週",
    "訂單確認後 3~4 週",
    "訂單確認後 4~6 週",
    "訂單確認後 6~8 週",
    "訂單確認後 8~12 週",
    "德國原廠空運約 3~4 週",
    "日本原廠進口約 4~6 週",
    "特殊規格，交期另議",
    "需提供設備規格後確認交期",
    "數量有限，售完為止"
  ];

  // 取得目前檢視/登入之業務姓名 (用於客戶、歷史報價、手機草稿、個人特殊品完全隔離)
  function getActiveSalesName() {
    return localStorage.getItem("shihlin_pwa_active_sales") || "曾維崧";
  }

  // 依據財務設定執行金額捨入 (對標舊版 PriceCalculator.apply_rounding)
  function applyQuoteRounding(rawVal) {
    const val = parseFloat(rawVal);
    if (isNaN(val) || val <= 0) return 0;
    const d = parseInt(quoteFinanceSettings.round_digits, 10) || 0;
    const mode = quoteFinanceSettings.round_mode || "無條件捨去";
    if (d <= 0) {
      if (mode === "無條件捨去" || mode === "FLOOR") return Math.floor(val);
      if (mode === "無條件進位" || mode === "CEIL") return Math.ceil(val);
      return Math.round(val);
    }
    const factor = Math.pow(10, d);
    if (mode === "無條件捨去" || mode === "FLOOR") {
      return Math.floor(val / factor) * factor;
    } else if (mode === "無條件進位" || mode === "CEIL") {
      return Math.ceil(val / factor) * factor;
    } else {
      return Math.round(val / factor) * factor;
    }
  }

  // 當業務人員身份切換時，重新載入該業務專屬客戶、產品牌價(含個人特殊品)、歷史報價與草稿
  async function onQuotationSalespersonChanged(salesName) {
    quoteProductsCache = [];
    currentSelectedCustomer = null;
    const spBadge = document.getElementById("specialProductOwnerBadge");
    if (spBadge) spBadge.textContent = `所屬業務：${salesName}`;
    await populateCustomerSelectDropdown(true);
    updateQuoteDraftsBadge();
    if (currentActiveQuoteTab === "history") {
      loadQuotationHistory("", isTrashViewActive);
    } else if (currentActiveQuoteTab === "drafts") {
      loadQuotationDraftsInbox();
    } else if (currentActiveQuoteTab === "special" || currentActiveQuoteTab === "products") {
      loadSpecialProductsManager();
      searchStandardProductsInMgmt();
    } else if (currentActiveQuoteTab === "dashboard") {
      loadQuotationDashboard();
    } else if (currentActiveQuoteTab === "customers") {
      loadQuotationCustomersPage();
    }
  }

  async function loadQuotationFinanceSettings() {
    try {
      const resp = await fetch("/api/quotation-settings");
      if (resp.ok) {
        const data = await resp.json();
        quoteFinanceSettings = {
          ...quoteFinanceSettings,
          ...data
        };
        document.querySelectorAll(".stock-update-time-text").forEach(el => {
          el.textContent = quoteFinanceSettings.stock_updated_at || "10/15 10:12";
        });
        document.querySelectorAll(".rounding-rule-text").forEach(el => {
          el.textContent = quoteFinanceSettings.round_summary_text || "捨入：百位數 / 無條件捨去";
        });
        const metaInfo = document.getElementById("quoteSidebarMetaInfo");
        if (metaInfo) metaInfo.textContent = quoteFinanceSettings.round_summary_text || "捨入：百位數 / 無條件捨去";
        const stockTime = document.getElementById("quoteSidebarStockTime");
        if (stockTime) stockTime.textContent = `L廠庫存更新：${quoteFinanceSettings.stock_updated_at || "10/15 10:12"}`;
      }
    } catch (err) {
      console.warn("讀取報價單財務與庫存設定失敗，使用預設值:", err);
    }
  }

  async function initQuotationModule() {
    await loadQuotationFinanceSettings();
    if (!isQuotationModuleInitialized) {
      setupQuotationEvents();
      setupQuotationModals();
      isQuotationModuleInitialized = true;
    }
    const qDate = document.getElementById("quoteDate");
    if (qDate && !qDate.value) {
      qDate.value = new Date().toISOString().split("T")[0];
    }
    const qValidDate = document.getElementById("quoteValidDate");
    if (qValidDate && !qValidDate.value) {
      const vDate = new Date();
      vDate.setDate(vDate.getDate() + 30);
      qValidDate.value = vDate.toISOString().split("T")[0];
    }
    await populateCustomerSelectDropdown(false);
    const itemsContainer = document.getElementById("itemsContainer");
    if (itemsContainer && itemsContainer.children.length === 0) {
      addBlankItem();
    }
    updateQuoteNoPreview();
    enforceDesktopQuotationView();
  }

  // 初始化並填充業務專屬客戶下拉清單
  async function populateCustomerSelectDropdown(forceReload = false) {
    const custSel = document.getElementById("customerSelect");
    const activeSales = getActiveSalesName();
    if (forceReload || !customersCache || customersCache.length === 0) {
      try {
        const resp = await fetch(`/api/customers?sales_name=${encodeURIComponent(activeSales)}`);
        if (resp.ok) {
          customersCache = await resp.json();
        }
      } catch (err) {
        console.warn("載入客戶清單失敗:", err);
      }
    }
    if (!custSel) return;
    const currentVal = custSel.value;
    custSel.innerHTML = '<option value="">-- 請選擇客戶 --</option>' +
      (customersCache || []).map(c => `<option value="${c.id}">${escapeHtml(c.name)}</option>`).join('');
    if (currentVal && (customersCache || []).some(c => String(c.id) === String(currentVal))) {
      custSel.value = currentVal;
    }
  }

  // 統一套用所選客戶至電腦端表頭與手機端詳細資料卡
  function applySelectedCustomerToQuoteUI(cust) {
    if (!cust) return;
    currentSelectedCustomer = cust;

    const custInput = document.getElementById("customerName");
    if (custInput) custInput.value = cust.name || "";

    const custSel = document.getElementById("customerSelect");
    if (custSel) custSel.value = String(cust.id || "");

    const contacts = Array.isArray(cust.contacts) && cust.contacts.length > 0
      ? cust.contacts
      : [{
          name: cust.default_contact || cust.contact_person || "主要窗口",
          phone: cust.phone || "",
          fax: cust.fax || "",
          address: cust.address || cust.delivery_address || ""
        }];
    const primaryContact = contacts[0] || {};

    // 1. 電腦端雙欄表頭連動
    const taxInput = document.getElementById("quoteTaxId");
    if (taxInput) taxInput.value = cust.tax_id || "";
    const phoneInput = document.getElementById("quotePhone");
    if (phoneInput) phoneInput.value = primaryContact.phone || cust.phone || "";
    const faxInput = document.getElementById("quoteFax");
    if (faxInput) faxInput.value = primaryContact.fax || cust.fax || "";
    const locInput = document.getElementById("quoteDeliveryLocation");
    if (locInput) locInput.value = primaryContact.address || cust.address || cust.delivery_address || "";
    const termsInput = document.getElementById("quotePaymentTerms");
    if (termsInput) termsInput.value = cust.payment_terms || "月結 30 天電匯";

    const contactSel = document.getElementById("quoteContactSelect");
    if (contactSel) {
      contactSel.innerHTML = contacts.map((ct, idx) =>
        `<option value="${escapeHtml(ct.name || '主要窗口')}" data-idx="${idx}">${escapeHtml(ct.name || '主要窗口')}${ct.department ? ' (' + escapeHtml(ct.department) + ')' : ''}</option>`
      ).join("");
      contactSel.onchange = () => {
        const selOpt = contactSel.options[contactSel.selectedIndex];
        const cIdx = selOpt ? parseInt(selOpt.getAttribute("data-idx"), 10) : 0;
        const chosen = contacts[cIdx] || primaryContact;
        if (phoneInput) phoneInput.value = chosen.phone || cust.phone || "";
        if (faxInput) faxInput.value = chosen.fax || cust.fax || "";
        if (locInput && chosen.address) locInput.value = chosen.address;
        const cdCont = document.getElementById("cdContact");
        if (cdCont) cdCont.textContent = chosen.name || "-";
        const cdPh = document.getElementById("cdPhone");
        if (cdPh) cdPh.textContent = chosen.phone || cust.phone || "-";
      };
    }

    // 單號前綴自動計算
    const prefixInput = document.getElementById("quotePrefixInput");
    if (prefixInput) {
      const now = new Date();
      const yy = String(now.getFullYear()).slice(-2);
      const mm = String(now.getMonth() + 1).padStart(2, "0");
      let cleanCust = (cust.short_name || cust.name || "").replace(/有限公司|股份有限公司|企業|實業|科技|工業/g, "").trim();
      const shortName = cleanCust.slice(0, 4) || "士電";
      prefixInput.value = `${shortName}${yy}${mm}`;
    }

    // 2. 手機端客戶詳細資料卡片連動
    const detailCard = document.getElementById("customerDetailCard");
    if (detailCard) detailCard.classList.remove("hidden");
    const comp = document.getElementById("cdCompany");
    if (comp) comp.textContent = cust.name || "-";
    const tax = document.getElementById("cdTaxId");
    if (tax) tax.textContent = cust.tax_id || "-";
    const cont = document.getElementById("cdContact");
    if (cont) cont.textContent = primaryContact.name || cust.default_contact || "-";
    const phone = document.getElementById("cdPhone");
    if (phone) phone.textContent = primaryContact.phone || cust.phone || "-";
    const addr = document.getElementById("cdAddress");
    if (addr) addr.textContent = primaryContact.address || cust.address || "-";

    const cdContactSelect = document.getElementById("cdContactSelect");
    if (cdContactSelect) {
      if (contacts.length > 1) {
        cdContactSelect.classList.remove("hidden");
        cdContactSelect.innerHTML = contacts.map((ct, idx) =>
          `<option value="${idx}">${escapeHtml(ct.name || '窗口')}</option>`
        ).join("");
        cdContactSelect.onchange = () => {
          const chosen = contacts[parseInt(cdContactSelect.value, 10)] || primaryContact;
          if (cont) cont.textContent = chosen.name || "-";
          if (phone) phone.textContent = chosen.phone || cust.phone || "-";
          if (addr && chosen.address) addr.textContent = chosen.address;
          if (contactSel) contactSel.value = chosen.name || "";
          if (phoneInput) phoneInput.value = chosen.phone || cust.phone || "";
        };
      } else {
        cdContactSelect.classList.add("hidden");
      }
    }

    updateQuoteNoPreview();
  }

  // 九大報價系統功能頁切換邏輯 (100% 對標電腦版 _QuotationApp 左側選單)
  function switchQuoteSubTab(tabName) {
    const normalizedTab = tabName === "special" ? "products" : tabName;
    currentActiveQuoteTab = normalizedTab;

    const tabCreate = document.getElementById("tabQuoteCreate");
    const tabHistory = document.getElementById("tabQuoteHistory");
    const tabDrafts = document.getElementById("tabQuoteDrafts");
    const tabAnalysis = document.getElementById("tabQuoteAnalysis");
    const tabSpecial = document.getElementById("tabQuoteSpecialProducts");

    const draftSec = document.getElementById("quoteDraftSection");
    const histSec = document.getElementById("quoteHistorySection");
    const draftsInboxSec = document.getElementById("quoteDraftsInboxSection");
    const analysisSec = document.getElementById("quoteAnalysisSection");
    const specialSec = document.getElementById("quoteSpecialProductsSection");
    const dashSec = document.getElementById("quoteDashboardSection");
    const custSec = document.getElementById("quoteCustomersSection");
    const setSec = document.getElementById("quoteSettingsSection");
    const helpSec = document.getElementById("quoteHelpSection");

    [tabCreate, tabHistory, tabDrafts, tabAnalysis, tabSpecial].forEach(btn => {
      if (btn) btn.classList.remove("active");
    });
    document.querySelectorAll("#quoteSidebarNavList .quote-nav-btn").forEach(btn => {
      btn.classList.toggle("active", btn.getAttribute("data-quote-page") === normalizedTab);
    });

    [draftSec, histSec, draftsInboxSec, analysisSec, specialSec, dashSec, custSec, setSec, helpSec].forEach(sec => {
      if (sec) sec.classList.add("hidden");
    });

    if (normalizedTab === "create") {
      if (tabCreate) tabCreate.classList.add("active");
      if (draftSec) draftSec.classList.remove("hidden");
    } else if (normalizedTab === "history") {
      if (tabHistory) tabHistory.classList.add("active");
      if (histSec) histSec.classList.remove("hidden");
      const kw = document.getElementById("inputSearchQuoteHistory")?.value.trim() || "";
      loadQuotationHistory(kw, isTrashViewActive);
    } else if (normalizedTab === "drafts") {
      if (tabDrafts) tabDrafts.classList.add("active");
      if (draftsInboxSec) draftsInboxSec.classList.remove("hidden");
      loadQuotationDraftsInbox();
    } else if (normalizedTab === "dashboard") {
      if (dashSec) dashSec.classList.remove("hidden");
      loadQuotationDashboard();
    } else if (normalizedTab === "analysis") {
      if (tabAnalysis) tabAnalysis.classList.add("active");
      if (analysisSec) analysisSec.classList.remove("hidden");
    } else if (normalizedTab === "customers") {
      if (custSec) custSec.classList.remove("hidden");
      loadQuotationCustomersPage();
    } else if (normalizedTab === "products") {
      if (tabSpecial) tabSpecial.classList.add("active");
      if (specialSec) specialSec.classList.remove("hidden");
      loadSpecialProductsManager();
      searchStandardProductsInMgmt();
    } else if (normalizedTab === "settings") {
      if (setSec) setSec.classList.remove("hidden");
      loadQuotationSettingsPage();
    } else if (normalizedTab === "help") {
      if (helpSec) helpSec.classList.remove("hidden");
    }

    closeDrawer();
  }

  // 更新三段式單號預覽 (Prefix-Seq-Rev)
  function updateQuoteNoPreview() {
    const custInput = document.getElementById("customerName");
    const custName = custInput?.value.trim() || "";
    const dateInput = document.getElementById("quoteDate");
    const qDate = dateInput?.value || new Date().toISOString().split("T")[0];

    const yy = qDate.slice(2, 4);
    const mm = qDate.slice(5, 7);
    const yymm = yy + mm;

    let prefix = "ST";
    if (custName) {
      let cleanCust = custName.replace(/有限公司|股份有限公司|企業|實業|科技|工業/g, "").trim();
      prefix = cleanCust.slice(0, 4) || "ST";
    }

    const badge = document.getElementById("quoteNoPreviewBadge");
    if (!badge) return;

    if (currentBaseQuoteNo) {
      badge.textContent = `單號預覽：${currentBaseQuoteNo} (增版)`;
    } else {
      badge.textContent = `單號預覽：${prefix}${yymm}-001-Rev0`;
    }
  }

  function setupQuotationEvents() {
    // 1. 五大子分頁按鈕監聽
    const tabCreate = document.getElementById("tabQuoteCreate");
    const tabHistory = document.getElementById("tabQuoteHistory");
    const tabDrafts = document.getElementById("tabQuoteDrafts");
    const tabAnalysis = document.getElementById("tabQuoteAnalysis");
    const tabSpecial = document.getElementById("tabQuoteSpecialProducts");

    if (tabCreate) tabCreate.addEventListener("click", () => switchQuoteSubTab('create'));
    if (tabHistory) tabHistory.addEventListener("click", () => switchQuoteSubTab('history'));
    if (tabDrafts) tabDrafts.addEventListener("click", () => switchQuoteSubTab('drafts'));
    if (tabAnalysis) tabAnalysis.addEventListener("click", () => switchQuoteSubTab('analysis'));
    if (tabSpecial) tabSpecial.addEventListener("click", () => switchQuoteSubTab('special'));

    // 2. _QuotationApp 頂部：收合選單按鈕 (左側側邊欄切換)
    const btnToggleSidebar = document.getElementById("btnToggleSidebar");
    if (btnToggleSidebar) {
      btnToggleSidebar.addEventListener("click", () => {
        const drawer = document.getElementById("appDrawer");
        if (drawer) drawer.classList.toggle("collapsed");
      });
    }

    // _QuotationApp 頂部：收合抬頭按鈕 (雙欄基本資料表頭切換)
    const btnToggleHeaderCompact = document.getElementById("btnToggleHeaderCompact");
    const quoteHeaderContainer = document.getElementById("quoteHeaderContainer");
    const headerToggleIcon = document.getElementById("headerToggleIcon");
    const headerToggleLabel = document.getElementById("headerToggleLabel");
    if (btnToggleHeaderCompact && quoteHeaderContainer) {
      btnToggleHeaderCompact.addEventListener("click", () => {
        quoteHeaderContainer.classList.toggle("collapsed");
        const isCollapsed = quoteHeaderContainer.classList.contains("collapsed");
        if (headerToggleIcon) headerToggleIcon.textContent = isCollapsed ? "▼" : "▲";
        if (headerToggleLabel) headerToggleLabel.textContent = isCollapsed ? "展開抬頭" : "收合抬頭";
      });
    }

    // 3. 電腦端公司名稱下拉選單更動 -> 同步手機與電腦端客戶資訊
    const custSelect = document.getElementById("customerSelect");
    if (custSelect) {
      custSelect.addEventListener("change", (e) => {
        const cid = parseInt(e.target.value, 10);
        const cust = (customersCache || []).find(c => c.id === cid);
        if (cust) {
          applySelectedCustomerToQuoteUI(cust);
        }
      });
    }

    // 4. 手機端客戶詳細資料展開/收合
    const btnToggleCustDetail = document.getElementById("btnToggleCustDetail");
    const customerDetailBody = document.getElementById("customerDetailBody");
    const custDetailArrow = document.getElementById("custDetailArrow");
    if (btnToggleCustDetail && customerDetailBody) {
      btnToggleCustDetail.addEventListener("click", () => {
        customerDetailBody.classList.toggle("hidden");
        const isHidden = customerDetailBody.classList.contains("hidden");
        if (custDetailArrow) custDetailArrow.textContent = isHidden ? "▼ 展開" : "▲ 收合";
      });
    }

    // 5. 電腦端與手機端關聯 CRM 商機按鈕
    const btnLinkCrm = document.getElementById("btnLinkCrmCase");
    if (btnLinkCrm) {
      btnLinkCrm.addEventListener("click", () => {
        if (dealsCache && dealsCache.length > 0) {
          const names = dealsCache.slice(0, 8).map((d, i) => `${i + 1}. ${d.project_name || d.name} (${d.customer_name})`).join("\n");
          const choice = prompt(`請輸入要關聯的商機編號 (1~${Math.min(dealsCache.length, 8)}):\n${names}`);
          const idx = parseInt(choice, 10) - 1;
          if (dealsCache[idx]) {
            const selectedDeal = dealsCache[idx];
            const badge = document.getElementById("linkedDealBadge");
            if (badge) {
              badge.textContent = `已關聯：${selectedDeal.project_name || selectedDeal.name}`;
              badge.classList.remove("hidden");
            }
            const projInput = document.getElementById("quoteProjectName");
            if (projInput && !projInput.value) {
              projInput.value = selectedDeal.project_name || selectedDeal.name || "";
            }
          }
        } else {
          alert("[商機清單空] 目前無可關聯的 Deal 商機。");
        }
      });
    }

    const btnMobileLinkCrm = document.getElementById("btnMobileLinkCrmCase");
    const crmWrap = document.getElementById("crmLinkedCasesWrap");
    const crmSelect = document.getElementById("crmCaseSelect");
    if (btnMobileLinkCrm && crmWrap && crmSelect) {
      btnMobileLinkCrm.addEventListener("click", async () => {
        if (!dealsCache || dealsCache.length === 0) {
          try {
            const r = await fetch("/api/deals");
            if (r.ok) dealsCache = await r.json();
          } catch (e) {}
        }
        crmSelect.innerHTML = '<option value="">-- 不關聯特定 CRM 案件 --</option>' +
          (dealsCache || []).map(d => `<option value="${escapeHtml(d.project_name || '')}" data-cust="${escapeHtml(d.customer_name || '')}">${escapeHtml(d.customer_name || '')} - ${escapeHtml(d.project_name || '未命名案件')}</option>`).join('');
        crmWrap.classList.toggle("hidden");
      });
      crmSelect.addEventListener("change", () => {
        const val = crmSelect.value;
        const projInput = document.getElementById("quoteProjectName");
        if (projInput && val) projInput.value = val;
      });
    }

    // 6. 電腦端表格上方動作按鈕 (新建報價單、儲存報價、複製 LINE 報價、匯出至Excel、列印 PDF)
    const btnNewTop = document.getElementById("btnNewQuoteTop");
    if (btnNewTop) {
      btnNewTop.addEventListener("click", () => {
        if (confirm("確定要開立新的空白報價單嗎？未儲存的內容將會重設。")) {
          resetQuotationEditor();
        }
      });
    }
    const btnSaveTop = document.getElementById("btnSaveQuoteTop");
    if (btnSaveTop) btnSaveTop.addEventListener("click", saveFormalQuotation);
    const btnLineTop = document.getElementById("btnLineQuoteTop");
    if (btnLineTop) btnLineTop.addEventListener("click", openLineQuoteModal);
    const btnExportTop = document.getElementById("btnExportExcelTop");
    if (btnExportTop) btnExportTop.addEventListener("click", exportQuoteToExcel);
    const btnPrintTop = document.getElementById("btnPrintPdfTop");
    if (btnPrintTop) btnPrintTop.addEventListener("click", printQuotePDF);

    // 取消列選取
    const btnClearActive = document.getElementById("btnClearActiveRow");
    if (btnClearActive) {
      btnClearActive.addEventListener("click", () => {
        const badge = document.getElementById("activeRowBadge");
        if (badge) badge.textContent = "第 1 列 | (尚未選擇產品)";
        const desktopTableBody = document.getElementById("quoteDesktopTableBody");
        if (desktopTableBody) {
          desktopTableBody.querySelectorAll("tr").forEach(tr => tr.classList.remove("active-row"));
        }
      });
    }

    // 7. 表格下方動作按鈕 (+ 新增產品明細, + 插入方案小計, 舊單匯入, 智慧BOM, 重整經銷價)
    const btnAddNew = document.getElementById("btnAddNewItem");
    if (btnAddNew) btnAddNew.addEventListener("click", () => addBlankItem());
    const btnInsertSub = document.getElementById("btnInsertSubtotal");
    if (btnInsertSub) btnInsertSub.addEventListener("click", () => insertSubtotalRow());

    const btnAddItem = document.getElementById("btnAddItem");
    if (btnAddItem) btnAddItem.addEventListener("click", () => addBlankItem());

    const btnImportOld = document.getElementById("btnImportOldQuote");
    if (btnImportOld) btnImportOld.addEventListener("click", openImportOldQuoteModal);

    const btnImportBom = document.getElementById("btnImportBom");
    if (btnImportBom) btnImportBom.addEventListener("click", openSmartBOMModal);
    const btnImportSmartBOM = document.getElementById("btnImportSmartBOM");
    if (btnImportSmartBOM) btnImportSmartBOM.addEventListener("click", openSmartBOMModal);

    const btnRefreshPrices = document.getElementById("btnRefreshPrices");
    if (btnRefreshPrices) btnRefreshPrices.addEventListener("click", refreshDealerPrices);
    const btnRefreshDealerPrices = document.getElementById("btnRefreshDealerPrices");
    if (btnRefreshDealerPrices) btnRefreshDealerPrices.addEventListener("click", refreshDealerPrices);

    // 8. 附加圖片上傳監聽
    const imgUploader = document.getElementById("quoteImageUploader");
    if (imgUploader) {
      imgUploader.addEventListener("change", handleImageUpload);
    }

    // 9. 手機端底部操作按鈕：送出草稿 (直接寫入資料庫手機報價草稿匣) / 複製 LINE 報價
    const btnSubmitDraft = document.getElementById("btnSubmitDraft");
    if (btnSubmitDraft) btnSubmitDraft.addEventListener("click", submitMobileQuotationDraftToDB);

    const btnDesktopSubmit = document.getElementById("btnDesktopSubmitDraft");
    if (btnDesktopSubmit) btnDesktopSubmit.addEventListener("click", saveFormalQuotation);

    const btnLineQuote = document.getElementById("btnLineQuote");
    if (btnLineQuote) btnLineQuote.addEventListener("click", openLineQuoteModal);

    const btnDesktopLine = document.getElementById("btnDesktopLineQuote");
    if (btnDesktopLine) btnDesktopLine.addEventListener("click", openLineQuoteModal);

    // 10. 歷史分頁 (Tab B) 垃圾桶切換與搜尋
    const inputSearchHistory = document.getElementById("inputSearchQuoteHistory");
    if (inputSearchHistory) {
      inputSearchHistory.addEventListener("input", (e) => {
        loadQuotationHistory(e.target.value.trim(), isTrashViewActive);
      });
    }

    const btnToggleTrash = document.getElementById("btnToggleTrashView");
    if (btnToggleTrash) {
      btnToggleTrash.addEventListener("click", () => {
        isTrashViewActive = !isTrashViewActive;
        const txt = document.getElementById("trashToggleText");
        if (txt) {
          txt.textContent = isTrashViewActive ? "返回正式報價列表" : "檢視垃圾桶 (保留 30 天)";
        }
        loadQuotationHistory(inputSearchHistory?.value.trim() || "", isTrashViewActive);
      });
    }

    const btnRefreshHistory = document.getElementById("btnRefreshQuoteHistory");
    if (btnRefreshHistory) {
      btnRefreshHistory.addEventListener("click", () => {
        loadQuotationHistory(inputSearchHistory?.value.trim() || "", isTrashViewActive);
      });
    }

    // 11. 手機報價草稿匣清空按鈕 (Tab C) - 二次確認防呆
    const btnClearDrafts = document.getElementById("btnClearDraftsInbox");
    if (btnClearDrafts) {
      btnClearDrafts.addEventListener("click", async () => {
        const activeSales = getActiveSalesName();
        if (!confirm(`【二次確認】確定要清空業務【${activeSales}】名下所有的手機暫存草稿嗎？此動作無法復原。`)) {
          return;
        }
        try {
          await fetch(`/api/quotation-drafts?sales_name=${encodeURIComponent(activeSales)}`, { method: "DELETE" });
          localStorage.removeItem("shihlin_mobile_quote_drafts");
          loadQuotationDraftsInbox();
        } catch (err) {
          alert("[清空草稿失敗] " + err.message);
        }
      });
    }

    // 12. 歷史單價查詢按鈕 (Tab D)
    const btnRunAnalysis = document.getElementById("btnRunPriceAnalysis");
    if (btnRunAnalysis) {
      btnRunAnalysis.addEventListener("click", runPriceAnalysis);
    }

    // 13. 個人特殊品新增/更新按鈕 (Tab E)
    const btnSaveSpecial = document.getElementById("btnSaveSpecialProduct");
    if (btnSaveSpecial) {
      btnSaveSpecial.addEventListener("click", savePersonalSpecialProductFromUI);
    }

    // 14. 綁定左側【報價管理系統】專屬 9 大功能頁按鈕與底部【離開報價單功能】按鈕
    document.querySelectorAll("#quoteSidebarNavList .quote-nav-btn").forEach(btn => {
      btn.addEventListener("click", () => {
        const page = btn.getAttribute("data-quote-page") || "create";
        switchQuoteSubTab(page);
      });
    });

    const btnExitQuoteWebTop = document.getElementById("btnExitQuotationWebTop");
    if (btnExitQuoteWebTop) {
      btnExitQuoteWebTop.addEventListener("click", () => switchPage(navItemHome, panelHome));
    }
    const btnExitQuoteWebBottom = document.getElementById("btnExitQuotationWebBottom");
    if (btnExitQuoteWebBottom) {
      btnExitQuoteWebBottom.addEventListener("click", () => switchPage(navItemHome, panelHome));
    }

    const btnExitQuote = document.getElementById("btnExitQuotationSystem");
    if (btnExitQuote) {
      btnExitQuote.addEventListener("click", () => {
        switchPage(navItemHome, panelHome);
      });
    }

    // 15. 報價單管理狀態篩選下拉選單
    const selHistStatus = document.getElementById("selectQuoteHistoryStatus");
    if (selHistStatus) {
      selHistStatus.addEventListener("change", () => {
        const kw = document.getElementById("inputSearchQuoteHistory")?.value.trim() || "";
        loadQuotationHistory(kw, isTrashViewActive);
      });
    }

    // 16. 大數據分析重新統計與最新版切換
    const btnRefreshDash = document.getElementById("btnRefreshQuoteDashboard");
    if (btnRefreshDash) btnRefreshDash.addEventListener("click", loadQuotationDashboard);
    const chkDashLatest = document.getElementById("chkQuoteDashLatestOnly");
    if (chkDashLatest) chkDashLatest.addEventListener("change", loadQuotationDashboard);

    // 17. 客戶管理頁面事件綁定
    const inputCustSearch = document.getElementById("inputQuoteCustSearch");
    if (inputCustSearch) {
      inputCustSearch.addEventListener("input", () => renderQuotationCustomersTable());
    }
    const btnCustClear = document.getElementById("btnQuoteCustClear");
    if (btnCustClear) btnCustClear.addEventListener("click", clearQuotationCustomerForm);
    const btnCustSave = document.getElementById("btnQuoteCustSave");
    if (btnCustSave) btnCustSave.addEventListener("click", () => saveQuotationCustomerFromUI(false));
    const btnCustCopy = document.getElementById("btnQuoteCustCopyContact");
    if (btnCustCopy) btnCustCopy.addEventListener("click", () => saveQuotationCustomerFromUI(true));
    const btnCustDel = document.getElementById("btnQuoteCustDelete");
    if (btnCustDel) btnCustDel.addEventListener("click", deleteSelectedQuotationCustomer);
    const btnSupSave = document.getElementById("btnQuoteSupplierSave");
    if (btnSupSave) btnSupSave.addEventListener("click", saveQuotationSupplierFromUI);

    // 18. 產品管理：共通標準牌價搜尋與快速折數計算機
    const btnProdSearch = document.getElementById("btnQuoteProdMgmtSearch");
    if (btnProdSearch) btnProdSearch.addEventListener("click", searchStandardProductsInMgmt);
    const inputProdSearch = document.getElementById("inputQuoteProdMgmtSearch");
    if (inputProdSearch) {
      inputProdSearch.addEventListener("keydown", (e) => {
        if (e.key === "Enter") {
          e.preventDefault();
          searchStandardProductsInMgmt();
        }
      });
    }
    const btnQuickCalc = document.getElementById("btnRunQuickCalc");
    if (btnQuickCalc) btnQuickCalc.addEventListener("click", runQuickExpressionCalc);
    const inputCalcBase = document.getElementById("inputCalcBasePrice");
    const inputCalcDisc = document.getElementById("inputCalcDiscountPct");
    if (inputCalcBase) inputCalcBase.addEventListener("input", updateQuickDiscountCalc);
    if (inputCalcDisc) inputCalcDisc.addEventListener("input", updateQuickDiscountCalc);

    // 19. 系統設定：儲存財務設定與報價條款
    const btnSaveSet = document.getElementById("btnSaveQuoteSettings");
    if (btnSaveSet) btnSaveSet.addEventListener("click", saveQuotationFinanceSettingsFromUI);
    const btnSaveTerms = document.getElementById("btnSaveQuoteTerms");
    if (btnSaveTerms) btnSaveTerms.addEventListener("click", saveQuotationTermsFromUI);
  }

  function resetQuotationEditor() {
    const desktopTableBody = document.getElementById("quoteDesktopTableBody");
    const itemsContainer = document.getElementById("itemsContainer");
    if (desktopTableBody) desktopTableBody.innerHTML = "";
    if (itemsContainer) itemsContainer.innerHTML = "";
    quoteItemCount = 0;
    currentBaseQuoteNo = null;
    currentLoadedDraftId = null;
    currentSelectedCustomer = null;

    const custSel = document.getElementById("customerSelect");
    if (custSel) custSel.value = "";
    const custInput = document.getElementById("customerName");
    if (custInput) custInput.value = "";
    const prefixInput = document.getElementById("quotePrefixInput");
    if (prefixInput) prefixInput.value = "-- 請選擇客戶 --";
    const taxInput = document.getElementById("quoteTaxId");
    if (taxInput) taxInput.value = "";
    const phoneInput = document.getElementById("quotePhone");
    if (phoneInput) phoneInput.value = "";
    const faxInput = document.getElementById("quoteFax");
    if (faxInput) faxInput.value = "";
    const locInput = document.getElementById("quoteDeliveryLocation");
    if (locInput) locInput.value = "";
    const projInput = document.getElementById("quoteProjectName");
    if (projInput) projInput.value = "";
    const detailCard = document.getElementById("customerDetailCard");
    if (detailCard) detailCard.classList.add("hidden");

    addBlankItem();
    calculateTotal();
  }

  // 視窗寬度自適應模式守護 (>= 768px 電腦端 Excel 試算表，< 768px 手機端圖三卡片)
  function enforceDesktopQuotationView() {
    const draftContainer = document.getElementById("quoteDraftSection");
    if (!draftContainer) return;
    if (window.innerWidth >= 768) {
      draftContainer.classList.remove("force-mobile-view");
    }
  }
  window.addEventListener("resize", enforceDesktopQuotationView);

  // 新增品項 (100% 同步手機端圖三卡片與電腦端圖二 10 欄試算表)
  function addBlankItem(initialData = null) {
    quoteItemCount++;
    const itemId = `item-${quoteItemCount}`;
    const itemsContainer = document.getElementById("itemsContainer");
    const desktopTableBody = document.getElementById("quoteDesktopTableBody");
    if (!itemsContainer) return;

    const currentCount = itemsContainer.querySelectorAll('.item-row').length + 1;
    const modelText = initialData?.model_no || "點擊選擇產品...";
    const hasModel = Boolean(initialData?.model_no);
    const specVal = initialData?.spec || "";
    const skuVal = initialData?.sku || "";
    const isOrderItem = Boolean(initialData?.is_order_item);
    const isSpecialProduct = Boolean(initialData?.is_special);

    const stockQtyVal = initialData?.stock_qty !== undefined && initialData?.stock_qty !== null
      ? initialData.stock_qty
      : (hasModel ? 0 : null);
    const stockDisplay = stockQtyVal !== null ? `${Number(stockQtyVal).toLocaleString()} 台` : "-";
    const desktopStockDisplay = stockQtyVal !== null ? `${Number(stockQtyVal).toLocaleString()}` : "-";

    const dealerPriceVal = parseFloat(initialData?.dealer_price) || 0;
    const dealerPriceDisplay = dealerPriceVal > 0 ? `$${dealerPriceVal.toLocaleString()}` : "-";

    const hasDiscount = initialData && initialData.discount_pct !== undefined && initialData.discount_pct !== null && initialData.discount_pct !== "";
    const discountVal = hasDiscount ? initialData.discount_pct : (hasModel ? 100 : "");

    let finalUnitPriceVal = 0;
    if (initialData && initialData.final_unit_price !== undefined && initialData.final_unit_price !== null && initialData.final_unit_price !== "") {
      finalUnitPriceVal = parseFloat(initialData.final_unit_price) || 0;
    } else if (dealerPriceVal > 0 && discountVal !== "") {
      finalUnitPriceVal = applyQuoteRounding(dealerPriceVal * (parseFloat(discountVal) / 100));
    }

    const sugPriceVal = finalUnitPriceVal > 0 ? finalUnitPriceVal : "";
    const finalPriceDisplay = finalUnitPriceVal > 0 ? `$${finalUnitPriceVal.toLocaleString()}` : "-";
    const qtyVal = parseInt(initialData?.quantity, 10) || 1;
    const rowSubtotalVal = initialData?.row_subtotal !== undefined
      ? (parseFloat(initialData.row_subtotal) || Math.round(finalUnitPriceVal * qtyVal))
      : Math.round(finalUnitPriceVal * qtyVal);
    const leadTimeVal = initialData?.note_lead_time || "";

    const stockUpdatedAtText = quoteFinanceSettings.stock_updated_at || "10/15 10:12";
    const roundingRuleText = quoteFinanceSettings.round_summary_text || "捨入：百位數 / 無條件捨去";

    // 1. 手機端卡片 HTML (100% 還原圖三配置：庫存數量、經銷價、黃色庫存提示框、建議折數、建議報價+捨入提示、需求數量、最終售價+未稅金額提示、交期/備註)
    const mobileItemHTML = `
      <div class="item-row" id="${itemId}" data-spec="${escapeHtml(specVal)}" data-stock="${stockQtyVal !== null ? stockQtyVal : ''}">
        <div class="item-header">
          <span class="item-index-badge">#${currentCount}</span>
          <button type="button" class="product-picker-input ${hasModel ? 'selected' : ''}" id="${itemId}-title">
            ${escapeHtml(modelText)}
          </button>
          <span class="order-item-badge ${isOrderItem ? '' : 'hidden'}" id="${itemId}-order-badge">受訂生產</span>
          <span class="special-product-tag ${isSpecialProduct ? '' : 'hidden'}" id="${itemId}-special-badge">個人特殊品</span>
          <button type="button" class="item-copy-btn ${hasModel ? '' : 'hidden'}" id="${itemId}-copy-btn" title="複製型號">
            <svg width="16" height="16" viewBox="0 0 16 16" fill="currentColor">
              <path d="M4 2a2 2 0 0 1 2-2h6a2 2 0 0 1 2 2v8a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V2zm2-1a1 1 0 0 0-1 1v8a1 1 0 0 0 1 1h6a1 1 0 0 0 1-1V2a1 1 0 0 0-1-1H6zM2 5a1 1 0 0 0-1 1v8a1 1 0 0 0 1 1h6a1 1 0 0 0 1-1v-1h1v1a2 2 0 0 1-2 2H2a2 2 0 0 1-2-2V6a2 2 0 0 1 2-2h1v1H2z"/>
            </svg>
          </button>
          <button type="button" class="item-remove" id="${itemId}-remove" title="刪除品項">&times;</button>
        </div>

        <div class="item-grid">
          <div>
            <label>庫存數量</label>
            <input type="text" id="${itemId}-stock" placeholder="-" readonly class="field-readonly stock-field" value="${escapeHtml(stockDisplay)}">
          </div>
          <div>
            <label>經銷價(未稅)</label>
            <input type="text" id="${itemId}-dealer-price" placeholder="-" readonly class="field-readonly price-field" value="${escapeHtml(dealerPriceDisplay)}">
          </div>

          <div class="stock-disclaimer-box">
            <div class="stock-disclaimer-time">
              <svg class="icon-svg-xs" style="width:13px; height:13px; margin-right:3px; vertical-align:middle;" viewBox="0 0 24 24"><path d="M11.99 2C6.47 2 2 6.48 2 12s4.47 10 9.99 10C17.52 22 22 17.52 22 12S17.52 2 11.99 2zM12 20c-4.42 0-8-3.58-8-8s3.58-8 8-8 8 3.58 8 8-3.58 8-8 8zm.5-13H11v6l5.25 3.15.75-1.23-4.5-2.67z"/></svg>
              <span>庫存更新時間：<span class="stock-update-time-text">${escapeHtml(stockUpdatedAtText)}</span></span>
              <span id="${itemId}-sku-wrap" class="${skuVal ? '' : 'hidden'}">
                ｜ 料號：<span id="${itemId}-sku-text" style="font-family:monospace;">${escapeHtml(skuVal)}</span>
              </span>
            </div>
            <div class="stock-disclaimer-text">[注意] 庫存非即時數量，若庫存數量偏低或屬大手案件，請再次查詢確認實際庫存。</div>
          </div>

          <div>
            <label>建議折數(%)</label>
            <input type="number" id="${itemId}-sug-discount" placeholder="輸入折數" step="any" value="${escapeHtml(String(discountVal))}">
          </div>
          <div>
            <label>建議報價(元)</label>
            <input type="number" id="${itemId}-sug-price" placeholder="輸入報價" step="any" value="${escapeHtml(String(sugPriceVal))}">
            <small class="field-hint-rounding">[提示] <span class="rounding-rule-text">${escapeHtml(roundingRuleText)}</span></small>
          </div>

          <div>
            <label>需求數量</label>
            <input type="number" id="${itemId}-qty" min="1" value="${qtyVal}" required>
          </div>
          <div>
            <label>最終售價(元)</label>
            <input type="text" id="${itemId}-final-price" placeholder="-" readonly class="field-readonly final-price-field" value="${escapeHtml(finalPriceDisplay)}">
            <small class="field-hint-tax">[提示] 未稅金額</small>
          </div>
        </div>

        <div class="item-delivery">
          <label>交期 / 備註</label>
          <input type="text" id="${itemId}-delivery" placeholder="點擊選擇交期 / 備註..." readonly style="cursor:pointer;" value="${escapeHtml(leadTimeVal)}">
        </div>

        <input type="hidden" id="${itemId}-code" value="${escapeHtml(initialData?.model_no || '')}">
        <input type="hidden" id="${itemId}-sku" value="${escapeHtml(skuVal)}">
        <input type="hidden" id="${itemId}-unit-price-num" value="${finalUnitPriceVal}">
      </div>
    `;
    itemsContainer.insertAdjacentHTML('beforeend', mobileItemHTML);

    // 2. 電腦端 Excel 試算表表格列 HTML (_QuotationApp 圖二 1:1 版面)
    if (desktopTableBody) {
      const desktopRowHTML = `
        <tr id="d-${itemId}" class="desktop-item-row" data-itemid="${itemId}">
          <td style="text-align:center; font-weight:700; color:var(--text-muted);" id="d-${itemId}-index">
            ${currentCount}
          </td>
          <td class="cell-stock-gray" id="d-${itemId}-stock">
            ${escapeHtml(desktopStockDisplay)}
          </td>
          <td>
            <button type="button" class="cell-model-combobox ${hasModel ? 'selected' : ''}" id="d-${itemId}-model">
              <span id="d-${itemId}-model-text">${escapeHtml(modelText)}</span>
              <span class="dropdown-arrow">▼</span>
            </button>
            <input type="hidden" id="d-${itemId}-code" value="${escapeHtml(initialData?.model_no || '')}">
          </td>
          <td>
            <input type="number" id="d-${itemId}-qty" class="table-cell-input number" min="1" value="${qtyVal}" style="width:100%;">
          </td>
          <td>
            <span id="d-${itemId}-spec" style="font-size:0.82rem; color:var(--text-secondary);">${escapeHtml(specVal || '-')}</span>
          </td>
          <td style="text-align:right;">
            <span id="d-${itemId}-dealer-price" class="table-cell-readonly money">${escapeHtml(dealerPriceVal > 0 ? '$ ' + dealerPriceVal.toLocaleString() : '-')}</span>
            <input type="hidden" id="d-${itemId}-price" value="${finalUnitPriceVal}">
          </td>
          <td>
            <input type="number" id="d-${itemId}-discount" class="table-cell-input number" step="any" value="${discountVal !== '' ? discountVal : 100}" style="width:100%;">
          </td>
          <td style="text-align:right;">
            <span id="d-${itemId}-subtotal" class="table-cell-readonly money" style="font-weight:700; color:#1e293b;">
              ${rowSubtotalVal > 0 ? '$ ' + rowSubtotalVal.toLocaleString() : '$ 0'}
            </span>
          </td>
          <td>
            <input type="text" id="d-${itemId}-delivery" class="table-cell-input" placeholder="點擊選取交期..." readonly style="cursor:pointer;" value="${escapeHtml(leadTimeVal)}">
          </td>
          <td style="text-align:center;">
            <button type="button" class="table-btn-remove" id="d-${itemId}-remove" title="刪除品項">&times;</button>
          </td>
        </tr>
      `;
      desktopTableBody.insertAdjacentHTML('beforeend', desktopRowHTML);

      const dRow = document.getElementById(`d-${itemId}`);
      if (dRow) {
        dRow.addEventListener("click", (e) => {
          if (e.target.closest("button") || e.target.closest("input")) return;
          desktopTableBody.querySelectorAll("tr").forEach(tr => tr.classList.remove("active-row"));
          dRow.classList.add("active-row");
          const idx = dRow.querySelector('[id$="-index"]')?.textContent?.trim() || "1";
          const model = document.getElementById(`${itemId}-code`)?.value || "";
          const badge = document.getElementById("activeRowBadge");
          if (badge) badge.textContent = `第 ${idx} 列 | ${model || '(尚未選擇產品)'}`;
        });
      }
    }

    refreshItemIndices();
    setupItemRowEvents(itemId);
    calculateTotal();
  }

  // 插入方案小計列 (橘色醒目分隔列，適配 10 欄桌面版面)
  function insertSubtotalRow(defaultTitle = "") {
    const desktopTableBody = document.getElementById("quoteDesktopTableBody");
    if (!desktopTableBody) return;

    const rowId = `subtotal-row-${Date.now()}`;
    const title = defaultTitle || `方案小計 (區段加總)`;

    const subtotalHTML = `
      <tr id="${rowId}" class="subtotal-divider-row">
        <td style="text-align:center;">
          <span class="subtotal-badge">小計</span>
        </td>
        <td colspan="6">
          <input type="text" class="subtotal-title-input" value="${escapeHtml(title)}" placeholder="輸入方案名稱 (如：方案 A：主驅動系統小計)">
        </td>
        <td style="text-align:right;">
          <span class="segment-subtotal-val" style="font-family:Consolas, monospace; font-size:1rem; font-weight:700; color:#c2410c;">$ 0</span>
        </td>
        <td>
          <span style="font-size:0.75rem; color:#9a3412;">[區段未稅金額]</span>
        </td>
        <td style="text-align:center;">
          <button type="button" class="table-btn-remove btn-remove-subtotal" title="刪除此小計列">&times;</button>
        </td>
      </tr>
    `;
    desktopTableBody.insertAdjacentHTML('beforeend', subtotalHTML);
    const rowEl = document.getElementById(rowId);
    const rmBtn = rowEl?.querySelector(".btn-remove-subtotal");
    if (rmBtn && rowEl) {
      rmBtn.addEventListener("click", () => {
        rowEl.remove();
        calculateTotal();
      });
    }
    calculateTotal();
  }

  // 更新各方案小計區段數值
  function updateSubtotalDividers() {
    const desktopTableBody = document.getElementById("quoteDesktopTableBody");
    if (!desktopTableBody) return;

    const allRows = desktopTableBody.querySelectorAll('tr');
    let currentSegmentSum = 0;

    allRows.forEach(row => {
      if (row.classList.contains('subtotal-divider-row')) {
        const segValSpan = row.querySelector('.segment-subtotal-val');
        if (segValSpan) {
          segValSpan.textContent = `$ ${currentSegmentSum.toLocaleString()}`;
        }
        currentSegmentSum = 0;
      } else if (row.classList.contains('desktop-item-row')) {
        const itemId = row.getAttribute('data-itemid');
        const price = parseFloat(document.getElementById(`d-${itemId}-price`)?.value) || 0;
        const qty = parseInt(document.getElementById(`d-${itemId}-qty`)?.value, 10) || 0;
        currentSegmentSum += Math.round(price * qty);
      }
    });
  }

  function setupItemRowEvents(itemId) {
    // 1. 手機端元件
    const mTitleBtn = document.getElementById(`${itemId}-title`);
    const mCopyBtn = document.getElementById(`${itemId}-copy-btn`);
    const mRemoveBtn = document.getElementById(`${itemId}-remove`);
    const mDeliveryInput = document.getElementById(`${itemId}-delivery`);
    const mSugPriceInput = document.getElementById(`${itemId}-sug-price`);
    const mSugDiscountInput = document.getElementById(`${itemId}-sug-discount`);
    const mDealerPriceInput = document.getElementById(`${itemId}-dealer-price`);
    const mQtyInput = document.getElementById(`${itemId}-qty`);
    const mFinalPriceInput = document.getElementById(`${itemId}-final-price`);
    const mUnitPriceNum = document.getElementById(`${itemId}-unit-price-num`);

    // 2. 電腦端表格元件
    const dModelBtn = document.getElementById(`d-${itemId}-model`);
    const dRemoveBtn = document.getElementById(`d-${itemId}-remove`);
    const dDeliveryInput = document.getElementById(`d-${itemId}-delivery`);
    const dPriceInput = document.getElementById(`d-${itemId}-price`);
    const dDiscountInput = document.getElementById(`d-${itemId}-discount`);
    const dQtyInput = document.getElementById(`d-${itemId}-qty`);
    const dSubtotalSpan = document.getElementById(`d-${itemId}-subtotal`);

    function triggerProductPicker() {
      currentEditingItemId = itemId;
      openProductModal();
    }
    if (mTitleBtn) mTitleBtn.addEventListener("click", triggerProductPicker);
    if (dModelBtn) dModelBtn.addEventListener("click", triggerProductPicker);

    function triggerCopyModel(btnEl) {
      const model = document.getElementById(`${itemId}-code`)?.value;
      if (model) {
        navigator.clipboard.writeText(model).then(() => {
          if (btnEl) btnEl.style.color = "#059669";
          setTimeout(() => { if (btnEl) btnEl.style.color = ""; }, 1200);
        });
      }
    }
    if (mCopyBtn) mCopyBtn.addEventListener("click", () => triggerCopyModel(mCopyBtn));

    function triggerRemoveItem() {
      const itemsContainer = document.getElementById("itemsContainer");
      const allMobileRows = itemsContainer ? itemsContainer.querySelectorAll('.item-row') : [];
      if (allMobileRows.length <= 1) {
        alert("報價單至少需保留一項產品欄位！");
        return;
      }
      const mRow = document.getElementById(itemId);
      const dRow = document.getElementById(`d-${itemId}`);
      if (mRow) mRow.remove();
      if (dRow) dRow.remove();
      calculateTotal();
      refreshItemIndices();
    }
    if (mRemoveBtn) mRemoveBtn.addEventListener("click", triggerRemoveItem);
    if (dRemoveBtn) dRemoveBtn.addEventListener("click", triggerRemoveItem);

    function triggerDeliveryPicker() {
      currentEditingItemId = itemId;
      openDeliveryModal();
    }
    if (mDeliveryInput) mDeliveryInput.addEventListener("click", triggerDeliveryPicker);
    if (dDeliveryInput) dDeliveryInput.addEventListener("click", triggerDeliveryPicker);

    function getDealerPrice() {
      const dpStr = (document.getElementById(`d-${itemId}-dealer-price`)?.textContent || mDealerPriceInput?.value || "").replace(/[^0-9.]/g, '');
      return parseFloat(dpStr) || 0;
    }

    function syncUnitAndRowSubtotal(finalUnitPrice) {
      const unitP = Math.max(0, Math.round(finalUnitPrice || 0));
      if (mUnitPriceNum) mUnitPriceNum.value = unitP;
      if (dPriceInput) dPriceInput.value = unitP;
      if (mFinalPriceInput) {
        mFinalPriceInput.value = unitP > 0 ? `$${unitP.toLocaleString()}` : "-";
      }
      const q = parseInt(dQtyInput?.value || mQtyInput?.value || "1", 10) || 1;
      const sub = Math.round(unitP * q);
      if (dSubtotalSpan) {
        dSubtotalSpan.textContent = sub > 0 ? `$ ${sub.toLocaleString()}` : "$ 0";
      }
      calculateTotal();
    }

    // 折數變更 -> 計算建議報價與捨入後最終售價
    function onDiscountChanged(val, fromSource) {
      const disc = parseFloat(val);
      const dp = getDealerPrice();

      if (fromSource === 'desktop' && mSugDiscountInput) mSugDiscountInput.value = val;
      if (fromSource === 'mobile' && dDiscountInput) dDiscountInput.value = val;

      if (!isNaN(disc) && dp > 0) {
        const rawPrice = dp * (disc / 100);
        const roundedPrice = applyQuoteRounding(rawPrice);
        if (mSugPriceInput) mSugPriceInput.value = Math.round(rawPrice);
        syncUnitAndRowSubtotal(roundedPrice);
      } else {
        syncUnitAndRowSubtotal(0);
      }
    }

    // 建議報價變更 -> 反推折數並計算捨入後最終售價
    function onSugPriceChanged(val) {
      const rawInput = parseFloat(val);
      const dp = getDealerPrice();
      if (!isNaN(rawInput) && rawInput > 0) {
        const roundedPrice = dp > 0 ? applyQuoteRounding(rawInput) : Math.round(rawInput);
        if (dp > 0) {
          const discStr = ((roundedPrice / dp) * 100).toFixed(1).replace(/\.0$/, '');
          if (mSugDiscountInput) mSugDiscountInput.value = discStr;
          if (dDiscountInput) dDiscountInput.value = discStr;
        }
        syncUnitAndRowSubtotal(roundedPrice);
      } else {
        syncUnitAndRowSubtotal(0);
      }
    }

    function onQtyChanged(val, fromSource) {
      if (fromSource === 'desktop' && mQtyInput) mQtyInput.value = val;
      if (fromSource === 'mobile' && dQtyInput) dQtyInput.value = val;
      const unitP = parseFloat(mUnitPriceNum?.value || dPriceInput?.value || "0") || 0;
      syncUnitAndRowSubtotal(unitP);
    }

    if (mSugDiscountInput) mSugDiscountInput.addEventListener("input", (e) => onDiscountChanged(e.target.value, 'mobile'));
    if (dDiscountInput) dDiscountInput.addEventListener("input", (e) => onDiscountChanged(e.target.value, 'desktop'));
    if (mSugPriceInput) mSugPriceInput.addEventListener("input", (e) => onSugPriceChanged(e.target.value));
    if (mQtyInput) mQtyInput.addEventListener("input", (e) => onQtyChanged(e.target.value, 'mobile'));
    if (dQtyInput) dQtyInput.addEventListener("input", (e) => onQtyChanged(e.target.value, 'desktop'));
  }

  // 核心計算總計：雙向同步手機端小計卡片與電腦端結算總額
  function calculateTotal() {
    const itemsContainer = document.getElementById("itemsContainer");
    let subtotal = 0;
    let itemCount = 0;

    const mRows = itemsContainer ? itemsContainer.querySelectorAll('.item-row') : [];
    mRows.forEach(r => {
      itemCount++;
      const itemId = r.id;
      const unitPriceEl = document.getElementById(`${itemId}-unit-price-num`);
      const dPriceEl = document.getElementById(`d-${itemId}-price`);
      const qtyEl = document.getElementById(`${itemId}-qty`) || document.getElementById(`d-${itemId}-qty`);

      const unitPrice = parseFloat(unitPriceEl?.value || dPriceEl?.value || "0") || 0;
      const qty = parseInt(qtyEl?.value || "1", 10) || 1;
      const rowSub = Math.round(unitPrice * qty);
      subtotal += rowSub;

      const subSpan = document.getElementById(`d-${itemId}-subtotal`);
      if (subSpan) subSpan.textContent = rowSub > 0 ? `$ ${rowSub.toLocaleString()}` : "$ 0";
      const mFinal = document.getElementById(`${itemId}-final-price`);
      if (mFinal) mFinal.value = unitPrice > 0 ? `$${unitPrice.toLocaleString()}` : "-";
    });

    const taxRate = (parseFloat(quoteFinanceSettings.tax_rate) || 5.0) / 100.0;
    const tax = Math.round(subtotal * taxRate);
    const grandTotal = subtotal + tax;

    // 1. 更新手機端小計卡片
    const mBadge = document.getElementById("draftItemBadge");
    if (mBadge) mBadge.textContent = `共 ${itemCount} 項`;
    const mTotal = document.getElementById("totalAmount");
    if (mTotal) mTotal.textContent = subtotal.toLocaleString();

    // 2. 更新電腦端相容結算 IDs
    const dBadge = document.getElementById("desktopItemCountBadge");
    if (dBadge) dBadge.textContent = `共 ${itemCount} 項`;
    const dSubtotal = document.getElementById("desktopSubtotal");
    if (dSubtotal) dSubtotal.textContent = `$ ${subtotal.toLocaleString()}`;
    const dTax = document.getElementById("desktopTax");
    if (dTax) dTax.textContent = `$ ${tax.toLocaleString()}`;
    const dGrandTotal = document.getElementById("desktopGrandTotal");
    if (dGrandTotal) dGrandTotal.textContent = `$ ${grandTotal.toLocaleString()}`;

    // 3. 更新 _QuotationApp 左下角結算總額 (合計未稅、營業稅5%、總金額含稅紅字)
    const sumSub = document.getElementById("summarySubtotal");
    if (sumSub) sumSub.textContent = subtotal.toLocaleString();
    const sumTax = document.getElementById("summaryTax");
    if (sumTax) sumTax.textContent = tax.toLocaleString();
    const sumGrand = document.getElementById("summaryGrandTotal");
    if (sumGrand) sumGrand.textContent = grandTotal.toLocaleString();

    // 4. 連動更新方案小計
    updateSubtotalDividers();
  }

  function refreshItemIndices() {
    const itemsContainer = document.getElementById("itemsContainer");
    if (itemsContainer) {
      const rows = itemsContainer.querySelectorAll('.item-row');
      rows.forEach((r, idx) => {
        const badge = r.querySelector('.item-index-badge');
        if (badge) badge.textContent = `#${idx + 1}`;
      });
    }

    const desktopTableBody = document.getElementById("quoteDesktopTableBody");
    if (desktopTableBody) {
      const dRows = desktopTableBody.querySelectorAll('.desktop-item-row');
      dRows.forEach((r, idx) => {
        const cell = r.querySelector('[id$="-index"]');
        if (cell) cell.textContent = `${idx + 1}`;
      });
    }
  }

  // 重整經銷價 (自資料庫同步最新共通標準牌價與個人特殊品牌價，並觸發方案 A 自動轉正檢查)
  async function refreshDealerPrices() {
    try {
      const activeSales = getActiveSalesName();
      const resp = await fetch(`/api/products?sales_name=${encodeURIComponent(activeSales)}`);
      if (resp.ok) {
        quoteProductsCache = await resp.json();
      }
      const itemsContainer = document.getElementById("itemsContainer");
      const mRows = itemsContainer ? itemsContainer.querySelectorAll('.item-row') : [];
      let updatedCount = 0;

      mRows.forEach(r => {
        const itemId = r.id;
        const codeInput = document.getElementById(`${itemId}-code`);
        const model = codeInput?.value;
        if (!model) return;

        const prod = (quoteProductsCache || []).find(p => p.model_no === model);
        if (prod && prod.dealer_price) {
          const newDP = prod.dealer_price;
          const dDP = document.getElementById(`d-${itemId}-dealer-price`);
          if (dDP) dDP.textContent = `$ ${newDP.toLocaleString()}`;
          const mDP = document.getElementById(`${itemId}-dealer-price`);
          if (mDP) mDP.value = `$${newDP.toLocaleString()}`;

          const disc = parseFloat(document.getElementById(`d-${itemId}-discount`)?.value || document.getElementById(`${itemId}-sug-discount`)?.value || "100") || 100;
          const rawPrice = newDP * (disc / 100);
          const newPrice = applyQuoteRounding(rawPrice);
          const dPrice = document.getElementById(`d-${itemId}-price`);
          if (dPrice) dPrice.value = newPrice;
          const mPrice = document.getElementById(`${itemId}-sug-price`);
          if (mPrice) mPrice.value = Math.round(rawPrice);
          const mUnitNum = document.getElementById(`${itemId}-unit-price-num`);
          if (mUnitNum) mUnitNum.value = newPrice;
          updatedCount++;
        }
      });
      calculateTotal();
      alert(`[經銷價重整完成] 已成功校對並同步 ${updatedCount} 項產品之最新經銷牌價！`);
    } catch (err) {
      alert("[牌價更新失敗] " + err.message);
    }
  }

  // 智慧 BOM 匯入對話框處理
  function openSmartBOMModal() {
    const modal = document.getElementById("smartBomModal");
    const closeBtn = document.getElementById("btnCloseSmartBomModal");
    const cancelBtn = document.getElementById("btnCancelSmartBom");
    const confirmBtn = document.getElementById("btnConfirmSmartBom");
    const textarea = document.getElementById("smartBomInput");
    const countSpan = document.getElementById("bomParsedCount");

    if (modal) modal.classList.remove("hidden");
    if (textarea) {
      textarea.value = "";
      textarea.focus();
    }
    if (countSpan) countSpan.textContent = "尚未解析";

    function closeModal() {
      if (modal) modal.classList.add("hidden");
    }
    if (closeBtn) closeBtn.onclick = closeModal;
    if (cancelBtn) cancelBtn.onclick = closeModal;

    if (textarea && countSpan) {
      textarea.oninput = () => {
        const lines = textarea.value.split("\n").filter(l => l.trim().length > 0);
        countSpan.textContent = `偵測到 ${lines.length} 行產品清單`;
      };
    }

    if (confirmBtn) {
      confirmBtn.onclick = async () => {
        const text = textarea?.value || "";
        const lines = text.split("\n").map(l => l.trim()).filter(Boolean);
        if (lines.length === 0) {
          alert("[請輸入 BOM 清單] 請貼入產品型號與數量！");
          return;
        }

        const activeSales = getActiveSalesName();
        let imported = 0;
        for (const line of lines) {
          const parts = line.split(/[\t ,]+/);
          const modelKeyword = parts[0];
          const qty = parseInt(parts[1], 10) || 1;

          let matched = null;
          try {
            const resp = await fetch(`/api/products?q=${encodeURIComponent(modelKeyword)}&sales_name=${encodeURIComponent(activeSales)}`);
            if (resp.ok) {
              const list = await resp.json();
              matched = (list || []).find(p => p.model_no && p.model_no.toUpperCase() === modelKeyword.toUpperCase()) || (list && list[0]);
            }
          } catch (e) {}

          if (matched) {
            addBlankItem({
              model_no: matched.model_no,
              spec: matched.spec || matched.series || "",
              sku: matched.sku || "",
              stock_qty: matched.stock_qty ?? 0,
              is_order_item: matched.is_order_item,
              is_special: matched.is_special,
              dealer_price: matched.dealer_price || 0,
              discount_pct: 100,
              final_unit_price: matched.dealer_price || 0,
              quantity: qty,
              note_lead_time: "現貨供應"
            });
            imported++;
          } else {
            addBlankItem({
              model_no: modelKeyword,
              spec: "自訂品項",
              dealer_price: 0,
              discount_pct: 100,
              final_unit_price: 0,
              quantity: qty,
              note_lead_time: ""
            });
            imported++;
          }
        }

        closeModal();
        calculateTotal();
        alert(`[BOM 匯入成功] 已成功批次匯入 ${imported} 項產品！`);
      };
    }
  }

  // 從歷史報價單匯入對話框處理 (嚴格依當前業務隔離)
  async function openImportOldQuoteModal() {
    const modal = document.getElementById("importOldQuoteModal");
    const closeBtn = document.getElementById("btnCloseImportOldQuoteModal");
    const cancelBtn = document.getElementById("btnCancelImportOldQuote");
    const listEl = document.getElementById("importOldQuoteList");

    if (modal) modal.classList.remove("hidden");
    function closeModal() {
      if (modal) modal.classList.add("hidden");
    }
    if (closeBtn) closeBtn.onclick = closeModal;
    if (cancelBtn) cancelBtn.onclick = closeModal;

    if (!listEl) return;
    listEl.innerHTML = '<div style="padding:16px; text-align:center; color:var(--text-muted);">正在載入個人歷史報價清單...</div>';

    try {
      const activeSales = getActiveSalesName();
      const resp = await fetch(`/api/quotations?sales_name=${encodeURIComponent(activeSales)}`);
      if (!resp.ok) throw new Error("讀取歷史報價單失敗");
      const quotes = await resp.json();

      if (!Array.isArray(quotes) || quotes.length === 0) {
        listEl.innerHTML = '<div style="padding:16px; text-align:center; color:var(--text-muted);">目前無過往報價單可供匯入</div>';
        return;
      }

      listEl.innerHTML = quotes.slice(0, 50).map(q => `
        <div class="import-quote-row" data-quoteno="${escapeHtml(q.quote_no)}" style="padding:10px 14px; border:1px solid var(--border-color); border-radius:6px; background:var(--bg-card); cursor:pointer; display:flex; justify-content:space-between; align-items:center;">
          <div>
            <div style="font-weight:700; color:var(--primary-color, #2563eb); font-family:monospace;">${escapeHtml(q.quote_no)}</div>
            <div style="font-size:0.85rem; color:var(--text-main); font-weight:600; margin-top:2px;">${escapeHtml(q.customer_name)} ｜ ${escapeHtml(q.project_name || '設備報價')}</div>
            <div style="font-size:0.75rem; color:var(--text-muted); margin-top:2px;">${escapeHtml(q.quote_date)} ｜ $ ${(q.grand_total || 0).toLocaleString()}</div>
          </div>
          <button type="button" class="fluent-btn secondary" style="font-size:0.8rem; padding:4px 10px;">匯入品項</button>
        </div>
      `).join("");

      listEl.querySelectorAll('.import-quote-row').forEach(row => {
        row.addEventListener("click", async () => {
          const qNo = row.dataset.quoteno;
          try {
            const detailResp = await fetch(`/api/quotations/${encodeURIComponent(qNo)}`);
            if (!detailResp.ok) throw new Error("讀取品項明細失敗");
            const data = await detailResp.json();
            const items = data.items || [];
            if (items.length === 0) {
              alert("[該單據無品項] 此報價單無任何產品明細。");
              return;
            }
            items.forEach(itm => addBlankItem(itm));
            closeModal();
            calculateTotal();
            alert(`[匯入完成] 已成功自單號【${qNo}】匯入 ${items.length} 項品項！`);
          } catch (e) {
            alert("[匯入失敗] " + e.message);
          }
        });
      });
    } catch (err) {
      listEl.innerHTML = `<div style="padding:16px; text-align:center; color:var(--color-danger, #ef4444);">載入失敗：${escapeHtml(err.message)}</div>`;
    }
  }

  // 附加圖片處理 (FileReader)
  function handleImageUpload(e) {
    const files = e.target.files;
    if (!files || files.length === 0) return;

    Array.from(files).forEach(file => {
      if (!file.type.startsWith("image/")) return;
      const reader = new FileReader();
      reader.onload = (evt) => {
        quoteAttachedImages.push({
          name: file.name,
          dataUrl: evt.target.result
        });
        renderAttachedImagesPreview();
      };
      reader.readAsDataURL(file);
    });
  }

  function renderAttachedImagesPreview() {
    const container = document.getElementById("quoteAttachmentsPreview");
    if (!container) return;

    if (quoteAttachedImages.length === 0) {
      container.innerHTML = '<span class="text-muted" style="font-size:0.82rem;" id="noAttachmentsHint">尚未附加圖片</span>';
      return;
    }

    container.innerHTML = quoteAttachedImages.map((img, idx) => `
      <div class="quote-attachment-card" title="${escapeHtml(img.name)}">
        <img src="${img.dataUrl}" alt="產品圖">
        <button type="button" class="quote-attachment-remove" data-idx="${idx}" title="刪除圖片">&times;</button>
      </div>
    `).join("");

    container.querySelectorAll('.quote-attachment-remove').forEach(btn => {
      btn.addEventListener("click", (e) => {
        e.stopPropagation();
        const idx = parseInt(btn.dataset.idx, 10);
        quoteAttachedImages.splice(idx, 1);
        renderAttachedImagesPreview();
      });
    });
  }

  // 匯出報價單為 Excel (CSV UTF-8 BOM 防亂碼格式)
  function exportQuoteToExcel() {
    const custName = document.getElementById("customerName")?.value.trim() || "士林電機報價客戶";
    const projName = document.getElementById("quoteProjectName")?.value.trim() || "專案設備";
    const qDate = document.getElementById("quoteDate")?.value || new Date().toISOString().split("T")[0];

    const itemsContainer = document.getElementById("itemsContainer");
    const mRows = itemsContainer ? itemsContainer.querySelectorAll('.item-row') : [];

    if (mRows.length === 0) {
      alert("[無品項可匯出] 報價單中目前無品項！");
      return;
    }

    let csvContent = "\uFEFF";
    csvContent += "士林電機 FA 業務報價單\n";
    csvContent += `客戶名稱,${custName},報價日期,${qDate}\n`;
    csvContent += `專案名稱,${projName},統一編號,${document.getElementById("quoteTaxId")?.value || ""}\n`;
    csvContent += `聯絡人,${document.getElementById("quoteContactSelect")?.value || ""},電話,${document.getElementById("quotePhone")?.value || ""}\n`;
    csvContent += `付款條件,${document.getElementById("quotePaymentTerms")?.value || ""},有效日期,${document.getElementById("quoteValidDate")?.value || ""}\n\n`;

    csvContent += "項次,產品型號,規格說明,經銷牌價,折數%,最終單價(未稅),數量,小計(未稅),備註交期\n";

    let subtotal = 0;
    mRows.forEach((r, idx) => {
      const itemId = r.id;
      const model = document.getElementById(`${itemId}-code`)?.value || "-";
      const spec = r.dataset.spec || "";
      const dpStr = (document.getElementById(`${itemId}-dealer-price`)?.value || "0").replace(/[^0-9.]/g, '');
      const disc = document.getElementById(`${itemId}-sug-discount`)?.value || "100";
      const price = parseFloat(document.getElementById(`${itemId}-unit-price-num`)?.value) || 0;
      const qty = parseInt(document.getElementById(`${itemId}-qty`)?.value, 10) || 1;
      const rowSub = Math.round(price * qty);
      subtotal += rowSub;
      const delivery = document.getElementById(`${itemId}-delivery`)?.value || "";

      csvContent += `${idx + 1},"${model}","${spec}",${dpStr},${disc}%,${price},${qty},${rowSub},"${delivery}"\n`;
    });

    const tax = Math.round(subtotal * 0.05);
    const grandTotal = subtotal + tax;

    csvContent += `\n未稅小計,,,,,,,${subtotal}\n`;
    csvContent += `營業稅(5%),,,,,,,${tax}\n`;
    csvContent += `含稅總計,,,,,,,${grandTotal}\n`;

    const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = `報價單_${custName}_${qDate}.csv`;
    link.click();
  }

  function printQuotePDF() {
    window.print();
  }

  // ============================================================================
  // 對話盒 (Modals) 控制與資料繫結
  // ============================================================================
  function setupQuotationModals() {
    // 1. 客戶搜尋 Modal 與雙向表頭連動
    const custInput = document.getElementById("customerName");
    const custModal = document.getElementById("customerModal");
    const custClose = document.getElementById("btnCloseCustomerModal");
    const custSearch = document.getElementById("customerModalSearch");
    const custResults = document.getElementById("customerModalResults");
    const custCount = document.getElementById("customerModalCount");

    if (custInput && custModal) {
      custInput.addEventListener("click", async () => {
        if (!customersCache || customersCache.length === 0) {
          await populateCustomerSelectDropdown(true);
        }
        if (custSearch) custSearch.value = "";
        renderCustomerModalList("");
        custModal.classList.remove("hidden");
        setTimeout(() => custSearch && custSearch.focus(), 100);
      });
    }

    if (custClose && custModal) {
      custClose.addEventListener("click", () => custModal.classList.add("hidden"));
    }

    if (custSearch) {
      custSearch.addEventListener("input", (e) => {
        renderCustomerModalList(e.target.value.trim());
      });
    }

    function renderCustomerModalList(query = "") {
      if (!custResults) return;
      const list = (customersCache || []).filter(c => {
        if (!query) return true;
        const q = query.toLowerCase();
        return (c.name && c.name.toLowerCase().includes(q)) ||
               (c.short_name && c.short_name.toLowerCase().includes(q)) ||
               (c.tax_id && c.tax_id.toLowerCase().includes(q)) ||
               (c.phone && c.phone.toLowerCase().includes(q));
      });

      if (custCount) custCount.textContent = list.length;

      if (list.length === 0) {
        custResults.innerHTML = '<div style="padding:16px; text-align:center; color:var(--text-muted); font-size:0.9rem;">查無符合的客戶資料</div>';
        return;
      }

      custResults.innerHTML = list.map(c => {
        const sub = [
          c.tax_id ? `統編: ${c.tax_id}` : "",
          c.default_contact ? `窗口: ${c.default_contact}` : "",
          c.phone ? `電話: ${c.phone}` : ""
        ].filter(Boolean).join(" ｜ ");
        return `
          <div class="customer-item" data-id="${c.id}">
            <div class="customer-item-name">${escapeHtml(c.name)}</div>
            ${sub ? `<div class="customer-item-sub">${escapeHtml(sub)}</div>` : ''}
          </div>
        `;
      }).join("");

      custResults.querySelectorAll('.customer-item').forEach(el => {
        el.addEventListener("click", () => {
          const cid = parseInt(el.dataset.id, 10);
          const found = (customersCache || []).find(c => c.id === cid);
          if (found) {
            applySelectedCustomerToQuoteUI(found);
          }
          if (custModal) custModal.classList.add("hidden");
        });
      });
    }

    // 2. 產品搜尋 Modal (即時向後端搜尋共通標準牌價 + 業務個人特殊品牌價)
    const prodModal = document.getElementById("productModal");
    const prodClose = document.getElementById("btnCloseModal");
    const prodSearch = document.getElementById("productSearch");
    let prodSearchTimer = null;

    if (prodClose && prodModal) {
      prodClose.addEventListener("click", () => prodModal.classList.add("hidden"));
    }

    if (prodSearch) {
      prodSearch.addEventListener("input", (e) => {
        const q = e.target.value.trim();
        if (prodSearchTimer) clearTimeout(prodSearchTimer);
        prodSearchTimer = setTimeout(() => {
          fetchAndRenderProductModalList(q);
        }, 180);
      });
    }

    // 3. 交期 / 備註 Modal
    const delivModal = document.getElementById("deliveryModal");
    const delivClose = document.getElementById("btnCloseDeliveryModal");
    const delivResults = document.getElementById("deliveryModalResults");
    const customDelivInput = document.getElementById("customDeliveryInput");
    const btnConfirmCustom = document.getElementById("btnConfirmCustomDelivery");

    if (delivClose && delivModal) {
      delivClose.addEventListener("click", () => delivModal.classList.add("hidden"));
    }

    if (delivResults) {
      delivResults.innerHTML = PRESET_DELIVERY_OPTIONS.map(opt => `
        <div class="delivery-item" data-text="${escapeHtml(opt)}">${escapeHtml(opt)}</div>
      `).join("");

      delivResults.querySelectorAll('.delivery-item').forEach(el => {
        el.addEventListener("click", () => {
          const txt = el.dataset.text;
          if (currentEditingItemId) {
            const input = document.getElementById(`${currentEditingItemId}-delivery`);
            if (input) input.value = txt;
            const dInput = document.getElementById(`d-${currentEditingItemId}-delivery`);
            if (dInput) dInput.value = txt;
          }
          if (delivModal) delivModal.classList.add("hidden");
        });
      });
    }

    if (btnConfirmCustom && customDelivInput) {
      btnConfirmCustom.addEventListener("click", () => {
        const txt = customDelivInput.value.trim();
        if (txt && currentEditingItemId) {
          const input = document.getElementById(`${currentEditingItemId}-delivery`);
          if (input) input.value = txt;
          const dInput = document.getElementById(`d-${currentEditingItemId}-delivery`);
          if (dInput) dInput.value = txt;
          customDelivInput.value = "";
        }
        if (delivModal) delivModal.classList.add("hidden");
      });
    }

    // 4. LINE 報價 Modal
    const lineModal = document.getElementById("lineModal");
    const lineClose = document.getElementById("btnCloseLineModal");
    const lineCancel = document.getElementById("btnCancelLineQuote");
    const lineCopy = document.getElementById("btnCopyLineQuote");

    if (lineClose && lineModal) {
      lineClose.addEventListener("click", () => lineModal.classList.add("hidden"));
    }
    if (lineCancel && lineModal) {
      lineCancel.addEventListener("click", () => lineModal.classList.add("hidden"));
    }
    if (lineCopy) {
      lineCopy.addEventListener("click", () => {
        const textarea = document.getElementById("lineQuoteText");
        if (textarea && textarea.value) {
          navigator.clipboard.writeText(textarea.value).then(() => {
            alert("[一鍵複製成功] 報價文字已存入剪貼簿，可直接貼到 LINE 聊天室發送！");
            if (lineModal) lineModal.classList.add("hidden");
          }).catch(err => {
            alert("[複製失敗] 請手動全選複製：" + err.message);
          });
        }
      });
    }

    // 5. 料號切換 Modal
    const skuModal = document.getElementById("skuModal");
    const skuClose = document.getElementById("btnCloseSkuModal");
    const skuCancel = document.getElementById("btnCancelSkuModal");
    if (skuClose && skuModal) skuClose.addEventListener("click", () => skuModal.classList.add("hidden"));
    if (skuCancel && skuModal) skuCancel.addEventListener("click", () => skuModal.classList.add("hidden"));

    // 6. 移至垃圾桶二次確認對話盒
    const delModal = document.getElementById("deleteConfirmModal");
    const btnCancelDel = document.getElementById("btnCancelDeleteConfirm");
    const btnConfirmDel = document.getElementById("btnExecuteDeleteConfirm");

    if (btnCancelDel && delModal) {
      btnCancelDel.addEventListener("click", () => delModal.classList.add("hidden"));
    }
    if (btnConfirmDel && delModal) {
      btnConfirmDel.addEventListener("click", async () => {
        if (!pendingDeleteQuoteNo) {
          delModal.classList.add("hidden");
          return;
        }
        btnConfirmDel.disabled = true;
        btnConfirmDel.textContent = "處理中...";
        try {
          const resp = await fetch(`/api/quotations/${encodeURIComponent(pendingDeleteQuoteNo)}`, {
            method: "DELETE"
          });
          if (!resp.ok) {
            const err = await resp.json().catch(() => ({}));
            throw new Error(err.detail || "伺服器刪除失敗");
          }
          delModal.classList.add("hidden");
          alert(`[已移至垃圾桶] 報價單【${pendingDeleteQuoteNo}】已安全移至垃圾桶！\n若需還原，可點擊「檢視垃圾桶」隨時一鍵復原。`);
          pendingDeleteQuoteNo = null;
          loadQuotationHistory("", false);
        } catch (err) {
          alert("[軟刪除失敗] " + err.message);
        } finally {
          btnConfirmDel.disabled = false;
          btnConfirmDel.textContent = "移至垃圾桶";
        }
      });
    }
  }

  async function openProductModal() {
    const modal = document.getElementById("productModal");
    const search = document.getElementById("productSearch");
    if (modal) modal.classList.remove("hidden");
    if (search) {
      search.value = "";
      setTimeout(() => search.focus(), 100);
    }
    await fetchAndRenderProductModalList("");
  }

  async function fetchAndRenderProductModalList(query = "") {
    const listEl = document.getElementById("productResults");
    if (!listEl) return;

    const activeSales = getActiveSalesName();
    try {
      const resp = await fetch(`/api/products?q=${encodeURIComponent(query)}&sales_name=${encodeURIComponent(activeSales)}`);
      if (resp.ok) {
        quoteProductsCache = await resp.json();
      }
    } catch (err) {
      console.warn("查詢產品庫失敗:", err);
    }

    const filtered = quoteProductsCache || [];
    if (filtered.length === 0) {
      listEl.innerHTML = '<div style="padding:16px; text-align:center; color:var(--text-muted); font-size:0.9rem;">查無符合的產品型號（可調整關鍵字或至「個人特殊品牌價」新增）</div>';
      return;
    }

    listEl.innerHTML = filtered.map((p, idx) => {
      const price = p.dealer_price ? `$ ${Number(p.dealer_price).toLocaleString()}` : "-";
      const stockNum = p.stock_qty !== undefined && p.stock_qty !== null ? Number(p.stock_qty) : 0;
      const isSpecial = Boolean(p.is_special);
      const isOrder = Boolean(p.is_order_item);
      return `
        <div class="product-item" data-idx="${idx}">
          <div style="display:flex; justify-content:space-between; align-items:center; gap:8px;">
            <div style="min-width:0; flex:1;">
              <div class="product-item-name" style="display:flex; align-items:center; flex-wrap:wrap; gap:6px;">
                <span>${escapeHtml(p.model_no)}</span>
                ${isSpecial ? '<span class="special-product-tag">個人特殊品</span>' : ''}
                ${isOrder ? '<span class="order-item-badge">受訂生產</span>' : ''}
              </div>
              <div class="product-item-sub">
                <span>${escapeHtml(p.spec || p.series || "士林標準系列")}</span>
                ${p.sku ? `<span>料號: ${escapeHtml(p.sku)}</span>` : ''}
              </div>
            </div>
            <div style="text-align:right; flex-shrink:0;">
              <div style="font-weight:700; color:var(--primary-color, #2563eb); font-size:0.95rem;">${price}</div>
              <div style="font-size:0.78rem; color:${stockNum > 0 ? '#059669' : '#64748b'}; font-weight:600; margin-top:2px;">
                庫存: ${stockNum.toLocaleString()} 台
              </div>
            </div>
          </div>
        </div>
      `;
    }).join("");

    listEl.querySelectorAll('.product-item').forEach(el => {
      el.addEventListener("click", () => {
        const idx = parseInt(el.dataset.idx, 10);
        const prod = filtered[idx];
        if (!prod) return;

        if (currentEditingItemId) {
          applyProductToItemRow(currentEditingItemId, prod);
        }
        const modal = document.getElementById("productModal");
        if (modal) modal.classList.add("hidden");
      });
    });
  }

  function applyProductToItemRow(itemId, prod) {
    const model = prod.model_no || "";
    const price = parseFloat(prod.dealer_price) || 0;
    const spec = prod.spec || prod.series || "";
    const sku = prod.sku || "";
    const stockQty = prod.stock_qty !== undefined && prod.stock_qty !== null ? Number(prod.stock_qty) : 0;
    const isOrder = Boolean(prod.is_order_item);
    const isSpecial = Boolean(prod.is_special);

    // 1. 手機端元件更新
    const titleBtn = document.getElementById(`${itemId}-title`);
    const codeInput = document.getElementById(`${itemId}-code`);
    const skuInput = document.getElementById(`${itemId}-sku`);
    const skuWrap = document.getElementById(`${itemId}-sku-wrap`);
    const skuText = document.getElementById(`${itemId}-sku-text`);
    const stockInput = document.getElementById(`${itemId}-stock`);
    const dealerPriceInput = document.getElementById(`${itemId}-dealer-price`);
    const copyBtn = document.getElementById(`${itemId}-copy-btn`);
    const orderBadge = document.getElementById(`${itemId}-order-badge`);
    const specialBadge = document.getElementById(`${itemId}-special-badge`);
    const sugDiscountInput = document.getElementById(`${itemId}-sug-discount`);
    const sugPriceInput = document.getElementById(`${itemId}-sug-price`);
    const qtyInput = document.getElementById(`${itemId}-qty`);
    const finalPriceInput = document.getElementById(`${itemId}-final-price`);
    const unitPriceNum = document.getElementById(`${itemId}-unit-price-num`);

    // 2. 電腦端表格元件更新
    const dModelBtn = document.getElementById(`d-${itemId}-model`);
    const dStock = document.getElementById(`d-${itemId}-stock`);
    const dSpec = document.getElementById(`d-${itemId}-spec`);
    const dDP = document.getElementById(`d-${itemId}-dealer-price`);
    const dPrice = document.getElementById(`d-${itemId}-price`);
    const dDiscount = document.getElementById(`d-${itemId}-discount`);
    const dSub = document.getElementById(`d-${itemId}-subtotal`);

    if (titleBtn) {
      titleBtn.textContent = model;
      titleBtn.classList.add("selected");
    }
    if (codeInput) codeInput.value = model;
    if (skuInput) skuInput.value = sku;
    if (skuWrap && skuText) {
      if (sku) {
        skuText.textContent = sku;
        skuWrap.classList.remove("hidden");
      } else {
        skuWrap.classList.add("hidden");
      }
    }
    if (stockInput) stockInput.value = `${stockQty.toLocaleString()} 台`;
    if (dealerPriceInput) dealerPriceInput.value = price > 0 ? `$${price.toLocaleString()}` : "-";
    if (copyBtn) copyBtn.classList.remove("hidden");
    if (orderBadge) orderBadge.classList.toggle("hidden", !isOrder);
    if (specialBadge) specialBadge.classList.toggle("hidden", !isSpecial);

    if (dModelBtn) {
      dModelBtn.innerHTML = `<span id="d-${itemId}-model-text">${escapeHtml(model)}</span> <span class="dropdown-arrow">▼</span>`;
      dModelBtn.classList.add("selected");
    }
    const dCodeInput = document.getElementById(`d-${itemId}-code`);
    if (dCodeInput) dCodeInput.value = model;
    if (dStock) dStock.textContent = `${stockQty.toLocaleString()}`;
    if (dSpec) dSpec.textContent = spec || "-";
    if (dDP) dDP.textContent = price > 0 ? `$ ${price.toLocaleString()}` : "-";

    // 同步上方作用列標籤
    const activeBadge = document.getElementById("activeRowBadge");
    const dRowEl = document.getElementById(`d-${itemId}`);
    const rowIdx = dRowEl?.querySelector('[id$="-index"]')?.textContent?.trim() || "1";
    if (activeBadge) activeBadge.textContent = `第 ${rowIdx} 列 | ${model}`;
    if (dRowEl) {
      const tableBody = document.getElementById("quoteDesktopTableBody");
      if (tableBody) tableBody.querySelectorAll("tr").forEach(tr => tr.classList.remove("active-row"));
      dRowEl.classList.add("active-row");
    }

    const discVal = sugDiscountInput && sugDiscountInput.value !== "" ? parseFloat(sugDiscountInput.value) : 100;
    const disc = isNaN(discVal) ? 100 : discVal;
    if (sugDiscountInput && sugDiscountInput.value === "") sugDiscountInput.value = String(disc);
    if (dDiscount) dDiscount.value = String(disc);

    const rawPrice = price * (disc / 100);
    const roundedPrice = disc === 100 ? Math.round(price) : applyQuoteRounding(rawPrice);
    if (sugPriceInput) sugPriceInput.value = Math.round(rawPrice);
    if (unitPriceNum) unitPriceNum.value = roundedPrice;
    if (dPrice) dPrice.value = roundedPrice;
    if (finalPriceInput) finalPriceInput.value = roundedPrice > 0 ? `$${roundedPrice.toLocaleString()}` : "-";

    const q = parseInt(qtyInput?.value || "1", 10) || 1;
    const sub = Math.round(roundedPrice * q);
    if (dSub) dSub.textContent = sub > 0 ? `$ ${sub.toLocaleString()}` : "$ 0";

    const mRow = document.getElementById(itemId);
    if (mRow) {
      mRow.dataset.spec = spec;
      mRow.dataset.stock = String(stockQty);
    }

    calculateTotal();
  }

  function openDeliveryModal() {
    const modal = document.getElementById("deliveryModal");
    if (modal) modal.classList.remove("hidden");
  }

  function openLineQuoteModal() {
    const custSel = document.getElementById("customerSelect");
    const custName = document.getElementById("customerName")?.value.trim() || (custSel && custSel.selectedIndex > 0 ? custSel.options[custSel.selectedIndex].text : "") || "貴客戶";
    const projName = document.getElementById("quoteProjectName")?.value.trim() || "設備報價";
    const qDate = document.getElementById("quoteDate")?.value || new Date().toISOString().split("T")[0];
    const activeSales = getActiveSalesName();

    const itemsContainer = document.getElementById("itemsContainer");
    const rows = itemsContainer ? itemsContainer.querySelectorAll('.item-row') : [];

    let lines = [];
    lines.push("【士林電機 FA 業務報價單】");
    lines.push(`客戶名稱：${custName}`);
    lines.push(`專案案場：${projName}`);
    lines.push(`報價日期：${qDate}`);
    lines.push("----------------------------------------");

    let subtotal = 0;
    if (rows.length === 0) {
      lines.push("(尚未加入任何品項)");
    } else {
      rows.forEach((r, idx) => {
        const itemId = r.id;
        const model = document.getElementById(`${itemId}-code`)?.value || "未指定型號";
        const sp = parseFloat(document.getElementById(`${itemId}-unit-price-num`)?.value || document.getElementById(`${itemId}-sug-price`)?.value || "0") || 0;
        const q = parseInt(document.getElementById(`${itemId}-qty`)?.value || "1", 10) || 1;
        const rowSub = Math.round(sp * q);
        subtotal += rowSub;
        const note = document.getElementById(`${itemId}-delivery`)?.value || "";
        const noteStr = note ? ` ｜ 備註：${note}` : "";

        lines.push(`${idx + 1}. ${model}`);
        lines.push(`   數量：${q} ｜ 單價：$${sp.toLocaleString()} ｜ 小計：$${rowSub.toLocaleString()}${noteStr}`);
      });
    }

    const tax = Math.round(subtotal * 0.05);
    const grandTotal = subtotal + tax;

    lines.push("----------------------------------------");
    lines.push(`未稅小計：$${subtotal.toLocaleString()}`);
    lines.push(`營業稅(5%)：$${tax.toLocaleString()}`);
    lines.push(`含稅總計：$${grandTotal.toLocaleString()}`);
    lines.push("----------------------------------------");
    lines.push(`業務代表：${activeSales}`);
    lines.push("※ 本報價單有效期限為 30 天，感謝您的支持與愛護！");

    const textarea = document.getElementById("lineQuoteText");
    if (textarea) textarea.value = lines.join("\n");

    const modal = document.getElementById("lineModal");
    if (modal) modal.classList.remove("hidden");
  }

  // 共用收集目前表單品項與客戶資料
  function collectCurrentQuoteFormData() {
    const custSel = document.getElementById("customerSelect");
    const custName = document.getElementById("customerName")?.value.trim() || (custSel && custSel.selectedIndex > 0 ? custSel.options[custSel.selectedIndex].text : "");
    if (!custName) {
      alert("[請填寫客戶名稱] 請先點擊選擇客戶名稱！");
      return null;
    }

    const itemsContainer = document.getElementById("itemsContainer");
    const rows = itemsContainer ? itemsContainer.querySelectorAll('.item-row') : [];
    if (rows.length === 0) {
      alert("[無品項] 報價單必須包含至少一項產品！");
      return null;
    }

    const itemsData = [];
    let subtotal = 0;
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      const itemId = r.id;
      const model = document.getElementById(`${itemId}-code`)?.value;
      if (!model) {
        alert(`[品項未指定] 第 ${i + 1} 項尚未選擇產品型號！`);
        return null;
      }
      const dpStr = (document.getElementById(`${itemId}-dealer-price`)?.value || "0").replace(/[^0-9.]/g, '');
      const dealerPrice = parseFloat(dpStr) || 0;
      const discount = parseFloat(document.getElementById(`${itemId}-sug-discount`)?.value || "100") || 100;
      const sp = parseFloat(document.getElementById(`${itemId}-unit-price-num`)?.value || document.getElementById(`${itemId}-sug-price`)?.value || "0") || 0;
      const q = parseInt(document.getElementById(`${itemId}-qty`)?.value || "1", 10) || 1;
      const rowSub = Math.round(sp * q);
      subtotal += rowSub;
      const leadTime = document.getElementById(`${itemId}-delivery`)?.value || "";
      const sku = document.getElementById(`${itemId}-sku`)?.value || "";
      const stockQty = r.dataset.stock !== "" && r.dataset.stock !== undefined ? parseInt(r.dataset.stock, 10) : 0;

      itemsData.push({
        model_no: model,
        sku: sku,
        spec: r.dataset.spec || "",
        stock_qty: stockQty,
        dealer_price: dealerPrice,
        discount_pct: discount,
        final_unit_price: sp,
        quantity: q,
        row_subtotal: rowSub,
        note_lead_time: leadTime
      });
    }

    const now = new Date();
    const yy = String(now.getFullYear()).slice(-2);
    const mm = String(now.getMonth() + 1).padStart(2, "0");
    const yymm = yy + mm;
    let cleanCust = custName.replace(/有限公司|股份有限公司|企業|實業|科技|工業/g, "").trim();
    let prefixName = cleanCust.slice(0, 4) || "士電";
    let prefixInput = document.getElementById("quotePrefixInput")?.value.trim();
    let prefix = (prefixInput && prefixInput !== "-- 請選擇客戶 --") ? prefixInput : `${prefixName}${yymm}`;

    const crmProject = document.getElementById("crmCaseSelect")?.value || "";
    const projName = document.getElementById("quoteProjectName")?.value.trim() || crmProject || "標準設備報價";
    const projType = document.getElementById("quoteProjectAttribute")?.value || "一次性專案";
    const qDate = document.getElementById("quoteDate")?.value || now.toISOString().split("T")[0];
    const vDate = document.getElementById("quoteValidDate")?.value || qDate;
    const tax = Math.round(subtotal * 0.05);
    const grandTotal = subtotal + tax;
    const activeSales = getActiveSalesName();

    return {
      sales_name: activeSales,
      prefix: prefix,
      quote_date: qDate,
      valid_date: vDate,
      customer_name: custName,
      tax_id: document.getElementById("quoteTaxId")?.value.trim() || currentSelectedCustomer?.tax_id || "",
      contact_name: document.getElementById("quoteContactSelect")?.value.trim() || currentSelectedCustomer?.default_contact || "",
      phone: document.getElementById("quotePhone")?.value.trim() || currentSelectedCustomer?.phone || "",
      fax: document.getElementById("quoteFax")?.value.trim() || currentSelectedCustomer?.fax || "",
      delivery_address: document.getElementById("quoteDeliveryLocation")?.value.trim() || currentSelectedCustomer?.address || "",
      payment_terms: document.getElementById("quotePaymentTerms")?.value.trim() || "月結 30 天電匯",
      project_name: projName,
      project_type: projType,
      subtotal: subtotal,
      tax_rate: 5.0,
      tax_amount: tax,
      grand_total: grandTotal,
      status: "追蹤中",
      items: itemsData
    };
  }

  // 手機端「送出草稿」：直接寫入資料庫 quotation_draft 表 (不需透過雲端硬碟)，回到電腦端即可於「手機報價草稿匣」一鍵載入
  async function submitMobileQuotationDraftToDB() {
    const payload = collectCurrentQuoteFormData();
    if (!payload) return;

    const btnSubmit = document.getElementById("btnSubmitDraft");
    if (btnSubmit) {
      btnSubmit.disabled = true;
      const span = btnSubmit.querySelector("span");
      if (span) span.textContent = "傳送草稿中...";
    }

    try {
      const resp = await fetch("/api/quotation-drafts", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload)
      });
      if (!resp.ok) {
        const err = await resp.json().catch(() => ({}));
        throw new Error(err.detail || "草稿寫入伺服器失敗");
      }
      alert(`[草稿已送出] 已直接將報價草稿傳送至資料庫「手機報價草稿匣」！\n客戶：${payload.customer_name}\n品項數：${payload.items.length} 項\n回到電腦桌前即可從「手機報價草稿匣」分頁一鍵載入並正式建立報價單。`);
      switchQuoteSubTab("drafts");
    } catch (err) {
      console.error("送出手機草稿失敗:", err);
      alert("[草稿傳送失敗] " + err.message);
    } finally {
      if (btnSubmit) {
        btnSubmit.disabled = false;
        const span = btnSubmit.querySelector("span");
        if (span) span.textContent = "送出草稿";
      }
    }
  }

  // 電腦端「儲存報價」：正式寫入資料庫 quotation_header + quotation_item (Append-Only)
  async function saveFormalQuotation() {
    const payload = collectCurrentQuoteFormData();
    if (!payload) return;

    const btnSaveTop = document.getElementById("btnSaveQuoteTop");
    if (btnSaveTop) {
      btnSaveTop.disabled = true;
      btnSaveTop.textContent = "儲存中...";
    }

    try {
      const resp = await fetch("/api/quotations", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload)
      });
      if (!resp.ok) {
        const err = await resp.json().catch(() => ({}));
        throw new Error(err.detail || "伺服器儲存失敗");
      }
      const result = await resp.json();
      currentBaseQuoteNo = result.quote_no;

      // 若此單是由手機草稿匣載入，自動清理該筆已轉正式單的草稿
      if (currentLoadedDraftId) {
        try {
          await fetch(`/api/quotation-drafts/${currentLoadedDraftId}?sales_name=${encodeURIComponent(payload.sales_name)}`, {
            method: "DELETE"
          });
        } catch (e) {}
        currentLoadedDraftId = null;
      }

      alert(`[儲存成功] 正式報價單已寫入資料庫！\n指派單號：【${result.quote_no}】\n所屬業務：${payload.sales_name}`);
      switchQuoteSubTab("history");
    } catch (err) {
      console.error("儲存報價單失敗:", err);
      alert("[儲存失敗] " + err.message);
    } finally {
      if (btnSaveTop) {
        btnSaveTop.disabled = false;
        btnSaveTop.textContent = "儲存報價";
      }
    }
  }

  // 載入歷史報價紀錄與垃圾桶 (嚴格依當前業務隔離，並支援標記已成交/未成交/不計入)
  async function loadQuotationHistory(keyword = "", isTrash = false) {
    const resultsContainer = document.getElementById("quoteHistoryResults");
    if (!resultsContainer) return;
    resultsContainer.innerHTML = '<div style="padding:24px; text-align:center; color:var(--text-muted); font-size:0.88rem;">正在載入個人歷史報價紀錄...</div>';

    try {
      const activeSales = getActiveSalesName();
      const statusFilter = document.getElementById("selectQuoteHistoryStatus")?.value || "ALL";
      const url = `/api/quotations?sales_name=${encodeURIComponent(activeSales)}&include_deleted=${isTrash ? 'true' : 'false'}${keyword ? `&keyword=${encodeURIComponent(keyword)}` : ''}${(!isTrash && statusFilter !== 'ALL') ? `&status=${encodeURIComponent(statusFilter)}` : ''}`;
      const resp = await fetch(url);
      if (!resp.ok) throw new Error("讀取報價單資料失敗");
      const allQuotes = await resp.json();

      const filtered = (allQuotes || []).filter(q => isTrash ? (q.is_deleted === 1) : (!q.is_deleted || q.is_deleted === 0));

      if (filtered.length === 0) {
        resultsContainer.innerHTML = `
          <div style="padding:32px; text-align:center; color:var(--text-muted); font-size:0.9rem;">
            ${isTrash ? `目前業務【${escapeHtml(activeSales)}】回收站中沒有已刪除的報價單。` : `目前業務【${escapeHtml(activeSales)}】查無符合條件的歷史報價單紀錄。`}
          </div>
        `;
        return;
      }

      resultsContainer.innerHTML = filtered.map(q => {
        const st = q.status || "追蹤中";
        const pillClass = st === "已成交" ? "success" : (st === "已失注" ? "danger" : "info");
        return `
        <div class="quote-history-card" data-quoteno="${escapeHtml(q.quote_no)}">
          <div style="display:flex; justify-content:space-between; align-items:flex-start; margin-bottom:6px; flex-wrap:wrap; gap:6px;">
            <div>
              <span style="font-family:monospace; font-weight:700; color:var(--primary-color, #2563eb); font-size:1rem;">${escapeHtml(q.quote_no)}</span>
              ${q.revision > 0 ? `<span style="display:inline-block; margin-left:6px; font-size:0.75rem; padding:1px 6px; background:#eff6ff; color:#2563eb; border-radius:4px; font-weight:600;">Rev.${q.revision}</span>` : ''}
            </div>
            <div style="display:flex; align-items:center; gap:6px; flex-wrap:wrap;">
              <span class="status-pill ${pillClass}">${escapeHtml(st)}</span>
              ${isTrash ? `
                <button type="button" class="btn-restore-quote secondary" data-quoteno="${escapeHtml(q.quote_no)}" style="padding:3px 10px; font-size:0.8rem; background:#10b981; color:#fff; border:none; border-radius:4px; cursor:pointer;">一鍵復原</button>
              ` : `
                <button type="button" class="status-mini-btn btn-set-quote-status" data-quoteno="${escapeHtml(q.quote_no)}" data-status="已成交" data-sub="${q.subtotal || 0}" title="標記為已成交">標記成交</button>
                <button type="button" class="status-mini-btn btn-set-quote-status" data-quoteno="${escapeHtml(q.quote_no)}" data-status="已失注" title="標記為未成交/已失注">未成交</button>
                <button type="button" class="status-mini-btn btn-set-quote-status" data-quoteno="${escapeHtml(q.quote_no)}" data-status="不跟進" title="標記為不計入統計">不計入</button>
                <button type="button" class="status-mini-btn btn-set-quote-status" data-quoteno="${escapeHtml(q.quote_no)}" data-status="追蹤中" title="恢復為追蹤中">追蹤中</button>
                <button type="button" class="btn-trash-quote danger" data-quoteno="${escapeHtml(q.quote_no)}" style="padding:3px 8px; font-size:0.78rem; background:transparent; color:#ef4444; border:1px solid #fecaca; border-radius:4px; cursor:pointer;">移至垃圾桶</button>
              `}
            </div>
          </div>
          <div style="font-size:0.92rem; color:var(--text-main, #1e293b); font-weight:600; margin-bottom:4px;">
            ${escapeHtml(q.customer_name)} ｜ ${escapeHtml(q.project_name || "標準設備報價")}
          </div>
          <div style="display:flex; justify-content:space-between; align-items:center; font-size:0.82rem; color:var(--text-muted, #64748b); flex-wrap:wrap; gap:6px;">
            <span>報價日期：${escapeHtml(q.quote_date)} ｜ 聯絡人：${escapeHtml(q.contact_name || '-')} ｜ 未稅：$ ${Number(q.subtotal || 0).toLocaleString()}</span>
            <span style="font-weight:700; color:var(--primary-color, #2563eb); font-size:1.05rem;">含稅總計：$ ${Number(q.grand_total || 0).toLocaleString()}</span>
          </div>
        </div>
      `;
      }).join("");

      resultsContainer.querySelectorAll('.quote-history-card').forEach(card => {
        card.addEventListener("click", (e) => {
          if (e.target.closest('.btn-trash-quote') || e.target.closest('.btn-restore-quote') || e.target.closest('.btn-set-quote-status')) return;
          const qNo = card.dataset.quoteno;
          loadQuotationToForm(qNo);
        });
      });

      resultsContainer.querySelectorAll('.btn-set-quote-status').forEach(btn => {
        btn.addEventListener("click", async (e) => {
          e.stopPropagation();
          const qNo = btn.dataset.quoteno;
          const newStatus = btn.dataset.status;
          const subAmt = parseFloat(btn.dataset.sub || "0") || 0;
          try {
            const r = await fetch(`/api/quotations/${encodeURIComponent(qNo)}/status`, {
              method: "POST",
              headers: { "Content-Type": "application/json" },
              body: JSON.stringify({
                status: newStatus,
                won_amount: newStatus === "已成交" ? subAmt : 0
              })
            });
            if (r.ok) {
              loadQuotationHistory(keyword, isTrash);
            }
          } catch (err) {
            alert("[更新狀態失敗] " + err.message);
          }
        });
      });

      resultsContainer.querySelectorAll('.btn-trash-quote').forEach(btn => {
        btn.addEventListener("click", (e) => {
          e.stopPropagation();
          pendingDeleteQuoteNo = btn.dataset.quoteno;
          const targetEl = document.getElementById("deleteConfirmQuoteNo");
          if (targetEl) targetEl.textContent = pendingDeleteQuoteNo;
          const delModal = document.getElementById("deleteConfirmModal");
          if (delModal) delModal.classList.remove("hidden");
        });
      });

      resultsContainer.querySelectorAll('.btn-restore-quote').forEach(btn => {
        btn.addEventListener("click", async (e) => {
          e.stopPropagation();
          const qNo = btn.dataset.quoteno;
          if (!confirm(`確定要將報價單【${qNo}】從回收站復原嗎？`)) return;
          try {
            const resp = await fetch(`/api/quotations/${encodeURIComponent(qNo)}/restore`, { method: "POST" });
            if (!resp.ok) {
              const err = await resp.json().catch(() => ({}));
              throw new Error(err.detail || "伺服器復原失敗");
            }
            alert(`[復原成功] 報價單【${qNo}】已成功復原！`);
            loadQuotationHistory(keyword, true);
          } catch (err) {
            alert("[復原失敗] " + err.message);
          }
        });
      });

    } catch (err) {
      resultsContainer.innerHTML = `<div style="padding:16px; text-align:center; color:var(--color-danger, #ef4444); font-size:0.88rem;">載入失敗：${escapeHtml(err.message)}</div>`;
    }
  }

  async function loadQuotationToForm(quoteNo) {
    try {
      const resp = await fetch(`/api/quotations/${encodeURIComponent(quoteNo)}`);
      if (!resp.ok) throw new Error("讀取報價單明細失敗");
      const data = await resp.json();

      const header = data.header || data;
      const items = data.items || [];
      currentBaseQuoteNo = header.quote_no;
      currentLoadedDraftId = null;

      // 嘗試比對客戶主檔
      const matchedCust = (customersCache || []).find(c => c.name === header.customer_name);
      if (matchedCust) {
        applySelectedCustomerToQuoteUI(matchedCust);
      } else {
        const custInput = document.getElementById("customerName");
        if (custInput) custInput.value = header.customer_name || "";
      }

      const projInput = document.getElementById("quoteProjectName");
      if (projInput) projInput.value = header.project_name || "";
      const projAttr = document.getElementById("quoteProjectAttribute");
      if (projAttr && header.project_type) projAttr.value = header.project_type;

      const dateInput = document.getElementById("quoteDate");
      if (dateInput) dateInput.value = (header.quote_date || "").replace(/\//g, "-") || new Date().toISOString().split("T")[0];
      const validDateInput = document.getElementById("quoteValidDate");
      if (validDateInput) validDateInput.value = (header.valid_date || header.quote_date || "").replace(/\//g, "-");

      const taxIdInput = document.getElementById("quoteTaxId");
      if (taxIdInput) taxIdInput.value = header.tax_id || "";
      const phoneInput = document.getElementById("quotePhone");
      if (phoneInput) phoneInput.value = header.phone || "";
      const faxInput = document.getElementById("quoteFax");
      if (faxInput) faxInput.value = header.fax || "";
      const termsInput = document.getElementById("quotePaymentTerms");
      if (termsInput) termsInput.value = header.payment_terms || "月結 30 天電匯";
      const addrInput = document.getElementById("quoteDeliveryLocation");
      if (addrInput) addrInput.value = header.delivery_address || "";

      if (header.quote_no) {
        const parts = header.quote_no.split("-");
        const prefixInput = document.getElementById("quotePrefixInput");
        if (prefixInput && parts.length > 0) prefixInput.value = parts[0];
        const seqInput = document.getElementById("quoteSeqInput");
        if (seqInput && parts.length > 1) seqInput.value = parts[1];
        const revInput = document.getElementById("quoteRevInput");
        if (revInput && parts.length > 2) revInput.value = parts[2].replace(/[^0-9]/g, "") || "00";
      }

      const detailCard = document.getElementById("customerDetailCard");
      if (detailCard) detailCard.classList.remove("hidden");
      const comp = document.getElementById("cdCompany");
      if (comp) comp.textContent = header.customer_name || "-";
      const tax = document.getElementById("cdTaxId");
      if (tax) tax.textContent = header.tax_id || "-";
      const cont = document.getElementById("cdContact");
      if (cont) cont.textContent = header.contact_name || "-";
      const phone = document.getElementById("cdPhone");
      if (phone) phone.textContent = header.phone || "-";
      const addr = document.getElementById("cdAddress");
      if (addr) addr.textContent = header.delivery_address || "-";

      const itemsContainer = document.getElementById("itemsContainer");
      if (itemsContainer) itemsContainer.innerHTML = "";
      const desktopTableBody = document.getElementById("quoteDesktopTableBody");
      if (desktopTableBody) desktopTableBody.innerHTML = "";
      quoteItemCount = 0;

      if (items.length === 0) {
        addBlankItem();
      } else {
        items.forEach(itm => addBlankItem(itm));
      }

      calculateTotal();
      updateQuoteNoPreview();
      switchQuoteSubTab("create");

      alert(`[已載入報價單] 已成功載入單號【${header.quote_no}】！\n若於電腦端修改後點擊「儲存報價」，系統將遵循 Append-Only 規則自動生成新版次。`);
    } catch (err) {
      alert("[載入失敗] " + err.message);
    }
  }

  // 子分頁 C：手機報價草稿匣 (直接從後端資料庫 /api/quotation-drafts 讀取，不需經由雲端硬碟)
  async function loadQuotationDraftsInbox() {
    const draftsContainer = document.getElementById("quoteDraftsResults");
    if (!draftsContainer) return;
    draftsContainer.innerHTML = '<div style="padding:24px; text-align:center; color:var(--text-muted); font-size:0.88rem;">正在從資料庫載入手機報價草稿...</div>';

    const activeSales = getActiveSalesName();
    let drafts = [];
    try {
      const resp = await fetch(`/api/quotation-drafts?sales_name=${encodeURIComponent(activeSales)}`);
      if (resp.ok) {
        drafts = await resp.json();
      }
    } catch (err) {
      console.warn("讀取資料庫草稿失敗:", err);
    }

    if (!Array.isArray(drafts) || drafts.length === 0) {
      draftsContainer.innerHTML = `
        <div style="padding:32px; text-align:center; color:var(--text-muted); font-size:0.9rem;">
          目前業務【${escapeHtml(activeSales)}】草稿匣內無任何待處理草稿。<br>
          <span style="font-size:0.8rem; color:var(--text-muted); display:inline-block; margin-top:6px;">
            在手機端點擊「送出草稿」後，資料將直接存入此草稿匣，回到電腦桌前即可在此一鍵載入正式報價單。
          </span>
        </div>
      `;
      return;
    }

    draftsContainer.innerHTML = drafts.map((d, idx) => `
      <div class="quote-history-card" style="margin-bottom:10px;">
        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:6px;">
          <span style="font-weight:700; color:var(--primary-color, #2563eb); font-size:0.96rem;">
            草稿 #${idx + 1} ｜ ${escapeHtml(d.customer_name || '未命名客戶')}
          </span>
          <span style="font-size:0.8rem; color:var(--text-muted); font-family:monospace;">${escapeHtml(d.created_at || '')}</span>
        </div>
        <div style="font-size:0.86rem; color:var(--text-main); margin-bottom:8px; display:flex; justify-content:space-between; align-items:center; flex-wrap:wrap; gap:8px;">
          <span>案場：${escapeHtml(d.project_name || '設備報價')} ｜ 品項數：<strong>${Array.isArray(d.items) ? d.items.length : 0}</strong> 項</span>
          <span style="font-weight:700; color:var(--primary-color, #2563eb);">未稅小計：$ ${Number(d.subtotal || 0).toLocaleString()}</span>
        </div>
        <div style="display:flex; justify-content:flex-end; gap:8px;">
          <button type="button" class="btn-delete-draft" data-id="${d.id}" style="padding:5px 12px; font-size:0.82rem; background:transparent; color:#ef4444; border:1px solid #fecaca; border-radius:4px; cursor:pointer;">刪除草稿</button>
          <button type="button" class="btn-load-draft" data-idx="${idx}" style="padding:5px 14px; font-size:0.82rem; background:var(--primary-color, #2563eb); color:#fff; border:none; border-radius:4px; cursor:pointer; font-weight:600;">載入至電腦編輯區</button>
        </div>
      </div>
    `).join("");

    draftsContainer.querySelectorAll('.btn-load-draft').forEach(btn => {
      btn.addEventListener("click", () => {
        const idx = parseInt(btn.dataset.idx, 10);
        const draft = drafts[idx];
        if (!draft) return;

        currentLoadedDraftId = draft.id;
        currentBaseQuoteNo = null;

        const matchedCust = (customersCache || []).find(c => c.name === draft.customer_name);
        if (matchedCust) {
          applySelectedCustomerToQuoteUI(matchedCust);
        } else {
          const custInput = document.getElementById("customerName");
          if (custInput) custInput.value = draft.customer_name || "";
          const taxInput = document.getElementById("quoteTaxId");
          if (taxInput) taxInput.value = draft.tax_id || "";
          const phoneInput = document.getElementById("quotePhone");
          if (phoneInput) phoneInput.value = draft.phone || "";
          const locInput = document.getElementById("quoteDeliveryLocation");
          if (locInput) locInput.value = draft.delivery_address || "";
        }

        const projInput = document.getElementById("quoteProjectName");
        if (projInput) projInput.value = draft.project_name || "";

        const itemsContainer = document.getElementById("itemsContainer");
        if (itemsContainer) itemsContainer.innerHTML = "";
        const desktopTableBody = document.getElementById("quoteDesktopTableBody");
        if (desktopTableBody) desktopTableBody.innerHTML = "";
        quoteItemCount = 0;

        if (Array.isArray(draft.items) && draft.items.length > 0) {
          draft.items.forEach(itm => addBlankItem(itm));
        } else {
          addBlankItem();
        }
        calculateTotal();
        switchQuoteSubTab("create");
        alert(`[已載入草稿] 手機草稿【${draft.customer_name}】已成功載入編輯區！\n確認無誤後即可點擊「儲存報價」正式產出報價單。`);
      });
    });

    draftsContainer.querySelectorAll('.btn-delete-draft').forEach(btn => {
      btn.addEventListener("click", async () => {
        const draftId = btn.dataset.id;
        if (!confirm("【二次確認】確定要刪除這筆手機報價草稿嗎？")) return;
        try {
          const resp = await fetch(`/api/quotation-drafts/${draftId}?sales_name=${encodeURIComponent(activeSales)}`, {
            method: "DELETE"
          });
          if (resp.ok) {
            loadQuotationDraftsInbox();
          }
        } catch (err) {
          alert("[刪除草稿失敗] " + err.message);
        }
      });
    });
  }

  // 子分頁 D：歷史單價分析 (直接呼叫後端 /api/price-analysis 進行高速 SQL 聚合與業務隔離查詢)
  async function runPriceAnalysis() {
    const resultsContainer = document.getElementById("quoteAnalysisResults");
    if (!resultsContainer) return;

    const modelQuery = document.getElementById("inputAnalysisModel")?.value.trim() || "";
    const custQuery = document.getElementById("inputAnalysisCustomer")?.value.trim() || "";
    const activeSales = getActiveSalesName();

    resultsContainer.innerHTML = '<div style="padding:24px; text-align:center; color:var(--text-muted); font-size:0.88rem;">正在分析個人歷史單價與折讓區間...</div>';

    try {
      const url = `/api/price-analysis?sales_name=${encodeURIComponent(activeSales)}&model_keyword=${encodeURIComponent(modelQuery)}&customer_keyword=${encodeURIComponent(custQuery)}`;
      const resp = await fetch(url);
      if (!resp.ok) throw new Error("讀取歷史單價分析失敗");
      const data = await resp.json();

      const matchedItems = Array.isArray(data) ? data : (data.records || []);
      const prices = matchedItems.map(m => parseFloat(m.final_unit_price) || 0).filter(v => v > 0);
      const summary = data.summary || {
        count: matchedItems.length,
        min_price: prices.length ? Math.min(...prices) : 0,
        max_price: prices.length ? Math.max(...prices) : 0
      };

      if (matchedItems.length === 0) {
        resultsContainer.innerHTML = `
          <div style="padding:32px; text-align:center; color:var(--text-muted); font-size:0.9rem;">
            目前業務【${escapeHtml(activeSales)}】查無相符的歷史單價紀錄。請嘗試調整型號或客戶關鍵字。
          </div>
        `;
        return;
      }

      const discounts = matchedItems.map(m => parseFloat(m.discount_pct) || 0).filter(d => d > 0);
      const avgDisc = discounts.length ? (discounts.reduce((a, b) => a + b, 0) / discounts.length).toFixed(1) : "100.0";

      resultsContainer.innerHTML = `
        <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(140px, 1fr)); gap:12px; padding:12px;">
          <div style="background:var(--bg-card); border:1px solid var(--border-color); border-radius:8px; padding:12px; text-align:center;">
            <div style="font-size:0.75rem; color:var(--text-muted);">歷史最高單價</div>
            <div style="font-size:1.15rem; font-weight:700; color:#ef4444; margin-top:2px;">$ ${Number(summary.max_price || 0).toLocaleString()}</div>
          </div>
          <div style="background:var(--bg-card); border:1px solid var(--border-color); border-radius:8px; padding:12px; text-align:center;">
            <div style="font-size:0.75rem; color:var(--text-muted);">歷史最低單價</div>
            <div style="font-size:1.15rem; font-weight:700; color:#10b981; margin-top:2px;">$ ${Number(summary.min_price || 0).toLocaleString()}</div>
          </div>
          <div style="background:var(--bg-card); border:1px solid var(--border-color); border-radius:8px; padding:12px; text-align:center;">
            <div style="font-size:0.75rem; color:var(--text-muted);">平均折讓率</div>
            <div style="font-size:1.15rem; font-weight:700; color:var(--primary-color, #2563eb); margin-top:2px;">${avgDisc} %</div>
          </div>
          <div style="background:var(--bg-card); border:1px solid var(--border-color); border-radius:8px; padding:12px; text-align:center;">
            <div style="font-size:0.75rem; color:var(--text-muted);">符合報價筆數</div>
            <div style="font-size:1.15rem; font-weight:700; color:var(--text-main); margin-top:2px;">${summary.count || matchedItems.length} 筆</div>
          </div>
        </div>

        <div style="overflow-x:auto;">
          <table style="width:100%; border-collapse:collapse; font-size:0.85rem;">
            <thead>
              <tr style="background:var(--table-header-bg, #f1f5f9); border-bottom:1px solid var(--border-color);">
                <th style="padding:8px 10px; text-align:left;">報價單號</th>
                <th style="padding:8px 10px; text-align:left;">客戶名稱</th>
                <th style="padding:8px 10px; text-align:left;">產品型號</th>
                <th style="padding:8px 10px; text-align:right;">經銷牌價</th>
                <th style="padding:8px 10px; text-align:center;">折數</th>
                <th style="padding:8px 10px; text-align:right;">報價單價</th>
                <th style="padding:8px 10px; text-align:right;">數量</th>
                <th style="padding:8px 10px; text-align:center;">日期</th>
              </tr>
            </thead>
            <tbody>
              ${matchedItems.map(m => `
                <tr style="border-bottom:1px solid var(--border-color);">
                  <td style="padding:8px 10px; font-family:monospace; font-weight:600; color:var(--primary-color);">${escapeHtml(m.quote_no)}</td>
                  <td style="padding:8px 10px;">${escapeHtml(m.customer_name)}</td>
                  <td style="padding:8px 10px; font-weight:600;">${escapeHtml(m.model_no)}</td>
                  <td style="padding:8px 10px; text-align:right; color:var(--text-muted);">$ ${Number(m.dealer_price || 0).toLocaleString()}</td>
                  <td style="padding:8px 10px; text-align:center; font-weight:600;">${Number(m.discount_pct || 100).toFixed(1).replace(/\.0$/, '')}%</td>
                  <td style="padding:8px 10px; text-align:right; font-weight:700; color:#10b981;">$ ${Number(m.final_unit_price || 0).toLocaleString()}</td>
                  <td style="padding:8px 10px; text-align:right;">${Number(m.quantity || 1)}</td>
                  <td style="padding:8px 10px; text-align:center; color:var(--text-muted); font-size:0.8rem;">${escapeHtml(m.quote_date)}</td>
                </tr>
              `).join("")}
            </tbody>
          </table>
        </div>
      `;
    } catch (err) {
      resultsContainer.innerHTML = `<div style="padding:16px; text-align:center; color:var(--color-danger, #ef4444); font-size:0.88rem;">分析失敗：${escapeHtml(err.message)}</div>`;
    }
  }

  // ============================================================================
  // 子分頁 E：個人特殊品經銷牌價表管理 (業務個人獨立 + 方案 A 自動轉正封存留紀錄)
  // ============================================================================
  async function loadSpecialProductsManager() {
    const activeContainer = document.getElementById("specialProductsActiveList");
    const archivedContainer = document.getElementById("specialProductsArchivedList");
    const ownerBadge = document.getElementById("specialProductOwnerBadge");
    const activeSales = getActiveSalesName();

    if (ownerBadge) ownerBadge.textContent = `所屬業務：${activeSales}`;
    if (!activeContainer || !archivedContainer) return;

    activeContainer.innerHTML = '<div style="padding:16px; text-align:center; color:var(--text-muted);">載入中...</div>';
    archivedContainer.innerHTML = '<div style="padding:16px; text-align:center; color:var(--text-muted);">載入中...</div>';

    try {
      const resp = await fetch(`/api/special-products?sales_name=${encodeURIComponent(activeSales)}&include_archived=true`);
      if (!resp.ok) throw new Error("讀取個人特殊品失敗");
      const allItems = await resp.json();

      const activeItems = (allItems || []).filter(item => !item.is_archived);
      const archivedItems = (allItems || []).filter(item => item.is_archived === 1);

      if (activeItems.length === 0) {
        activeContainer.innerHTML = `<div style="padding:20px; text-align:center; color:var(--text-muted); font-size:0.88rem;">目前業務【${escapeHtml(activeSales)}】名下尚無生效中之個人特殊品。可於上方表單新增。</div>`;
      } else {
        activeContainer.innerHTML = `
          <table style="width:100%; border-collapse:collapse; font-size:0.86rem;">
            <thead>
              <tr style="background:var(--table-header-bg, #f1f5f9); border-bottom:1px solid var(--border-color);">
                <th style="padding:8px 10px; text-align:left;">特殊品型號</th>
                <th style="padding:8px 10px; text-align:left;">規格 / 說明</th>
                <th style="padding:8px 10px; text-align:right;">個人特殊經銷價(未稅)</th>
                <th style="padding:8px 10px; text-align:right;">可用庫存</th>
                <th style="padding:8px 10px; text-align:left;">備註</th>
                <th style="padding:8px 10px; text-align:center;">操作</th>
              </tr>
            </thead>
            <tbody>
              ${activeItems.map(item => `
                <tr style="border-bottom:1px solid var(--border-color);">
                  <td style="padding:8px 10px; font-weight:700; font-family:monospace; color:var(--primary-color, #2563eb);">${escapeHtml(item.model_no)}</td>
                  <td style="padding:8px 10px;">${escapeHtml(item.spec || '-')}</td>
                  <td style="padding:8px 10px; text-align:right; font-weight:700;">$ ${Number(item.dealer_price || 0).toLocaleString()}</td>
                  <td style="padding:8px 10px; text-align:right;">${Number(item.stock_qty || 0).toLocaleString()}</td>
                  <td style="padding:8px 10px; color:var(--text-muted); font-size:0.82rem;">${escapeHtml(item.note || item.notes || '-')}</td>
                  <td style="padding:8px 10px; text-align:center; white-space:nowrap;">
                    <button type="button" class="btn-use-special-prod fluent-btn secondary" data-model="${escapeHtml(item.model_no)}" data-spec="${escapeHtml(item.spec || '')}" data-price="${item.dealer_price || 0}" data-stock="${item.stock_qty || 0}" style="font-size:0.78rem; padding:3px 8px; margin-right:4px;">帶入報價單</button>
                    <button type="button" class="btn-promote-special-prod fluent-btn primary" data-id="${item.id}" data-model="${escapeHtml(item.model_no)}" data-spec="${escapeHtml(item.spec || '')}" data-price="${item.dealer_price || 0}" data-stock="${item.stock_qty || 0}" style="font-size:0.78rem; padding:3px 8px; margin-right:4px;" title="模擬或執行公司正式發行此型號為共通標準品（方案A自動取代並封存）">轉標準品發行(方案A)</button>
                    <button type="button" class="btn-del-special-prod fluent-btn danger" data-id="${item.id}" data-model="${escapeHtml(item.model_no)}" style="font-size:0.78rem; padding:3px 8px;">刪除</button>
                  </td>
                </tr>
              `).join("")}
            </tbody>
          </table>
        `;

        activeContainer.querySelectorAll(".btn-use-special-prod").forEach(btn => {
          btn.addEventListener("click", () => {
            addBlankItem({
              model_no: btn.dataset.model,
              spec: btn.dataset.spec,
              dealer_price: parseFloat(btn.dataset.price) || 0,
              stock_qty: parseInt(btn.dataset.stock, 10) || 0,
              is_special: true,
              discount_pct: 100,
              final_unit_price: parseFloat(btn.dataset.price) || 0,
              quantity: 1
            });
            switchQuoteSubTab("create");
            alert(`[已帶入報價單] 個人特殊品【${btn.dataset.model}】已加入目前報價單明細！`);
          });
        });

        activeContainer.querySelectorAll(".btn-promote-special-prod").forEach(btn => {
          btn.addEventListener("click", async () => {
            const modelNo = btn.dataset.model;
            const origPrice = parseFloat(btn.dataset.price) || 0;
            const inputStdPrice = prompt(`【方案 A：特殊品轉公司共通標準品發行確認】\n\n即將把型號【${modelNo}】正式發行至「公司共通標準牌價表」。\n發行後系統將自動以公司標準品取代，並將各業務個人特殊品封存留紀錄。\n\n請輸入公司正式發行之「標準經銷牌價(未稅)」：`, String(origPrice));
            if (inputStdPrice === null) return;
            const stdPrice = parseFloat(inputStdPrice);
            if (isNaN(stdPrice) || stdPrice <= 0) {
              alert("請輸入有效的標準經銷牌價數字！");
              return;
            }
            if (!confirm(`【二次確認】確定要將【${modelNo}】以標準牌價 $${stdPrice.toLocaleString()} 正式發行至共通標準牌價表，並自動封存個人特殊品嗎？`)) {
              return;
            }
            try {
              const spId = btn.dataset.id;
              const r = await fetch(`/api/special-products/${spId}/promote`, {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify({
                  standard_dealer_price: stdPrice
                })
              });
              const resData = await r.json();
              if (!r.ok) throw new Error(resData.detail || "轉正發行失敗");
              quoteProductsCache = [];
              alert(`[方案 A 執行成功] ${resData.message}`);
              loadSpecialProductsManager();
            } catch (e) {
              alert("[執行失敗] " + e.message);
            }
          });
        });

        activeContainer.querySelectorAll(".btn-del-special-prod").forEach(btn => {
          btn.addEventListener("click", async () => {
            const spId = btn.dataset.id;
            const modelNo = btn.dataset.model;
            if (!confirm(`【二次確認】確定要刪除個人特殊品【${modelNo}】嗎？`)) return;
            try {
              const r = await fetch(`/api/special-products/${spId}?sales_name=${encodeURIComponent(activeSales)}`, {
                method: "DELETE"
              });
              if (r.ok) {
                quoteProductsCache = [];
                loadSpecialProductsManager();
              }
            } catch (e) {
              alert("[刪除失敗] " + e.message);
            }
          });
        });
      }

      if (archivedItems.length === 0) {
        archivedContainer.innerHTML = `<div style="padding:16px; text-align:center; color:var(--text-muted); font-size:0.86rem;">尚無封存紀錄</div>`;
      } else {
        archivedContainer.innerHTML = `
          <table style="width:100%; border-collapse:collapse; font-size:0.84rem;">
            <thead>
              <tr style="background:var(--table-header-bg, #f1f5f9); border-bottom:1px solid var(--border-color);">
                <th style="padding:8px 10px; text-align:left;">型號</th>
                <th style="padding:8px 10px; text-align:right;">原個人特殊價</th>
                <th style="padding:8px 10px; text-align:right;">轉正時公司標準牌價</th>
                <th style="padding:8px 10px; text-align:center;">封存時間</th>
                <th style="padding:8px 10px; text-align:left;">方案 A 封存備註</th>
              </tr>
            </thead>
            <tbody>
              ${archivedItems.map(item => `
                <tr style="border-bottom:1px solid var(--border-color); opacity:0.88;">
                  <td style="padding:8px 10px; font-weight:700; font-family:monospace;">${escapeHtml(item.model_no)}</td>
                  <td style="padding:8px 10px; text-align:right; text-decoration:line-through; color:var(--text-muted);">$ ${Number(item.dealer_price || 0).toLocaleString()}</td>
                  <td style="padding:8px 10px; text-align:right; font-weight:700; color:#059669;">$ ${Number(item.archived_standard_price || 0).toLocaleString()}</td>
                  <td style="padding:8px 10px; text-align:center; font-family:monospace; font-size:0.8rem;">${escapeHtml(item.archived_at || '-')}</td>
                  <td style="padding:8px 10px; color:var(--text-secondary);">${escapeHtml(item.archived_note || '')}</td>
                </tr>
              `).join("")}
            </tbody>
          </table>
        `;
      }
    } catch (err) {
      activeContainer.innerHTML = `<div style="padding:16px; text-align:center; color:var(--color-danger);">${escapeHtml(err.message)}</div>`;
    }
  }

  async function savePersonalSpecialProductFromUI() {
    const modelNo = document.getElementById("inputSpecialModelNo")?.value.trim() || "";
    const spec = document.getElementById("inputSpecialSpec")?.value.trim() || "";
    const priceVal = parseFloat(document.getElementById("inputSpecialPrice")?.value || "0");
    const stockVal = parseInt(document.getElementById("inputSpecialStockQty")?.value || "0", 10) || 0;
    const notes = document.getElementById("inputSpecialNotes")?.value.trim() || "";
    const activeSales = getActiveSalesName();

    if (!modelNo) {
      alert("請輸入特殊品型號！");
      return;
    }
    if (isNaN(priceVal) || priceVal <= 0) {
      alert("請輸入大於 0 的個人特殊經銷價(未稅)！");
      return;
    }

    try {
      const resp = await fetch("/api/special-products", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          owner_sales: activeSales,
          model_no: modelNo,
          spec: spec,
          dealer_price: priceVal,
          stock_qty: stockVal,
          note: notes
        })
      });
      const resData = await resp.json();
      if (!resp.ok) {
        throw new Error(resData.detail || "儲存特殊品失敗");
      }

      quoteProductsCache = [];
      document.getElementById("inputSpecialModelNo").value = "";
      document.getElementById("inputSpecialSpec").value = "";
      document.getElementById("inputSpecialPrice").value = "";
      document.getElementById("inputSpecialStockQty").value = "0";
      document.getElementById("inputSpecialNotes").value = "";

      alert(`[${resData.status === 'archived_by_standard' ? '方案 A 自動轉正提醒' : '儲存成功'}] ${resData.message}`);
      loadSpecialProductsManager();
    } catch (err) {
      alert("[儲存失敗] " + err.message);
    }
  }

  // ============================================================================
  // 報價管理系統新增完整功能頁：手機草稿徽章、大數據分析、客戶管理、產品搜尋與計算機、系統設定
  // ============================================================================
  async function updateQuoteDraftsBadge() {
    const badge = document.getElementById("quoteDraftsCountBadge");
    if (!badge) return;
    try {
      const activeSales = getActiveSalesName();
      const r = await fetch(`/api/quotation-drafts?sales_name=${encodeURIComponent(activeSales)}`);
      if (r.ok) {
        const list = await r.json();
        const cnt = Array.isArray(list) ? list.length : 0;
        badge.textContent = String(cnt);
        badge.classList.toggle("hidden", cnt === 0);
      }
    } catch (e) {}
  }

  // 子分頁 F：大數據分析 (_QuotationApp gui_dashboard.py)
  async function loadQuotationDashboard() {
    const activeSales = getActiveSalesName();
    const latestOnly = document.getElementById("chkQuoteDashLatestOnly")?.checked !== false;
    try {
      const r = await fetch(`/api/quotation-dashboard?sales_name=${encodeURIComponent(activeSales)}&latest_only=${latestOnly}`);
      if (!r.ok) throw new Error("讀取大數據統計失敗");
      const d = await r.json();

      const elTotalQuoted = document.getElementById("qDashTotalQuoted");
      const elTotalCount = document.getElementById("qDashTotalCount");
      const elWonAmount = document.getElementById("qDashWonAmount");
      const elWonCount = document.getElementById("qDashWonCount");
      const elWinRateAmt = document.getElementById("qDashWinRateAmt");
      const elWinRateCnt = document.getElementById("qDashWinRateCnt");
      const elPendingAmt = document.getElementById("qDashPendingAmount");
      const elPendingCnt = document.getElementById("qDashPendingCount");

      if (elTotalQuoted) elTotalQuoted.textContent = `$ ${Number(d.total_quoted_amount || 0).toLocaleString()}`;
      if (elTotalCount) elTotalCount.textContent = `有效報價單：${d.total_count || 0} 筆（排除不計入 ${d.ignored_count || 0} 筆）`;
      if (elWonAmount) elWonAmount.textContent = `$ ${Number(d.won_amount || 0).toLocaleString()}`;
      if (elWonCount) elWonCount.textContent = `已成交筆數：${d.won_count || 0} 筆`;
      if (elWinRateAmt) elWinRateAmt.textContent = `${Number(d.win_rate_amount_pct || 0).toFixed(1)}%`;
      if (elWinRateCnt) elWinRateCnt.textContent = `件數成交率：${Number(d.win_rate_count_pct || 0).toFixed(1)}%`;
      if (elPendingAmt) elPendingAmt.textContent = `$ ${Number(d.pending_amount || 0).toLocaleString()}`;
      if (elPendingCnt) elPendingCnt.textContent = `追蹤中：${d.pending_count || 0} 筆 ｜ 已失注：${d.lost_count || 0} 筆`;

      const custBody = document.getElementById("qDashCustomerRankBody");
      if (custBody) {
        const topCusts = d.top_customers || [];
        if (topCusts.length === 0) {
          custBody.innerHTML = `<tr><td colspan="7" style="text-align:center; padding:16px; color:var(--text-muted);">目前尚無報價統計資料</td></tr>`;
        } else {
          custBody.innerHTML = topCusts.map((c, idx) => `
            <tr>
              <td style="text-align:center; font-weight:700;">${idx + 1}</td>
              <td style="font-weight:600;">${escapeHtml(c.customer_name)}</td>
              <td style="text-align:right;">${c.quote_count}</td>
              <td style="text-align:right; font-family:monospace;">$ ${Number(c.total_quoted || 0).toLocaleString()}</td>
              <td style="text-align:right; font-family:monospace; color:#34d399; font-weight:700;">$ ${Number(c.won_amount || 0).toLocaleString()}</td>
              <td style="text-align:right; font-family:monospace; color:#c084fc;">$ ${Number(c.pending_amount || 0).toLocaleString()}</td>
              <td style="text-align:right; font-weight:700;">${Number(c.win_rate_pct || 0).toFixed(1)}%</td>
            </tr>
          `).join("");
        }
      }

      const projBody = document.getElementById("qDashProjectTypeBody");
      if (projBody) {
        const pTypes = d.project_types || [];
        if (pTypes.length === 0) {
          projBody.innerHTML = `<tr><td colspan="4" style="text-align:center; padding:16px; color:var(--text-muted);">尚無資料</td></tr>`;
        } else {
          projBody.innerHTML = pTypes.map(p => `
            <tr>
              <td style="font-weight:600;">${escapeHtml(p.project_type)}</td>
              <td style="text-align:right;">${p.quote_count}</td>
              <td style="text-align:right; font-family:monospace;">$ ${Number(p.total_quoted || 0).toLocaleString()}</td>
              <td style="text-align:right; font-family:monospace; color:#34d399;">$ ${Number(p.won_amount || 0).toLocaleString()}</td>
            </tr>
          `).join("");
        }
      }
    } catch (err) {
      console.warn("大數據分析載入失敗:", err);
    }
  }

  // 子分頁 G：客戶管理 (_QuotationApp gui_page4.py)
  let selectedMgmtCustomerId = null;

  async function loadQuotationCustomersPage() {
    await populateCustomerSelectDropdown(true);
    renderQuotationCustomersTable();
    await loadQuotationSupplierInfo();
  }

  function renderQuotationCustomersTable() {
    const tbody = document.getElementById("quoteCustMgmtTableBody");
    const badge = document.getElementById("qCustCountBadge");
    const kw = (document.getElementById("inputQuoteCustSearch")?.value || "").trim().toLowerCase();
    if (!tbody) return;

    const list = (customersCache || []).filter(c => {
      if (!kw) return true;
      return (
        (c.name || "").toLowerCase().includes(kw) ||
        (c.tax_id || "").toLowerCase().includes(kw) ||
        (c.default_contact || "").toLowerCase().includes(kw) ||
        (c.phone || "").toLowerCase().includes(kw)
      );
    });

    if (badge) badge.textContent = `共 ${list.length} 家客戶`;

    if (list.length === 0) {
      tbody.innerHTML = `<tr><td colspan="5" style="text-align:center; padding:18px; color:var(--text-muted);">查無相符的客戶資料</td></tr>`;
      return;
    }

    tbody.innerHTML = list.map(c => {
      const isFav = Number(c.is_favorite || 0) > 0;
      const isSel = Number(c.id) === Number(selectedMgmtCustomerId);
      return `
        <tr class="quote-cust-row ${isSel ? 'selected-row' : ''}" data-cid="${c.id}">
          <td style="text-align:center; font-weight:700; color:${isFav ? '#fbbf24' : 'var(--text-muted)'};">${isFav ? '[常用]' : '-'}</td>
          <td style="font-weight:600;">${escapeHtml(c.name)}</td>
          <td style="font-family:monospace;">${escapeHtml(c.tax_id || '-')}</td>
          <td>${escapeHtml(c.default_contact || '-')}</td>
          <td style="font-family:monospace;">${escapeHtml(c.phone || '-')}</td>
        </tr>
      `;
    }).join("");

    tbody.querySelectorAll(".quote-cust-row").forEach(tr => {
      tr.addEventListener("click", () => {
        const cid = parseInt(tr.getAttribute("data-cid"), 10);
        const cust = (customersCache || []).find(x => Number(x.id) === cid);
        if (cust) {
          selectCustomerInMgmtForm(cust);
        }
      });
    });
  }

  function selectCustomerInMgmtForm(cust) {
    selectedMgmtCustomerId = cust.id;
    renderQuotationCustomersTable();
    const hint = document.getElementById("qCustSelectedHint");
    if (hint) hint.textContent = `已選取：${cust.name} (ID: ${cust.id})`;

    const contacts = Array.isArray(cust.contacts) && cust.contacts.length > 0
      ? cust.contacts
      : [{ name: cust.default_contact || "", email: "", mobile: "" }];
    const primary = contacts[0] || {};

    document.getElementById("qCustEditId").value = String(cust.id || "");
    document.getElementById("qCustName").value = cust.name || "";
    document.getElementById("qCustTax").value = cust.tax_id || "";
    document.getElementById("qCustContact").value = primary.name || cust.default_contact || "";
    document.getElementById("qCustMobile").value = primary.mobile || "";
    document.getElementById("qCustEmail").value = primary.email || "";
    document.getElementById("qCustPhone").value = cust.phone || "";
    document.getElementById("qCustFax").value = cust.fax || "";
    document.getElementById("qCustPayment").value = cust.payment_terms || "月結 30 天電匯";
    document.getElementById("qCustAddress").value = cust.address || "";
    document.getElementById("qCustFavorite").checked = Number(cust.is_favorite || 0) > 0;

    const pillsBox = document.getElementById("qCustContactsPills");
    if (pillsBox) {
      if (contacts.length > 0 && contacts[0].name) {
        pillsBox.innerHTML = `<span style="font-size:0.76rem; color:var(--text-muted); align-self:center;">已建檔窗口（點擊切換編輯）：</span>` +
          contacts.map((ct, idx) => `
            <button type="button" class="contact-pill-btn" data-idx="${idx}">
              ${escapeHtml(ct.name)}${ct.mobile ? ' (' + escapeHtml(ct.mobile) + ')' : ''}
            </button>
          `).join("");
        pillsBox.querySelectorAll(".contact-pill-btn").forEach(btn => {
          btn.addEventListener("click", () => {
            const cObj = contacts[parseInt(btn.getAttribute("data-idx"), 10)];
            if (cObj) {
              document.getElementById("qCustContact").value = cObj.name || "";
              document.getElementById("qCustMobile").value = cObj.mobile || "";
              document.getElementById("qCustEmail").value = cObj.email || "";
            }
          });
        });
      } else {
        pillsBox.innerHTML = "";
      }
    }
  }

  function clearQuotationCustomerForm() {
    selectedMgmtCustomerId = null;
    document.getElementById("qCustEditId").value = "";
    document.getElementById("qCustName").value = "";
    document.getElementById("qCustTax").value = "";
    document.getElementById("qCustContact").value = "";
    document.getElementById("qCustMobile").value = "";
    document.getElementById("qCustEmail").value = "";
    document.getElementById("qCustPhone").value = "";
    document.getElementById("qCustFax").value = "";
    document.getElementById("qCustPayment").value = "月結 30 天電匯";
    document.getElementById("qCustAddress").value = "";
    document.getElementById("qCustFavorite").checked = false;
    const pillsBox = document.getElementById("qCustContactsPills");
    if (pillsBox) pillsBox.innerHTML = "";
    const hint = document.getElementById("qCustSelectedHint");
    if (hint) hint.textContent = "尚未選取客戶（填寫右側表單可直接新增）";
    renderQuotationCustomersTable();
  }

  async function saveQuotationCustomerFromUI(copyAsNewContact = false) {
    const name = document.getElementById("qCustName")?.value.trim() || "";
    if (!name) {
      alert("請輸入公司名稱！");
      return;
    }
    const contactName = document.getElementById("qCustContact")?.value.trim() || "";
    if (copyAsNewContact && !contactName) {
      alert("若要複製為新聯絡人，請先在「聯絡人姓名」輸入新窗口姓名！");
      return;
    }
    const activeSales = getActiveSalesName();
    const payload = {
      name: name,
      owner_sales: activeSales,
      tax_id: document.getElementById("qCustTax")?.value.trim() || "",
      contact_name: contactName,
      contact_mobile: document.getElementById("qCustMobile")?.value.trim() || "",
      contact_email: document.getElementById("qCustEmail")?.value.trim() || "",
      phone: document.getElementById("qCustPhone")?.value.trim() || "",
      fax: document.getElementById("qCustFax")?.value.trim() || "",
      payment_terms: document.getElementById("qCustPayment")?.value.trim() || "月結 30 天電匯",
      address: document.getElementById("qCustAddress")?.value.trim() || "",
      is_favorite: !!document.getElementById("qCustFavorite")?.checked,
      copy_as_new_contact: copyAsNewContact
    };

    try {
      const r = await fetch("/api/quotation-customers", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload)
      });
      const res = await r.json();
      if (!r.ok) throw new Error(res.detail || "儲存客戶失敗");
      await populateCustomerSelectDropdown(true);
      const savedCust = (customersCache || []).find(c => Number(c.id) === Number(res.id) || c.name === name);
      if (savedCust) selectCustomerInMgmtForm(savedCust);
      else renderQuotationCustomersTable();
      alert(`[儲存成功] ${res.message}`);
    } catch (err) {
      alert("[儲存失敗] " + err.message);
    }
  }

  async function deleteSelectedQuotationCustomer() {
    if (!selectedMgmtCustomerId) {
      alert("請先從左側列表點選要刪除的客戶公司！");
      return;
    }
    const cust = (customersCache || []).find(c => Number(c.id) === Number(selectedMgmtCustomerId));
    const custName = cust ? cust.name : `ID ${selectedMgmtCustomerId}`;
    if (!confirm(`【二次確認】確定要刪除客戶「${custName}」及其聯絡人資料嗎？此動作無法復原。`)) {
      return;
    }
    try {
      const activeSales = getActiveSalesName();
      const r = await fetch(`/api/quotation-customers/${selectedMgmtCustomerId}?sales_name=${encodeURIComponent(activeSales)}`, {
        method: "DELETE"
      });
      if (!r.ok) {
        const err = await r.json().catch(() => ({}));
        throw new Error(err.detail || "刪除失敗");
      }
      clearQuotationCustomerForm();
      await populateCustomerSelectDropdown(true);
      renderQuotationCustomersTable();
      alert(`[已刪除] 客戶「${custName}」已成功刪除。`);
    } catch (err) {
      alert("[刪除失敗] " + err.message);
    }
  }

  async function loadQuotationSupplierInfo() {
    try {
      const activeSales = getActiveSalesName();
      const r = await fetch(`/api/quotation-supplier?sales_name=${encodeURIComponent(activeSales)}`);
      if (r.ok) {
        const sup = await r.json();
        document.getElementById("qSupCompany").value = sup.company || "士林電機廠股份有限公司";
        document.getElementById("qSupTax").value = sup.tax || "11039306";
        document.getElementById("qSupSales").value = sup.sales || activeSales;
        document.getElementById("qSupMobile").value = sup.mobile || "";
        document.getElementById("qSupPhone").value = sup.phone || "";
        document.getElementById("qSupFax").value = sup.fax || "";
        document.getElementById("qSupEmail").value = sup.email || "";
        document.getElementById("qSupLine").value = sup.line || "";
        document.getElementById("qSupAddress").value = sup.address || "";
      }
    } catch (e) {}
  }

  async function saveQuotationSupplierFromUI() {
    const activeSales = getActiveSalesName();
    const payload = {
      sales_name: activeSales,
      company: document.getElementById("qSupCompany")?.value.trim() || "",
      tax: document.getElementById("qSupTax")?.value.trim() || "",
      sales: document.getElementById("qSupSales")?.value.trim() || activeSales,
      mobile: document.getElementById("qSupMobile")?.value.trim() || "",
      phone: document.getElementById("qSupPhone")?.value.trim() || "",
      fax: document.getElementById("qSupFax")?.value.trim() || "",
      email: document.getElementById("qSupEmail")?.value.trim() || "",
      line: document.getElementById("qSupLine")?.value.trim() || "",
      address: document.getElementById("qSupAddress")?.value.trim() || ""
    };
    try {
      const r = await fetch("/api/quotation-supplier", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload)
      });
      const res = await r.json();
      if (!r.ok) throw new Error(res.detail || "儲存供應商設定失敗");
      alert(`[儲存成功] ${res.message}`);
    } catch (err) {
      alert("[儲存失敗] " + err.message);
    }
  }

  // 子分頁 E 擴充：產品管理標準品檢索與內建快速計算機 (_QuotationApp gui_page5.py)
  async function searchStandardProductsInMgmt() {
    const tbody = document.getElementById("quoteProdMgmtTableBody");
    if (!tbody) return;
    const q = document.getElementById("inputQuoteProdMgmtSearch")?.value.trim() || "";
    const activeSales = getActiveSalesName();
    tbody.innerHTML = `<tr><td colspan="5" style="text-align:center; padding:16px; color:var(--text-muted);">搜尋中...</td></tr>`;
    try {
      const r = await fetch(`/api/products?q=${encodeURIComponent(q)}&limit=80&sales_name=${encodeURIComponent(activeSales)}`);
      if (!r.ok) throw new Error("搜尋產品失敗");
      const list = await r.json();
      if (!Array.isArray(list) || list.length === 0) {
        tbody.innerHTML = `<tr><td colspan="5" style="text-align:center; padding:16px; color:var(--text-muted);">查無相符的產品型號</td></tr>`;
        return;
      }
      tbody.innerHTML = list.map(p => `
        <tr class="quote-cust-row" data-model="${escapeHtml(p.model_no)}" data-spec="${escapeHtml(p.spec || '')}" data-price="${p.dealer_price || 0}" data-stock="${p.stock_qty || 0}" title="點擊可帶入上方特殊品編輯區或計算機">
          <td><span class="status-pill ${p.is_special ? 'warning' : 'info'}" style="font-size:0.72rem;">${escapeHtml(p.source_label || '標準品')}</span></td>
          <td style="font-family:monospace; font-weight:700; color:#60a5fa;">${escapeHtml(p.model_no)}</td>
          <td style="font-size:0.82rem;">${escapeHtml(p.spec || '-')}</td>
          <td style="text-align:right; font-family:monospace; font-weight:700;">$ ${Number(p.dealer_price || 0).toLocaleString()}</td>
          <td style="text-align:right; font-family:monospace; color:${Number(p.stock_qty || 0) > 0 ? '#34d399' : 'var(--text-muted)'};">${Number(p.stock_qty || 0).toLocaleString()}</td>
        </tr>
      `).join("");

      tbody.querySelectorAll("tr.quote-cust-row").forEach(tr => {
        tr.addEventListener("click", () => {
          const baseInput = document.getElementById("inputCalcBasePrice");
          if (baseInput) {
            baseInput.value = tr.getAttribute("data-price") || "0";
            updateQuickDiscountCalc();
          }
        });
      });
    } catch (err) {
      tbody.innerHTML = `<tr><td colspan="5" style="text-align:center; padding:16px; color:var(--color-danger);">${escapeHtml(err.message)}</td></tr>`;
    }
  }

  function runQuickExpressionCalc() {
    const expr = (document.getElementById("inputQuickCalcExpr")?.value || "").replace(/[^0-9+\-*/().\s]/g, "").trim();
    const disp = document.getElementById("quickCalcResultDisplay");
    if (!disp) return;
    if (!expr) {
      disp.textContent = "計算結果：0";
      return;
    }
    try {
      const val = Function(`"use strict"; return (${expr});`)();
      const num = Number(val);
      if (isFinite(num)) {
        const rounded = applyQuoteRounding(num);
        disp.textContent = `原始值：${num.toLocaleString(undefined, { maximumFractionDigits: 2 })} ｜ 系統捨入後：$ ${rounded.toLocaleString()}`;
      } else {
        disp.textContent = "計算結果：算式無效";
      }
    } catch (e) {
      disp.textContent = "計算結果：算式格式錯誤";
    }
  }

  function updateQuickDiscountCalc() {
    const base = parseFloat(document.getElementById("inputCalcBasePrice")?.value || "0") || 0;
    const disc = parseFloat(document.getElementById("inputCalcDiscountPct")?.value || "100") || 0;
    const raw = base * (disc / 100.0);
    const untaxed = applyQuoteRounding(raw);
    const taxed = Math.round(untaxed * (1 + (parseFloat(quoteFinanceSettings.tax_rate) || 5.0) / 100.0));
    const summaryEl = document.getElementById("quickCalcDiscountSummary");
    if (summaryEl) {
      summaryEl.textContent = `折後未稅價：$ ${untaxed.toLocaleString()} ｜ 含稅價：$ ${taxed.toLocaleString()}`;
    }
  }

  // 子分頁 H：系統設定 (_QuotationApp gui_page6.py)
  async function loadQuotationSettingsPage() {
    await loadQuotationFinanceSettings();
    const taxInput = document.getElementById("qSetTaxRate");
    const digitsSel = document.getElementById("qSetRoundDigits");
    const modeSel = document.getElementById("qSetRoundMode");
    const sortSel = document.getElementById("qSetCustSortMode");
    const termsTxt = document.getElementById("qSetTermsTextarea");

    if (taxInput) taxInput.value = String(quoteFinanceSettings.tax_rate ?? 5.0);
    if (digitsSel) digitsSel.value = String(quoteFinanceSettings.round_digits ?? 2);
    if (modeSel) modeSel.value = quoteFinanceSettings.round_mode || "無條件捨去";
    if (sortSel) sortSel.value = quoteFinanceSettings.customer_sort_mode || "依筆畫/字母排序";
    if (termsTxt) termsTxt.value = quoteFinanceSettings.quote_terms || "";
  }

  async function saveQuotationFinanceSettingsFromUI() {
    const payload = {
      tax_rate: parseFloat(document.getElementById("qSetTaxRate")?.value || "5.0"),
      round_digits: parseInt(document.getElementById("qSetRoundDigits")?.value || "2", 10),
      round_mode: document.getElementById("qSetRoundMode")?.value || "無條件捨去",
      customer_sort_mode: document.getElementById("qSetCustSortMode")?.value || "依筆畫/字母排序"
    };
    try {
      const r = await fetch("/api/quotation-settings", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(payload)
      });
      const res = await r.json();
      if (!r.ok) throw new Error(res.detail || "儲存財務設定失敗");
      await loadQuotationFinanceSettings();
      calculateTotal();
      alert(`[儲存成功] ${res.message}（最新捨入規則：${quoteFinanceSettings.round_summary_text}）`);
    } catch (err) {
      alert("[儲存失敗] " + err.message);
    }
  }

  async function saveQuotationTermsFromUI() {
    const terms = document.getElementById("qSetTermsTextarea")?.value || "";
    try {
      const r = await fetch("/api/quotation-terms", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ quote_terms: terms })
      });
      const res = await r.json();
      if (!r.ok) throw new Error(res.detail || "儲存條款失敗");
      quoteFinanceSettings.quote_terms = res.quote_terms || terms;
      alert(`[儲存成功] ${res.message}`);
    } catch (err) {
      alert("[儲存失敗] " + err.message);
    }
  }

  // 初始載入設定與數據
  initSettings();
  setupScoreboardScopeSwitch();
  setupGlobalDriverToggle();
  renderHomeDashboard();
  setupDriverModalEvents();

  const curSales = localStorage.getItem("shihlin_pwa_active_sales") || "曾維崧";
  loadAndRenderDriverMetrics(curSales);
  renderManagerReviewPanel(curSales);
  updateManagerUIVisibility(curSales);

  // 初始載入客戶主檔與後端日報
  loadCustomers();
  loadDailyReports();
  loadDeals();
  initQuotationModule();

  // ============================================================================
  // 15. Google 帳號身分驗證 (Google Identity Services) 與管理員身分模擬切換器
  //     (對齊舊版 _QuotationApp/pwa 驗證架構，並提供一鍵切換全體業務身分功能)
  // ============================================================================
  const GOOGLE_WEB_CLIENT_ID = "668571991428-ffjs6ud0apusi7akb0lmptae24qqtbto.apps.googleusercontent.com";
  const GOOGLE_AUTH_SCOPES = "https://www.googleapis.com/auth/drive.file profile email";
  const GAS_WHITELIST_URL = "https://script.google.com/macros/s/AKfycbx48PMWVPysN9gd4OcPq1JmzqkRzwu494C2jFxK71Al13Q4Lr2y5KP3RbS80pgs8CYxGg/exec";
  const SUPER_ADMIN_EMAILS = ["tsengweisung@gmail.com", "faldn3@gmail.com"];

  let googleTokenClient = null;
  let googleAccessToken = localStorage.getItem("google_access_token") || null;
  let authToastTimer = null;

  function isLocalOrTestEnvironment() {
    const host = window.location.hostname;
    return (
      host === "localhost" ||
      host === "127.0.0.1" ||
      host.startsWith("192.168.") ||
      host.startsWith("172.") ||
      host.startsWith("10.") ||
      host.endsWith(".local")
    );
  }

  function showAuthToast(msg) {
    const toast = document.getElementById("authToastBanner");
    if (!toast) return;
    toast.textContent = msg;
    toast.classList.remove("hidden");
    if (authToastTimer) clearTimeout(authToastTimer);
    authToastTimer = setTimeout(() => {
      toast.classList.add("hidden");
    }, 2800);
  }

  function setLoginStatusMessage(msg, isError = false) {
    const box = document.getElementById("loginStatusMessage");
    if (!box) return;
    if (!msg) {
      box.classList.add("hidden");
      box.textContent = "";
      return;
    }
    box.textContent = msg;
    box.classList.toggle("error", !!isError);
    box.classList.remove("hidden");
  }

  // 渲染頂部「管理員身分模擬」下拉選單
  function renderAdminImpersonateOptions() {
    const adminSelect = document.getElementById("adminImpersonateSelect");
    if (!adminSelect) return;
    const activeSales = localStorage.getItem("shihlin_pwa_active_sales") || "曾維崧";
    const list = getOrderedSalespersons();
    adminSelect.innerHTML = list.map(s => {
      const selected = s.name === activeSales ? "selected" : "";
      return `<option value="${escapeHtml(s.name)}" ${selected}>${escapeHtml(s.name)}</option>`;
    }).join("");
    if (list.some(s => s.name === activeSales)) {
      adminSelect.value = activeSales;
    }
  }

  // 評估是否顯示頂部「管理員身分模擬」工具列
  function setupAdminImpersonator() {
    const adminBar = document.getElementById("adminImpersonateBar");
    if (!adminBar) return;

    let profEmail = "";
    try {
      const pStr = localStorage.getItem("saved_user_profile");
      if (pStr) profEmail = (JSON.parse(pStr).email || "").trim().toLowerCase();
    } catch (e) {}

    const isAdmin =
      localStorage.getItem("is_admin_user") === "true" ||
      SUPER_ADMIN_EMAILS.includes(profEmail);

    if (isAdmin) {
      localStorage.setItem("is_admin_user", "true");
      adminBar.classList.remove("hidden");
      renderAdminImpersonateOptions();
    } else {
      adminBar.classList.add("hidden");
    }
  }

  // 核心：切換模擬業務身分，同步刷新全系統所有模組視角與資料隔離
  function switchSimulatedSalesperson(targetSales, showNotice = true) {
    if (!targetSales) return;
    localStorage.setItem("shihlin_pwa_active_sales", targetSales);
    localStorage.setItem("saved_display_name", targetSales);
    currentViewingDriverSales = targetSales;

    const selectActiveSales = document.getElementById("selectActiveSales");
    if (selectActiveSales && selectActiveSales.value !== targetSales) {
      selectActiveSales.value = targetSales;
    }
    const adminSelect = document.getElementById("adminImpersonateSelect");
    if (adminSelect && adminSelect.value !== targetSales) {
      adminSelect.value = targetSales;
    }

    // 1. 更新頂部頭像、首頁戰情計分板、業績預測與個人特殊品徽章
    updateActiveSalesUI(targetSales);
    // 2. 依被模擬業務的主管權限動態顯示/隱藏「業績數據資料庫」與「領先指標審核中心」
    updateManagerUIVisibility(targetSales);
    // 3. 重新載入領先指標與主管審核面板
    loadAndRenderDriverMetrics(targetSales);
    renderManagerReviewPanel(targetSales);
    if (panelDriverMgmt && !panelDriverMgmt.classList.contains("hidden")) {
      renderDriverManagementPage(targetSales);
    }
    // 4. 重新載入客戶、日報與案件
    loadCustomers();
    loadDailyReports();
    loadDeals();
    // 5. 重新載入報價單模組（獨立客戶、個人特殊品、歷史報價單、手機草稿匣）
    if (typeof onQuotationSalespersonChanged === "function") {
      onQuotationSalespersonChanged(targetSales);
    }

    if (showNotice) {
      showAuthToast(`[身分切換] 已切換至「${targetSales}」視角（客戶、報價單與權限已同步隔離）`);
    }
  }

  // 呼叫 Google UserInfo API 取得登入者姓名與 Email
  async function fetchGoogleUserProfile(token) {
    const res = await fetch("https://www.googleapis.com/oauth2/v3/userinfo", {
      headers: { Authorization: `Bearer ${token}` }
    });
    if (!res.ok) {
      throw new Error("無法取得 Google 帳號個人資訊 (HTTP " + res.status + ")");
    }
    return await res.json();
  }

  // 檢查 Email 是否在白名單內 (對齊舊版 checkWhitelist)
  async function checkGoogleWhitelist(email) {
    if (!email) return { allowed: false, msg: "未提供 Google 帳號 Email" };
    const lowerEmail = email.trim().toLowerCase();

    // 最高管理員帳號直接授權放行
    if (SUPER_ADMIN_EMAILS.includes(lowerEmail)) {
      return {
        allowed: true,
        name: "曾維崧",
        is_admin: true
      };
    }

    try {
      const res = await fetch(`${GAS_WHITELIST_URL}?action=check_whitelist&email=${encodeURIComponent(lowerEmail)}`);
      if (!res.ok) {
        return { allowed: true };
      }
      const data = await res.json();
      if (data.status === "ok") {
        return { allowed: true, name: data.name, is_admin: !!data.is_admin };
      } else if (data.status === "rejected") {
        return { allowed: false, msg: data.msg || "[存取受限] 您的帳號尚未通過管理員審核。" };
      }
      return { allowed: false, msg: data.msg || "白名單驗證失敗" };
    } catch (err) {
      console.warn("[Auth] 白名單連線檢查異常，允許離線存取:", err);
      return { allowed: true };
    }
  }

  // 將登入者名稱對應至系統內的業務人員名單
  function resolveSalespersonName(rawName) {
    const list = getOrderedSalespersons();
    if (!rawName) return "曾維崧";
    const clean = String(rawName).replace(/\(.*?\)/g, "").trim();
    const matched = list.find(s => s.name === clean || clean.includes(s.name) || s.name.includes(clean));
    return matched ? matched.name : clean;
  }

  // Google OAuth2 授權成功回呼
  async function handleGoogleTokenResponse(response) {
    if (response.error) {
      console.warn("[Auth] Google 授權回應錯誤:", response.error, response.error_description);
      setLoginStatusMessage(`Google 授權未完成（${response.error_description || response.error}）。若於本機測試，可直接點擊下方「管理員測試權限直接進入」。`, true);
      return;
    }

    googleAccessToken = response.access_token;
    const expiresIn = parseInt(response.expires_in, 10) || 3600;
    const expiresAt = Date.now() + (expiresIn - 60) * 1000;
    localStorage.setItem("google_access_token", googleAccessToken);
    localStorage.setItem("google_token_expires_at", expiresAt.toString());

    setLoginStatusMessage("正在驗證您的 Google 帳號與系統白名單權限...", false);

    try {
      const profile = await fetchGoogleUserProfile(googleAccessToken);
      const checkResult = await checkGoogleWhitelist(profile.email);
      if (!checkResult.allowed) {
        googleAccessToken = null;
        localStorage.removeItem("google_access_token");
        localStorage.removeItem("google_token_expires_at");
        localStorage.removeItem("is_admin_user");
        setLoginStatusMessage(checkResult.msg || "[存取受限] 您的 Google 帳號尚未授權。", true);
        return;
      }

      const userEmail = (profile.email || "").trim().toLowerCase();
      const isAdmin = SUPER_ADMIN_EMAILS.includes(userEmail) || checkResult.is_admin === true;
      if (isAdmin) {
        localStorage.setItem("is_admin_user", "true");
      } else {
        localStorage.removeItem("is_admin_user");
      }

      const resolvedSales = resolveSalespersonName(checkResult.name || profile.name || "曾維崧");
      localStorage.setItem("has_logged_in", "true");
      localStorage.setItem("saved_display_name", resolvedSales);
      localStorage.setItem("saved_user_profile", JSON.stringify(profile));
      localStorage.setItem("last_active_user_email", userEmail);

      const loginOverlay = document.getElementById("loginOverlaySection");
      if (loginOverlay) loginOverlay.classList.add("hidden");
      setLoginStatusMessage("");

      setupAdminImpersonator();
      switchSimulatedSalesperson(resolvedSales, false);
      showAuthToast(`[驗證成功] 歡迎 ${resolvedSales}（${profile.email}）`);
    } catch (err) {
      console.error("[Auth] Google 驗證程序異常:", err);
      setLoginStatusMessage("驗證過程發生錯誤：" + (err.message || err), true);
    }
  }

  function initGoogleIdentityServices() {
    if (typeof google === "undefined" || !google.accounts || !google.accounts.oauth2) {
      return false;
    }
    googleTokenClient = google.accounts.oauth2.initTokenClient({
      client_id: GOOGLE_WEB_CLIENT_ID,
      scope: GOOGLE_AUTH_SCOPES,
      callback: handleGoogleTokenResponse
    });
    return true;
  }

  let gisPollCount = 0;
  function ensureGisClientReady() {
    if (initGoogleIdentityServices()) return;
    if (gisPollCount < 40) {
      gisPollCount++;
      setTimeout(ensureGisClientReady, 250);
    }
  }
  ensureGisClientReady();

  // 開啟登入驗證畫面 (供初次進入或點擊「切換帳號」時呼叫)
  function openLoginOverlay() {
    const loginOverlay = document.getElementById("loginOverlaySection");
    if (loginOverlay) {
      setLoginStatusMessage("");
      loginOverlay.classList.remove("hidden");
    }
  }

  function setupAuthAndImpersonationEvents() {
    const urlParams = new URLSearchParams(window.location.search);
    if (urlParams.get("logout") === "true") {
      localStorage.removeItem("has_logged_in");
      localStorage.removeItem("saved_display_name");
      localStorage.removeItem("saved_user_profile");
      localStorage.removeItem("google_access_token");
      localStorage.removeItem("google_token_expires_at");
      localStorage.removeItem("is_admin_user");
    }

    const btnGoogleLogin = document.getElementById("btnGoogleLogin");
    if (btnGoogleLogin) {
      btnGoogleLogin.addEventListener("click", () => {
        if (!googleTokenClient && !initGoogleIdentityServices()) {
          setLoginStatusMessage("Google 驗證元件載入中，若網路受限可點擊下方「管理員測試權限直接進入」。", true);
          return;
        }
        try {
          setLoginStatusMessage("正在開啟 Google 帳號授權視窗...", false);
          googleTokenClient.requestAccessToken({ prompt: "select_account" });
        } catch (err) {
          setLoginStatusMessage("啟動 Google 登入視窗失敗：" + (err.message || err), true);
        }
      });
    }

    const btnTestEnvQuickLogin = document.getElementById("btnTestEnvQuickLogin");
    if (btnTestEnvQuickLogin) {
      btnTestEnvQuickLogin.addEventListener("click", () => {
        const defaultAdminProfile = {
          name: "曾維崧",
          email: "tsengweisung@gmail.com",
          mode: "admin_test"
        };
        localStorage.setItem("has_logged_in", "true");
        localStorage.setItem("is_admin_user", "true");
        localStorage.setItem("saved_display_name", "曾維崧");
        localStorage.setItem("saved_user_profile", JSON.stringify(defaultAdminProfile));

        const loginOverlay = document.getElementById("loginOverlaySection");
        if (loginOverlay) loginOverlay.classList.add("hidden");
        setLoginStatusMessage("");

        const targetSales = localStorage.getItem("shihlin_pwa_active_sales") || "曾維崧";
        setupAdminImpersonator();
        switchSimulatedSalesperson(targetSales, false);
        showAuthToast("[測試權限已啟用] 可由頂部「管理員身分模擬」任意切換全體業務身分");
      });
    }

    const adminSelect = document.getElementById("adminImpersonateSelect");
    if (adminSelect) {
      adminSelect.addEventListener("change", () => {
        switchSimulatedSalesperson(adminSelect.value, true);
      });
    }

    const btnOpenPermModal = document.getElementById("btnOpenPermModal");
    if (btnOpenPermModal) {
      btnOpenPermModal.addEventListener("click", () => {
        switchPage(navItemSettings, panelSettings);
        setTimeout(() => {
          const salesOrderTable = document.getElementById("tableSalesOrder") || document.getElementById("containerSalesOrderMobile");
          if (salesOrderTable) {
            salesOrderTable.scrollIntoView({ behavior: "smooth", block: "center" });
          }
        }, 120);
      });
    }

    const btnSwitchGoogleAccount = document.getElementById("btnSwitchGoogleAccount");
    if (btnSwitchGoogleAccount) {
      btnSwitchGoogleAccount.addEventListener("click", () => {
        openLoginOverlay();
      });
    }

    const headerUserBadge = document.getElementById("headerUserBadge");
    if (headerUserBadge) {
      headerUserBadge.addEventListener("click", () => {
        openLoginOverlay();
      });
    }

    // 判斷是否已經完成登入驗證：若尚未登入，立即顯示 Google 身分驗證全螢幕遮罩
    const hasLoggedIn = localStorage.getItem("has_logged_in") === "true";
    if (!hasLoggedIn) {
      openLoginOverlay();
    } else {
      setupAdminImpersonator();
    }
  }

  setupAuthAndImpersonationEvents();

  // ============================================================================
  // 報價單自適應模式 (1280px 分界線)
  // 視窗 >= 1280px：電腦版 9 大功能頁 + 左側報價單側欄 (版面由 CSS 控制)
  // 視窗 <  1280px：網頁版報價單，只保留報價擬稿與離開按鈕，其餘頁面全部隱藏
  // ============================================================================
  // 目前是否正在報價單頁面內
  function isQuotationPanelActive() {
    return typeof panelQuotes !== "undefined" && panelQuotes && !panelQuotes.classList.contains("hidden");
  }

  // 若在報價單內且視窗偏窄，強制回到「新建報價單」頁，避免停留在已被隱藏的頁面造成空白畫面
  function forceQuoteWebModeIfNarrow() {
    if (!isQuotationPanelActive()) return;
    if (window.matchMedia("(max-width: 1279px)").matches && currentActiveQuoteTab !== "create") {
      switchQuoteSubTab("create");
    }
  }

  // 使用者拉動視窗大小時，即時判斷是否需要切換
  window.addEventListener("resize", forceQuoteWebModeIfNarrow);

});

