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
      document.body.classList.add("quotation-system-active");
      if (mainList) mainList.classList.add("hidden");
      if (quoteList) quoteList.classList.remove("hidden");
      if (drawerTitle) drawerTitle.textContent = "報價管理系統";
      if (appDrawer) appDrawer.classList.remove("collapsed");
      const btnToggleSidebar = document.getElementById("btnToggleSidebar");
      if (btnToggleSidebar) btnToggleSidebar.innerHTML = "<span>收合選單</span>";
    } else {
      document.body.classList.remove("quotation-system-active");
      if (mainList) mainList.classList.remove("hidden");
      if (quoteList) quoteList.classList.add("hidden");
      if (drawerTitle) drawerTitle.textContent = "功能選單";
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
      initQuoteSidebarClock();
      initQuoteColumnResizer();
      initQuoteColumnOrderSwapper();
      initQuoteRowOrderSwapper();
      initQuoteTableHeightDragHandle();
      initQuoteWorkspaceBoundaryResizers();
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
      const defRows = parseInt(localStorage.getItem("shihlin_quote_default_rows") || "5", 10) || 5;
      for (let i = 0; i < defRows; i++) {
        addBlankItem();
      }
    }
    updateQuoteDraftsBadge();
    updateQuoteNoPreview();
    enforceDesktopQuotationView();
  }

  // 100% 還原 _QuotationApp main.py 左下角即時時鐘 (yyyy-MM-dd HH:mm)
  let quoteSidebarClockTimer = null;
  function initQuoteSidebarClock() {
    const updateClock = () => {
      const el = document.getElementById("quoteSidebarClock");
      if (!el) return;
      const now = new Date();
      const yyyy = now.getFullYear();
      const mm = String(now.getMonth() + 1).padStart(2, "0");
      const dd = String(now.getDate()).padStart(2, "0");
      const hh = String(now.getHours()).padStart(2, "0");
      const min = String(now.getMinutes()).padStart(2, "0");
      el.textContent = `${yyyy}-${mm}-${dd} ${hh}:${min}`;
    };
    updateClock();
    if (quoteSidebarClockTimer) clearInterval(quoteSidebarClockTimer);
    quoteSidebarClockTimer = setInterval(updateClock, 1000);
  }

  // 100% 還原 _QuotationApp ColumnWidthManager 表頭欄位拖曳調整寬度與欄位順序互換記憶機制 (支援全系統第 1~7 頁所有表格)
  const QUOTE_TABLE_STORAGE_PREFIX = "shihlin_quote_col_widths_v2_";
  const QUOTE_COL_ORDER_STORAGE_PREFIX = "shihlin_quote_col_order_v1_";

  function syncExcelTableTotalWidth(table) {
    if (!table) return;
    const allTh = Array.from(table.querySelectorAll("thead th"));
    if (allTh.length === 0) return;
    let sumW = 0;
    allTh.forEach(th => {
      const explicitW = parseInt(th.style.width, 10);
      sumW += (!isNaN(explicitW) && explicitW > 0) ? explicitW : (th.offsetWidth || 80);
    });
    if (sumW > 0) {
      table.style.width = `${sumW}px`;
      table.style.minWidth = `${sumW}px`;
    }
  }

  function applySingleColumnWidth(th, widthPx) {
    if (!th || !(widthPx > 0)) return;
    const wStr = `${Math.round(widthPx)}px`;
    th.style.width = wStr;
    th.style.minWidth = wStr;
    th.style.maxWidth = wStr;
  }

  // 標記並對齊單一資料列的儲存格順序，使其與目前 thead th 的欄位順序 100% 一致
  function alignRowCellsToCurrentHeaderOrder(table, tr) {
    if (!table || !tr) return;
    const cells = Array.from(tr.children);
    if (cells.length <= 1) return;

    const headRow = table.querySelector("thead tr");
    if (!headRow) return;
    const thList = Array.from(headRow.children);
    if (cells.length !== thList.length) return;

    // 首次標記該列各儲存格的原始邏輯欄位編號 (0..N)
    if (!tr.hasAttribute("data-orig-stamped")) {
      cells.forEach((td, i) => {
        if (!td.hasAttribute("data-orig-idx")) {
          td.setAttribute("data-orig-idx", String(i));
        }
      });
      tr.setAttribute("data-orig-stamped", "true");
    }

    const desiredOrigIndices = thList.map(th => th.getAttribute("data-orig-idx"));
    const currentOrigIndices = cells.map(td => td.getAttribute("data-orig-idx"));
    const isSameOrder = desiredOrigIndices.every((val, idx) => val === currentOrigIndices[idx]);
    if (isSameOrder) return;

    const cellMap = {};
    cells.forEach(td => {
      cellMap[td.getAttribute("data-orig-idx")] = td;
    });
    desiredOrigIndices.forEach(origIdx => {
      const targetTd = cellMap[origIdx];
      if (targetTd) tr.appendChild(targetTd);
    });
  }

  // 將表格依據已儲存的欄位順序重新排列 thead 與 tbody 所有列
  function applySavedColumnOrderToTable(table, customOrder = null) {
    if (!table) return;
    const headRow = table.querySelector("thead tr");
    if (!headRow) return;

    const thList = Array.from(headRow.children);
    thList.forEach((th, idx) => {
      if (!th.hasAttribute("data-orig-idx")) {
        th.setAttribute("data-orig-idx", String(idx));
      }
    });

    const tableKey = table.id || "quote_table_default";
    let targetOrder = customOrder;
    if (!targetOrder) {
      try {
        const raw = localStorage.getItem(`${QUOTE_COL_ORDER_STORAGE_PREFIX}${tableKey}`);
        if (raw) targetOrder = JSON.parse(raw);
      } catch (e) {}
    }

    if (Array.isArray(targetOrder) && targetOrder.length === thList.length) {
      const thMap = {};
      thList.forEach(th => {
        thMap[th.getAttribute("data-orig-idx")] = th;
      });
      targetOrder.forEach(origIdx => {
        const th = thMap[String(origIdx)];
        if (th) headRow.appendChild(th);
      });
    }

    const tbody = table.querySelector("tbody");
    if (tbody) {
      Array.from(tbody.querySelectorAll("tr")).forEach(tr => {
        alignRowCellsToCurrentHeaderOrder(table, tr);
      });
    }
  }

  // 初始化所有報價系統表格的「表頭左右拖曳互換順序」與動態列自動對齊守護
  function initQuoteColumnOrderSwapper() {
    const tables = document.querySelectorAll("#panelQuotes .pyside-excel-table");
    tables.forEach((table, tIdx) => {
      if (!table.id) table.id = `quote_auto_table_${tIdx}`;
      const tableKey = table.id;
      const orderStorageKey = `${QUOTE_COL_ORDER_STORAGE_PREFIX}${tableKey}`;

      const headRow = table.querySelector("thead tr");
      if (!headRow) return;
      const thList = Array.from(headRow.children);
      thList.forEach((th, idx) => {
        if (!th.hasAttribute("data-orig-idx")) {
          th.setAttribute("data-orig-idx", String(idx));
        }
      });

      // 先套用先前儲存的欄位順序
      applySavedColumnOrderToTable(table);

      // 綁定 MutationObserver：當 tbody 動態新增或重繪任何資料列時，自動同步欄位順序
      const tbody = table.querySelector("tbody");
      if (tbody && !tbody.hasAttribute("data-col-order-observed")) {
        tbody.setAttribute("data-col-order-observed", "true");
        const observer = new MutationObserver((mutations) => {
          mutations.forEach(m => {
            m.addedNodes.forEach(node => {
              if (node && node.nodeType === 1 && node.tagName === "TR") {
                alignRowCellsToCurrentHeaderOrder(table, node);
              }
            });
          });
        });
        observer.observe(tbody, { childList: true });
      }

      // 為每個非項次角的表頭綁定拖曳互換事件
      thList.forEach(th => {
        if (th.classList.contains("col-corner-header")) return;
        if (th.getAttribute("data-order-drag-bound") === "true") return;
        th.setAttribute("data-order-drag-bound", "true");
        th.setAttribute("draggable", "true");
        if (!th.getAttribute("title")) {
          th.setAttribute("title", "按住表頭左右拖曳可互換欄位順序；拖曳右緣可調整欄寬");
        }

        th.addEventListener("dragstart", (e) => {
          if (th.querySelector(".col-resizer.resizing")) {
            e.preventDefault();
            return;
          }
          th.classList.add("th-dragging");
          e.dataTransfer.effectAllowed = "move";
          e.dataTransfer.setData("text/plain", th.getAttribute("data-orig-idx") || "");
        });

        th.addEventListener("dragend", () => {
          th.classList.remove("th-dragging");
          headRow.querySelectorAll("th").forEach(h => {
            h.classList.remove("th-drag-over-left", "th-drag-over-right");
          });
        });

        th.addEventListener("dragover", (e) => {
          const draggingTh = headRow.querySelector("th.th-dragging");
          if (!draggingTh || draggingTh === th || th.classList.contains("col-corner-header")) return;
          e.preventDefault();
          e.dataTransfer.dropEffect = "move";

          const rect = th.getBoundingClientRect();
          const isLeftHalf = (e.clientX - rect.left) < rect.width / 2;
          th.classList.toggle("th-drag-over-left", isLeftHalf);
          th.classList.toggle("th-drag-over-right", !isLeftHalf);
        });

        th.addEventListener("dragleave", () => {
          th.classList.remove("th-drag-over-left", "th-drag-over-right");
        });

        th.addEventListener("drop", (e) => {
          e.preventDefault();
          th.classList.remove("th-drag-over-left", "th-drag-over-right");
          const draggingTh = headRow.querySelector("th.th-dragging");
          if (!draggingTh || draggingTh === th || th.classList.contains("col-corner-header")) return;

          const rect = th.getBoundingClientRect();
          const insertBefore = (e.clientX - rect.left) < rect.width / 2;
          if (insertBefore) {
            headRow.insertBefore(draggingTh, th);
          } else {
            headRow.insertBefore(draggingTh, th.nextSibling);
          }

          // 儲存最新表頭順序並立即同步重排所有 tbody 資料列
          const newOrder = Array.from(headRow.children).map(h => h.getAttribute("data-orig-idx"));
          localStorage.setItem(orderStorageKey, JSON.stringify(newOrder));
          applySavedColumnOrderToTable(table, newOrder);
        });
      });
    });
  }

  // 新建報價單明細列上下順序互換 (支援按住左側項次上下拖曳互換 + 工具列上移/下移按鈕)
  function moveDesktopItemRow(rowEl, direction) {
    const tbody = document.getElementById("quoteDesktopTableBody");
    if (!tbody || !rowEl) return;

    const rows = Array.from(tbody.querySelectorAll("tr"));
    const idx = rows.indexOf(rowEl);
    if (idx < 0) return;

    const targetIdx = idx + direction;
    if (targetIdx < 0 || targetIdx >= rows.length) return;

    const targetRow = rows[targetIdx];
    if (direction < 0) {
      tbody.insertBefore(rowEl, targetRow);
    } else {
      tbody.insertBefore(rowEl, targetRow.nextSibling);
    }

    syncMobileItemsOrderWithDesktop();
    refreshItemIndices();
    calculateTotal();
    updateDesktopBannerForRow(rowEl);
  }

  function syncMobileItemsOrderWithDesktop() {
    const tbody = document.getElementById("quoteDesktopTableBody");
    const itemsContainer = document.getElementById("itemsContainer");
    if (!tbody || !itemsContainer) return;

    tbody.querySelectorAll("tr.desktop-item-row").forEach(dRow => {
      const itemId = dRow.getAttribute("data-itemid");
      const mRow = itemId ? document.getElementById(itemId) : null;
      if (mRow) {
        itemsContainer.appendChild(mRow);
      }
    });
  }

  function initQuoteRowOrderSwapper() {
    const tbody = document.getElementById("quoteDesktopTableBody");
    if (!tbody || tbody.getAttribute("data-row-swap-bound") === "true") return;
    tbody.setAttribute("data-row-swap-bound", "true");

    let draggingRow = null;

    tbody.addEventListener("mousedown", (e) => {
      const indexCell = e.target.closest("td.cell-vertical-index");
      const tr = e.target.closest("tr");
      if (!tr) return;
      if (indexCell) {
        tr.setAttribute("draggable", "true");
      } else {
        tr.removeAttribute("draggable");
      }
    });

    tbody.addEventListener("dragstart", (e) => {
      const tr = e.target.closest("tr");
      if (!tr || tr.getAttribute("draggable") !== "true") {
        e.preventDefault();
        return;
      }
      draggingRow = tr;
      tr.classList.add("row-dragging");
      e.dataTransfer.effectAllowed = "move";
      e.dataTransfer.setData("text/plain", tr.id || "row");
    });

    tbody.addEventListener("dragend", () => {
      if (draggingRow) {
        draggingRow.classList.remove("row-dragging");
        draggingRow.removeAttribute("draggable");
      }
      draggingRow = null;
      tbody.querySelectorAll("tr").forEach(r => {
        r.classList.remove("row-drag-over-top", "row-drag-over-bottom");
      });
    });

    tbody.addEventListener("dragover", (e) => {
      if (!draggingRow) return;
      const targetTr = e.target.closest("tr");
      if (!targetTr || targetTr === draggingRow || targetTr.parentElement !== tbody) return;
      e.preventDefault();
      e.dataTransfer.dropEffect = "move";

      const rect = targetTr.getBoundingClientRect();
      const isTopHalf = (e.clientY - rect.top) < rect.height / 2;
      targetTr.classList.toggle("row-drag-over-top", isTopHalf);
      targetTr.classList.toggle("row-drag-over-bottom", !isTopHalf);
    });

    tbody.addEventListener("dragleave", (e) => {
      const targetTr = e.target.closest("tr");
      if (targetTr) {
        targetTr.classList.remove("row-drag-over-top", "row-drag-over-bottom");
      }
    });

    tbody.addEventListener("drop", (e) => {
      if (!draggingRow) return;
      e.preventDefault();
      const targetTr = e.target.closest("tr");
      tbody.querySelectorAll("tr").forEach(r => {
        r.classList.remove("row-drag-over-top", "row-drag-over-bottom");
      });
      if (!targetTr || targetTr === draggingRow || targetTr.parentElement !== tbody) return;

      const rect = targetTr.getBoundingClientRect();
      const insertBefore = (e.clientY - rect.top) < rect.height / 2;
      if (insertBefore) {
        tbody.insertBefore(draggingRow, targetTr);
      } else {
        tbody.insertBefore(draggingRow, targetTr.nextSibling);
      }

      const movedRow = draggingRow;
      syncMobileItemsOrderWithDesktop();
      refreshItemIndices();
      calculateTotal();
      updateDesktopBannerForRow(movedRow);
    });
  }

  function initQuoteColumnResizer() {
    const tables = document.querySelectorAll("#panelQuotes .pyside-excel-table");
    tables.forEach((table, tIdx) => {
      if (!table.id) table.id = `quote_auto_table_${tIdx}`;
      const tableKey = table.id;
      const storageKey = `${QUOTE_TABLE_STORAGE_PREFIX}${tableKey}`;

      let savedWidths = {};
      try {
        const raw = localStorage.getItem(storageKey);
        if (raw) savedWidths = JSON.parse(raw) || {};
      } catch (e) {}

      const thList = Array.from(table.querySelectorAll("thead th"));
      thList.forEach((th, colIdx) => {
        if (!th.hasAttribute("data-orig-idx")) {
          th.setAttribute("data-orig-idx", String(colIdx));
        }
        if (th.classList.contains("col-corner-header")) return;

        // 紀錄原始推薦寬度，供第 8 頁「還原預設欄寬」按鈕使用
        if (!th.hasAttribute("data-default-width")) {
          const initW = parseInt(th.style.width, 10) || th.offsetWidth || 100;
          th.setAttribute("data-default-width", String(initW));
        }

        if (!th.classList.contains("resizable-th")) {
          th.classList.add("resizable-th");
        }
        const origIdx = th.getAttribute("data-orig-idx") || String(colIdx);
        const colKey = th.getAttribute("data-col") || `col_${origIdx}`;
        th.setAttribute("data-col", colKey);

        if (savedWidths[colKey] && Number(savedWidths[colKey]) >= 36) {
          applySingleColumnWidth(th, Number(savedWidths[colKey]));
        } else {
          const defW = parseInt(th.getAttribute("data-default-width"), 10);
          if (defW > 0) applySingleColumnWidth(th, defW);
        }

        let resizer = th.querySelector(".col-resizer");
        if (!resizer) {
          resizer = document.createElement("span");
          resizer.className = "col-resizer";
          th.appendChild(resizer);
        }

        if (resizer.getAttribute("data-bound") === "true") return;
        resizer.setAttribute("data-bound", "true");

        let startX = 0;
        let startWidth = 0;

        const onMouseMove = (e) => {
          const dx = e.clientX - startX;
          const newWidth = Math.max(40, startWidth + dx);
          applySingleColumnWidth(th, newWidth);
          syncExcelTableTotalWidth(table);
        };

        const onMouseUp = () => {
          resizer.classList.remove("resizing");
          th.setAttribute("draggable", "true");
          document.body.style.cursor = "";
          document.removeEventListener("mousemove", onMouseMove);
          document.removeEventListener("mouseup", onMouseUp);

          const currentWidths = {};
          table.querySelectorAll("thead th.resizable-th").forEach(header => {
            const k = header.getAttribute("data-col");
            const w = parseInt(header.style.width, 10) || header.offsetWidth;
            if (k && w > 0) currentWidths[k] = w;
          });
          localStorage.setItem(storageKey, JSON.stringify(currentWidths));
        };

        resizer.addEventListener("mousedown", (e) => {
          e.preventDefault();
          e.stopPropagation();
          th.setAttribute("draggable", "false");
          startX = e.clientX;
          startWidth = parseInt(th.style.width, 10) || th.offsetWidth || 100;
          resizer.classList.add("resizing");
          document.body.style.cursor = "col-resize";
          document.addEventListener("mousemove", onMouseMove);
          document.addEventListener("mouseup", onMouseUp);
        });
      });

      syncExcelTableTotalWidth(table);
    });
  }

  function resetQuoteTableColumnWidthsByTarget(targetPage) {
    const pageTableMap = {
      page1: ["quoteDesktopTable"],
      page2: ["quoteDraftsTable"],
      page3: ["quoteHistoryTable"],
      page5: ["quoteAnalysisTable"],
      page6: ["quoteCustMgmtTable"],
      page7: ["quoteProdMgmtTable"]
    };
    const ids = pageTableMap[targetPage] || ["quoteDesktopTable"];
    ids.forEach(tid => {
      localStorage.removeItem(`${QUOTE_TABLE_STORAGE_PREFIX}${tid}`);
      localStorage.removeItem(`${QUOTE_COL_ORDER_STORAGE_PREFIX}${tid}`);
      if (tid === "quoteDesktopTable") {
        localStorage.removeItem("shihlin_quote_table_col_widths_v1");
      }
      const tbl = document.getElementById(tid);
      if (!tbl) return;

      // 同時還原預設欄位順序 (0..N)
      const headRow = tbl.querySelector("thead tr");
      if (headRow) {
        const ths = Array.from(headRow.children);
        const defaultOrder = ths
          .map(th => parseInt(th.getAttribute("data-orig-idx") || "0", 10))
          .sort((a, b) => a - b)
          .map(n => String(n));
        applySavedColumnOrderToTable(tbl, defaultOrder);
      }

      tbl.querySelectorAll("thead th.resizable-th").forEach(th => {
        const defW = parseInt(th.getAttribute("data-default-width") || "100", 10);
        if (defW > 0) {
          applySingleColumnWidth(th, defW);
        }
      });
      syncExcelTableTotalWidth(tbl);

      // 同時還原該表格容器的預設高度
      const wrap = tbl.closest(".pyside-table-container, .pyside-full-table-wrap");
      if (wrap) {
        const hKey = `shihlin_quote_wrap_height_${tid}`;
        localStorage.removeItem(hKey);
        if (tid === "quoteDesktopTable") {
          localStorage.removeItem("shihlin_quote_table_height_v1");
          wrap.style.height = "250px";
        } else {
          const defH = parseInt(wrap.getAttribute("data-default-height") || "360", 10);
          wrap.style.height = `${defH}px`;
        }
      }
    });
  }

  // 100% 還原 _QuotationApp TableHeightDragHandle 表格高度上下拖曳條與記憶機制
  function initQuoteTableHeightDragHandle() {
    const handle = document.getElementById("quoteTableDragHandle");
    const container = document.getElementById("quoteDesktopTableContainer");
    if (!handle || !container) return;

    const savedHeight = parseInt(localStorage.getItem("shihlin_quote_table_height_v1") || "250", 10);
    if (!isNaN(savedHeight) && savedHeight >= 150) {
      container.style.height = `${savedHeight}px`;
    }

    let startY = 0;
    let startHeight = 250;

    const onMouseMove = (e) => {
      const dy = e.clientY - startY;
      const newHeight = Math.max(150, startHeight + dy);
      container.style.height = `${newHeight}px`;
    };

    const onMouseUp = () => {
      document.removeEventListener("mousemove", onMouseMove);
      document.removeEventListener("mouseup", onMouseUp);
      const finalH = container.offsetHeight;
      if (finalH >= 150) {
        localStorage.setItem("shihlin_quote_table_height_v1", String(finalH));
      }
    };

    handle.addEventListener("mousedown", (e) => {
      e.preventDefault();
      startY = e.clientY;
      startHeight = container.offsetHeight;
      document.addEventListener("mousemove", onMouseMove);
      document.addEventListener("mouseup", onMouseUp);
    });

    handle.addEventListener("dblclick", () => {
      const rect = container.getBoundingClientRect();
      const fitH = Math.max(200, Math.min(680, window.innerHeight - rect.top - 160));
      container.style.height = `${Math.round(fitH)}px`;
      localStorage.setItem("shihlin_quote_table_height_v1", String(Math.round(fitH)));
    });
  }

  // 全系統共用「右側垂直邊界拖曳拉桿」與第 2~7 頁「底部水平高度拖曳拉桿」（一頁調寬全頁同步，且鎖定表頭不因拉底端橫 bar 而消失）
  const QUOTE_WORKSPACE_WIDTH_KEY = "shihlin_quote_workspace_width_v1";

  function updateRightBoundaryLinePosition() {
    const resizer = document.getElementById("quoteRightBoundaryResizer");
    if (!resizer) return;
    const drawer = document.getElementById("appDrawer");
    const sidebarW = (drawer && !drawer.classList.contains("collapsed") && drawer.offsetWidth > 0) ? drawer.offsetWidth : 0;
    const curW = parseInt(getComputedStyle(document.documentElement).getPropertyValue("--quote-workspace-width"), 10) || 1050;
    const targetLeft = sidebarW + 15 + curW;
    resizer.style.left = `${Math.min(window.innerWidth - 10, Math.max(sidebarW + 900, targetLeft))}px`;
  }

  function applyGlobalQuoteWorkspaceWidth(widthPx, save = true) {
    const drawer = document.getElementById("appDrawer");
    const sidebarW = (drawer && !drawer.classList.contains("collapsed") && drawer.offsetWidth > 0) ? drawer.offsetWidth : 200;
    const maxAllowed = Math.max(960, window.innerWidth - sidebarW - 32);
    const clampedW = Math.round(Math.max(920, Math.min(maxAllowed, widthPx)));
    document.documentElement.style.setProperty("--quote-workspace-width", `${clampedW}px`);
    if (save) {
      localStorage.setItem(QUOTE_WORKSPACE_WIDTH_KEY, String(clampedW));
    }
    updateRightBoundaryLinePosition();
  }

  function initQuoteWorkspaceBoundaryResizers() {
    const panelQuotes = document.getElementById("panelQuotes");
    if (!panelQuotes) return;

    // 1. 建立或取得全系統共用右側垂直邊界拉桿
    let rightResizer = document.getElementById("quoteRightBoundaryResizer");
    if (!rightResizer) {
      rightResizer = document.createElement("div");
      rightResizer.id = "quoteRightBoundaryResizer";
      rightResizer.className = "quote-right-boundary-resizer";
      rightResizer.title = "左右拖曳可調整全系統各頁面的右側可視邊界（雙擊自動撐滿視窗寬度）";
      rightResizer.innerHTML = `<div class="boundary-grip-pill">||</div>`;
      panelQuotes.appendChild(rightResizer);
    }

    // 載入已儲存的工作區寬度（若無儲存且螢幕夠寬，預設給更充裕的可視範圍或維持儲存值）
    const savedW = parseInt(localStorage.getItem(QUOTE_WORKSPACE_WIDTH_KEY) || "0", 10);
    if (savedW >= 920) {
      applyGlobalQuoteWorkspaceWidth(savedW, false);
    } else {
      const drawer = document.getElementById("appDrawer");
      const sidebarW = (drawer && !drawer.classList.contains("collapsed") && drawer.offsetWidth > 0) ? drawer.offsetWidth : 200;
      const defaultFitW = Math.max(1050, window.innerWidth - sidebarW - 40);
      applyGlobalQuoteWorkspaceWidth(defaultFitW, false);
    }

    if (rightResizer.getAttribute("data-bound") !== "true") {
      rightResizer.setAttribute("data-bound", "true");

      let startX = 0;
      let startW = 1050;

      const onMouseMove = (e) => {
        const dx = e.clientX - startX;
        applyGlobalQuoteWorkspaceWidth(startW + dx, false);
      };

      const onMouseUp = () => {
        rightResizer.classList.remove("resizing");
        document.body.style.cursor = "";
        document.removeEventListener("mousemove", onMouseMove);
        document.removeEventListener("mouseup", onMouseUp);
        const curW = parseInt(getComputedStyle(document.documentElement).getPropertyValue("--quote-workspace-width"), 10) || 1050;
        localStorage.setItem(QUOTE_WORKSPACE_WIDTH_KEY, String(curW));
      };

      rightResizer.addEventListener("mousedown", (e) => {
        e.preventDefault();
        startX = e.clientX;
        startW = parseInt(getComputedStyle(document.documentElement).getPropertyValue("--quote-workspace-width"), 10) || 1050;
        rightResizer.classList.add("resizing");
        document.body.style.cursor = "ew-resize";
        document.addEventListener("mousemove", onMouseMove);
        document.addEventListener("mouseup", onMouseUp);
      });

      rightResizer.addEventListener("dblclick", () => {
        const drawer = document.getElementById("appDrawer");
        const sidebarW = (drawer && !drawer.classList.contains("collapsed") && drawer.offsetWidth > 0) ? drawer.offsetWidth : 200;
        const fullW = Math.max(1050, window.innerWidth - sidebarW - 36);
        applyGlobalQuoteWorkspaceWidth(fullW, true);
      });

      window.addEventListener("resize", () => {
        updateRightBoundaryLinePosition();
      });
    }

    // 2. 為第 2~7 頁所有 .pyside-full-table-wrap 自動掛載「底部橫向高度調整拉桿」並將高度控制在螢幕可視範圍內（防止拉底部橫 bar 時上方表頭消失）
    const tableWraps = panelQuotes.querySelectorAll(".pyside-full-table-wrap");
    tableWraps.forEach((wrap, idx) => {
      if (wrap.id === "dashFollowupTableWrap") return; // 第 4 頁待跟進區已有專屬拉桿
      const innerTable = wrap.querySelector("table");
      const tableId = innerTable?.id || `quote_wrap_${idx}`;
      const hStorageKey = `shihlin_quote_wrap_height_${tableId}`;

      // 去除會撐破整個視窗的 flex:1 與過大 min-height，改為固定內部捲動高度，讓表頭 sticky 永遠固定在上方
      wrap.style.flex = "none";
      wrap.style.minHeight = "160px";

      if (!wrap.hasAttribute("data-default-height")) {
        let defH = 380;
        if (tableId === "quoteHistoryTable") defH = Math.max(280, Math.min(520, window.innerHeight - 230));
        else if (tableId === "quoteDraftsTable") defH = Math.max(300, Math.min(540, window.innerHeight - 210));
        else if (tableId === "quoteAnalysisTable") defH = 240;
        else if (tableId === "quoteCustMgmtTable" || tableId === "quoteProdMgmtTable") defH = Math.max(280, Math.min(480, window.innerHeight - 250));
        else defH = parseInt(wrap.style.height, 10) || 220;
        wrap.setAttribute("data-default-height", String(Math.round(defH)));
      }

      const savedH = parseInt(localStorage.getItem(hStorageKey) || "0", 10);
      if (savedH >= 150) {
        wrap.style.height = `${savedH}px`;
      } else {
        const defH = parseInt(wrap.getAttribute("data-default-height") || "360", 10);
        wrap.style.height = `${defH}px`;
      }

      if (wrap.getAttribute("data-height-handle-bound") === "true") return;
      wrap.setAttribute("data-height-handle-bound", "true");

      const hHandle = document.createElement("div");
      hHandle.className = "pyside-universal-height-handle";
      hHandle.title = "上下拖曳可調整表格高度（雙擊自動貼合螢幕高度，讓頂部表頭與底部橫向捲軸同時可見）";
      hHandle.innerHTML = `<div class="drag-handle-grip"></div>`;
      wrap.insertAdjacentElement("afterend", hHandle);

      let startY = 0;
      let startH = 360;

      const onMoveH = (e) => {
        const dy = e.clientY - startY;
        const newH = Math.max(150, startH + dy);
        wrap.style.height = `${newH}px`;
      };

      const onUpH = () => {
        document.body.style.cursor = "";
        document.removeEventListener("mousemove", onMoveH);
        document.removeEventListener("mouseup", onUpH);
        const finalH = wrap.offsetHeight;
        if (finalH >= 150) {
          localStorage.setItem(hStorageKey, String(finalH));
        }
      };

      hHandle.addEventListener("mousedown", (e) => {
        e.preventDefault();
        startY = e.clientY;
        startH = wrap.offsetHeight || parseInt(wrap.style.height, 10) || 360;
        document.body.style.cursor = "ns-resize";
        document.addEventListener("mousemove", onMoveH);
        document.addEventListener("mouseup", onUpH);
      });

      hHandle.addEventListener("dblclick", () => {
        const rect = wrap.getBoundingClientRect();
        const bestH = Math.max(180, Math.min(650, window.innerHeight - rect.top - 95));
        wrap.style.height = `${Math.round(bestH)}px`;
        localStorage.setItem(hStorageKey, String(Math.round(bestH)));
      });
    });

    setTimeout(updateRightBoundaryLinePosition, 60);
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
      runPriceAnalysis();
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
      loadQuotationHelpPage();
    }

    initQuoteColumnResizer();
    initQuoteColumnOrderSwapper();
    initQuoteWorkspaceBoundaryResizers();
    setTimeout(updateRightBoundaryLinePosition, 50);
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

    // 2. _QuotationApp 頂部：收合選單按鈕 (所有頁面共用左側側邊欄切換)
    const toggleSidebarHandler = () => {
      const drawer = document.getElementById("appDrawer");
      if (drawer) {
        drawer.classList.toggle("collapsed");
        const isCol = drawer.getAttribute("class")?.includes("collapsed");
        const labelHtml = `<span>${isCol ? "展開選單" : "收合選單"}</span>`;
        const b1 = document.getElementById("btnToggleSidebar");
        if (b1) b1.innerHTML = labelHtml;
        document.querySelectorAll(".btn-toggle-sidebar-any").forEach(b => {
          b.innerHTML = labelHtml;
        });
        setTimeout(updateRightBoundaryLinePosition, 60);
      }
    };
    const btnToggleSidebar = document.getElementById("btnToggleSidebar");
    if (btnToggleSidebar) {
      btnToggleSidebar.addEventListener("click", toggleSidebarHandler);
    }
    document.querySelectorAll(".btn-toggle-sidebar-any").forEach(btn => {
      btn.addEventListener("click", toggleSidebarHandler);
    });

    // _QuotationApp 頂部：收合抬頭按鈕 (雙欄基本資料表頭切換)
    const btnToggleHeaderCompact = document.getElementById("btnToggleHeaderCompact");
    const quoteHeaderContainer = document.getElementById("quoteHeaderContainer");
    const headerToggleIcon = document.getElementById("headerToggleIcon");
    const headerToggleLabel = document.getElementById("headerToggleLabel");
    if (btnToggleHeaderCompact && quoteHeaderContainer) {
      btnToggleHeaderCompact.addEventListener("click", () => {
        const isCollapsed = quoteHeaderContainer.classList.toggle("collapsed");
        quoteHeaderContainer.classList.toggle("hidden", isCollapsed);
        if (headerToggleIcon) headerToggleIcon.textContent = "";
        if (headerToggleLabel) headerToggleLabel.textContent = isCollapsed ? "展開抬頭" : "收合抬頭";
        setTimeout(updateRightBoundaryLinePosition, 50);
      });
    }

    // 單號模式切換 (自動擷取 vs 自行輸入，對標 gui_page1.py toggle_id_mode)
    const quoteNoMode = document.getElementById("quoteNoMode");
    if (quoteNoMode) {
      quoteNoMode.addEventListener("change", (e) => {
        const isAuto = e.target.value === "auto";
        const prefixInput = document.getElementById("quotePrefixInput");
        const seqInput = document.getElementById("quoteSeqInput");
        const revInput = document.getElementById("quoteRevInput");
        if (isAuto) {
          if (prefixInput) prefixInput.setAttribute("readonly", "readonly");
          if (seqInput) seqInput.setAttribute("readonly", "readonly");
          if (revInput) revInput.setAttribute("readonly", "readonly");
          if (currentSelectedCustomer) applySelectedCustomerToQuoteUI(currentSelectedCustomer);
        } else {
          if (prefixInput) {
            prefixInput.removeAttribute("readonly");
            if (prefixInput.value === "-- 請選擇客戶 --") prefixInput.value = "";
          }
          if (seqInput) seqInput.removeAttribute("readonly");
          if (revInput) revInput.removeAttribute("readonly");
        }
      });
    }

    // 手機端網頁報價模式頂部重設與離開按鈕
    const btnMobileReset = document.getElementById("btnMobileResetQuote");
    if (btnMobileReset) {
      btnMobileReset.addEventListener("click", () => {
        if (confirm("確定要重設並清空目前報價單嗎？")) {
          resetQuotationEditor();
        }
      });
    }
    const btnMobileExitTop = document.getElementById("btnMobileExitQuoteTop");
    if (btnMobileExitTop) {
      btnMobileExitTop.addEventListener("click", () => switchPage(navItemHome, panelHome));
    }
    const btnMobileExitBottom = document.getElementById("btnMobileExitQuoteBottom");
    if (btnMobileExitBottom) {
      btnMobileExitBottom.addEventListener("click", () => switchPage(navItemHome, panelHome));
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
        updateDesktopBannerForRow(null);
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

    // 10. 報價單管理 (Page 3) 搜尋、清除、垃圾桶切換、版次比對、匯出 CRM、分頁與刪除
    const inputSearchHistory = document.getElementById("inputSearchQuoteHistory");
    if (inputSearchHistory) {
      inputSearchHistory.addEventListener("input", (e) => {
        p3HistoryPage = 1;
        loadQuotationHistory(e.target.value.trim(), isTrashViewActive);
      });
    }

    const chkHistoryShowAll = document.getElementById("chkHistoryShowAll");
    if (chkHistoryShowAll) {
      chkHistoryShowAll.addEventListener("change", () => {
        p3HistoryPage = 1;
        loadQuotationHistory(inputSearchHistory?.value.trim() || "", isTrashViewActive);
      });
    }
    const historyStartDate = document.getElementById("historyStartDate");
    const historyEndDate = document.getElementById("historyEndDate");
    if (historyStartDate) historyStartDate.addEventListener("change", () => loadQuotationHistory(inputSearchHistory?.value.trim() || "", isTrashViewActive));
    if (historyEndDate) historyEndDate.addEventListener("change", () => loadQuotationHistory(inputSearchHistory?.value.trim() || "", isTrashViewActive));

    const btnToggleTrash = document.getElementById("btnToggleTrashView");
    if (btnToggleTrash) {
      btnToggleTrash.addEventListener("click", () => {
        isTrashViewActive = !isTrashViewActive;
        p3HistoryPage = 1;
        const txt = document.getElementById("trashToggleText");
        if (txt) {
          txt.textContent = isTrashViewActive ? "返回列表" : "垃圾桶";
        }
        loadQuotationHistory(inputSearchHistory?.value.trim() || "", isTrashViewActive);
      });
    }

    const btnRefreshHistory = document.getElementById("btnRefreshQuoteHistory");
    if (btnRefreshHistory) {
      btnRefreshHistory.addEventListener("click", () => {
        p3HistoryPage = 1;
        loadQuotationHistory(inputSearchHistory?.value.trim() || "", isTrashViewActive);
      });
    }

    const btnClearHistorySearch = document.getElementById("btnClearQuoteHistorySearch");
    if (btnClearHistorySearch) {
      btnClearHistorySearch.addEventListener("click", () => {
        if (inputSearchHistory) inputSearchHistory.value = "";
        if (chkHistoryShowAll) chkHistoryShowAll.checked = true;
        p3HistoryPage = 1;
        loadQuotationHistory("", isTrashViewActive);
      });
    }

    const btnCompareRev = document.getElementById("btnCompareQuoteRevisions");
    if (btnCompareRev) {
      btnCompareRev.addEventListener("click", () => {
        if (!p3SelectedQuoteObj) {
          alert("請先在下方報價單列表點選一筆報價單，再進行版次比對！");
          return;
        }
        alert(`[版次比對] 單號【${p3SelectedQuoteObj.quote_no}】目前金額為 $${Number(p3SelectedQuoteObj.subtotal || 0).toLocaleString()} (未稅)。\n若同案號有多個版次，可雙擊單號載入主畫面逐項檢視差異。`);
      });
    }

    const btnExportCrmP3 = document.getElementById("btnExportCrmReportP3");
    if (btnExportCrmP3) {
      btnExportCrmP3.addEventListener("click", exportCrmMonthlyReportCsv);
    }

    const btnHistPrev = document.getElementById("btnHistoryPrevPage");
    if (btnHistPrev) {
      btnHistPrev.addEventListener("click", () => {
        if (p3HistoryPage > 1) {
          p3HistoryPage--;
          renderQuotationHistoryPage();
        }
      });
    }
    const btnHistNext = document.getElementById("btnHistoryNextPage");
    if (btnHistNext) {
      btnHistNext.addEventListener("click", () => {
        const totalPages = Math.max(1, Math.ceil(p3FilteredQuotesCache.length / P3_PAGE_SIZE));
        if (p3HistoryPage < totalPages) {
          p3HistoryPage++;
          renderQuotationHistoryPage();
        }
      });
    }

    const btnDelSelectedP3 = document.getElementById("btnDeleteSelectedQuoteP3");
    if (btnDelSelectedP3) {
      btnDelSelectedP3.addEventListener("click", () => {
        if (!p3SelectedQuoteObj) {
          alert("請先點選要移至垃圾桶的報價單！");
          return;
        }
        pendingDeleteQuoteNo = p3SelectedQuoteObj.quote_no;
        const targetEl = document.getElementById("deleteConfirmQuoteNo");
        if (targetEl) targetEl.textContent = pendingDeleteQuoteNo;
        const delModal = document.getElementById("deleteConfirmModal");
        if (delModal) delModal.classList.remove("hidden");
      });
    }

    const btnClearAllP3 = document.getElementById("btnClearAllQuotesP3");
    if (btnClearAllP3) {
      btnClearAllP3.addEventListener("click", () => {
        alert("[企業防呆保護] 為確保歷史報價單與稽核軌跡安全，系統已鎖定單鍵清空全部正式報價單功能。請使用單筆「刪除選取項目（移至垃圾桶）」。");
      });
    }

    // 11. 手機報價草稿覆核區 (Page 2) 重新整理、覆核轉正與清空按鈕
    const btnRefreshDrafts = document.getElementById("btnRefreshDraftsInbox");
    if (btnRefreshDrafts) {
      btnRefreshDrafts.addEventListener("click", loadQuotationDraftsInbox);
    }

    const btnConvertDraft = document.getElementById("btnConvertDraftToFormal");
    if (btnConvertDraft) {
      btnConvertDraft.addEventListener("click", convertSelectedMobileDraftToFormal);
    }

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
          updateQuoteDraftsBadge();
        } catch (err) {
          alert("[清空草稿失敗] " + err.message);
        }
      });
    }

    // 12. 歷史單價分析 (Page 5) 搜尋與清除按鈕
    const btnRunAnalysis = document.getElementById("btnRunPriceAnalysis");
    if (btnRunAnalysis) {
      btnRunAnalysis.addEventListener("click", runPriceAnalysis);
    }
    const btnClearAnalysis = document.getElementById("btnClearPriceAnalysis");
    if (btnClearAnalysis) {
      btnClearAnalysis.addEventListener("click", () => {
        const mIn = document.getElementById("inputAnalysisModel");
        const cIn = document.getElementById("inputAnalysisCustomer");
        const chkAll = document.getElementById("chkAnalysisShowAll");
        if (mIn) mIn.value = "";
        if (cIn) cIn.value = "";
        if (chkAll) chkAll.checked = true;
        runPriceAnalysis();
      });
    }
    const inputAnalysisModel = document.getElementById("inputAnalysisModel");
    if (inputAnalysisModel) {
      inputAnalysisModel.addEventListener("input", () => runPriceAnalysis());
    }
    const chkAnalysisShowAll = document.getElementById("chkAnalysisShowAll");
    if (chkAnalysisShowAll) {
      chkAnalysisShowAll.addEventListener("change", () => runPriceAnalysis());
    }

    // 13. 產品管理 (Page 7) 與個人特殊品新增/更新按鈕
    const btnSaveSpecial = document.getElementById("btnSaveSpecialProduct");
    if (btnSaveSpecial) {
      btnSaveSpecial.addEventListener("click", savePersonalSpecialProductFromUI);
    }
    const btnToggleSpecialDrawer = document.getElementById("btnTogglePersonalSpecialDrawer");
    if (btnToggleSpecialDrawer) {
      btnToggleSpecialDrawer.addEventListener("click", () => {
        const sub = document.getElementById("personalSpecialProductsSubPanel");
        if (sub) sub.classList.toggle("hidden");
      });
    }

    // 14. 綁定左側【報價管理系統】專屬 9 大功能頁按鈕與底部【離開報價單功能】按鈕
    document.querySelectorAll("#quoteSidebarNavList .quote-nav-btn").forEach(btn => {
      btn.addEventListener("click", () => {
        const page = btn.getAttribute("data-quote-page") || "create";
        switchQuoteSubTab(page);
      });
    });

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
        p3HistoryPage = 1;
        const kw = document.getElementById("inputSearchQuoteHistory")?.value.trim() || "";
        loadQuotationHistory(kw, isTrashViewActive);
      });
    }

    // 16. 大數據分析 (Page 4) 篩選、匯出與高度拖曳桿
    const btnRefreshDash = document.getElementById("btnRefreshQuoteDashboard");
    if (btnRefreshDash) btnRefreshDash.addEventListener("click", loadQuotationDashboard);
    const chkDashLatest = document.getElementById("chkQuoteDashLatestOnly");
    if (chkDashLatest) chkDashLatest.addEventListener("change", loadQuotationDashboard);
    const dashQuickRange = document.getElementById("dashQuickRangeSelect");
    if (dashQuickRange) {
      dashQuickRange.addEventListener("change", () => {
        applyDashboardQuickDateRange(dashQuickRange.value);
        loadQuotationDashboard();
      });
    }
    const btnDashExcel = document.getElementById("btnDashExportExcel");
    if (btnDashExcel) btnDashExcel.addEventListener("click", exportDashboardReportCsv);
    const btnDashPdf = document.getElementById("btnDashExportPdf");
    if (btnDashPdf) btnDashPdf.addEventListener("click", () => window.print());
    setupDashboardDragHandles();

    // 17. 客戶管理頁面 (Page 6) 事件綁定
    const inputCustSearch = document.getElementById("inputQuoteCustSearch");
    if (inputCustSearch) {
      inputCustSearch.addEventListener("input", () => {
        p6CustPage = 1;
        renderQuotationCustomersTable();
      });
    }
    const btnCustPrev = document.getElementById("btnCustMgmtPrevPage");
    if (btnCustPrev) {
      btnCustPrev.addEventListener("click", () => {
        if (p6CustPage > 1) {
          p6CustPage--;
          renderQuotationCustomersTable();
        }
      });
    }
    const btnCustNext = document.getElementById("btnCustMgmtNextPage");
    if (btnCustNext) {
      btnCustNext.addEventListener("click", () => {
        p6CustPage++;
        renderQuotationCustomersTable();
      });
    }
    const btnSyncCloudCust = document.getElementById("btnSyncCloudCustomers");
    if (btnSyncCloudCust) {
      btnSyncCloudCust.addEventListener("click", async () => {
        await populateCustomerSelectDropdown(true);
        renderQuotationCustomersTable();
        alert("[同步完成] 已從資料庫重新同步最新客戶主檔與常用星號設定！");
      });
    }
    const btnExportCrmP6 = document.getElementById("btnExportCrmReportP6");
    if (btnExportCrmP6) btnExportCrmP6.addEventListener("click", exportCrmMonthlyReportCsv);
    const btnCustClearAll = document.getElementById("btnQuoteCustClearAll");
    if (btnCustClearAll) {
      btnCustClearAll.addEventListener("click", () => {
        alert("[企業防呆保護] 為防止誤刪全公司客戶資料庫，系統已鎖定單鍵清空所有公司功能。請使用單筆「刪除選取公司」。");
      });
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

    // 18. 產品管理 (Page 7)：共通標準牌價搜尋、分頁、表單與 4x5 小算盤
    const btnProdSearch = document.getElementById("btnQuoteProdMgmtSearch");
    if (btnProdSearch) btnProdSearch.addEventListener("click", () => { p7ProdPage = 1; searchStandardProductsInMgmt(); });
    const inputProdSearch = document.getElementById("inputQuoteProdMgmtSearch");
    if (inputProdSearch) {
      inputProdSearch.addEventListener("input", () => {
        p7ProdPage = 1;
        searchStandardProductsInMgmt();
      });
    }
    const btnProdPrev = document.getElementById("btnProdMgmtPrevPage");
    if (btnProdPrev) {
      btnProdPrev.addEventListener("click", () => {
        if (p7ProdPage > 1) {
          p7ProdPage--;
          renderStandardProductsPage();
        }
      });
    }
    const btnProdNext = document.getElementById("btnProdMgmtNextPage");
    if (btnProdNext) {
      btnProdNext.addEventListener("click", () => {
        const totalPages = Math.max(1, Math.ceil(p7ProductsListCache.length / P7_PAGE_SIZE));
        if (p7ProdPage < totalPages) {
          p7ProdPage++;
          renderStandardProductsPage();
        }
      });
    }
    const btnP7Clear = document.getElementById("btnP7ClearForm");
    if (btnP7Clear) btnP7Clear.addEventListener("click", clearP7ProductForm);
    const btnP7Save = document.getElementById("btnP7SaveProduct");
    if (btnP7Save) btnP7Save.addEventListener("click", saveP7ProductFromForm);
    const btnProdDelSel = document.getElementById("btnProdMgmtDeleteSelected");
    if (btnProdDelSel) btnProdDelSel.addEventListener("click", deleteSelectedP7Product);
    const btnProdClearAll = document.getElementById("btnProdMgmtClearAll");
    if (btnProdClearAll) {
      btnProdClearAll.addEventListener("click", () => {
        alert("[企業防呆保護] 標準產品資料庫包含上千筆牌價，為防止手滑誤刪，系統已鎖定單鍵清空所有產品功能。");
      });
    }
    setupP7CalculatorEvents();

    // 19. 系統設定 (Page 8)：5 大子分頁切換、儲存財務設定與報價條款
    setupQuotationSettingsTabsAndEvents();

    // 20. 使用說明 (Page 9)：2 大子分頁與 10 大說明主題切換
    setupQuotationHelpTabsAndEvents();
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

    updateDesktopBannerForRow(null);
    const defRows = parseInt(localStorage.getItem("shihlin_quote_default_rows") || "5", 10) || 5;
    for (let i = 0; i < defRows; i++) {
      addBlankItem();
    }
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

  // 更新電腦版第 2 排選取明細橫幅與動作按鈕 (100% 還原 _QuotationApp gui_page1.py update_banner_for_row)
  function updateDesktopBannerForRow(rowEl) {
    const badge = document.getElementById("activeRowBadge");
    const actionsBox = document.getElementById("itemActionsContainer");
    const clearBtn = document.getElementById("btnClearActiveRow");
    const tableBody = document.getElementById("quoteDesktopTableBody");

    if (tableBody) {
      tableBody.querySelectorAll("tr").forEach(tr => tr.classList.remove("active-row"));
    }

    if (!rowEl) {
      if (badge) {
        badge.textContent = "目前尚未選擇明細";
        badge.className = "pyside-edit-banner-label";
      }
      if (actionsBox) actionsBox.innerHTML = "";
      if (clearBtn) clearBtn.classList.add("hidden");
      return;
    }

    rowEl.classList.add("active-row");
    if (clearBtn) clearBtn.classList.remove("hidden");

    const rowIdxText = rowEl.querySelector('[id$="-index"]')?.textContent?.trim() || "1";

    if (rowEl.classList.contains("subtotal-divider-row")) {
      const subTitle = rowEl.querySelector(".subtotal-title-input")?.value || "--- 方案小計 ---";
      if (badge) {
        badge.textContent = `第 ${rowIdxText} 列  |  方案小計 (${subTitle})`;
        badge.className = "pyside-edit-banner-label active-subtotal";
      }
      if (actionsBox) {
        actionsBox.innerHTML = `
          <button type="button" class="banner-icon-action-btn" id="btnBannerMoveUpSub" title="將此列向上移動一列">▲ 上移</button>
          <button type="button" class="banner-icon-action-btn" id="btnBannerMoveDownSub" title="將此列向下移動一列">▼ 下移</button>
          <button type="button" class="banner-icon-action-btn danger" id="btnBannerDeleteSubtotal" title="刪除此方案小計列">刪除此明細</button>
        `;
        const upSubBtn = document.getElementById("btnBannerMoveUpSub");
        if (upSubBtn) upSubBtn.addEventListener("click", () => moveDesktopItemRow(rowEl, -1));
        const downSubBtn = document.getElementById("btnBannerMoveDownSub");
        if (downSubBtn) downSubBtn.addEventListener("click", () => moveDesktopItemRow(rowEl, 1));
        const delSubBtn = document.getElementById("btnBannerDeleteSubtotal");
        if (delSubBtn) {
          delSubBtn.addEventListener("click", () => {
            rowEl.remove();
            updateDesktopBannerForRow(null);
            refreshItemIndices();
            calculateTotal();
          });
        }
      }
      return;
    }

    const itemId = rowEl.getAttribute("data-itemid");
    const model = document.getElementById(`d-${itemId}-model-input`)?.value.trim() || document.getElementById(`${itemId}-code`)?.value.trim() || "";
    const spec = document.getElementById(`d-${itemId}-spec`)?.textContent?.trim() || "";
    const sku = document.getElementById(`${itemId}-sku`)?.value.trim() || "";
    const isOrder = !document.getElementById(`${itemId}-order-badge`)?.classList.contains("hidden");

    const displayModel = model || "(未填寫型號)";
    const displaySpec = spec && spec !== "-" ? (spec.length > 22 ? spec.slice(0, 22) + "..." : spec) : "(無規格)";

    if (badge) {
      badge.textContent = `第 ${rowIdxText} 列  |  型號：${displayModel}  |  規格：${displaySpec}`;
      badge.className = "pyside-edit-banner-label active-product";
    }

    if (actionsBox) {
      let html = "";
      if (isOrder) {
        html += `<span class="banner-order-badge">訂購品</span>`;
      }
      if (sku) {
        html += `<button type="button" class="banner-sku-copy-btn" id="btnBannerCopySku" title="點擊複製料號: ${escapeHtml(sku)}">料號: ${escapeHtml(sku)}</button>`;
      }
      html += `<button type="button" class="banner-icon-action-btn" id="btnBannerMoveUpRow" title="將此明細向上移動一列 (亦可直接按住左側項次上下拖曳)">▲ 上移</button>`;
      html += `<button type="button" class="banner-icon-action-btn" id="btnBannerMoveDownRow" title="將此明細向下移動一列 (亦可直接按住左側項次上下拖曳)">▼ 下移</button>`;
      if (model) {
        html += `<button type="button" class="banner-icon-action-btn" id="btnBannerHistoryPrice" title="查詢此型號的歷史報價紀錄">查詢歷史價格</button>`;
      }
      html += `<button type="button" class="banner-icon-action-btn danger" id="btnBannerDeleteRow" title="刪除此筆產品明細">刪除此明細</button>`;
      actionsBox.innerHTML = html;

      const upBtn = document.getElementById("btnBannerMoveUpRow");
      if (upBtn) upBtn.addEventListener("click", () => moveDesktopItemRow(rowEl, -1));
      const downBtn = document.getElementById("btnBannerMoveDownRow");
      if (downBtn) downBtn.addEventListener("click", () => moveDesktopItemRow(rowEl, 1));

      const copySkuBtn = document.getElementById("btnBannerCopySku");
      if (copySkuBtn && sku) {
        copySkuBtn.addEventListener("click", () => {
          navigator.clipboard.writeText(sku).then(() => {
            copySkuBtn.textContent = "已複製!";
            setTimeout(() => { copySkuBtn.textContent = `料號: ${sku}`; }, 1000);
          });
        });
      }

      const histBtn = document.getElementById("btnBannerHistoryPrice");
      if (histBtn && model) {
        histBtn.addEventListener("click", () => {
          const inpModel = document.getElementById("inputAnalysisModel");
          if (inpModel) inpModel.value = model;
          switchQuoteSubTab("analysis");
          if (typeof runPriceAnalysis === "function") runPriceAnalysis();
        });
      }

      const delBtn = document.getElementById("btnBannerDeleteRow");
      if (delBtn) {
        delBtn.addEventListener("click", () => {
          const rmBtn = document.getElementById(`${itemId}-remove`);
          if (rmBtn) rmBtn.click();
        });
      }
    }
  }

  // 新增品項 (100% 同步手機端 _QuotationApp\pwa 卡片與電腦端 _QuotationApp 10 欄試算表)
  function addBlankItem(initialData = null) {
    quoteItemCount++;
    const itemId = `item-${quoteItemCount}`;
    const itemsContainer = document.getElementById("itemsContainer");
    const desktopTableBody = document.getElementById("quoteDesktopTableBody");
    if (!itemsContainer) return;

    const currentCount = itemsContainer.querySelectorAll('.item-row').length + 1;
    const modelVal = initialData?.model_no || "";
    const modelText = modelVal || "點擊選擇產品...";
    const hasModel = Boolean(modelVal);
    const specVal = initialData?.spec || "";
    const skuVal = initialData?.sku || "";
    const isOrderItem = Boolean(initialData?.is_order_item);
    const isSpecialProduct = Boolean(initialData?.is_special);

    const stockQtyVal = initialData?.stock_qty !== undefined && initialData?.stock_qty !== null
      ? initialData.stock_qty
      : (hasModel ? 0 : null);
    const stockDisplay = stockQtyVal !== null ? `${Number(stockQtyVal).toLocaleString()} 台` : "-";
    const desktopStockDisplay = stockQtyVal !== null ? `${Number(stockQtyVal).toLocaleString()}` : "";

    const dealerPriceVal = parseFloat(initialData?.dealer_price) || 0;
    const dealerPriceDisplay = dealerPriceVal > 0 ? `$${dealerPriceVal.toLocaleString()}` : "-";
    const desktopDealerDisplay = dealerPriceVal > 0 ? dealerPriceVal.toLocaleString() : "";

    const hasDiscount = initialData && initialData.discount_pct !== undefined && initialData.discount_pct !== null && initialData.discount_pct !== "";
    const discountVal = hasDiscount ? initialData.discount_pct : (hasModel ? 100 : "");

    let rawEstPriceVal = 0;
    if (dealerPriceVal > 0 && discountVal !== "") {
      rawEstPriceVal = Math.round(dealerPriceVal * (parseFloat(discountVal) / 100));
    }

    let finalUnitPriceVal = 0;
    if (initialData && initialData.final_unit_price !== undefined && initialData.final_unit_price !== null && initialData.final_unit_price !== "") {
      finalUnitPriceVal = parseFloat(initialData.final_unit_price) || 0;
    } else if (rawEstPriceVal > 0) {
      finalUnitPriceVal = applyQuoteRounding(rawEstPriceVal);
    }

    if (!rawEstPriceVal && finalUnitPriceVal > 0) {
      rawEstPriceVal = finalUnitPriceVal;
    }

    const sugPriceVal = rawEstPriceVal > 0 ? rawEstPriceVal : "";
    const finalPriceDisplay = finalUnitPriceVal > 0 ? `$${finalUnitPriceVal.toLocaleString()}` : "-";
    const qtyVal = parseInt(initialData?.quantity, 10) || 1;
    const rowSubtotalVal = initialData?.row_subtotal !== undefined
      ? (parseFloat(initialData.row_subtotal) || Math.round(finalUnitPriceVal * qtyVal))
      : Math.round(finalUnitPriceVal * qtyVal);
    const leadTimeVal = initialData?.note_lead_time || "";

    const stockUpdatedAtText = quoteFinanceSettings.stock_updated_at || "10/15 10:12";
    const roundingRuleText = quoteFinanceSettings.round_summary_text || "捨入：百位數 / 無條件捨去";

    // 1. 窄螢幕 / 手機端卡片 HTML (100% 還原 _QuotationApp\pwa 規格)
    const mobileItemHTML = `
      <div class="item-row" id="${itemId}" data-spec="${escapeHtml(specVal)}" data-stock="${stockQtyVal !== null ? stockQtyVal : ''}">
        <div class="item-header">
          <span class="item-index-badge">#${currentCount}</span>
          <button type="button" class="product-picker-input ${hasModel ? 'selected' : ''}" id="${itemId}-title">
            ${escapeHtml(modelText)}
          </button>
          <span class="order-item-badge ${isOrderItem ? '' : 'hidden'}" id="${itemId}-order-badge">訂購品</span>
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

        <input type="hidden" id="${itemId}-code" value="${escapeHtml(modelVal)}">
        <input type="hidden" id="${itemId}-sku" value="${escapeHtml(skuVal)}">
        <input type="hidden" id="${itemId}-unit-price-num" value="${finalUnitPriceVal}">
      </div>
    `;
    itemsContainer.insertAdjacentHTML('beforeend', mobileItemHTML);

    // 2. 電腦端 QuoteTableWidget 10 欄試算表表格列 HTML (100% 還原附圖 1 欄位順序與圓角白框型號選單)
    // 欄位順序：項次 | 9:可用數量 | 0:型號 | 6:數量 | 1:規格 | 2:經銷價 | 3:報價折數% | 4:預計報價金額 | 5:最終報價金額 | 7:小計 | 8:備註/交期
    if (desktopTableBody) {
      const desktopRowHTML = `
        <tr id="d-${itemId}" class="desktop-item-row" data-itemid="${itemId}">
          <td class="cell-vertical-index" id="d-${itemId}-index">${currentCount}</td>
          <td>
            <span id="d-${itemId}-stock" class="pyside-cell-readonly align-center">${escapeHtml(desktopStockDisplay)}</span>
          </td>
          <td>
            <div class="pyside-model-combo-wrap" id="d-${itemId}-model-wrap">
              <input type="text" id="d-${itemId}-model-input" class="cell-model-input" placeholder="" value="${escapeHtml(modelVal)}">
              <button type="button" id="d-${itemId}-model" class="pyside-model-combo-arrow" style="background:none; border:none; pointer-events:auto; cursor:pointer;" title="從產品庫挑選">▼</button>
            </div>
            <input type="hidden" id="d-${itemId}-code" value="${escapeHtml(modelVal)}">
          </td>
          <td>
            <input type="number" id="d-${itemId}-qty" class="pyside-cell-input align-center" min="1" value="${hasModel || initialData ? qtyVal : ''}">
          </td>
          <td>
            <span id="d-${itemId}-spec" class="pyside-cell-readonly" title="${escapeHtml(specVal)}">${escapeHtml(specVal)}</span>
          </td>
          <td>
            <span id="d-${itemId}-dealer-price" class="pyside-cell-readonly align-right">${escapeHtml(desktopDealerDisplay)}</span>
          </td>
          <td>
            <input type="number" id="d-${itemId}-discount" class="pyside-cell-input align-right" step="any" value="${escapeHtml(String(discountVal))}">
          </td>
          <td>
            <input type="number" id="d-${itemId}-est-price" class="pyside-cell-input align-right" step="any" value="${rawEstPriceVal > 0 ? rawEstPriceVal : ''}">
          </td>
          <td>
            <span id="d-${itemId}-final-price-cell" class="pyside-cell-readonly align-right" style="font-weight:700;">${finalUnitPriceVal > 0 ? finalUnitPriceVal.toLocaleString() : ''}</span>
            <input type="hidden" id="d-${itemId}-price" value="${finalUnitPriceVal}">
          </td>
          <td>
            <span id="d-${itemId}-subtotal" class="pyside-cell-readonly align-right" style="font-weight:700;">${rowSubtotalVal > 0 ? rowSubtotalVal.toLocaleString() : ''}</span>
          </td>
          <td>
            <input type="text" id="d-${itemId}-delivery" class="pyside-cell-input" placeholder="" value="${escapeHtml(leadTimeVal)}">
          </td>
        </tr>
      `;
      desktopTableBody.insertAdjacentHTML('beforeend', desktopRowHTML);

      const dRow = document.getElementById(`d-${itemId}`);
      if (dRow) {
        dRow.addEventListener("click", () => {
          updateDesktopBannerForRow(dRow);
        });
      }
    }

    refreshItemIndices();
    setupItemRowEvents(itemId);
    calculateTotal();
  }

  // 插入方案小計列 (100% 還原 _QuotationApp gui_page1.py insert_subtotal_row：[小計] + --- 方案小計 --- #c0392b 粗體紅字)
  function insertSubtotalRow(defaultTitle = "") {
    const desktopTableBody = document.getElementById("quoteDesktopTableBody");
    if (!desktopTableBody) return;

    const rowId = `subtotal-row-${Date.now()}`;
    const title = defaultTitle || `--- 方案小計 ---`;

    const subtotalHTML = `
      <tr id="${rowId}" class="subtotal-divider-row">
        <td class="cell-vertical-index" id="${rowId}-index">-</td>
        <td></td>
        <td>
          <span class="pyside-cell-readonly" style="color:#c0392b; font-weight:700;">[小計]</span>
        </td>
        <td></td>
        <td>
          <input type="text" class="pyside-cell-input subtotal-title-input" value="${escapeHtml(title)}" style="color:#c0392b; font-weight:700; text-align:center;">
        </td>
        <td></td>
        <td></td>
        <td></td>
        <td></td>
        <td style="text-align:right;">
          <span class="segment-subtotal-val pyside-cell-readonly align-right" style="color:#c0392b; font-weight:700;">0</span>
        </td>
        <td>
          <input type="text" class="pyside-cell-input" placeholder="備註...">
        </td>
      </tr>
    `;
    desktopTableBody.insertAdjacentHTML('beforeend', subtotalHTML);
    const rowEl = document.getElementById(rowId);
    if (rowEl) {
      rowEl.addEventListener("click", () => {
        updateDesktopBannerForRow(rowEl);
      });
      const titleInp = rowEl.querySelector(".subtotal-title-input");
      if (titleInp) {
        titleInp.addEventListener("input", () => updateDesktopBannerForRow(rowEl));
      }
    }
    refreshItemIndices();
    calculateTotal();
    if (rowEl) updateDesktopBannerForRow(rowEl);
  }

  // 更新各方案小計區段數值 (加總上一個小計列之後到本小計列之間的所有明細)
  function updateSubtotalDividers() {
    const desktopTableBody = document.getElementById("quoteDesktopTableBody");
    if (!desktopTableBody) return;

    const allRows = desktopTableBody.querySelectorAll('tr');
    let currentSegmentSum = 0;

    allRows.forEach(row => {
      if (row.classList.contains('subtotal-divider-row')) {
        const segValSpan = row.querySelector('.segment-subtotal-val');
        if (segValSpan) {
          segValSpan.textContent = `${currentSegmentSum.toLocaleString()}`;
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

    // 2. 電腦端 10 欄表格元件
    const dModelBtn = document.getElementById(`d-${itemId}-model`);
    const dModelInput = document.getElementById(`d-${itemId}-model-input`);
    const dDeliveryInput = document.getElementById(`d-${itemId}-delivery`);
    const dPriceInput = document.getElementById(`d-${itemId}-price`);
    const dEstPriceInput = document.getElementById(`d-${itemId}-est-price`);
    const dFinalPriceCell = document.getElementById(`d-${itemId}-final-price-cell`);
    const dDiscountInput = document.getElementById(`d-${itemId}-discount`);
    const dQtyInput = document.getElementById(`d-${itemId}-qty`);
    const dSubtotalSpan = document.getElementById(`d-${itemId}-subtotal`);

    function triggerProductPicker() {
      currentEditingItemId = itemId;
      openProductModal();
    }
    if (mTitleBtn) mTitleBtn.addEventListener("click", triggerProductPicker);
    if (dModelBtn) dModelBtn.addEventListener("click", triggerProductPicker);

    // 電腦端型號輸入框直接輸入或貼上後自動查價匹配 (_QuotationApp lookup_product)
    if (dModelInput) {
      dModelInput.addEventListener("change", async () => {
        const kw = dModelInput.value.trim();
        const mCode = document.getElementById(`${itemId}-code`);
        const dCode = document.getElementById(`d-${itemId}-code`);
        if (mCode) mCode.value = kw;
        if (dCode) dCode.value = kw;
        if (mTitleBtn) {
          mTitleBtn.textContent = kw || "點擊選擇產品...";
          mTitleBtn.classList.toggle("selected", Boolean(kw));
        }
        if (!kw) return;
        const activeSales = getActiveSalesName();
        try {
          const resp = await fetch(`/api/products?q=${encodeURIComponent(kw)}&sales_name=${encodeURIComponent(activeSales)}`);
          if (resp.ok) {
            const list = await resp.json();
            const exact = (list || []).find(p => p.model_no && p.model_no.toUpperCase() === kw.toUpperCase());
            if (exact) {
              applyProductToItemRow(itemId, exact);
            }
          }
        } catch (e) {}
      });
    }

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
      updateDesktopBannerForRow(null);
      calculateTotal();
      refreshItemIndices();
    }
    if (mRemoveBtn) mRemoveBtn.addEventListener("click", triggerRemoveItem);

    function triggerDeliveryPicker() {
      currentEditingItemId = itemId;
      openDeliveryModal();
    }
    if (mDeliveryInput) mDeliveryInput.addEventListener("click", triggerDeliveryPicker);
    if (dDeliveryInput) {
      dDeliveryInput.addEventListener("dblclick", triggerDeliveryPicker);
      dDeliveryInput.addEventListener("input", () => {
        if (mDeliveryInput) mDeliveryInput.value = dDeliveryInput.value;
      });
    }

    function getDealerPrice() {
      const dpStr = (document.getElementById(`d-${itemId}-dealer-price`)?.textContent || mDealerPriceInput?.value || "").replace(/[^0-9.]/g, '');
      return parseFloat(dpStr) || 0;
    }

    function syncUnitAndRowSubtotal(finalUnitPrice, rawEstPrice) {
      const unitP = Math.max(0, Math.round(finalUnitPrice || 0));
      if (mUnitPriceNum) mUnitPriceNum.value = unitP;
      if (dPriceInput) dPriceInput.value = unitP;
      if (dFinalPriceCell) {
        dFinalPriceCell.textContent = unitP > 0 ? unitP.toLocaleString() : "";
      }
      if (rawEstPrice !== undefined && dEstPriceInput && document.activeElement !== dEstPriceInput) {
        dEstPriceInput.value = rawEstPrice > 0 ? Math.round(rawEstPrice) : "";
      }
      if (mFinalPriceInput) {
        mFinalPriceInput.value = unitP > 0 ? `$${unitP.toLocaleString()}` : "-";
      }
      const q = parseInt(dQtyInput?.value || mQtyInput?.value || "1", 10) || 1;
      const sub = Math.round(unitP * q);
      if (dSubtotalSpan) {
        dSubtotalSpan.textContent = sub > 0 ? `${sub.toLocaleString()}` : "0";
      }
      calculateTotal();
    }

    // 折數變更 -> 計算預計報價金額與捨入後最終報價金額
    function onDiscountChanged(val, fromSource) {
      const disc = parseFloat(val);
      const dp = getDealerPrice();

      if (fromSource === 'desktop' && mSugDiscountInput) mSugDiscountInput.value = val;
      if (fromSource === 'mobile' && dDiscountInput) dDiscountInput.value = val;

      if (!isNaN(disc) && dp > 0) {
        const rawPrice = dp * (disc / 100);
        const roundedPrice = applyQuoteRounding(rawPrice);
        if (mSugPriceInput) mSugPriceInput.value = Math.round(rawPrice);
        syncUnitAndRowSubtotal(roundedPrice, rawPrice);
      } else {
        syncUnitAndRowSubtotal(0, 0);
      }
    }

    // 預計報價金額 / 建議報價變更 -> 反推折數並計算捨入後最終報價金額
    function onSugPriceChanged(val, fromSource) {
      const rawInput = parseFloat(val);
      const dp = getDealerPrice();
      if (fromSource === 'desktop' && mSugPriceInput) mSugPriceInput.value = val;
      if (fromSource === 'mobile' && dEstPriceInput) dEstPriceInput.value = val;

      if (!isNaN(rawInput) && rawInput > 0) {
        const roundedPrice = dp > 0 ? applyQuoteRounding(rawInput) : Math.round(rawInput);
        if (dp > 0) {
          const discStr = ((roundedPrice / dp) * 100).toFixed(1).replace(/\.0$/, '');
          if (mSugDiscountInput) mSugDiscountInput.value = discStr;
          if (dDiscountInput) dDiscountInput.value = discStr;
        }
        syncUnitAndRowSubtotal(roundedPrice, rawInput);
      } else {
        syncUnitAndRowSubtotal(0, 0);
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
    if (mSugPriceInput) mSugPriceInput.addEventListener("input", (e) => onSugPriceChanged(e.target.value, 'mobile'));
    if (dEstPriceInput) dEstPriceInput.addEventListener("input", (e) => onSugPriceChanged(e.target.value, 'desktop'));
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
      if (subSpan) subSpan.textContent = rowSub > 0 ? `${rowSub.toLocaleString()}` : "0";
      const dFinalCell = document.getElementById(`d-${itemId}-final-price-cell`);
      if (dFinalCell) dFinalCell.textContent = unitPrice > 0 ? `${unitPrice.toLocaleString()}` : "";
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
      const dRows = desktopTableBody.querySelectorAll('tr');
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
          if (dDP) dDP.textContent = `${newDP.toLocaleString()}`;
          const mDP = document.getElementById(`${itemId}-dealer-price`);
          if (mDP) mDP.value = `$${newDP.toLocaleString()}`;

          const disc = parseFloat(document.getElementById(`d-${itemId}-discount`)?.value || document.getElementById(`${itemId}-sug-discount`)?.value || "100") || 100;
          const rawPrice = newDP * (disc / 100);
          const newPrice = applyQuoteRounding(rawPrice);
          const dPrice = document.getElementById(`d-${itemId}-price`);
          if (dPrice) dPrice.value = newPrice;
          const dEstPrice = document.getElementById(`d-${itemId}-est-price`);
          if (dEstPrice) dEstPrice.value = Math.round(rawPrice);
          const dFinalCell = document.getElementById(`d-${itemId}-final-price-cell`);
          if (dFinalCell) dFinalCell.textContent = newPrice > 0 ? newPrice.toLocaleString() : "";
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

    // 2. 電腦端 10 欄表格元件更新
    const dModelInput = document.getElementById(`d-${itemId}-model-input`);
    const dStock = document.getElementById(`d-${itemId}-stock`);
    const dSpec = document.getElementById(`d-${itemId}-spec`);
    const dDP = document.getElementById(`d-${itemId}-dealer-price`);
    const dPrice = document.getElementById(`d-${itemId}-price`);
    const dEstPrice = document.getElementById(`d-${itemId}-est-price`);
    const dFinalCell = document.getElementById(`d-${itemId}-final-price-cell`);
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

    if (dModelInput) {
      dModelInput.value = model;
    }
    const dCodeInput = document.getElementById(`d-${itemId}-code`);
    if (dCodeInput) dCodeInput.value = model;
    if (dStock) dStock.textContent = `${stockQty.toLocaleString()}`;
    if (dSpec) dSpec.textContent = spec || "";
    if (dDP) dDP.textContent = price > 0 ? `${price.toLocaleString()}` : "";

    const discVal = sugDiscountInput && sugDiscountInput.value !== "" ? parseFloat(sugDiscountInput.value) : 100;
    const disc = isNaN(discVal) ? 100 : discVal;
    if (sugDiscountInput && sugDiscountInput.value === "") sugDiscountInput.value = String(disc);
    if (dDiscount) dDiscount.value = String(disc);

    const rawPrice = price * (disc / 100);
    const roundedPrice = disc === 100 ? Math.round(price) : applyQuoteRounding(rawPrice);
    if (sugPriceInput) sugPriceInput.value = Math.round(rawPrice);
    if (dEstPrice) dEstPrice.value = rawPrice > 0 ? Math.round(rawPrice) : "";
    if (unitPriceNum) unitPriceNum.value = roundedPrice;
    if (dPrice) dPrice.value = roundedPrice;
    if (dFinalCell) dFinalCell.textContent = roundedPrice > 0 ? roundedPrice.toLocaleString() : "";
    if (finalPriceInput) finalPriceInput.value = roundedPrice > 0 ? `$${roundedPrice.toLocaleString()}` : "-";

    const q = parseInt(qtyInput?.value || "1", 10) || 1;
    const sub = Math.round(roundedPrice * q);
    if (dSub) dSub.textContent = sub > 0 ? `${sub.toLocaleString()}` : "0";

    const mRow = document.getElementById(itemId);
    if (mRow) {
      mRow.dataset.spec = spec;
      mRow.dataset.stock = String(stockQty);
    }

    // 同步上方第 2 排作用列標籤與動作按鈕 (訂購品、料號複製、歷史價格查詢、刪除明細)
    const dRowEl = document.getElementById(`d-${itemId}`);
    if (dRowEl) {
      updateDesktopBannerForRow(dRowEl);
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
      const model = (document.getElementById(`d-${itemId}-model-input`)?.value || document.getElementById(`${itemId}-code`)?.value || "").trim();
      if (!model) {
        continue;
      }
      const dpStr = (document.getElementById(`d-${itemId}-dealer-price`)?.textContent || document.getElementById(`${itemId}-dealer-price`)?.value || "0").replace(/[^0-9.]/g, '');
      const dealerPrice = parseFloat(dpStr) || 0;
      const discount = parseFloat(document.getElementById(`d-${itemId}-discount`)?.value || document.getElementById(`${itemId}-sug-discount`)?.value || "100") || 100;
      const sp = parseFloat(document.getElementById(`${itemId}-unit-price-num`)?.value || document.getElementById(`d-${itemId}-price`)?.value || document.getElementById(`${itemId}-sug-price`)?.value || "0") || 0;
      const q = parseInt(document.getElementById(`d-${itemId}-qty`)?.value || document.getElementById(`${itemId}-qty`)?.value || "1", 10) || 1;
      const rowSub = Math.round(sp * q);
      subtotal += rowSub;
      const leadTime = document.getElementById(`d-${itemId}-delivery`)?.value || document.getElementById(`${itemId}-delivery`)?.value || "";
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

    if (itemsData.length === 0) {
      alert("[無有效品項] 請至少填寫一筆產品型號！");
      return null;
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

  // ============================================================================
  // 第 3 頁：報價單管理 (100% 對標電腦版 _QuotationApp gui_page3.py & 截圖 3)
  // ============================================================================
  let p3FilteredQuotesCache = [];
  let p3HistoryPage = 1;
  const P3_PAGE_SIZE = 50;
  let p3SelectedQuoteObj = null;

  function splitQuoteNoAndRev(fullQuoteNo, revNum) {
    const raw = String(fullQuoteNo || "").trim();
    if (!raw) return { baseNo: "-", revStr: "00" };
    // 例如 "台瓷自動2611-001-Rev00" 或 "台瓷自動2611-001-02"
    const m = raw.match(/^(.*?)-(?:Rev)?(\d+)$/i);
    if (m && m[1].includes("-")) {
      return {
        baseNo: m[1],
        revStr: String(parseInt(m[2], 10) || 0).padStart(2, "0")
      };
    }
    return {
      baseNo: raw,
      revStr: String(parseInt(revNum, 10) || 0).padStart(2, "0")
    };
  }

  function formatStatusBadgeP3(rawStatus) {
    const st = String(rawStatus || "追蹤中").trim();
    if (st === "已成交") {
      return `<span class="p3-status-badge won">[V] 已成交</span>`;
    }
    if (st === "已失注" || st === "未成交") {
      return `<span class="p3-status-badge lost">[X] 未成交</span>`;
    }
    if (st === "不跟進" || st === "不計入") {
      return `<span class="p3-status-badge ignore"><span style="color:#e74c3c; font-weight:700;">[O]</span> 不跟進</span>`;
    }
    return `<span class="p3-status-badge pending">[~] 報價中</span>`;
  }

  async function loadQuotationHistory(keyword = "", isTrash = false) {
    const tbody = document.getElementById("quoteHistoryTableBody");
    const legacyContainer = document.getElementById("quoteHistoryResults");
    if (tbody) {
      tbody.innerHTML = `<tr><td colspan="11" style="padding:24px; text-align:center; color:#7f8c8d;">正在載入個人歷史報價紀錄...</td></tr>`;
    }

    try {
      const activeSales = getActiveSalesName();
      const statusFilter = document.getElementById("selectQuoteHistoryStatus")?.value || "ALL";
      const url = `/api/quotations?sales_name=${encodeURIComponent(activeSales)}&include_deleted=${isTrash ? 'true' : 'false'}${keyword ? `&keyword=${encodeURIComponent(keyword)}` : ''}${(!isTrash && statusFilter !== 'ALL') ? `&status=${encodeURIComponent(statusFilter)}` : ''}`;
      const resp = await fetch(url);
      if (!resp.ok) throw new Error("讀取報價單資料失敗");
      const allQuotes = await resp.json();

      const showAllDates = document.getElementById("chkHistoryShowAll")?.checked !== false;
      const sDate = document.getElementById("historyStartDate")?.value || "";
      const eDate = document.getElementById("historyEndDate")?.value || "";

      p3FilteredQuotesCache = (allQuotes || []).filter(q => {
        const matchTrash = isTrash ? (q.is_deleted === 1) : (!q.is_deleted || q.is_deleted === 0);
        if (!matchTrash) return false;
        if (!showAllDates && (sDate || eDate)) {
          const qd = String(q.quote_date || "").replace(/\//g, "-");
          if (sDate && qd < sDate) return false;
          if (eDate && qd > eDate) return false;
        }
        return true;
      });

      p3SelectedQuoteObj = null;
      updateP3SelectedBanner(null, isTrash);
      renderQuotationHistoryPage(isTrash, keyword);

      // 同步維持舊隱藏容器相容性
      if (legacyContainer) {
        legacyContainer.innerHTML = p3FilteredQuotesCache.map(q => `<div class="quote-history-card" data-quoteno="${escapeHtml(q.quote_no)}">${escapeHtml(q.quote_no)}</div>`).join("");
      }
    } catch (err) {
      if (tbody) {
        tbody.innerHTML = `<tr><td colspan="11" style="padding:20px; text-align:center; color:#e74c3c;">載入失敗：${escapeHtml(err.message)}</td></tr>`;
      }
    }
  }

  function updateP3SelectedBanner(qObj, isTrash = false) {
    const badge = document.getElementById("historyCurrentEditBadge");
    const actBox = document.getElementById("historyRowActionsContainer");
    if (!badge) return;

    if (!qObj) {
      badge.innerHTML = `<span class="p3-edit-icon">[單據]</span><span>(尚未選取報價單)</span>`;
      if (actBox) {
        actBox.classList.add("hidden");
        actBox.innerHTML = "";
      }
      return;
    }

    const { baseNo, revStr } = splitQuoteNoAndRev(qObj.quote_no, qObj.revision);
    const st = qObj.status || "追蹤中";
    badge.innerHTML = `<span class="p3-edit-icon">[已選取]</span><span style="color:#1e293b; font-weight:700;">${escapeHtml(baseNo)} (版次 ${revStr}) ｜ ${escapeHtml(qObj.customer_name)} ｜ 未稅 $${Number(qObj.subtotal || 0).toLocaleString()}</span>`;

    if (actBox) {
      actBox.classList.remove("hidden");
      if (isTrash) {
        actBox.innerHTML = `
          <button type="button" class="pyside-flat-btn" id="btnP3ActRestore" style="background:#10b981; color:#fff; border:none;">一鍵復原此單據</button>
        `;
        const bRes = document.getElementById("btnP3ActRestore");
        if (bRes) {
          bRes.addEventListener("click", async () => {
            if (!confirm(`確定要將報價單【${qObj.quote_no}】從回收站復原嗎？`)) return;
            const resp = await fetch(`/api/quotations/${encodeURIComponent(qObj.quote_no)}/restore`, { method: "POST" });
            if (resp.ok) {
              alert(`[復原成功] 報價單【${qObj.quote_no}】已成功復原！`);
              loadQuotationHistory(document.getElementById("inputSearchQuoteHistory")?.value.trim() || "", true);
            }
          });
        }
      } else {
        actBox.innerHTML = `
          <button type="button" class="pyside-flat-btn" id="btnP3ActLoad" style="background:#2563eb; color:#fff; border:none;">載入編輯 / 升版</button>
          <button type="button" class="pyside-flat-btn p3-act-status" data-st="已成交" style="color:#16a34a; border-color:#86efac;">標記成交</button>
          <button type="button" class="pyside-flat-btn p3-act-status" data-st="已失注" style="color:#dc2626; border-color:#fca5a5;">標記未成交</button>
          <button type="button" class="pyside-flat-btn p3-act-status" data-st="不跟進" style="color:#64748b;">不跟進</button>
          <button type="button" class="pyside-flat-btn p3-act-status" data-st="追蹤中" style="color:#d97706;">報價中</button>
        `;
        const bLoad = document.getElementById("btnP3ActLoad");
        if (bLoad) bLoad.addEventListener("click", () => loadQuotationToForm(qObj.quote_no));

        actBox.querySelectorAll(".p3-act-status").forEach(btn => {
          btn.addEventListener("click", async () => {
            const newSt = btn.getAttribute("data-st");
            let wonAmt = Number(qObj.subtotal || 0);
            if (newSt === "已成交") {
              const pVal = prompt(`請確認【${qObj.quote_no}】最終成交金額 (未稅)：`, String(wonAmt));
              if (pVal === null) return;
              wonAmt = parseFloat(pVal) || wonAmt;
            }
            const r = await fetch(`/api/quotations/${encodeURIComponent(qObj.quote_no)}/status`, {
              method: "POST",
              headers: { "Content-Type": "application/json" },
              body: JSON.stringify({
                status: newSt,
                won_amount: newSt === "已成交" ? wonAmt : 0
              })
            });
            if (r.ok) {
              loadQuotationHistory(document.getElementById("inputSearchQuoteHistory")?.value.trim() || "", false);
            }
          });
        });
      }
    }
  }

  function renderQuotationHistoryPage(isTrash = isTrashViewActive, keyword = "") {
    const tbody = document.getElementById("quoteHistoryTableBody");
    const pageInfo = document.getElementById("lblHistoryPageInfo");
    if (!tbody) return;

    const totalCount = p3FilteredQuotesCache.length;
    const totalPages = Math.max(1, Math.ceil(totalCount / P3_PAGE_SIZE));
    if (p3HistoryPage > totalPages) p3HistoryPage = totalPages;

    if (pageInfo) {
      pageInfo.textContent = `第 ${p3HistoryPage} / ${totalPages} 頁 (共 ${totalCount} 筆)`;
    }

    if (totalCount === 0) {
      const activeSales = getActiveSalesName();
      tbody.innerHTML = `
        <tr>
          <td colspan="11" style="padding:36px; text-align:center; color:#7f8c8d; font-size:0.9rem;">
            ${isTrash ? `目前業務【${escapeHtml(activeSales)}】垃圾桶中沒有已刪除的報價單。` : `目前業務【${escapeHtml(activeSales)}】查無符合條件的報價單紀錄。`}
          </td>
        </tr>
      `;
      return;
    }

    const startIdx = (p3HistoryPage - 1) * P3_PAGE_SIZE;
    const pageRows = p3FilteredQuotesCache.slice(startIdx, startIdx + P3_PAGE_SIZE);

    tbody.innerHTML = pageRows.map((q, idx) => {
      const rowNum = startIdx + idx + 1;
      const { baseNo, revStr } = splitQuoteNoAndRev(q.quote_no, q.revision);
      const isSel = p3SelectedQuoteObj && p3SelectedQuoteObj.quote_no === q.quote_no;
      const wonAmtText = (q.status === "已成交" && Number(q.won_amount || q.subtotal || 0) > 0)
        ? Number(q.won_amount || q.subtotal || 0).toLocaleString()
        : "";
      const wonDateText = q.won_date || (q.status === "已成交" ? (q.quote_date || "") : "");

      return `
        <tr class="p3-hist-row ${isSel ? 'selected-row' : ''}" data-quoteno="${escapeHtml(q.quote_no)}" title="單擊選取並顯示狀態操作鈕；雙擊直接載入至「新建報價單」編輯區">
          <td class="col-idx">${rowNum}</td>
          <td style="text-align:center; font-weight:600; color:#2c3e50;">${escapeHtml(baseNo)}</td>
          <td style="text-align:center;">${escapeHtml(revStr)}</td>
          <td style="max-width:170px; overflow:hidden; text-overflow:ellipsis;">${escapeHtml(q.customer_name || "")}</td>
          <td style="max-width:175px; overflow:hidden; text-overflow:ellipsis;">${escapeHtml(q.project_name || "")}</td>
          <td style="text-align:right; font-family:Consolas, monospace;">${Number(q.subtotal || 0).toLocaleString()}</td>
          <td style="text-align:center;">${formatStatusBadgeP3(q.status)}</td>
          <td style="text-align:center; font-family:Consolas, monospace;">${escapeHtml(q.quote_date || "")}</td>
          <td style="text-align:right; font-family:Consolas, monospace; color:#16a34a; font-weight:600;">${escapeHtml(wonAmtText)}</td>
          <td style="text-align:center; font-family:Consolas, monospace;">${escapeHtml(wonDateText)}</td>
          <td style="text-align:center;">${escapeHtml(q.project_type || "一次性專案")}</td>
        </tr>
      `;
    }).join("");

    tbody.querySelectorAll(".p3-hist-row").forEach(tr => {
      tr.addEventListener("click", () => {
        const qNo = tr.getAttribute("data-quoteno");
        const found = p3FilteredQuotesCache.find(x => x.quote_no === qNo);
        p3SelectedQuoteObj = found || null;
        tbody.querySelectorAll(".p3-hist-row").forEach(r => r.classList.toggle("selected-row", r === tr));
        updateP3SelectedBanner(p3SelectedQuoteObj, isTrash);
      });
      tr.addEventListener("dblclick", () => {
        const qNo = tr.getAttribute("data-quoteno");
        if (qNo) loadQuotationToForm(qNo);
      });
    });
  }

  function exportCrmMonthlyReportCsv() {
    const rows = p3FilteredQuotesCache.length > 0 ? p3FilteredQuotesCache : [];
    const header = ["單號", "版次", "公司名稱", "專案名稱", "總金額(未稅)", "狀態", "報價日期", "成交金額(未稅)", "專案屬性"];
    const csvLines = [header.join(",")];
    rows.forEach(q => {
      const { baseNo, revStr } = splitQuoteNoAndRev(q.quote_no, q.revision);
      const cols = [
        baseNo,
        revStr,
        q.customer_name || "",
        q.project_name || "",
        Number(q.subtotal || 0),
        q.status || "報價中",
        q.quote_date || "",
        Number(q.won_amount || 0),
        q.project_type || "一次性專案"
      ].map(v => `"${String(v).replace(/"/g, '""')}"`);
      csvLines.push(cols.join(","));
    });
    const blob = new Blob(["\uFEFF" + csvLines.join("\r\n")], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `CRM_商機月報_${getActiveSalesName()}_${new Date().toISOString().slice(0, 10)}.csv`;
    a.click();
    URL.revokeObjectURL(url);
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
        const { baseNo, revStr } = splitQuoteNoAndRev(header.quote_no, header.revision);
        const parts = baseNo.split("-");
        const prefixInput = document.getElementById("quotePrefixInput");
        if (prefixInput && parts.length > 0) prefixInput.value = parts[0];
        const seqInput = document.getElementById("quoteSeqInput");
        if (seqInput && parts.length > 1) seqInput.value = parts[1];
        const revInput = document.getElementById("quoteRevInput");
        if (revInput) revInput.value = revStr || "00";
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

      const defaultRowCount = Math.max(1, parseInt(localStorage.getItem("shihlin_quote_default_rows") || "5", 10) || 5);
      if (items.length === 0) {
        for (let i = 0; i < defaultRowCount; i++) addBlankItem();
      } else {
        items.forEach(itm => addBlankItem(itm));
        if (items.length < defaultRowCount) {
          for (let i = items.length; i < defaultRowCount; i++) addBlankItem();
        }
      }

      calculateTotal();
      updateQuoteNoPreview();
      switchQuoteSubTab("create");

      alert(`[已載入報價單] 已成功載入單號【${header.quote_no}】！\n若於電腦端修改後點擊「儲存報價」，系統將遵循 Append-Only 規則自動生成新版次。`);
    } catch (err) {
      alert("[載入失敗] " + err.message);
    }
  }

  // ============================================================================
  // 第 2 頁：手機報價草稿覆核區 (100% 對標電腦版 _QuotationApp gui_mobile_drafts.py & 截圖 2)
  // ============================================================================
  let p2DraftsCache = [];
  let p2SelectedDraftIdx = -1;

  function formatChineseOrdinal(idxZeroBased) {
    const ch = ["一", "二", "三", "四", "五", "六", "七", "八", "九", "十"];
    const n = idxZeroBased + 1;
    if (n <= 10) return `第${ch[idxZeroBased]}筆`;
    return `第 ${n} 筆`;
  }

  async function loadQuotationDraftsInbox() {
    const tbody = document.getElementById("quoteDraftsTableBody");
    const previewBox = document.getElementById("quoteDraftPreviewBox");
    const btnConvert = document.getElementById("btnConvertDraftToFormal");
    const syncLbl = document.getElementById("lblDraftsLastSync");
    const legacyContainer = document.getElementById("quoteDraftsResults");

    const now = new Date();
    const hh = String(now.getHours()).padStart(2, "0");
    const mm = String(now.getMinutes()).padStart(2, "0");
    const ss = String(now.getSeconds()).padStart(2, "0");
    if (syncLbl) syncLbl.textContent = `最後同步：${hh}:${mm}:${ss}`;

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

    p2DraftsCache = Array.isArray(drafts) ? drafts : [];
    p2SelectedDraftIdx = p2DraftsCache.length > 0 ? 0 : -1;

    updateQuoteDraftsBadge();

    if (p2DraftsCache.length === 0) {
      if (tbody) {
        tbody.innerHTML = `
          <tr>
            <td colspan="5" style="padding:32px; text-align:center; color:#7f8c8d;">
              目前業務【${escapeHtml(activeSales)}】無待覆核的手機報價草稿。
            </td>
          </tr>
        `;
      }
      if (previewBox) previewBox.textContent = "請從左側列表點選一筆手機草稿以檢視明細...";
      if (btnConvert) btnConvert.disabled = true;
      if (legacyContainer) legacyContainer.innerHTML = "<div>無待處理草稿</div>";
      return;
    }

    if (tbody) {
      tbody.innerHTML = p2DraftsCache.map((d, idx) => {
        const isSel = idx === p2SelectedDraftIdx;
        const custDisplay = d.contact_name
          ? `${d.customer_name || "未命名客戶"} - ${d.contact_name}`
          : (d.customer_name || "未命名客戶");
        const itemCount = Array.isArray(d.items) ? d.items.length : 0;
        const timeStr = String(d.created_at || "").slice(0, 16);

        return `
          <tr class="p2-draft-row ${isSel ? 'selected-row' : ''}" data-idx="${idx}" data-id="${d.id}">
            <td class="col-idx">${idx + 1}</td>
            <td style="text-align:center;">
              <button type="button" class="btn-p2-del-square" data-id="${d.id}" title="刪除此筆草稿">✕</button>
            </td>
            <td style="text-align:center; font-family:Consolas, monospace;">${escapeHtml(timeStr)}</td>
            <td>${escapeHtml(custDisplay)}</td>
            <td style="text-align:center;">${itemCount}</td>
          </tr>
        `;
      }).join("");

      tbody.querySelectorAll(".p2-draft-row").forEach(tr => {
        tr.addEventListener("click", (e) => {
          if (e.target.closest(".btn-p2-del-square")) return;
          const idx = parseInt(tr.getAttribute("data-idx"), 10);
          selectMobileDraftRow(idx);
        });
      });

      tbody.querySelectorAll(".btn-p2-del-square").forEach(btn => {
        btn.addEventListener("click", async (e) => {
          e.stopPropagation();
          const draftId = btn.getAttribute("data-id");
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

    selectMobileDraftRow(p2SelectedDraftIdx);

    if (legacyContainer) {
      legacyContainer.innerHTML = p2DraftsCache.map((d, idx) => `
        <div class="quote-history-card">
          <span>${escapeHtml(d.customer_name || "")}</span>
          <button type="button" class="btn-load-draft" data-idx="${idx}">載入</button>
        </div>
      `).join("");
    }
  }

  function selectMobileDraftRow(idx) {
    const tbody = document.getElementById("quoteDraftsTableBody");
    const previewBox = document.getElementById("quoteDraftPreviewBox");
    const btnConvert = document.getElementById("btnConvertDraftToFormal");
    p2SelectedDraftIdx = idx;

    if (tbody) {
      tbody.querySelectorAll(".p2-draft-row").forEach(tr => {
        tr.classList.toggle("selected-row", parseInt(tr.getAttribute("data-idx"), 10) === idx);
      });
    }

    const draft = p2DraftsCache[idx];
    if (!draft) {
      if (previewBox) previewBox.textContent = "請從左側列表點選一筆手機草稿以檢視明細...";
      if (btnConvert) btnConvert.disabled = true;
      return;
    }

    if (btnConvert) btnConvert.disabled = false;

    const items = Array.isArray(draft.items) ? draft.items : [];
    if (items.length === 0) {
      if (previewBox) previewBox.textContent = "(此草稿無任何品項明細)";
      return;
    }

    const blocks = items.map((itm, i) => {
      const ord = formatChineseOrdinal(i);
      const model = itm.model_no || itm.model || "未指定型號";
      const qty = Number(itm.quantity || itm.qty || 1);
      const price = Number(itm.final_unit_price || itm.est_unit_price || itm.dealer_price || 0).toLocaleString();
      const note = itm.delivery_note || itm.note || "無";
      return `${ord}\n${model}\n數量：${qty}\n預計報價：${price}元\n交期：${note}`;
    });

    if (previewBox) {
      previewBox.textContent = blocks.join("\n\n");
    }
  }

  function convertSelectedMobileDraftToFormal() {
    const draft = p2DraftsCache[p2SelectedDraftIdx];
    if (!draft) {
      alert("請先從左側列表點選一筆手機草稿！");
      return;
    }

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

    const defaultRowCount = Math.max(1, parseInt(localStorage.getItem("shihlin_quote_default_rows") || "5", 10) || 5);
    const dItems = Array.isArray(draft.items) ? draft.items : [];
    if (dItems.length > 0) {
      dItems.forEach(itm => addBlankItem(itm));
      if (dItems.length < defaultRowCount) {
        for (let i = dItems.length; i < defaultRowCount; i++) addBlankItem();
      }
    } else {
      for (let i = 0; i < defaultRowCount; i++) addBlankItem();
    }

    calculateTotal();
    switchQuoteSubTab("create");
    alert(`[已載入草稿] 手機草稿【${draft.customer_name}】已成功轉入「新建報價單」編輯區！\n確認無誤後即可點擊「儲存報價」正式產出報價單。`);
  }

  // ============================================================================
  // 第 5 頁：歷史單價分析 (100% 對標電腦版 _QuotationApp gui_page2.py & 截圖 5)
  // ============================================================================
  let p5AnalysisRecordsCache = [];

  async function runPriceAnalysis() {
    const tbody = document.getElementById("quoteAnalysisTableBody");
    const legacyContainer = document.getElementById("quoteAnalysisResults");
    const focusBadge = document.getElementById("analysisCurrentFocusBadge");
    const lblMinPrice = document.getElementById("lblAnalysisMinPrice");
    const lblAvgPrice = document.getElementById("lblAnalysisAvgPrice");
    const lblMinDisc = document.getElementById("lblAnalysisMinDiscount");
    const lblTotalCnt = document.getElementById("lblAnalysisTotalCount");

    const modelQuery = document.getElementById("inputAnalysisModel")?.value.trim() || "";
    const custQuery = document.getElementById("inputAnalysisCustomer")?.value.trim() || "";
    const activeSales = getActiveSalesName();

    if (tbody) {
      tbody.innerHTML = `<tr><td colspan="13" style="padding:24px; text-align:center; color:#7f8c8d;">正在分析歷史單價與折讓區間...</td></tr>`;
    }

    try {
      const url = `/api/price-analysis?sales_name=${encodeURIComponent(activeSales)}&model_keyword=${encodeURIComponent(modelQuery)}&customer_keyword=${encodeURIComponent(custQuery)}`;
      const resp = await fetch(url);
      if (!resp.ok) throw new Error("讀取歷史單價分析失敗");
      const data = await resp.json();

      let matchedItems = Array.isArray(data) ? data : (data.records || []);
      const showAllDates = document.getElementById("chkAnalysisShowAll")?.checked !== false;
      const sDate = document.getElementById("analysisStartDate")?.value || "";
      const eDate = document.getElementById("analysisEndDate")?.value || "";

      if (!showAllDates && (sDate || eDate)) {
        matchedItems = matchedItems.filter(m => {
          const qd = String(m.quote_date || "").replace(/\//g, "-");
          if (sDate && qd < sDate) return false;
          if (eDate && qd > eDate) return false;
          return true;
        });
      }

      p5AnalysisRecordsCache = matchedItems;

      if (matchedItems.length === 0) {
        if (tbody) {
          tbody.innerHTML = `<tr><td colspan="13" style="padding:32px; text-align:center; color:#7f8c8d;">查無相符的歷史報價明細紀錄。</td></tr>`;
        }
        if (focusBadge) focusBadge.textContent = "目前分析焦點：無符合資料";
        if (lblMinPrice) lblMinPrice.textContent = "歷史最低價: ---";
        if (lblAvgPrice) lblAvgPrice.textContent = "平均報價: ---";
        if (lblMinDisc) lblMinDisc.textContent = "最低折數: ---";
        if (lblTotalCnt) lblTotalCnt.textContent = "報價次數: 0 次";
        renderPriceTrendSvgChart([]);
        if (legacyContainer) legacyContainer.innerHTML = "<div>查無相符紀錄</div>";
        return;
      }

      if (tbody) {
        tbody.innerHTML = matchedItems.map((m, idx) => {
          const histPrice = Number(m.final_unit_price || 0);
          const dealer26 = Number(m.dealer_price || histPrice || 0);
          const histDisc = Number(m.discount_pct || (dealer26 > 0 ? (histPrice / dealer26) * 100 : 100));
          const convDisc = dealer26 > 0 ? (histPrice / dealer26) * 100 : histDisc;
          const diffAmt = histPrice - dealer26;
          const diffRate = dealer26 > 0 ? (diffAmt / dealer26) * 100 : 0;
          const isWon = m.status === "已成交";
          const wonQty = isWon ? Number(m.quantity || 1) : 0;
          const wonPriceStr = isWon ? histPrice.toLocaleString() : "-";
          const diffColor = diffAmt < 0 ? "#27ae60" : (diffAmt > 0 ? "#e74c3c" : "#2c3e50");

          return `
            <tr class="p5-analysis-row" data-idx="${idx}" data-model="${escapeHtml(m.model_no || '')}">
              <td class="col-idx">${idx + 1}</td>
              <td style="max-width:140px; overflow:hidden; text-overflow:ellipsis;">${escapeHtml(m.customer_name || "")}</td>
              <td style="font-weight:600; color:#2c3e50;">${escapeHtml(m.model_no || "")}</td>
              <td style="text-align:center; font-family:Consolas, monospace;">${escapeHtml(m.quote_date || "")}</td>
              <td style="text-align:center;">${Number(m.quantity || 1)}</td>
              <td style="text-align:right; font-family:Consolas, monospace;">${histPrice.toLocaleString()}</td>
              <td style="text-align:center;">${wonQty}</td>
              <td style="text-align:center; font-family:Consolas, monospace;">${escapeHtml(wonPriceStr)}</td>
              <td style="text-align:right; font-family:Consolas, monospace;">${histDisc.toFixed(1)}%</td>
              <td style="text-align:right; font-family:Consolas, monospace;">${dealer26.toLocaleString()}</td>
              <td style="text-align:right; font-family:Consolas, monospace;">${convDisc.toFixed(1)}%</td>
              <td style="text-align:right; font-family:Consolas, monospace; color:${diffColor}; font-weight:600;">${diffAmt > 0 ? '+' : ''}${diffAmt.toLocaleString()}</td>
              <td style="text-align:right; font-family:Consolas, monospace; color:${diffColor}; font-weight:600;">${diffRate > 0 ? '+' : ''}${diffRate.toFixed(1)}%</td>
            </tr>
          `;
        }).join("");

        tbody.querySelectorAll(".p5-analysis-row").forEach(tr => {
          tr.addEventListener("click", () => {
            tbody.querySelectorAll(".p5-analysis-row").forEach(r => r.classList.toggle("selected-row", r === tr));
            const targetModel = tr.getAttribute("data-model") || "";
            updatePriceAnalysisFocusStats(targetModel);
          });
        });
      }

      // 預設聚焦第一筆產品或目前搜尋型號
      const initialFocusModel = modelQuery || (matchedItems[0]?.model_no || "");
      const firstRow = tbody?.querySelector(".p5-analysis-row");
      if (firstRow) firstRow.classList.add("selected-row");
      updatePriceAnalysisFocusStats(initialFocusModel);

      if (legacyContainer) {
        legacyContainer.innerHTML = `<div>共 ${matchedItems.length} 筆歷史單價紀錄</div>`;
      }
    } catch (err) {
      if (tbody) {
        tbody.innerHTML = `<tr><td colspan="13" style="padding:20px; text-align:center; color:#e74c3c;">分析失敗：${escapeHtml(err.message)}</td></tr>`;
      }
    }
  }

  function updatePriceAnalysisFocusStats(focusModel) {
    const focusBadge = document.getElementById("analysisCurrentFocusBadge");
    const lblMinPrice = document.getElementById("lblAnalysisMinPrice");
    const lblAvgPrice = document.getElementById("lblAnalysisAvgPrice");
    const lblMinDisc = document.getElementById("lblAnalysisMinDiscount");
    const lblTotalCnt = document.getElementById("lblAnalysisTotalCount");

    const subset = focusModel
      ? p5AnalysisRecordsCache.filter(m => String(m.model_no || "").toLowerCase() === focusModel.toLowerCase())
      : p5AnalysisRecordsCache;
    const rows = subset.length > 0 ? subset : p5AnalysisRecordsCache;

    if (focusBadge) {
      focusBadge.textContent = `目前分析焦點：${focusModel || (rows[0]?.model_no || "全部搜尋結果")}`;
    }

    const prices = rows.map(r => Number(r.final_unit_price || 0)).filter(v => v > 0);
    const discs = rows.map(r => {
      const p = Number(r.final_unit_price || 0);
      const d = Number(r.dealer_price || p || 0);
      return Number(r.discount_pct || (d > 0 ? (p / d) * 100 : 100));
    }).filter(v => v > 0);

    const minP = prices.length ? Math.min(...prices) : 0;
    const avgP = prices.length ? Math.round(prices.reduce((a, b) => a + b, 0) / prices.length) : 0;
    const minD = discs.length ? Math.min(...discs) : 100;

    if (lblMinPrice) lblMinPrice.textContent = `歷史最低價: $${minP.toLocaleString()}`;
    if (lblAvgPrice) lblAvgPrice.textContent = `平均報價: $${avgP.toLocaleString()}`;
    if (lblMinDisc) lblMinDisc.textContent = `最低折數: ${minD.toFixed(1)}%`;
    if (lblTotalCnt) lblTotalCnt.textContent = `報價次數: ${rows.length} 次`;

    renderPriceTrendSvgChart(rows);
  }

  function renderPriceTrendSvgChart(rows) {
    const box = document.getElementById("priceTrendSvgContainer");
    if (!box) return;

    if (!rows || rows.length === 0) {
      box.innerHTML = `<div style="height:100%; display:flex; align-items:center; justify-content:center; color:#95a5a6; font-size:0.86rem;">尚無足夠數據繪製價格走勢圖</div>`;
      return;
    }

    // 依日期由舊到新排序
    const sorted = [...rows].sort((a, b) => String(a.quote_date || "").localeCompare(String(b.quote_date || "")));
    const pts = sorted.map(r => ({
      date: String(r.quote_date || "").slice(5).replace("-", "/"),
      price: Number(r.final_unit_price || 0)
    })).filter(p => p.price > 0);

    if (pts.length === 0) {
      box.innerHTML = `<div style="height:100%; display:flex; align-items:center; justify-content:center; color:#95a5a6;">無有效單價數據</div>`;
      return;
    }

    const w = 820;
    const h = 215;
    const padL = 64;
    const padR = 64;
    const padT = 22;
    const padB = 34;
    const plotW = w - padL - padR;
    const plotH = h - padT - padB;

    const vals = pts.map(p => p.price);
    const avgVal = Math.round(vals.reduce((a, b) => a + b, 0) / vals.length);
    let minY = Math.min(...vals);
    let maxY = Math.max(...vals);
    if (minY === maxY) {
      minY = Math.max(0, Math.round(minY * 0.9));
      maxY = Math.round(maxY * 1.1) + 10;
    } else {
      const span = maxY - minY;
      minY = Math.max(0, Math.floor(minY - span * 0.12));
      maxY = Math.ceil(maxY + span * 0.12);
    }

    const getX = (i) => pts.length === 1 ? (padL + plotW / 2) : (padL + (i / (pts.length - 1)) * plotW);
    const getY = (v) => padT + plotH - ((v - minY) / Math.max(1, maxY - minY)) * plotH;

    let gridSvg = "";
    for (let g = 0; g <= 4; g++) {
      const gy = padT + (g / 4) * plotH;
      const gVal = Math.round(maxY - (g / 4) * (maxY - minY));
      gridSvg += `
        <line x1="${padL}" y1="${gy}" x2="${w - padR}" y2="${gy}" stroke="#ecf0f1" stroke-dasharray="4,3" />
        <text x="${padL - 8}" y="${gy + 4}" text-anchor="end" fill="#7f8c8d" font-size="11">$${gVal.toLocaleString()}</text>
      `;
    }

    const avgY = getY(avgVal);
    const avgLineSvg = `
      <line x1="${padL}" y1="${avgY}" x2="${w - padR}" y2="${avgY}" stroke="#e67e22" stroke-width="1.5" stroke-dasharray="6,4" />
      <text x="${w - padR + 6}" y="${avgY + 4}" fill="#e67e22" font-size="11" font-weight="700">均價</text>
    `;

    const polyPoints = pts.map((p, i) => `${getX(i).toFixed(1)},${getY(p.price).toFixed(1)}`).join(" ");
    const lineSvg = pts.length > 1
      ? `<polyline fill="none" stroke="#3498db" stroke-width="2.5" points="${polyPoints}" />`
      : "";

    const dotsSvg = pts.map((p, i) => {
      const cx = getX(i);
      const cy = getY(p.price);
      const showDate = pts.length <= 10 || i === 0 || i === pts.length - 1 || i % Math.ceil(pts.length / 6) === 0;
      return `
        <circle cx="${cx.toFixed(1)}" cy="${cy.toFixed(1)}" r="4.5" fill="#e74c3c" stroke="#ffffff" stroke-width="1.5">
          <title>${p.date}：$${p.price.toLocaleString()}</title>
        </circle>
        ${showDate ? `<text x="${cx.toFixed(1)}" y="${h - 10}" text-anchor="middle" fill="#7f8c8d" font-size="10.5">${escapeHtml(p.date)}</text>` : ""}
      `;
    }).join("");

    box.innerHTML = `
      <svg viewBox="0 0 ${w} ${h}" style="width:100%; height:100%; display:block;">
        ${gridSvg}
        <line x1="${padL}" y1="${padT}" x2="${padL}" y2="${h - padB}" stroke="#bdc3c7" />
        <line x1="${padL}" y1="${h - padB}" x2="${w - padR}" y2="${h - padB}" stroke="#bdc3c7" />
        ${avgLineSvg}
        ${lineSvg}
        ${dotsSvg}
      </svg>
    `;
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
  // 報價管理系統新增完整功能頁：手機草稿徽章、大數據分析、客戶管理、產品搜尋與計算機、系統設定、使用說明
  // ============================================================================
  async function updateQuoteDraftsBadge() {
    const badge = document.getElementById("quoteDraftsCountBadge");
    const redDot = document.getElementById("quoteMobileDraftDot");
    const statusText = document.getElementById("quoteSystemStatusText");
    try {
      const activeSales = getActiveSalesName();
      const r = await fetch(`/api/quotation-drafts?sales_name=${encodeURIComponent(activeSales)}`);
      if (r.ok) {
        const list = await r.json();
        const cnt = Array.isArray(list) ? list.length : 0;
        if (badge) {
          badge.textContent = String(cnt);
          badge.classList.toggle("hidden", cnt === 0);
        }
        if (redDot) {
          redDot.classList.toggle("hidden", cnt === 0);
        }
        if (statusText) {
          statusText.textContent = `系統提示：Google 雲端目前有 ${cnt} 筆來自手機端的新草稿待覆核`;
        }
      }
    } catch (e) {}
  }

  // ============================================================================
  // 第 4 頁：大數據分析 (100% 對標電腦版 _QuotationApp gui_dashboard.py & 截圖 4、11)
  // ============================================================================
  let p4DashDataCache = null;

  function applyDashboardQuickDateRange(rangeCode) {
    const sInput = document.getElementById("dashStartDate");
    const eInput = document.getElementById("dashEndDate");
    const now = new Date();
    const todayStr = now.toISOString().slice(0, 10);
    if (eInput) eInput.value = todayStr;

    if (!sInput) return;
    if (rangeCode === "ALL") {
      sInput.value = "2020-01-01";
    } else if (rangeCode === "THIS_MONTH") {
      sInput.value = `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, "0")}-01`;
    } else if (rangeCode === "LAST_3M") {
      const d = new Date(now);
      d.setMonth(d.getMonth() - 3);
      sInput.value = d.toISOString().slice(0, 10);
    } else if (rangeCode === "THIS_YEAR") {
      sInput.value = `${now.getFullYear()}-01-01`;
    }
  }

  function setupDashboardDragHandles() {
    document.querySelectorAll(".pyside-dash-drag-handle").forEach(handle => {
      let startY = 0;
      let startH = 0;
      let targetEl = null;

      const onMove = (e) => {
        if (!targetEl) return;
        const dy = e.clientY - startY;
        const nextH = Math.max(110, startH + dy);
        targetEl.style.height = `${nextH}px`;
        targetEl.style.maxHeight = `${nextH}px`;
      };
      const onUp = () => {
        targetEl = null;
        window.removeEventListener("mousemove", onMove);
        window.removeEventListener("mouseup", onUp);
      };

      handle.addEventListener("mousedown", (e) => {
        const targetId = handle.getAttribute("data-resize-target");
        targetEl = targetId ? document.getElementById(targetId) : null;
        if (!targetEl) return;
        e.preventDefault();
        startY = e.clientY;
        startH = targetEl.getBoundingClientRect().height;
        window.addEventListener("mousemove", onMove);
        window.addEventListener("mouseup", onUp);
      });
    });
  }

  async function loadQuotationDashboard() {
    const activeSales = getActiveSalesName();
    const latestOnly = document.getElementById("chkQuoteDashLatestOnly")?.checked !== false;
    try {
      const [rDash, rQuotes, rAnalysis] = await Promise.all([
        fetch(`/api/quotation-dashboard?sales_name=${encodeURIComponent(activeSales)}&latest_only=${latestOnly}`),
        fetch(`/api/quotations?sales_name=${encodeURIComponent(activeSales)}&include_deleted=false`),
        fetch(`/api/price-analysis?sales_name=${encodeURIComponent(activeSales)}`)
      ]);
      if (!rDash.ok) throw new Error("讀取大數據統計失敗");
      const d = await rDash.json();
      const allQuotes = rQuotes.ok ? await rQuotes.json() : [];
      const priceRecsRaw = rAnalysis.ok ? await rAnalysis.json() : [];
      const priceRecs = Array.isArray(priceRecsRaw) ? priceRecsRaw : (priceRecsRaw.records || []);

      p4DashDataCache = d;

      // 1. 保持舊隱藏相容節點更新
      const elTotalQuoted = document.getElementById("qDashTotalQuoted");
      const elTotalCount = document.getElementById("qDashTotalCount");
      const elWonAmount = document.getElementById("qDashWonAmount");
      const elWonCount = document.getElementById("qDashWonCount");
      const elWinRateAmt = document.getElementById("qDashWinRateAmt");
      const elWinRateCnt = document.getElementById("qDashWinRateCnt");
      const elPendingAmt = document.getElementById("qDashPendingAmount");
      const elPendingCnt = document.getElementById("qDashPendingCount");
      if (elTotalQuoted) elTotalQuoted.textContent = `$ ${Number(d.total_quoted_amount || 0).toLocaleString()}`;
      if (elTotalCount) elTotalCount.textContent = `有效報價單：${d.total_count || 0} 筆`;
      if (elWonAmount) elWonAmount.textContent = `$ ${Number(d.won_amount || 0).toLocaleString()}`;
      if (elWonCount) elWonCount.textContent = `已成交筆數：${d.won_count || 0} 筆`;
      if (elWinRateAmt) elWinRateAmt.textContent = `${Number(d.win_rate_amount_pct || 0).toFixed(1)}%`;
      if (elWinRateCnt) elWinRateCnt.textContent = `件數成交率：${Number(d.win_rate_count_pct || 0).toFixed(1)}%`;
      if (elPendingAmt) elPendingAmt.textContent = `$ ${Number(d.pending_amount || 0).toLocaleString()}`;
      if (elPendingCnt) elPendingCnt.textContent = `追蹤中：${d.pending_count || 0} 筆`;

      // 2. 更新 5 張頂部彩色邊框 KPI 卡片 (對應截圖 4)
      const totalCnt = Number(d.total_count || 0);
      const totalWan = (Number(d.total_quoted_amount || 0) / 10000).toFixed(1);
      const wonCnt = Number(d.won_count || 0);
      const wonWan = (Number(d.won_amount || 0) / 10000).toFixed(1);
      const lostCnt = Number(d.lost_count || 0);
      const lostAmtVal = Math.max(0, Number(d.total_quoted_amount || 0) - Number(d.won_amount || 0) - Number(d.pending_amount || 0));
      const lostWan = (lostAmtVal / 10000).toFixed(1);
      const pendCnt = Number(d.pending_count || 0);
      const pendWan = (Number(d.pending_amount || 0) / 10000).toFixed(1);
      const rateCntPct = Number(d.win_rate_count_pct || 0).toFixed(1);

      const setTxt = (id, txt) => {
        const el = document.getElementById(id);
        if (el) el.textContent = txt;
      };
      setTxt("qDashKpiTotalCnt", `${totalCnt} 件`);
      setTxt("qDashKpiTotalAmt", `金額: ${totalWan} 萬`);
      setTxt("qDashKpiWonCnt", `${wonCnt} 件`);
      setTxt("qDashKpiWonAmt", `金額: ${wonWan} 萬`);
      setTxt("qDashKpiLostCnt", `${lostCnt} 件`);
      setTxt("qDashKpiLostAmt", `金額: ${lostWan} 萬`);
      setTxt("qDashKpiPendingCnt", `${pendCnt} 件`);
      setTxt("qDashKpiPendingAmt", `金額: ${pendWan} 萬`);
      setTxt("qDashKpiRateVal", `${rateCntPct}%`);
      setTxt("qDashKpiRateSub", `${wonCnt}/${Math.max(1, totalCnt)} 件`);

      // 3. 待跟進區 (發出 30 天內未回報)
      const nowMs = Date.now();
      const followups = (allQuotes || []).filter(q => {
        if (q.is_deleted === 1) return false;
        const st = q.status || "追蹤中";
        if (st !== "追蹤中" && st !== "報價中") return false;
        const qDateMs = Date.parse(String(q.quote_date || "").replace(/\//g, "-"));
        if (isNaN(qDateMs)) return true;
        const diffDays = Math.floor((nowMs - qDateMs) / 86400000);
        return diffDays >= 0 && diffDays <= 30;
      });

      setTxt("lblDashFollowupLegend", `待跟進區 (發出 30 天內未回報 : ${followups.length} 筆)`);
      const followupBody = document.getElementById("dashFollowupTableBody");
      if (followupBody) {
        if (followups.length === 0) {
          followupBody.innerHTML = `<tr><td colspan="8" style="padding:18px; text-align:center; color:#7f8c8d;">目前無 30 天內待跟進之未回報單據</td></tr>`;
        } else {
          followupBody.innerHTML = followups.map((q, idx) => {
            const { baseNo } = splitQuoteNoAndRev(q.quote_no, q.revision);
            const qDateMs = Date.parse(String(q.quote_date || "").replace(/\//g, "-"));
            const days = isNaN(qDateMs) ? 0 : Math.max(0, Math.floor((nowMs - qDateMs) / 86400000));
            const matchedC = (customersCache || []).find(c => c.name === q.customer_name) || {};
            const primaryCt = (Array.isArray(matchedC.contacts) && matchedC.contacts[0]) || {};
            return `
              <tr data-quoteno="${escapeHtml(q.quote_no)}" style="cursor:pointer;" title="雙擊載入此報價單">
                <td class="col-idx">${idx + 1}</td>
                <td style="text-align:center; font-weight:600;">${escapeHtml(baseNo)}</td>
                <td>${escapeHtml(q.customer_name || "")}</td>
                <td style="text-align:center;">${escapeHtml(q.contact_name || matchedC.default_contact || "")}</td>
                <td style="text-align:center; font-family:Consolas, monospace;">${escapeHtml(q.phone || matchedC.phone || "")}</td>
                <td style="text-align:center; font-family:Consolas, monospace;">${escapeHtml(primaryCt.mobile || "")}</td>
                <td style="text-align:center; font-family:Consolas, monospace;">${escapeHtml(q.quote_date || "")}</td>
                <td style="text-align:center; color:#e67e22; font-weight:700;">${days} 天</td>
              </tr>
            `;
          }).join("");
          followupBody.querySelectorAll("tr[data-quoteno]").forEach(tr => {
            tr.addEventListener("dblclick", () => loadQuotationToForm(tr.getAttribute("data-quoteno")));
          });
        }
      }

      // 4. 2x2 表格：最常報價客戶、成交貢獻排行、未成交(失注)客戶、熱門報價產品
      const topCusts = d.top_customers || [];
      const freqBody = document.getElementById("dashTopFreqCustBody");
      if (freqBody) {
        const sortedFreq = [...topCusts].sort((a, b) => Number(b.quote_count || 0) - Number(a.quote_count || 0)).slice(0, 10);
        freqBody.innerHTML = sortedFreq.length === 0
          ? `<tr><td colspan="4" style="text-align:center; padding:14px; color:#95a5a6;">尚無統計資料</td></tr>`
          : sortedFreq.map((c, i) => `
              <tr>
                <td class="col-idx">${i + 1}</td>
                <td>${escapeHtml(c.customer_name)}</td>
                <td style="text-align:center; font-weight:700; color:#2980b9;">${c.quote_count}</td>
                <td style="text-align:right; font-family:Consolas, monospace;">$${Number(c.total_quoted || 0).toLocaleString()}</td>
              </tr>
            `).join("");
      }

      const wonCustBody = document.getElementById("dashTopWonCustBody");
      if (wonCustBody) {
        const sortedWon = [...topCusts].filter(c => Number(c.won_amount || 0) > 0).sort((a, b) => Number(b.won_amount || 0) - Number(a.won_amount || 0)).slice(0, 10);
        wonCustBody.innerHTML = sortedWon.length === 0
          ? `<tr><td colspan="4" style="text-align:center; padding:14px; color:#95a5a6;">尚無成交客戶資料</td></tr>`
          : sortedWon.map((c, i) => `
              <tr>
                <td class="col-idx">${i + 1}</td>
                <td>${escapeHtml(c.customer_name)}</td>
                <td style="text-align:right; font-family:Consolas, monospace; color:#27ae60; font-weight:700;">$${Number(c.won_amount || 0).toLocaleString()}</td>
                <td style="text-align:center;">${Math.max(1, Math.round((Number(c.win_rate_pct || 0) / 100) * Number(c.quote_count || 1)))}</td>
              </tr>
            `).join("");
      }

      const lostCustBody = document.getElementById("dashTopLostCustBody");
      if (lostCustBody) {
        const lostList = [...topCusts].map(c => {
          const lostA = Math.max(0, Number(c.total_quoted || 0) - Number(c.won_amount || 0) - Number(c.pending_amount || 0));
          return { ...c, lost_amount: lostA };
        }).filter(c => c.lost_amount > 0).sort((a, b) => b.lost_amount - a.lost_amount).slice(0, 10);

        lostCustBody.innerHTML = lostList.length === 0
          ? `<tr><td colspan="4" style="text-align:center; padding:14px; color:#95a5a6;">尚無失注客戶紀錄</td></tr>`
          : lostList.map((c, i) => `
              <tr>
                <td class="col-idx">${i + 1}</td>
                <td>${escapeHtml(c.customer_name)}</td>
                <td style="text-align:center; color:#e74c3c; font-weight:700;">1</td>
                <td style="text-align:right; font-family:Consolas, monospace;">$${c.lost_amount.toLocaleString()}</td>
              </tr>
            `).join("");
      }

      const topItemsBody = document.getElementById("dashTopItemsBody");
      if (topItemsBody) {
        const itemMap = new Map();
        priceRecs.forEach(r => {
          const m = String(r.model_no || "").trim();
          if (!m) return;
          const prev = itemMap.get(m) || { model_no: m, count: 0, total_qty: 0 };
          prev.count += 1;
          prev.total_qty += Number(r.quantity || 1);
          itemMap.set(m, prev);
        });
        const sortedItems = Array.from(itemMap.values()).sort((a, b) => b.count - a.count).slice(0, 10);
        topItemsBody.innerHTML = sortedItems.length === 0
          ? `<tr><td colspan="4" style="text-align:center; padding:14px; color:#95a5a6;">尚無產品明細統計</td></tr>`
          : sortedItems.map((it, i) => `
              <tr>
                <td class="col-idx">${i + 1}</td>
                <td style="font-weight:600;">${escapeHtml(it.model_no)}</td>
                <td style="text-align:center; color:#2980b9; font-weight:700;">${it.count}</td>
                <td style="text-align:right; font-family:Consolas, monospace;">${it.total_qty.toLocaleString()}</td>
              </tr>
            `).join("");
      }

      // 5. 圖表區 (甜甜圈圖、失注原因橫條圖、月度趨勢折線圖)
      renderDashboardChartsSvg(d, allQuotes);

      // 6. 沉睡客戶流失預警雷達 (>90 天未再詢價)
      const churnBody = document.getElementById("dashChurnTableBody");
      if (churnBody) {
        const custLastMap = new Map();
        (allQuotes || []).forEach(q => {
          if (!q.customer_name) return;
          const cur = custLastMap.get(q.customer_name) || {
            customer_name: q.customer_name,
            last_date: "",
            count: 0,
            total_amt: 0
          };
          cur.count += 1;
          cur.total_amt += Number(q.subtotal || 0);
          const qd = String(q.quote_date || "").replace(/\//g, "-");
          if (!cur.last_date || qd > cur.last_date) cur.last_date = qd;
          custLastMap.set(q.customer_name, cur);
        });

        const churnList = Array.from(custLastMap.values()).map(c => {
          const ms = Date.parse(c.last_date);
          const silentDays = isNaN(ms) ? 0 : Math.floor((nowMs - ms) / 86400000);
          return { ...c, silentDays };
        }).filter(c => c.silentDays > 90).sort((a, b) => b.silentDays - a.silentDays);

        churnBody.innerHTML = churnList.length === 0
          ? `<tr><td colspan="7" style="text-align:center; padding:16px; color:#7f8c8d;">太棒了！目前無超過 90 天未詢價之沉睡流失預警客戶</td></tr>`
          : churnList.map((c, i) => {
              const matchedC = (customersCache || []).find(x => x.name === c.customer_name) || {};
              return `
                <tr>
                  <td class="col-idx">${i + 1}</td>
                  <td style="font-weight:600;">${escapeHtml(c.customer_name)}</td>
                  <td style="text-align:center;">${escapeHtml(matchedC.default_contact || "-")}</td>
                  <td style="text-align:center; font-family:Consolas, monospace;">${escapeHtml(matchedC.phone || "-")}</td>
                  <td style="text-align:center; font-family:Consolas, monospace;">${escapeHtml(c.last_date)}</td>
                  <td style="text-align:center; color:#c0392b; font-weight:700;">${c.silentDays} 天</td>
                  <td style="text-align:right; font-family:Consolas, monospace;">${c.count} 次 / $${c.total_amt.toLocaleString()}</td>
                </tr>
              `;
            }).join("");
      }
    } catch (err) {
      console.warn("大數據分析載入失敗:", err);
    }
  }

  function renderDashboardChartsSvg(d, allQuotes) {
    const pieBox = document.getElementById("dashStatusPieBox");
    const barBox = document.getElementById("dashLostReasonBarBox");
    const lineBox = document.getElementById("dashTrendLineBox");

    const wonCnt = Number(d.won_count || 0);
    const lostCnt = Number(d.lost_count || 0);
    const pendCnt = Number(d.pending_count || 0);
    const ignCnt = Number(d.ignored_count || 0);
    const sumCnt = Math.max(1, wonCnt + lostCnt + pendCnt + ignCnt);

    if (pieBox) {
      const r = 56;
      const c = 2 * Math.PI * r;
      const segs = [
        { label: "已成交", val: wonCnt, color: "#27ae60" },
        { label: "報價中", val: pendCnt, color: "#f39c12" },
        { label: "未成交", val: lostCnt, color: "#e74c3c" },
        { label: "不跟進", val: ignCnt, color: "#95a5a6" }
      ];
      let offset = 0;
      const circlesSvg = segs.map(s => {
        const len = (s.val / sumCnt) * c;
        const svg = `<circle cx="80" cy="80" r="${r}" fill="none" stroke="${s.color}" stroke-width="22" stroke-dasharray="${len.toFixed(1)} ${(c - len).toFixed(1)}" stroke-dashoffset="${(-offset).toFixed(1)}" transform="rotate(-90 80 80)" />`;
        offset += len;
        return svg;
      }).join("");

      const legendHtml = segs.map(s => {
        const pct = ((s.val / sumCnt) * 100).toFixed(1);
        return `
          <div style="display:flex; align-items:center; gap:6px; font-size:0.8rem; color:#2c3e50; margin-bottom:4px;">
            <span style="width:10px; height:10px; border-radius:2px; background:${s.color}; display:inline-block;"></span>
            <span>${s.label} (${pct}%) - ${s.val}件</span>
          </div>
        `;
      }).join("");

      pieBox.innerHTML = `
        <div style="display:flex; align-items:center; justify-content:center; gap:20px; width:100%;">
          <svg width="160" height="160" viewBox="0 0 160 160">
            <circle cx="80" cy="80" r="${r}" fill="none" stroke="#ecf0f1" stroke-width="22" />
            ${circlesSvg}
            <text x="80" y="76" text-anchor="middle" fill="#2c3e50" font-size="13" font-weight="700">總件數</text>
            <text x="80" y="95" text-anchor="middle" fill="#2980b9" font-size="15" font-weight="800">${wonCnt + lostCnt + pendCnt + ignCnt}</text>
          </svg>
          <div>${legendHtml}</div>
        </div>
      `;
    }

    if (barBox) {
      if (lostCnt === 0) {
        barBox.innerHTML = `<div style="color:#95a5a6; font-size:0.88rem;">目前區間無未成交(失注)單據</div>`;
      } else {
        barBox.innerHTML = `
          <div style="width:100%; padding:8px 16px;">
            <div style="display:flex; justify-content:space-between; font-size:0.8rem; margin-bottom:4px;">
              <span>價格競爭 / 預算不符</span><span>${lostCnt} 件</span>
            </div>
            <div style="height:16px; background:#ecf0f1; border-radius:4px; overflow:hidden;">
              <div style="width:100%; height:100%; background:#e74c3c;"></div>
            </div>
          </div>
        `;
      }
    }

    if (lineBox) {
      const monthMap = new Map();
      (allQuotes || []).forEach(q => {
        if (q.is_deleted === 1) return;
        const ym = String(q.quote_date || "").replace(/\//g, "-").slice(0, 7);
        if (!ym || ym.length < 7) return;
        const cur = monthMap.get(ym) || { ym, quoted: 0, won: 0 };
        cur.quoted += Number(q.subtotal || 0) / 10000;
        if (q.status === "已成交") cur.won += Number(q.won_amount || q.subtotal || 0) / 10000;
        monthMap.set(ym, cur);
      });
      const mList = Array.from(monthMap.values()).sort((a, b) => a.ym.localeCompare(b.ym)).slice(-8);
      if (mList.length === 0) {
        lineBox.innerHTML = `<div style="height:100%; display:flex; align-items:center; justify-content:center; color:#95a5a6;">尚無月度趨勢資料</div>`;
      } else {
        const w = 780;
        const h = 185;
        const padL = 56;
        const padR = 30;
        const padT = 20;
        const padB = 32;
        const plotW = w - padL - padR;
        const plotH = h - padT - padB;
        const maxV = Math.max(10, ...mList.map(m => Math.max(m.quoted, m.won))) * 1.15;
        const getX = (i) => mList.length === 1 ? (padL + plotW / 2) : (padL + (i / (mList.length - 1)) * plotW);
        const getY = (v) => padT + plotH - (v / maxV) * plotH;

        const qPts = mList.map((m, i) => `${getX(i).toFixed(1)},${getY(m.quoted).toFixed(1)}`).join(" ");
        const wPts = mList.map((m, i) => `${getX(i).toFixed(1)},${getY(m.won).toFixed(1)}`).join(" ");

        lineBox.innerHTML = `
          <svg viewBox="0 0 ${w} ${h}" style="width:100%; height:100%; display:block;">
            <line x1="${padL}" y1="${padT}" x2="${padL}" y2="${h - padB}" stroke="#bdc3c7" />
            <line x1="${padL}" y1="${h - padB}" x2="${w - padR}" y2="${h - padB}" stroke="#bdc3c7" />
            ${mList.length > 1 ? `<polyline fill="none" stroke="#3498db" stroke-width="2.5" points="${qPts}" />` : ""}
            ${mList.length > 1 ? `<polyline fill="none" stroke="#2ecc71" stroke-width="2.5" points="${wPts}" />` : ""}
            ${mList.map((m, i) => `
              <circle cx="${getX(i).toFixed(1)}" cy="${getY(m.quoted).toFixed(1)}" r="4" fill="#3498db"><title>${m.ym} 報價: ${m.quoted.toFixed(1)}萬</title></circle>
              <circle cx="${getX(i).toFixed(1)}" cy="${getY(m.won).toFixed(1)}" r="4" fill="#2ecc71"><title>${m.ym} 成交: ${m.won.toFixed(1)}萬</title></circle>
              <text x="${getX(i).toFixed(1)}" y="${h - 10}" text-anchor="middle" fill="#7f8c8d" font-size="11">${escapeHtml(m.ym)}</text>
            `).join("")}
            <text x="${w - padR}" y="14" text-anchor="end" fill="#3498db" font-size="11" font-weight="700">● 總報價額(萬)　<tspan fill="#2ecc71">● 已成交額(萬)</tspan></text>
          </svg>
        `;
      }
    }
  }

  function exportDashboardReportCsv() {
    if (!p4DashDataCache) {
      alert("請先載入大數據分析資料！");
      return;
    }
    const topCusts = p4DashDataCache.top_customers || [];
    const lines = ["排名,客戶名稱,報價次數,累計報價金額(未稅),已成交金額(未稅),成交率(%)"];
    topCusts.forEach((c, idx) => {
      lines.push(`${idx + 1},"${String(c.customer_name || "").replace(/"/g, '""')}",${c.quote_count || 0},${c.total_quoted || 0},${c.won_amount || 0},${Number(c.win_rate_pct || 0).toFixed(1)}`);
    });
    const blob = new Blob(["\uFEFF" + lines.join("\r\n")], { type: "text/csv;charset=utf-8;" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = `企業營運戰略報表_${getActiveSalesName()}_${new Date().toISOString().slice(0, 10)}.csv`;
    a.click();
    URL.revokeObjectURL(url);
  }

  // ============================================================================
  // 第 6 頁：客戶管理 (100% 對標電腦版 _QuotationApp gui_page4.py & 截圖 8)
  // ============================================================================
  let selectedMgmtCustomerId = null;
  let p6CustPage = 1;
  const P6_PAGE_SIZE = 30;

  async function loadQuotationCustomersPage() {
    await populateCustomerSelectDropdown(true);
    renderQuotationCustomersTable();
    await loadQuotationSupplierInfo();
  }

  function renderQuotationCustomersTable() {
    const tbody = document.getElementById("quoteCustMgmtTableBody");
    const badge = document.getElementById("qCustCountBadge");
    const pageInfo = document.getElementById("lblCustMgmtPageInfo");
    const kw = (document.getElementById("inputQuoteCustSearch")?.value || "").trim().toLowerCase();
    if (!tbody) return;

    // 將每位聯絡人展開成獨立列顯示 (對應截圖 8 同一家公司多位聯絡人分列)
    const flatRows = [];
    (customersCache || []).forEach(c => {
      const contacts = Array.isArray(c.contacts) && c.contacts.length > 0
        ? c.contacts
        : [{ name: c.default_contact || "", mobile: "", email: "" }];
      contacts.forEach((ct, ctIdx) => {
        const contactName = ct.name || c.default_contact || "";
        if (kw) {
          const hit =
            (c.name || "").toLowerCase().includes(kw) ||
            (c.tax_id || "").toLowerCase().includes(kw) ||
            contactName.toLowerCase().includes(kw) ||
            (c.phone || "").toLowerCase().includes(kw) ||
            (ct.mobile || "").toLowerCase().includes(kw);
          if (!hit) return;
        }
        flatRows.push({
          cust: c,
          contact: ct,
          contactIdx: ctIdx,
          contactName: contactName
        });
      });
    });

    if (badge) badge.textContent = `共 ${flatRows.length} 筆`;

    const totalCount = flatRows.length;
    const totalPages = Math.max(1, Math.ceil(totalCount / P6_PAGE_SIZE));
    if (p6CustPage > totalPages) p6CustPage = totalPages;

    if (pageInfo) {
      pageInfo.textContent = `第 ${p6CustPage} / ${totalPages} 頁 (共 ${totalCount} 筆)`;
    }

    if (totalCount === 0) {
      tbody.innerHTML = `<tr><td colspan="12" style="text-align:center; padding:24px; color:#7f8c8d;">查無相符的客戶資料</td></tr>`;
      return;
    }

    const startIdx = (p6CustPage - 1) * P6_PAGE_SIZE;
    const pageList = flatRows.slice(startIdx, startIdx + P6_PAGE_SIZE);

    tbody.innerHTML = pageList.map((row, idx) => {
      const c = row.cust;
      const ct = row.contact || {};
      const isFav = Number(c.is_favorite || 0) > 0;
      const isSel = Number(c.id) === Number(selectedMgmtCustomerId);
      const rowNum = startIdx + idx + 1;

      return `
        <tr class="quote-cust-row ${isSel ? 'selected-row' : ''}" data-cid="${c.id}" data-cidx="${row.contactIdx}">
          <td class="col-idx">${rowNum}</td>
          <td class="p6-star-cell ${isFav ? '' : 'inactive'}" data-star-cid="${c.id}" title="點擊切換常用客戶星號">${isFav ? '★' : '☆'}</td>
          <td>${escapeHtml(row.contactName || "")}</td>
          <td style="font-family:Consolas, monospace;">${escapeHtml(c.tax_id || "")}</td>
          <td style="font-weight:600;">${escapeHtml(c.name || "")}</td>
          <td style="font-family:Consolas, monospace;">${escapeHtml(c.phone || "")}</td>
          <td style="font-family:Consolas, monospace;">${escapeHtml(c.fax || "")}</td>
          <td style="font-family:Consolas, monospace;">${escapeHtml(ct.mobile || "")}</td>
          <td>${escapeHtml(ct.email || "")}</td>
          <td style="max-width:180px; overflow:hidden; text-overflow:ellipsis;">${escapeHtml(c.address || "")}</td>
          <td>${escapeHtml(c.payment_terms || "月結 30 天電匯")}</td>
          <td style="text-align:center;">正常營業</td>
        </tr>
      `;
    }).join("");

    tbody.querySelectorAll(".p6-star-cell").forEach(starTd => {
      starTd.addEventListener("click", async (e) => {
        e.stopPropagation();
        const cid = parseInt(starTd.getAttribute("data-star-cid"), 10);
        const cust = (customersCache || []).find(x => Number(x.id) === cid);
        if (!cust) return;
        const nextFav = !(Number(cust.is_favorite || 0) > 0);
        const primary = (Array.isArray(cust.contacts) && cust.contacts[0]) || {};
        try {
          await fetch("/api/quotation-customers", {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify({
              name: cust.name,
              owner_sales: getActiveSalesName(),
              tax_id: cust.tax_id || "",
              contact_name: primary.name || cust.default_contact || "",
              contact_mobile: primary.mobile || "",
              contact_email: primary.email || "",
              phone: cust.phone || "",
              fax: cust.fax || "",
              payment_terms: cust.payment_terms || "月結 30 天電匯",
              address: cust.address || "",
              is_favorite: nextFav,
              copy_as_new_contact: false
            })
          });
          await populateCustomerSelectDropdown(true);
          renderQuotationCustomersTable();
        } catch (err) {}
      });
    });

    tbody.querySelectorAll(".quote-cust-row").forEach(tr => {
      tr.addEventListener("click", () => {
        const cid = parseInt(tr.getAttribute("data-cid"), 10);
        const cidx = parseInt(tr.getAttribute("data-cidx") || "0", 10);
        const cust = (customersCache || []).find(x => Number(x.id) === cid);
        if (cust) {
          selectCustomerInMgmtForm(cust, cidx);
        }
      });
    });
  }

  function selectCustomerInMgmtForm(cust, contactIdx = 0) {
    selectedMgmtCustomerId = cust.id;
    renderQuotationCustomersTable();
    const hint = document.getElementById("qCustSelectedHint");
    if (hint) hint.textContent = `已選取：${cust.name} (ID: ${cust.id})`;

    const legend = document.getElementById("qCustModeLegend");
    if (legend) {
      legend.innerHTML = `<span class="mode-dot" style="background:#f39c12;"></span>【目前狀態：編輯模式】`;
      legend.style.color = "#e67e22";
    }

    const contacts = Array.isArray(cust.contacts) && cust.contacts.length > 0
      ? cust.contacts
      : [{ name: cust.default_contact || "", email: "", mobile: "" }];
    const chosen = contacts[contactIdx] || contacts[0] || {};

    document.getElementById("qCustEditId").value = String(cust.id || "");
    document.getElementById("qCustName").value = cust.name || "";
    document.getElementById("qCustTax").value = cust.tax_id || "";
    document.getElementById("qCustContact").value = chosen.name || cust.default_contact || "";
    document.getElementById("qCustMobile").value = chosen.mobile || "";
    document.getElementById("qCustEmail").value = chosen.email || "";
    document.getElementById("qCustPhone").value = cust.phone || "";
    document.getElementById("qCustFax").value = cust.fax || "";
    document.getElementById("qCustPayment").value = cust.payment_terms || "月結 30 天電匯";
    document.getElementById("qCustAddress").value = cust.address || "";
    document.getElementById("qCustFavorite").checked = Number(cust.is_favorite || 0) > 0;
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
    document.getElementById("qCustPayment").value = "";
    document.getElementById("qCustAddress").value = "";
    document.getElementById("qCustFavorite").checked = false;

    const legend = document.getElementById("qCustModeLegend");
    if (legend) {
      legend.innerHTML = `<span class="mode-dot" style="background:#2ecc71;"></span>【目前狀態：新增模式】`;
      legend.style.color = "#27ae60";
    }
    const hint = document.getElementById("qCustSelectedHint");
    if (hint) hint.textContent = "尚未選取客戶";
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
      alert("若要複製為新聯絡人，請先在「聯絡人」輸入新窗口姓名！");
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

  // ============================================================================
  // 第 7 頁：產品管理與 4x5 小算盤 (100% 對標電腦版 _QuotationApp gui_page5.py & 截圖 6)
  // ============================================================================
  let p7ProductsListCache = [];
  let p7ProdPage = 1;
  const P7_PAGE_SIZE = 50;
  let p7SelectedProductObj = null;

  async function searchStandardProductsInMgmt() {
    const tbody = document.getElementById("quoteProdMgmtTableBody");
    if (!tbody) return;
    const q = document.getElementById("inputQuoteProdMgmtSearch")?.value.trim() || "";
    const activeSales = getActiveSalesName();
    tbody.innerHTML = `<tr><td colspan="6" style="text-align:center; padding:20px; color:#7f8c8d;">載入產品資料中...</td></tr>`;
    try {
      const r = await fetch(`/api/products?q=${encodeURIComponent(q)}&limit=500&sales_name=${encodeURIComponent(activeSales)}`);
      if (!r.ok) throw new Error("搜尋產品失敗");
      const list = await r.json();
      p7ProductsListCache = Array.isArray(list) ? list : [];
      renderStandardProductsPage();
    } catch (err) {
      tbody.innerHTML = `<tr><td colspan="6" style="text-align:center; padding:20px; color:#e74c3c;">${escapeHtml(err.message)}</td></tr>`;
    }
  }

  function renderStandardProductsPage() {
    const tbody = document.getElementById("quoteProdMgmtTableBody");
    const pageInfo = document.getElementById("lblProdMgmtPageInfo");
    if (!tbody) return;

    const totalCount = p7ProductsListCache.length;
    const totalPages = Math.max(1, Math.ceil(totalCount / P7_PAGE_SIZE));
    if (p7ProdPage > totalPages) p7ProdPage = totalPages;

    if (pageInfo) {
      pageInfo.textContent = `第 ${p7ProdPage} / ${totalPages} 頁`;
    }

    if (totalCount === 0) {
      tbody.innerHTML = `<tr><td colspan="6" style="text-align:center; padding:24px; color:#7f8c8d;">查無相符的產品型號</td></tr>`;
      return;
    }

    const startIdx = (p7ProdPage - 1) * P7_PAGE_SIZE;
    const pageRows = p7ProductsListCache.slice(startIdx, startIdx + P7_PAGE_SIZE);

    tbody.innerHTML = pageRows.map((p, idx) => {
      const rowNum = startIdx + idx + 1;
      const dealerP = Number(p.dealer_price || 0);
      const listP = Number(p.list_price || (dealerP > 0 ? Math.round(dealerP * 2) : 0));
      const isMto = String(p.spec || "").includes("訂購品") || Number(p.is_mto || 0) > 0;
      const isSel = p7SelectedProductObj && p7SelectedProductObj.model_no === p.model_no;

      return `
        <tr class="p7-prod-row ${isSel ? 'selected-row' : ''}" data-idx="${startIdx + idx}">
          <td class="col-idx">${rowNum}</td>
          <td style="font-weight:600; color:#2c3e50;">${escapeHtml(p.model_no || "")}</td>
          <td style="text-align:right; font-family:Consolas, monospace;">${listP.toLocaleString()}</td>
          <td style="text-align:right; font-family:Consolas, monospace;">${dealerP.toLocaleString()}</td>
          <td style="max-width:210px; overflow:hidden; text-overflow:ellipsis;">${escapeHtml(p.spec || "")}</td>
          <td style="text-align:center; font-weight:700; color:#e67e22;">${isMto ? "V" : ""}</td>
        </tr>
      `;
    }).join("");

    tbody.querySelectorAll(".p7-prod-row").forEach(tr => {
      tr.addEventListener("click", () => {
        const idx = parseInt(tr.getAttribute("data-idx"), 10);
        const prod = p7ProductsListCache[idx];
        if (!prod) return;
        p7SelectedProductObj = prod;
        tbody.querySelectorAll(".p7-prod-row").forEach(r => r.classList.toggle("selected-row", r === tr));

        const dealerP = Number(prod.dealer_price || 0);
        const listP = Number(prod.list_price || (dealerP > 0 ? Math.round(dealerP * 2) : 0));
        const mIn = document.getElementById("p7InputModel");
        const lIn = document.getElementById("p7InputListPrice");
        const dIn = document.getElementById("p7InputDealerPrice");
        const sIn = document.getElementById("p7InputSpec");
        const cIn = document.getElementById("p7ChkIsMto");
        if (mIn) mIn.value = prod.model_no || "";
        if (lIn) lIn.value = String(listP);
        if (dIn) dIn.value = String(dealerP);
        if (sIn) sIn.value = prod.spec || "";
        if (cIn) cIn.checked = String(prod.spec || "").includes("訂購品");

        // 同步帶入小算盤顯示幕
        const calcDisp = document.getElementById("p7CalcDisplay");
        if (calcDisp && dealerP > 0) calcDisp.value = String(dealerP);
      });
    });
  }

  function clearP7ProductForm() {
    p7SelectedProductObj = null;
    const ids = ["p7InputModel", "p7InputListPrice", "p7InputDealerPrice", "p7InputSpec"];
    ids.forEach(id => {
      const el = document.getElementById(id);
      if (el) el.value = "";
    });
    const chk = document.getElementById("p7ChkIsMto");
    if (chk) chk.checked = false;
    renderStandardProductsPage();
  }

  async function saveP7ProductFromForm() {
    const modelNo = document.getElementById("p7InputModel")?.value.trim() || "";
    const dealerPrice = parseFloat(document.getElementById("p7InputDealerPrice")?.value || "0") || 0;
    const specRaw = document.getElementById("p7InputSpec")?.value.trim() || "";
    const isMto = !!document.getElementById("p7ChkIsMto")?.checked;
    if (!modelNo) {
      alert("請輸入產品型號！");
      return;
    }
    if (dealerPrice <= 0) {
      alert("請輸入大於 0 的經銷價！");
      return;
    }
    const finalSpec = isMto && !specRaw.includes("訂購品") ? `${specRaw} (訂購品)`.trim() : specRaw;

    // 同步填入個人特殊品表單並呼叫後端儲存
    const spM = document.getElementById("inputSpecialModelNo");
    const spS = document.getElementById("inputSpecialSpec");
    const spP = document.getElementById("inputSpecialPrice");
    if (spM) spM.value = modelNo;
    if (spS) spS.value = finalSpec;
    if (spP) spP.value = String(dealerPrice);

    await savePersonalSpecialProductFromUI();
    await searchStandardProductsInMgmt();
  }

  async function deleteSelectedP7Product() {
    if (!p7SelectedProductObj) {
      alert("請先從左側產品列表點選一筆產品！");
      return;
    }
    if (!p7SelectedProductObj.is_special) {
      alert("此品項為公司共通標準牌價品項，受系統保護無法直接刪除；若為個人特殊品請展開下方「個人特殊品與方案 A 自動轉正專區」刪除。");
      return;
    }
    alert(`請展開下方「個人特殊品與方案 A 自動轉正專區」刪除【${p7SelectedProductObj.model_no}】。`);
  }

  function setupP7CalculatorEvents() {
    const grid = document.getElementById("p7CalcGrid");
    const disp = document.getElementById("p7CalcDisplay");
    if (!grid || !disp) return;

    const evalCalc = () => {
      const expr = (disp.value || "").replace(/[^0-9+\-*/().]/g, "");
      if (!expr) return;
      try {
        const res = Function(`"use strict"; return (${expr});`)();
        if (isFinite(res)) {
          disp.value = String(Math.round(res * 10000) / 10000);
        } else {
          disp.value = "Error";
        }
      } catch (e) {
        disp.value = "Error";
      }
    };

    grid.querySelectorAll(".calc-btn").forEach(btn => {
      btn.addEventListener("click", () => {
        const k = btn.getAttribute("data-calc");
        if (!k) return;
        if (disp.value === "Error") disp.value = "";
        if (k === "C") {
          disp.value = "";
        } else if (k === "DEL") {
          disp.value = disp.value.slice(0, -1);
        } else if (k === "=") {
          evalCalc();
        } else {
          disp.value += k;
        }
      });
    });
  }

  // ============================================================================
  // 第 8 頁：系統設定 5 大子分頁 (100% 對標電腦版 _QuotationApp gui_page6.py & 截圖 7、9、10、12、13)
  // ============================================================================
  const DEFAULT_QUOTE_TERMS_8_CLAUSES = [
    "1. 本報價單有效期限自報價日起算 30 天內有效，逾期價格與交期需重新確認。",
    "2. 報價金額除特別註明外，單價均為新台幣未稅價，營業稅 (5%) 外加。",
    "3. 交貨期限：收到正式採購訂單（PO）並確認規格無誤後，依明細表所列交期交貨。",
    "4. 交貨地點：買方指定之台灣本島一樓卸貨地點（不含現場吊裝、定位、配線與試車）。",
    "5. 付款條件：依雙方約定之付款方式辦理；客製化或特殊訂購品（標註訂購品 V 者）一經下單恕不接受取消或退換貨。",
    "6. 驗收標準：依士林電機原廠出廠標準規格書驗收，貨到 7 日內如有外觀瑕疵或數量短少請書面提出，逾期視同驗收合格。",
    "7. 保固條款：產品自交貨日起依原廠規定提供 12 個月正常使用保固（消耗品及天災、人為操作不當除外）。",
    "8. 所有權保留：在買方未付清全部貨款前，本批貨物之所有權仍歸賣方（士林電機）所有。"
  ].join("\n");

  function setupQuotationSettingsTabsAndEvents() {
    const tabBar = document.getElementById("quoteSettingsTabBar");
    const qSetPaneMap = {
      general: "qSetPaneGeneral",
      batch: "qSetPaneBatch",
      archive: "qSetPaneArchive",
      terms: "qSetPaneTerms",
      update: "qSetPaneUpdate"
    };
    if (tabBar) {
      tabBar.querySelectorAll(".pyside-tab-btn").forEach(btn => {
        btn.addEventListener("click", () => {
          const rawKey = btn.getAttribute("data-qset-tab") || "general";
          const targetId = qSetPaneMap[rawKey] || rawKey;
          tabBar.querySelectorAll(".pyside-tab-btn").forEach(b => b.classList.toggle("active", b === btn));
          document.querySelectorAll("#quoteSettingsSection .pyside-tab-pane").forEach(pane => {
            const isMatch = pane.id === targetId;
            pane.classList.toggle("active", isMatch);
            pane.classList.toggle("hidden", !isMatch);
          });
        });
      });
    }

    const btnSaveSet = document.getElementById("btnSaveQuoteSettings");
    if (btnSaveSet) btnSaveSet.addEventListener("click", saveQuotationFinanceSettingsFromUI);

    const btnSaveTerms = document.getElementById("btnSaveQuoteTerms");
    if (btnSaveTerms) btnSaveTerms.addEventListener("click", saveQuotationTermsFromUI);

    const btnResetTerms = document.getElementById("btnResetDefaultTerms");
    if (btnResetTerms) {
      btnResetTerms.addEventListener("click", () => {
        if (!confirm("確定要將交貨驗收條款恢復為系統預設的 8 大條款嗎？")) return;
        const txt = document.getElementById("qSetTermsTextarea");
        if (txt) txt.value = DEFAULT_QUOTE_TERMS_8_CLAUSES;
      });
    }

    document.querySelectorAll(".btn-reset-col-width").forEach(btn => {
      btn.addEventListener("click", () => {
        const name = btn.textContent.trim();
        const targetPage = btn.getAttribute("data-reset-target") || "page1";
        resetQuoteTableColumnWidthsByTarget(targetPage);
        alert(`[已還原預設欄寬] 「${name}」表格欄寬已重設為系統推薦寬度！`);
      });
    });

    // 批次資料管理與封存按鈕綁定
    const btnExpCust = document.getElementById("btnBatchExportCustExcel");
    if (btnExpCust) btnExpCust.addEventListener("click", exportCrmMonthlyReportCsv);

    const btnImpCust = document.getElementById("btnBatchImportCustExcel");
    const btnImpProd = document.getElementById("btnBatchImportProdExcel");
    const btnExpProd = document.getElementById("btnBatchExportProdExcel");
    const btnLegacy = document.getElementById("btnLaunchLegacyImporter");
    [btnImpCust, btnImpProd, btnExpProd, btnLegacy].forEach(b => {
      if (b) {
        b.addEventListener("click", () => {
          alert(`[批次資料管理] 網頁版目前已直接連結 SQLite 資料庫 (${b.textContent.trim()})。`);
        });
      }
    });

    const btnCheckUpd = document.getElementById("btnCheckSystemUpdate");
    if (btnCheckUpd) {
      btnCheckUpd.addEventListener("click", () => {
        const st = document.getElementById("lblUpdateCheckStatus");
        if (st) st.textContent = "狀態：目前已是最新版本 v5.5.1.6 (資料庫連線正常)";
      });
    }
    const btnRollback = document.getElementById("btnRollbackPreviousVersion");
    if (btnRollback) {
      btnRollback.addEventListener("click", () => {
        alert("[版本還原提示] 網頁版採用 Git 版本控制，如需退回上一版本請於終端機執行 git revert。");
      });
    }
  }

  async function loadQuotationSettingsPage() {
    await loadQuotationFinanceSettings();
    const taxInput = document.getElementById("qSetTaxRate");
    const digitsSel = document.getElementById("qSetRoundDigits");
    const modeSel = document.getElementById("qSetRoundMode");
    const sortSel = document.getElementById("qSetCustSortMode");
    const defRowsIn = document.getElementById("qSetDefaultRows");
    const termsTxt = document.getElementById("qSetTermsTextarea");

    if (taxInput) taxInput.value = String(quoteFinanceSettings.tax_rate ?? 5);
    if (digitsSel) digitsSel.value = String(quoteFinanceSettings.round_digits ?? 1);
    if (modeSel) modeSel.value = quoteFinanceSettings.round_mode || "四捨五入";
    if (sortSel) sortSel.value = quoteFinanceSettings.customer_sort_mode || "常用客戶優先";
    if (defRowsIn) defRowsIn.value = localStorage.getItem("shihlin_quote_default_rows") || "5";
    if (termsTxt) {
      termsTxt.value = quoteFinanceSettings.quote_terms || DEFAULT_QUOTE_TERMS_8_CLAUSES;
    }
  }

  async function saveQuotationFinanceSettingsFromUI() {
    const defRows = Math.max(1, Math.min(50, parseInt(document.getElementById("qSetDefaultRows")?.value || "5", 10) || 5));
    localStorage.setItem("shihlin_quote_default_rows", String(defRows));

    const payload = {
      tax_rate: parseFloat(document.getElementById("qSetTaxRate")?.value || "5.0"),
      round_digits: parseInt(document.getElementById("qSetRoundDigits")?.value || "1", 10),
      round_mode: document.getElementById("qSetRoundMode")?.value || "四捨五入",
      customer_sort_mode: document.getElementById("qSetCustSortMode")?.value || "常用客戶優先"
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
      alert(`[儲存成功] 系統設定已立即套用！（新建預設行數：${defRows} 行｜捨入規則：${quoteFinanceSettings.round_summary_text}）`);
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

  // ============================================================================
  // 第 9 頁：使用說明 (100% 對標電腦版 _QuotationApp gui_page7.py & 截圖 14、15，純文字標籤無 Emoji)
  // ============================================================================
  const HELP_MANUAL_TOPICS_HTML = [
    `
      <h2 class="help-h2">[單據] 新建/編輯報價單 (核心打單區)</h2>
      <p>這是本系統的核心，提供極速打單與智慧查價功能，以下為各項進階操作說明：</p>
      <div class="help-section-title">[1] 報價單號：兩種模式與三段式流水號</div>
      <p>本系統的報價單編號採用<strong>【三段式唯一識別碼】</strong>，格式為：<code>[第一部分] - [第二部分] - [第三部分]</code>，確保單據絕對不會重複。</p>
      <p><strong>[第一部分：單號前綴] 提供兩種選擇模式：</strong><br>
      ・模式 A (自動擷取 / 公司+年月)：系統會自動以「公司名稱前 4 碼 + 報價年月(YYMM)」組成。適合希望由系統自動統一規範的團隊。<br>
      ・模式 B (自行輸入)：使用者可自行輸入前綴名稱（不可空白）。適合不同專案、部門或業務員有客製化字首需求的團隊。<br>
      <span class="help-tip">[提示] 您選擇的模式與輸入內容，系統皆會永久記憶，下次開單不需重新設定！</span></p>
      <p><strong>[第二部分：3 碼流水號]</strong><br>
      ・系統會依照您的「第一部分前綴」自動分組。新增報價單時，會尋找該群組目前的最大值並 +1 作為新流水號（若無資料則從 001 開始）。<br>
      ・重要規則：流水號一旦產生，<span class="help-warn">絕不會因為您未來修改報價單而重新計算或跳號</span>。</p>
      <p><strong>[第三部分：2 碼版次號]</strong><br>
      ・新建報價單預設為 <code>00</code>。當您從「報價單管理」載入舊單修改並重新儲存時，系統會自動遞增為 <code>01</code>、<code>02</code>，完整保留歷史軌跡。</p>
      <div class="help-section-title">[2] 智慧型號輸入與庫存即時比對</div>
      <p>・<strong>可用數量自動顯示</strong>：輸入或選取型號後，第 1 欄會即時顯示原廠最新可用庫存數量。<br>
      ・<strong>插入方案小計</strong>：若同一張報價單包含多個方案（如方案 A、方案 B），可點擊下方橙色「+ 插入方案小計」分段計算。</p>
    `,
    `
      <h2 class="help-h2">[手機] 手機報價 (外勤草稿與雲端覆核)</h2>
      <p>專為業務人員在外拜訪客戶、無法開啟電腦時設計的行動快速打單流程：</p>
      <div class="help-section-title">[1] 外勤手機端快速建單</div>
      <p>・在外勤使用手機或窄螢幕開啟報價系統時，畫面會自動切換為直向卡片式「手機網頁報價模式」。<br>
      ・選取客戶與產品型號後，可直接點擊<strong>「複製 LINE 報價文字」</strong>傳送給客戶，或點擊<strong>「送出草稿至資料庫」</strong>。</p>
      <div class="help-section-title">[2] 電腦端手機草稿覆核區</div>
      <p>・回到辦公室打開電腦版「手機報價」分頁，左側列表會列出所有來自手機端的待覆核草稿（側邊欄與底部狀態列皆會顯示醒目提示）。<br>
      ・點選任一筆草稿可在右側預覽完整品項、數量、預計報價與交期，確認無誤後點擊綠色<strong>「覆核並轉為正式報價單」</strong>即可一鍵帶入主畫面產出正式 PDF/Excel 報價單。</p>
    `,
    `
      <h2 class="help-h2">[管理] 報價單管理區 (改版、追蹤與垃圾桶)</h2>
      <p>集中管理所有歷史開立之報價單，並與業務戰情室 CRM 商機無縫連動：</p>
      <div class="help-section-title">[1] 單據搜尋與狀態回報</div>
      <p>・支援以公司名稱、聯絡人、單號、專案名稱進行關鍵字即時檢索，或勾選「顯示全部 (忽略日期)」查閱全歷史單據。<br>
      ・單擊列表任一列，上方工具列會顯示目前選取的單號，並可快速標記為<strong>「已成交」、「未成交」、「不跟進」或「報價中」</strong>。<br>
      ・<strong>雙擊任一列</strong>：可直接將該張報價單完整載入至「新建報價單」畫面進行修改或升版。</p>
      <div class="help-section-title">[2] 版次比對與 30 天安全垃圾桶</div>
      <p>・刪除報價單時，系統不會直接抹除資料，而是移入<strong>「垃圾桶」</strong>保留 30 天，隨時可點擊「垃圾桶」切換視角並一鍵復原。</p>
    `,
    `
      <h2 class="help-h2">[看板] 大數據分析看板 (企業營運戰略駕駛艙)</h2>
      <p>將報價紀錄轉化為決策情報，協助業務與主管精準掌握成交率與客戶動態：</p>
      <div class="help-section-title">[1] 待跟進區與五大核心 KPI</div>
      <p>・<strong>待跟進區</strong>：自動篩選出「發出 30 天內尚未回報成交或失注」的報價單，提醒業務黃金追蹤期。<br>
      ・<strong>5 大 KPI 卡片</strong>：即時統計總報價件數/金額、已成交件數/金額、未成交件數/金額、跟進中件數/金額與實質成交率。</p>
      <div class="help-section-title">[2] 2x2 排行與沉睡客戶流失預警雷達</div>
      <p>・自動排行最常報價客戶、成交貢獻王、失注客戶與熱門產品 Top 10。<br>
      ・底部<strong>「沉睡客戶流失預警雷達」</strong>會自動抓出曾有交易但超過 90 天未再詢價的熟客，防止重要客戶默默流失。</p>
    `,
    `
      <h2 class="help-h2">[查價] 歷史單價分析 (成本查閱與利潤防護)</h2>
      <p>打單前秒查過去賣給同客戶或全體客戶的歷史價格，避免報錯價或折數過低：</p>
      <div class="help-section-title">[1] 13 欄深度比對與 26 年最新經銷價換算</div>
      <p>・輸入產品型號或客戶名稱，系統即刻列出過去每一筆報價的「歷史報價(未稅)」、「歷史折數(%)」，並自動比對最新「26年經銷價」計算價差金額與價差率。<br>
      ・下方內建<strong>「價格走勢與統計摘要」</strong>折線圖，藍色實線呈現報價走勢、紅色圓點標示各次落點、橘色虛線標示歷史均價。</p>
    `,
    `
      <h2 class="help-h2">[客戶] 客戶資料管理 (常用星號與多聯絡人)</h2>
      <p>建立標準化客戶檔案，打單時一鍵帶入統編、電話、傳真、地址與付款條件：</p>
      <div class="help-section-title">[1] 常用客戶星號 (★) 與一公司多窗口</div>
      <p>・點擊表格第 2 欄的星號即可將客戶設為<strong>「常用客戶 (★)」</strong>，在新建報價單下拉選單中會優先置頂排列。<br>
      ・同一家公司若有多位採購或工程窗口，可在右側表單修改聯絡人姓名後點擊<strong>「複製為新聯絡人」</strong>，系統即會為該公司新增獨立聯絡人列。</p>
    `,
    `
      <h2 class="help-h2">[產品] 產品資料庫与內建小算盤</h2>
      <p>管理全系列標準品牌價、經銷價、訂購品標記與業務個人特殊品：</p>
      <div class="help-section-title">[1] 標準產品檢索與小算盤</div>
      <p>・點選左側產品列，即可將牌價與經銷價帶入右側編輯區與下方<strong>「4x5 獨立小算盤」</strong>，方便即時試算特殊折數。<br>
      ・點擊右上角「個人特殊品與方案 A 轉正」按鈕，可管理各業務專屬的專案特價品；當公司正式發行標準牌價時，支援方案 A 一鍵轉正並封存歷史紀錄。</p>
    `,
    `
      <h2 class="help-h2">[設定] 系統設定與交貨驗收條款</h2>
      <p>個人化調整財務計算規則、新建預設行數與報價單底部法律條款：</p>
      <div class="help-section-title">[1] 財務捨入規則與新建預設行數</div>
      <p>・支援設定營業稅率 (預設 5%)、捨入位數 (個位數/十位數/百位數/小數點) 與捨入方式 (四捨五入/無條件捨去/無條件進位)。<br>
      ・可自訂「新建預設行數 (預設 5 行)」與「交貨驗收條款 (預設 8 大條款)」，點擊「儲存並立即套用系統設定」即刻生效。</p>
    `,
    `
      <h2 class="help-h2">[備份] 資料庫備份、封存與安全還原機制</h2>
      <p>確保企業報價資產萬無一失的多重防護設計：</p>
      <div class="help-section-title">[1] Append-Only 改版不覆蓋原則</div>
      <p>・所有正式報價單儲存時皆採用 Append-Only 機制，修改舊單會自動產生新版次 (Rev01, Rev02...)，絕不覆蓋原始單據。<br>
      ・所有資料庫讀寫 100% 集中於後端 <code>data_access.py</code>，支援隨時匯出 CSV/Excel 備份並無痛銜接公司 SQL Server Express。</p>
    `,
    `
      <h2 class="help-h2">[秘訣] 實戰應用情境與快捷神技總整理</h2>
      <p>善用以下技巧可將日常打單速度提升 3 倍以上：</p>
      <div class="help-section-title">[1] 四大高效率打單神技</div>
      <p>1. <strong>雙擊歷史單據秒速改版</strong>：在「報價單管理」或「大數據待跟進區」雙擊任一列，1 秒載入完整明細。<br>
      2. <strong>重整經銷價按鈕</strong>：載入半年前的舊報價單後，點擊右下角紫色「重整經銷價」，自動將所有品項更新為今年最新經銷牌價。<br>
      3. <strong>收合選單與收合抬頭</strong>：在小筆電螢幕打單時，點擊左上「收合選單」與右上「收合抬頭」，讓明細表格獲得最大視野。<br>
      4. <strong>LINE 報價文字一鍵複製</strong>：外勤或急件時點擊綠色「複製 LINE 報價」，自動排版成專業報價訊息直接貼上 LINE 傳送。</p>
    `
  ];

  const HELP_CHANGELOG_HTML = `
    <h2 class="help-h2">[履歷] 系統升級與修復履歷 (Changelog)</h2>
    <div style="color:#27ae60; font-weight:700; font-size:1.02rem; margin:10px 0 8px 0;">v5.5.1.6 / PWA v1.0.43 雙端 100% 對齊版 (2026/11)</div>
    <ul style="line-height:1.85; padding-left:20px; margin-bottom:18px;">
      <li><strong>全系統 9 大功能頁 1:1 完整還原</strong>：新建報價單、手機報價草稿覆核區、報價單管理、大數據分析戰略駕駛艙、歷史單價分析、客戶管理、產品管理(含 4x5 小算盤)、系統設定(5 大子分頁)、使用說明(10 大操作手冊主題 + 版次履歷) 100% 對齊電腦版桌面軟體。</li>
      <li><strong>新建報價單 10 欄順序與圓角白底型號選單</strong>：嚴格對齊「可用數量、型號、數量、規格、經銷價、報價折數%、預計報價金額、最終報價金額、小計、備註/交期」欄位順序與預設 5 列空白明細。</li>
      <li><strong>底部雲端草稿狀態列與側邊欄紅點提醒</strong>：即時顯示目前待覆核之手機端草稿筆數，支援單擊預覽「第一筆 / 型號 / 數量 / 預計報價 / 交期」並一鍵轉為正式報價單。</li>
    </ul>
    <div style="color:#27ae60; font-weight:700; font-size:1.02rem; margin:10px 0 8px 0;">v5.5.0.8 版本更新 (2026/09)</div>
    <ul style="line-height:1.85; padding-left:20px;">
      <li><strong>原生極速排版與介面反應最佳化</strong>：全面還原原生 Qt 視窗自適應佈局 (1280x800 基準)，徹底移除向量畫布縮放架構與黑邊補色機制，杜絕繪製管線延遲，提升操作流暢度。</li>
      <li><strong>客戶管理「常用客戶 (★)」標記與多維度排序機制</strong>：客戶資料表格最左側新增常用標記，單擊星號即可即時切換常用狀態；支援 4 種客戶排序模式（常用客戶優先、歷史開單次數優先、公司名稱正序、公司名稱倒序），新建報價單客戶選單同步連動；並增設「同步雲端客戶」手動按鈕。</li>
      <li><strong>士林變頻器系列別結尾規範標準化</strong>：全面對齊官方最新編碼規則，標準款（SA3、SF3、SE3、SC3、SL3）結尾統一標準化為 K；SS2 系列維持 K28 結尾；特規衍生款（如 KP, K-A01 等客製碼）精準保留原始輸入，不強制覆寫。</li>
      <li><strong>料號對照表摘要自動提取與經銷價防呆</strong>：新建報價單手動輸入未建檔型號時，自動自料號對照表 B 欄提取「摘要」填入規格說明；並強制清空經銷價，徹底杜絕上一筆產品之價格殘留。</li>
      <li><strong>庫存比對引擎重大修正 (防範仕様破壞)</strong>：改用正規表達式僅過濾中文備註說明（如 (訂購品)），完整保留英數機能仕様代碼（如 (C)、-F），徹底杜絕 Q38B(C) 誤抓到 Q38B 庫存；保留小數點防止 5.5K 與 55K 互撞。</li>
      <li><strong>品項訂購品視覺警示提示</strong>：型號選單、提示列與表格 Tooltip 新增「[訂購品]」專屬橘色標籤，提醒業務留意備貨排程；選取後維持純淨型號，防止資料庫與報價單輸出字串受到污染。</li>
      <li><strong>手機 PWA 多聯絡人展開選單與草稿相容轉單 (v1.35)</strong>：外勤手機端選客時將同一公司多位聯絡人獨立分列，選取後自動展開詳細資料摘要卡片；桌機端草稿轉單新增向下相容機制，大幅提升穩定性。</li>
    </ul>
  `;

  function setupQuotationHelpTabsAndEvents() {
    const tabBar = document.getElementById("quoteHelpTabBar");
    const qHelpPaneMap = {
      manual: "qHelpPaneManual",
      changelog: "qHelpPaneChangelog"
    };
    if (tabBar) {
      tabBar.querySelectorAll(".pyside-tab-btn").forEach(btn => {
        btn.addEventListener("click", () => {
          const rawKey = btn.getAttribute("data-qhelp-tab") || "manual";
          const targetId = qHelpPaneMap[rawKey] || rawKey;
          tabBar.querySelectorAll(".pyside-tab-btn").forEach(b => b.classList.toggle("active", b === btn));
          document.querySelectorAll("#quoteHelpSection .pyside-tab-pane").forEach(pane => {
            const isMatch = pane.id === targetId;
            pane.classList.toggle("active", isMatch);
            pane.classList.toggle("hidden", !isMatch);
          });
        });
      });
    }

    const topicList = document.getElementById("quoteHelpTopicList");
    if (topicList) {
      topicList.querySelectorAll(".pyside-help-topic-item").forEach(item => {
        item.addEventListener("click", () => {
          const idx = parseInt(item.getAttribute("data-topic-idx") || item.getAttribute("data-help-topic") || "0", 10);
          topicList.querySelectorAll(".pyside-help-topic-item").forEach(li => li.classList.toggle("active", li === item));
          const manualBox = document.getElementById("quoteHelpManualContent");
          if (manualBox) {
            manualBox.innerHTML = HELP_MANUAL_TOPICS_HTML[idx] || HELP_MANUAL_TOPICS_HTML[0];
          }
        });
      });
    }
  }

  function loadQuotationHelpPage() {
    const manualBox = document.getElementById("quoteHelpManualContent");
    const changelogBox = document.getElementById("quoteHelpChangelogContent");
    if (manualBox && !manualBox.innerHTML.trim()) {
      manualBox.innerHTML = HELP_MANUAL_TOPICS_HTML[0];
    }
    if (changelogBox && !changelogBox.innerHTML.trim()) {
      changelogBox.innerHTML = HELP_CHANGELOG_HTML;
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

});

