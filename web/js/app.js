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
  const navItemSettings = document.getElementById("navItemSettings");
  const navItemAdminTargets = document.getElementById("navItemAdminTargets");
  const navItemDriverMgmt = document.getElementById("navItemDriverMgmt");

  const panelHome = document.getElementById("panelHome");
  const panelCases = document.getElementById("panelCases");
  const panelQuotes = document.getElementById("panelQuotes");
  const panelSpecials = document.getElementById("panelSpecials");
  const panelSettings = document.getElementById("panelSettings");
  const panelAdminTargets = document.getElementById("panelAdminTargets");
  const panelDriverMgmt = document.getElementById("panelDriverMgmt");

  const navItems = [
    { btn: navItemHome, panel: panelHome },
    { btn: navItemCases, panel: panelCases },
    { btn: navItemQuotes, panel: panelQuotes },
    { btn: navItemSpecials, panel: panelSpecials },
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
        const curSales = localStorage.getItem("shihlin_pwa_active_sales") || "曾維崧";
        currentViewingDriverSales = currentViewingDriverSales || curSales;
        renderDriverManagementPage(currentViewingDriverSales);
      }

      if (activePanel === panelQuotes) {
        initQuotationModule();
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
        localStorage.setItem("shihlin_pwa_active_sales", sName);
        currentViewingDriverSales = sName;
        updateActiveSalesUI(sName);
        updateManagerUIVisibility(sName);
        loadAndRenderDriverMetrics(sName);
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
  // 13. 報價單管理模組 (舊版報價單完整移植版)
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
  let targetQuoteNoToDelete = null;
  let currentBaseQuoteNo = null;

  const PRESET_DELIVERY_OPTIONS = [
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

  function initQuotationModule() {
    if (!isQuotationModuleInitialized) {
      setupQuotationEvents();
      setupQuotationModals();
      isQuotationModuleInitialized = true;
    }
    const qDate = document.getElementById("quoteDate");
    if (qDate && !qDate.value) {
      qDate.value = new Date().toISOString().split("T")[0];
    }
    const itemsContainer = document.getElementById("itemsContainer");
    if (itemsContainer && itemsContainer.children.length === 0) {
      addBlankItem();
    }
    updateQuoteNoPreview();
  }

  // 四大子分頁切換邏輯
  function switchQuoteSubTab(tabName) {
    currentActiveQuoteTab = tabName;
    const tabCreate = document.getElementById("tabQuoteCreate");
    const tabHistory = document.getElementById("tabQuoteHistory");
    const tabDrafts = document.getElementById("tabQuoteDrafts");
    const tabAnalysis = document.getElementById("tabQuoteAnalysis");

    const draftSec = document.getElementById("quoteDraftSection");
    const histSec = document.getElementById("quoteHistorySection");
    const draftsInboxSec = document.getElementById("quoteDraftsInboxSection");
    const analysisSec = document.getElementById("quoteAnalysisSection");

    [tabCreate, tabHistory, tabDrafts, tabAnalysis].forEach(btn => {
      if (btn) btn.classList.remove("active");
    });
    [draftSec, histSec, draftsInboxSec, analysisSec].forEach(sec => {
      if (sec) sec.classList.add("hidden");
    });

    if (tabName === 'create') {
      if (tabCreate) tabCreate.classList.add("active");
      if (draftSec) draftSec.classList.remove("hidden");
    } else if (tabName === 'history') {
      if (tabHistory) tabHistory.classList.add("active");
      if (histSec) histSec.classList.remove("hidden");
      loadQuotationHistory("", isTrashViewActive);
    } else if (tabName === 'drafts') {
      if (tabDrafts) tabDrafts.classList.add("active");
      if (draftsInboxSec) draftsInboxSec.classList.remove("hidden");
      loadQuotationDraftsInbox();
    } else if (tabName === 'analysis') {
      if (tabAnalysis) tabAnalysis.classList.add("active");
      if (analysisSec) analysisSec.classList.remove("hidden");
    }
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
    // 1. 四大子分頁按鈕監聽
    const tabCreate = document.getElementById("tabQuoteCreate");
    const tabHistory = document.getElementById("tabQuoteHistory");
    const tabDrafts = document.getElementById("tabQuoteDrafts");
    const tabAnalysis = document.getElementById("tabQuoteAnalysis");

    if (tabCreate) tabCreate.addEventListener("click", () => switchQuoteSubTab('create'));
    if (tabHistory) tabHistory.addEventListener("click", () => switchQuoteSubTab('history'));
    if (tabDrafts) tabDrafts.addEventListener("click", () => switchQuoteSubTab('drafts'));
    if (tabAnalysis) tabAnalysis.addEventListener("click", () => switchQuoteSubTab('analysis'));

    // 2. 表頭展開 / 收合切換
    const btnToggleHeader = document.getElementById("btnToggleQuoteHeader");
    const headerContent = document.getElementById("quoteHeaderContent");
    const headerToggleText = document.getElementById("quoteHeaderToggleText");
    if (btnToggleHeader && headerContent) {
      btnToggleHeader.addEventListener("click", () => {
        if (headerContent.classList.contains("hidden")) {
          headerContent.classList.remove("hidden");
          if (headerToggleText) headerToggleText.textContent = "收合表頭";
        } else {
          headerContent.classList.add("hidden");
          if (headerToggleText) headerToggleText.textContent = "展開表頭";
        }
      });
    }

    // 3. 客戶挑選輔助按鈕
    const btnBrowseCust = document.getElementById("btnBrowseCustomerModal");
    if (btnBrowseCust) {
      btnBrowseCust.addEventListener("click", () => {
        const custModal = document.getElementById("customerModal");
        const custSearch = document.getElementById("customerModalSearch");
        if (custSearch) custSearch.value = "";
        renderCustomerModalList("");
        if (custModal) custModal.classList.remove("hidden");
        setTimeout(() => custSearch && custSearch.focus(), 100);
      });
    }

    // 4. 關聯商機按鈕 (Deal)
    const btnLinkDeal = document.getElementById("btnLinkDeal");
    if (btnLinkDeal) {
      btnLinkDeal.addEventListener("click", () => {
        if (dealsCache && dealsCache.length > 0) {
          const names = dealsCache.slice(0, 5).map((d, i) => `${i + 1}. ${d.name} (${d.customer_name})`).join("\n");
          const choice = prompt(`請輸入要關聯的商機編號 (1~${Math.min(dealsCache.length, 5)}):\n${names}`);
          const idx = parseInt(choice, 10) - 1;
          if (dealsCache[idx]) {
            const selectedDeal = dealsCache[idx];
            const dealNameEl = document.getElementById("linkedDealName");
            if (dealNameEl) {
              dealNameEl.textContent = `已關聯：${selectedDeal.name}`;
              dealNameEl.style.color = "var(--primary-color, #2563eb)";
              dealNameEl.style.fontWeight = "600";
            }
          }
        } else {
          alert("[商機清單空] 目前無可關聯的 Deal 商機，可至儀表板或後續模組建立。");
        }
      });
    }

    // 5. 日期與客戶名稱更動連動單號
    const custInput = document.getElementById("customerName");
    if (custInput) custInput.addEventListener("input", updateQuoteNoPreview);
    const dateInput = document.getElementById("quoteDate");
    if (dateInput) dateInput.addEventListener("change", updateQuoteNoPreview);

    // 6. 新增品項按鈕 (手機端與電腦端)
    const btnAddItem = document.getElementById("btnAddItem");
    if (btnAddItem) btnAddItem.addEventListener("click", () => addBlankItem());

    const btnDesktopAdd = document.getElementById("btnDesktopAddItem");
    if (btnDesktopAdd) btnDesktopAdd.addEventListener("click", () => addBlankItem());

    // 7. 電腦端從產品庫挑選
    const btnDesktopBrowse = document.getElementById("btnDesktopBrowseProducts");
    if (btnDesktopBrowse) {
      btnDesktopBrowse.addEventListener("click", () => {
        addBlankItem();
        currentEditingItemId = `item-${quoteItemCount}`;
        openProductModal();
      });
    }

    // 8. 插入方案小計列
    const btnInsertSubtotal = document.getElementById("btnInsertSubtotalRow");
    if (btnInsertSubtotal) {
      btnInsertSubtotal.addEventListener("click", () => insertSubtotalRow());
    }

    // 9. 從舊報價單匯入
    const btnImportOld = document.getElementById("btnImportOldQuote");
    if (btnImportOld) {
      btnImportOld.addEventListener("click", openImportOldQuoteModal);
    }

    // 10. 智慧 BOM 匯入
    const btnImportBOM = document.getElementById("btnImportSmartBOM");
    if (btnImportBOM) {
      btnImportBOM.addEventListener("click", openSmartBOMModal);
    }

    // 11. 重整經銷價
    const btnRefreshPrices = document.getElementById("btnRefreshDealerPrices");
    if (btnRefreshPrices) {
      btnRefreshPrices.addEventListener("click", refreshDealerPrices);
    }

    // 12. 清空品項 (防呆二次確認)
    const btnDesktopClear = document.getElementById("btnDesktopClearItems");
    const clearModal = document.getElementById("clearConfirmModal");
    const btnCancelClear = document.getElementById("btnCancelClearConfirm");
    const btnExecuteClear = document.getElementById("btnExecuteClearConfirm");

    if (btnDesktopClear && clearModal) {
      btnDesktopClear.addEventListener("click", () => {
        clearModal.classList.remove("hidden");
      });
    }
    if (btnCancelClear && clearModal) {
      btnCancelClear.addEventListener("click", () => {
        clearModal.classList.add("hidden");
      });
    }
    if (btnExecuteClear && clearModal) {
      btnExecuteClear.addEventListener("click", () => {
        const itemsContainer = document.getElementById("itemsContainer");
        const desktopTableBody = document.getElementById("quoteDesktopTableBody");
        if (itemsContainer) itemsContainer.innerHTML = "";
        if (desktopTableBody) desktopTableBody.innerHTML = "";
        quoteItemCount = 0;
        addBlankItem();
        clearModal.classList.add("hidden");
      });
    }

    // 13. 手動視角切換器 (手機卡片 vs 電腦試算表)
    const btnToggleView = document.getElementById("btnToggleManualView");
    if (btnToggleView) {
      btnToggleView.addEventListener("click", toggleQuoteViewMode);
    }

    // 14. 附加圖片上傳監聽
    const imgUploader = document.getElementById("quoteImageUploader");
    if (imgUploader) {
      imgUploader.addEventListener("change", handleImageUpload);
    }

    // 15. 底部操作按鈕：儲存報價單 (Append-Only)
    const btnSubmitDraft = document.getElementById("btnSubmitDraft");
    if (btnSubmitDraft) btnSubmitDraft.addEventListener("click", saveQuotationDraft);

    const btnDesktopSubmit = document.getElementById("btnDesktopSubmitDraft");
    if (btnDesktopSubmit) btnDesktopSubmit.addEventListener("click", saveQuotationDraft);

    // 16. 底部操作按鈕：複製 LINE 報價
    const btnLineQuote = document.getElementById("btnLineQuote");
    if (btnLineQuote) btnLineQuote.addEventListener("click", openLineQuoteModal);

    const btnDesktopLine = document.getElementById("btnDesktopLineQuote");
    if (btnDesktopLine) btnDesktopLine.addEventListener("click", openLineQuoteModal);

    // 17. 底部操作按鈕：匯出 Excel
    const btnExportExcel = document.getElementById("btnExportExcel");
    if (btnExportExcel) btnExportExcel.addEventListener("click", exportQuoteToExcel);

    // 18. 底部操作按鈕：列印 PDF
    const btnPrintPDF = document.getElementById("btnPrintPDF");
    if (btnPrintPDF) btnPrintPDF.addEventListener("click", printQuotePDF);

    // 19. 底部操作按鈕：發起專案售價申請
    const btnApplySpecial = document.getElementById("btnApplySpecialPrice");
    if (btnApplySpecial) btnApplySpecial.addEventListener("click", applySpecialPrice);

    // 20. 歷史分頁 (Tab B) 垃圾桶切換與搜尋
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

    // 21. 草稿匣清空按鈕 (Tab C)
    const btnClearDrafts = document.getElementById("btnClearDraftsInbox");
    if (btnClearDrafts) {
      btnClearDrafts.addEventListener("click", () => {
        if (confirm("確定要清空所有行動裝置暫存草稿嗎？")) {
          localStorage.removeItem("shihlin_mobile_quote_drafts");
          loadQuotationDraftsInbox();
        }
      });
    }

    // 22. 歷史單價查詢按鈕 (Tab D)
    const btnRunAnalysis = document.getElementById("btnRunPriceAnalysis");
    if (btnRunAnalysis) {
      btnRunAnalysis.addEventListener("click", runPriceAnalysis);
    }
  }

  // 手動視角切換器
  function toggleQuoteViewMode() {
    const draftContainer = document.getElementById("quoteDraftSection");
    const indicator = document.getElementById("viewModeIndicator");
    if (!draftContainer) return;

    if (draftContainer.classList.contains("force-mobile-view")) {
      draftContainer.classList.remove("force-mobile-view");
      if (indicator) indicator.textContent = "電腦端 Excel 試算表視角";
    } else {
      draftContainer.classList.add("force-mobile-view");
      if (indicator) indicator.textContent = "手機端觸控卡片視角 (手動切換)";
    }
  }

  // 新增品項 (同步手機與電腦端)
  function addBlankItem(initialData = null) {
    quoteItemCount++;
    const itemId = `item-${quoteItemCount}`;
    const itemsContainer = document.getElementById("itemsContainer");
    const desktopTableBody = document.getElementById("quoteDesktopTableBody");
    if (!itemsContainer) return;

    const currentCount = itemsContainer.querySelectorAll('.item-row').length + 1;
    const modelText = initialData?.model_no || "點擊選擇產品...";
    const hasModel = initialData?.model_no ? true : false;
    const specVal = initialData?.spec || "";
    const dealerPriceVal = initialData?.dealer_price || 0;
    const dealerPriceDisplay = dealerPriceVal > 0 ? `$ ${dealerPriceVal.toLocaleString()}` : "-";
    const discountVal = initialData?.discount_pct !== undefined ? initialData.discount_pct : 100;
    const unitPriceVal = initialData?.final_unit_price !== undefined ? initialData.final_unit_price : (dealerPriceVal > 0 ? Math.round(dealerPriceVal * (discountVal / 100)) : 0);
    const qtyVal = initialData?.quantity || 1;
    const subtotalVal = initialData?.row_subtotal !== undefined ? initialData.row_subtotal : Math.round(unitPriceVal * qtyVal);
    const leadTimeVal = initialData?.note_lead_time || "";

    // 1. 手機端卡片 HTML
    const mobileItemHTML = `
      <div class="item-row" id="${itemId}" data-spec="${escapeHtml(specVal)}">
        <div class="item-header">
          <span class="item-index-badge">#${currentCount}</span>
          <button type="button" class="product-picker-input ${hasModel ? 'selected' : ''}" id="${itemId}-title">
            ${escapeHtml(modelText)}
          </button>
          <button type="button" class="item-copy-btn ${hasModel ? '' : 'hidden'}" id="${itemId}-copy-btn" title="複製型號">
            <svg width="16" height="16" viewBox="0 0 16 16" fill="currentColor">
              <path d="M4 2a2 2 0 0 1 2-2h6a2 2 0 0 1 2 2v8a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V2zm2-1a1 1 0 0 0-1 1v8a1 1 0 0 0 1 1h6a1 1 0 0 0 1-1V2a1 1 0 0 0-1-1H6zM2 5a1 1 0 0 0-1 1v8a1 1 0 0 0 1 1h6a1 1 0 0 0 1-1v-1h1v1a2 2 0 0 1-2 2H2a2 2 0 0 1-2-2V6a2 2 0 0 1 2-2h1v1H2z"/>
            </svg>
          </button>
          <button type="button" class="item-remove" id="${itemId}-remove" title="刪除品項">&times;</button>
        </div>

        <div class="item-grid" style="margin-top:8px;">
          <div>
            <label>經銷牌價(未稅)</label>
            <input type="text" id="${itemId}-dealer-price" placeholder="-" readonly class="field-readonly price-field" value="${dealerPriceDisplay}">
          </div>
          <div>
            <label>建議折數(%)</label>
            <input type="number" id="${itemId}-sug-discount" placeholder="100" step="any" value="${discountVal}">
          </div>
          <div>
            <label>預計單價(元)</label>
            <input type="number" id="${itemId}-sug-price" placeholder="0" step="any" value="${unitPriceVal}">
          </div>
          <div>
            <label>需求數量</label>
            <input type="number" id="${itemId}-qty" min="1" value="${qtyVal}" required>
          </div>
          <div>
            <label>小計(未稅)</label>
            <input type="text" id="${itemId}-final-price" placeholder="0" readonly class="field-readonly final-price-field" value="${subtotalVal > 0 ? subtotalVal.toLocaleString() : '0'}">
          </div>
        </div>

        <div class="item-delivery" style="margin-top:8px;">
          <label>交期 / 備註</label>
          <input type="text" id="${itemId}-delivery" placeholder="點擊選擇交期 / 備註..." readonly style="cursor:pointer;" value="${escapeHtml(leadTimeVal)}">
        </div>

        <input type="hidden" id="${itemId}-code" value="${escapeHtml(initialData?.model_no || '')}">
        <input type="hidden" id="${itemId}-sku" value="${escapeHtml(initialData?.sku || '')}">
      </div>
    `;
    itemsContainer.insertAdjacentHTML('beforeend', mobileItemHTML);

    // 2. 電腦端 Excel 試算表表格列 HTML
    if (desktopTableBody) {
      const desktopRowHTML = `
        <tr id="d-${itemId}" class="desktop-item-row" data-itemid="${itemId}">
          <td style="text-align:center; font-weight:700; color:var(--text-muted);" id="d-${itemId}-index">
            ${currentCount}
          </td>
          <td>
            <div style="display:flex; align-items:center; gap:4px;">
              <button type="button" class="table-cell-model-btn ${hasModel ? 'selected' : ''}" id="d-${itemId}-model">
                ${escapeHtml(modelText)}
              </button>
              <button type="button" class="item-copy-btn ${hasModel ? '' : 'hidden'}" id="d-${itemId}-copy-btn" title="複製型號">
                <svg width="14" height="14" viewBox="0 0 16 16" fill="currentColor">
                  <path d="M4 2a2 2 0 0 1 2-2h6a2 2 0 0 1 2 2v8a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V2zm2-1a1 1 0 0 0-1 1v8a1 1 0 0 0 1 1h6a1 1 0 0 0 1-1V2a1 1 0 0 0-1-1H6zM2 5a1 1 0 0 0-1 1v8a1 1 0 0 0 1 1h6a1 1 0 0 0 1-1v-1h1v1a2 2 0 0 1-2 2H2a2 2 0 0 1-2-2V6a2 2 0 0 1 2-2h1v1H2z"/>
                </svg>
              </button>
            </div>
          </td>
          <td>
            <span id="d-${itemId}-spec" style="font-size:0.82rem; color:var(--text-secondary);">${escapeHtml(specVal || '-')}</span>
          </td>
          <td style="text-align:center;">
            <span id="d-${itemId}-stock" class="stock-badge stock-in" style="font-size:0.75rem;">現貨充足</span>
          </td>
          <td style="text-align:right;">
            <span id="d-${itemId}-dealer-price" class="table-cell-readonly money">${dealerPriceDisplay}</span>
          </td>
          <td>
            <input type="number" id="d-${itemId}-discount" class="table-cell-input number" step="any" value="${discountVal}">
          </td>
          <td>
            <input type="number" id="d-${itemId}-price" class="table-cell-input number" step="any" value="${unitPriceVal}">
          </td>
          <td>
            <input type="number" id="d-${itemId}-qty" class="table-cell-input number" min="1" value="${qtyVal}">
          </td>
          <td style="text-align:right;">
            <span id="d-${itemId}-subtotal" class="table-cell-readonly money" style="color:var(--primary-color, #2563eb); font-size:0.95rem;">
              ${subtotalVal > 0 ? '$ ' + subtotalVal.toLocaleString() : '$ 0'}
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
    }

    refreshItemIndices();
    setupItemRowEvents(itemId);
    calculateTotal();
  }

  // 插入方案小計列 (橘色醒目分隔列)
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
        <td colspan="7">
          <input type="text" class="subtotal-title-input" value="${escapeHtml(title)}" placeholder="輸入方案名稱 (如：方案 A：主驅動系統小計)">
        </td>
        <td style="text-align:right;">
          <span class="segment-subtotal-val" style="font-family:Consolas, monospace; font-size:1rem; font-weight:700; color:#c2410c;">$ 0</span>
        </td>
        <td>
          <span style="font-size:0.75rem; color:#9a3412;">[區段未稅金額]</span>
        </td>
        <td style="text-align:center;">
          <button type="button" class="table-btn-remove" onclick="document.getElementById('${rowId}').remove(); calculateTotal();" title="刪除此小計列">&times;</button>
        </td>
      </tr>
    `;
    desktopTableBody.insertAdjacentHTML('beforeend', subtotalHTML);
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
        currentSegmentSum = 0; // 重設供下一區段累加
      } else if (row.classList.contains('desktop-item-row')) {
        const itemId = row.getAttribute('data-itemid');
        const price = parseFloat(document.getElementById(`d-${itemId}-price`)?.value) || 0;
        const qty = parseInt(document.getElementById(`d-${itemId}-qty`)?.value) || 0;
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

    // 2. 電腦端表格元件
    const dModelBtn = document.getElementById(`d-${itemId}-model`);
    const dCopyBtn = document.getElementById(`d-${itemId}-copy-btn`);
    const dRemoveBtn = document.getElementById(`d-${itemId}-remove`);
    const dDeliveryInput = document.getElementById(`d-${itemId}-delivery`);
    const dPriceInput = document.getElementById(`d-${itemId}-price`);
    const dDiscountInput = document.getElementById(`d-${itemId}-discount`);
    const dQtyInput = document.getElementById(`d-${itemId}-qty`);
    const dSubtotalSpan = document.getElementById(`d-${itemId}-subtotal`);

    // 挑選產品庫
    function triggerProductPicker() {
      currentEditingItemId = itemId;
      openProductModal();
    }
    if (mTitleBtn) mTitleBtn.addEventListener("click", triggerProductPicker);
    if (dModelBtn) dModelBtn.addEventListener("click", triggerProductPicker);

    // 複製型號
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
    if (dCopyBtn) dCopyBtn.addEventListener("click", () => triggerCopyModel(dCopyBtn));

    // 刪除品項
    function triggerRemoveItem() {
      const mRow = document.getElementById(itemId);
      const dRow = document.getElementById(`d-${itemId}`);
      if (mRow) mRow.remove();
      if (dRow) dRow.remove();
      calculateTotal();
      refreshItemIndices();
    }
    if (mRemoveBtn) mRemoveBtn.addEventListener("click", triggerRemoveItem);
    if (dRemoveBtn) dRemoveBtn.addEventListener("click", triggerRemoveItem);

    // 選取交期
    function triggerDeliveryPicker() {
      currentEditingItemId = itemId;
      openDeliveryModal();
    }
    if (mDeliveryInput) mDeliveryInput.addEventListener("click", triggerDeliveryPicker);
    if (dDeliveryInput) dDeliveryInput.addEventListener("click", triggerDeliveryPicker);

    // 計算小計與同步
    function updateRowSubtotal() {
      const sp = parseFloat(dPriceInput?.value !== undefined && dPriceInput.value !== "" ? dPriceInput.value : mSugPriceInput?.value) || 0;
      const q = parseInt(dQtyInput?.value !== undefined && dQtyInput.value !== "" ? dQtyInput.value : mQtyInput?.value) || 0;
      const sub = Math.round(sp * q);

      if (mFinalPriceInput) {
        mFinalPriceInput.value = sub > 0 ? sub.toLocaleString() : "0";
      }
      if (dSubtotalSpan) {
        dSubtotalSpan.textContent = sub > 0 ? `$ ${sub.toLocaleString()}` : "$ 0";
      }
      calculateTotal();
    }

    // 單價變更 (雙向同步折數與小計)
    function onPriceChanged(val, fromSource) {
      const sp = parseFloat(val) || 0;
      const dpStr = (mDealerPriceInput?.value || "").replace(/[^0-9.]/g, '');
      const dp = parseFloat(dpStr) || 0;

      if (fromSource === 'desktop' && mSugPriceInput) mSugPriceInput.value = val;
      if (fromSource === 'mobile' && dPriceInput) dPriceInput.value = val;

      if (dp > 0) {
        const discount = (sp / dp) * 100;
        const discStr = discount.toFixed(1);
        if (mSugDiscountInput) mSugDiscountInput.value = discStr;
        if (dDiscountInput) dDiscountInput.value = discStr;
      }
      updateRowSubtotal();
    }

    // 折數變更 (雙向同步單價與小計)
    function onDiscountChanged(val, fromSource) {
      const disc = parseFloat(val);
      const dpStr = (mDealerPriceInput?.value || "").replace(/[^0-9.]/g, '');
      const dp = parseFloat(dpStr) || 0;

      if (fromSource === 'desktop' && mSugDiscountInput) mSugDiscountInput.value = val;
      if (fromSource === 'mobile' && dDiscountInput) dDiscountInput.value = val;

      if (!isNaN(disc) && dp > 0) {
        const sp = Math.round(dp * (disc / 100));
        if (mSugPriceInput) mSugPriceInput.value = sp;
        if (dPriceInput) dPriceInput.value = sp;
      }
      updateRowSubtotal();
    }

    // 數量變更
    function onQtyChanged(val, fromSource) {
      if (fromSource === 'desktop' && mQtyInput) mQtyInput.value = val;
      if (fromSource === 'mobile' && dQtyInput) dQtyInput.value = val;
      updateRowSubtotal();
    }

    if (mSugPriceInput) mSugPriceInput.addEventListener("input", (e) => onPriceChanged(e.target.value, 'mobile'));
    if (dPriceInput) dPriceInput.addEventListener("input", (e) => onPriceChanged(e.target.value, 'desktop'));

    if (mSugDiscountInput) mSugDiscountInput.addEventListener("input", (e) => onDiscountChanged(e.target.value, 'mobile'));
    if (dDiscountInput) dDiscountInput.addEventListener("input", (e) => onDiscountChanged(e.target.value, 'desktop'));

    if (mQtyInput) mQtyInput.addEventListener("input", (e) => onQtyChanged(e.target.value, 'mobile'));
    if (dQtyInput) dQtyInput.addEventListener("input", (e) => onQtyChanged(e.target.value, 'desktop'));
  }

  // 核心計算總計：徹底解決右下角 $0 Bug，雙向即時連動
  function calculateTotal() {
    const desktopTableBody = document.getElementById("quoteDesktopTableBody");
    const itemsContainer = document.getElementById("itemsContainer");
    let subtotal = 0;
    let itemCount = 0;

    const dRows = desktopTableBody ? desktopTableBody.querySelectorAll('.desktop-item-row') : [];
    const mRows = itemsContainer ? itemsContainer.querySelectorAll('.item-row') : [];

    if (dRows.length > 0) {
      dRows.forEach(r => {
        if (r.classList.contains('subtotal-divider-row')) return;
        itemCount++;
        const itemId = r.getAttribute('data-itemid');
        const priceInput = document.getElementById(`d-${itemId}-price`);
        const qtyInput = document.getElementById(`d-${itemId}-qty`);
        const price = parseFloat(priceInput?.value) || 0;
        const qty = parseInt(qtyInput?.value) || 0;
        const rowSub = Math.round(price * qty);
        subtotal += rowSub;

        const subSpan = document.getElementById(`d-${itemId}-subtotal`);
        if (subSpan) subSpan.textContent = rowSub > 0 ? `$ ${rowSub.toLocaleString()}` : "$ 0";

        const mFinal = document.getElementById(`${itemId}-final-price`);
        if (mFinal) mFinal.value = rowSub > 0 ? rowSub.toLocaleString() : "0";
      });
    } else if (mRows.length > 0) {
      mRows.forEach(r => {
        itemCount++;
        const priceInput = r.querySelector('[id$="-sug-price"]');
        const qtyInput = r.querySelector('[id$="-qty"]');
        const price = parseFloat(priceInput?.value) || 0;
        const qty = parseInt(qtyInput?.value) || 0;
        const rowSub = Math.round(price * qty);
        subtotal += rowSub;

        const mFinal = r.querySelector('[id$="-final-price"]');
        if (mFinal) mFinal.value = rowSub > 0 ? rowSub.toLocaleString() : "0";
      });
    }

    const tax = Math.round(subtotal * 0.05);
    const grandTotal = subtotal + tax;

    // 1. 更新手機端小計卡片
    const mBadge = document.getElementById("draftItemBadge");
    if (mBadge) mBadge.textContent = `共 ${itemCount} 項`;
    const mTotal = document.getElementById("totalAmount");
    if (mTotal) mTotal.textContent = subtotal.toLocaleString();

    // 2. 更新電腦端 Excel 結算卡片 (未稅、營業稅、含稅)
    const dBadge = document.getElementById("desktopItemCountBadge");
    if (dBadge) dBadge.textContent = `共 ${itemCount} 項`;
    const dSubtotal = document.getElementById("desktopSubtotal");
    if (dSubtotal) dSubtotal.textContent = `$ ${subtotal.toLocaleString()}`;
    const dTax = document.getElementById("desktopTax");
    if (dTax) dTax.textContent = `$ ${tax.toLocaleString()}`;
    const dGrandTotal = document.getElementById("desktopGrandTotal");
    if (dGrandTotal) dGrandTotal.textContent = `$ ${grandTotal.toLocaleString()}`;

    // 3. 連動更新方案小計
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

  // 重整經銷價 (自資料庫同步最新牌價並更新)
  async function refreshDealerPrices() {
    try {
      const resp = await fetch("/api/products");
      if (resp.ok) {
        quoteProductsCache = await resp.json();
      }
      const desktopTableBody = document.getElementById("quoteDesktopTableBody");
      const dRows = desktopTableBody ? desktopTableBody.querySelectorAll('.desktop-item-row') : [];
      let updatedCount = 0;

      dRows.forEach(r => {
        const itemId = r.getAttribute('data-itemid');
        const codeInput = document.getElementById(`${itemId}-code`);
        const model = codeInput?.value;
        if (!model) return;

        const prod = (quoteProductsCache || []).find(p => p.model_no === model);
        if (prod && prod.dealer_price) {
          const newDP = prod.dealer_price;
          const dDP = document.getElementById(`d-${itemId}-dealer-price`);
          if (dDP) dDP.textContent = `$ ${newDP.toLocaleString()}`;
          const mDP = document.getElementById(`${itemId}-dealer-price`);
          if (mDP) mDP.value = `$ ${newDP.toLocaleString()}`;

          const disc = parseFloat(document.getElementById(`d-${itemId}-discount`)?.value) || 100;
          const newPrice = Math.round(newDP * (disc / 100));
          const dPrice = document.getElementById(`d-${itemId}-price`);
          if (dPrice) dPrice.value = newPrice;
          const mPrice = document.getElementById(`${itemId}-sug-price`);
          if (mPrice) mPrice.value = newPrice;
          updatedCount++;
        }
      });
      calculateTotal();
      alert(`[經銷價重整完成] 已成功校對並同步 ${updatedCount} 項產品之最新牌價！`);
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

        if (quoteProductsCache.length === 0) {
          try {
            const resp = await fetch("/api/products");
            if (resp.ok) quoteProductsCache = await resp.json();
          } catch (e) {}
        }

        let imported = 0;
        lines.forEach(line => {
          const parts = line.split(/[\t ,]+/);
          const modelKeyword = parts[0];
          const qty = parseInt(parts[1], 10) || 1;

          const matched = (quoteProductsCache || []).find(p =>
            p.model_no && p.model_no.toLowerCase().includes(modelKeyword.toLowerCase())
          );

          if (matched) {
            addBlankItem({
              model_no: matched.model_no,
              spec: matched.spec || matched.series || "",
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
        });

        closeModal();
        calculateTotal();
        alert(`[BOM 匯入成功] 已成功批次匯入 ${imported} 項產品！`);
      };
    }
  }

  // 從歷史報價單匯入對話框處理
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
    listEl.innerHTML = '<div style="padding:16px; text-align:center; color:var(--text-muted);">正在載入歷史報價清單...</div>';

    try {
      const resp = await fetch("/api/quotations");
      if (!resp.ok) throw new Error("讀取歷史報價單失敗");
      const quotes = await resp.json();

      if (!Array.isArray(quotes) || quotes.length === 0) {
        listEl.innerHTML = '<div style="padding:16px; text-align:center; color:var(--text-muted);">目前無過往報價單可供匯入</div>';
        return;
      }

      listEl.innerHTML = quotes.map(q => `
        <div class="import-quote-row" data-quoteno="${escapeHtml(q.quote_no)}" style="padding:10px 14px; border:1px solid var(--border-color); border-radius:6px; background:var(--bg-card); cursor:pointer; display:flex; justify-content:space-between; align-items:center;">
          <div>
            <div style="font-weight:700; color:var(--primary-color, #2563eb); font-family:monospace;">${escapeHtml(q.quote_no)}</div>
            <div style="font-size:0.85rem; color:var(--text-main); font-weight:600; margin-top:2px;">${escapeHtml(q.customer_name)} ｜ ${escapeHtml(q.project_name || '設備報價')}</div>
            <div style="font-size:0.75rem; color:var(--text-muted); margin-top:2px;">${escapeHtml(q.quote_date)} ｜ ${escapeHtml(q.items_count || 0)} 項產品</div>
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
      container.innerHTML = '<span class="text-muted" style="font-size:0.82rem;" id="noAttachmentsHint">尚未附加圖片 (可上傳產品實機照片、尺寸圖面供客戶參考)</span>';
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

    const desktopTableBody = document.getElementById("quoteDesktopTableBody");
    const dRows = desktopTableBody ? desktopTableBody.querySelectorAll('.desktop-item-row') : [];

    if (dRows.length === 0) {
      alert("[無品項可匯出] 報價單中目前無品項！");
      return;
    }

    let csvContent = "\uFEFF"; // UTF-8 BOM
    csvContent += "士林電機 FA 業務報價單\n";
    csvContent += `客戶名稱,${custName},報價日期,${qDate}\n`;
    csvContent += `專案名稱,${projName},統一編號,${document.getElementById("quoteTaxId")?.value || ""}\n`;
    csvContent += `聯絡人,${document.getElementById("quoteContact")?.value || ""},電話,${document.getElementById("quotePhone")?.value || ""}\n`;
    csvContent += `付款條件,${document.getElementById("quotePaymentTerms")?.value || ""},有效天數,${document.getElementById("quoteValidDays")?.value || "30"}天\n\n`;

    csvContent += "項次,產品型號,規格說明,牌價,折數%,預計單價,數量,小計,備註交期\n";

    let subtotal = 0;
    dRows.forEach((r, idx) => {
      const itemId = r.getAttribute('data-itemid');
      const model = document.getElementById(`${itemId}-code`)?.value || "-";
      const spec = document.getElementById(`d-${itemId}-spec`)?.textContent || "";
      const dpStr = (document.getElementById(`d-${itemId}-dealer-price`)?.textContent || "0").replace(/[^0-9.]/g, '');
      const disc = document.getElementById(`d-${itemId}-discount`)?.value || "100";
      const price = parseFloat(document.getElementById(`d-${itemId}-price`)?.value) || 0;
      const qty = parseInt(document.getElementById(`d-${itemId}-qty`)?.value) || 1;
      const rowSub = Math.round(price * qty);
      subtotal += rowSub;
      const delivery = document.getElementById(`d-${itemId}-delivery`)?.value || "";

      csvContent += `${idx + 1},"${model}","${spec}",${dpStr},${disc}%,${price},${qty},${rowSub},"${delivery}"\n`;
    });

    const tax = Math.round(subtotal * 0.05);
    const grandTotal = subtotal + tax;

    csvContent += `\n未稅小計,,,,,${subtotal}\n`;
    csvContent += `營業稅(5%),,,,,${tax}\n`;
    csvContent += `含稅總計,,,,,${grandTotal}\n`;

    const blob = new Blob([csvContent], { type: "text/csv;charset=utf-8;" });
    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = `報價單_${custName}_${qDate}.csv`;
    link.click();
  }

  // 列印 PDF 格式處理
  function printQuotePDF() {
    window.print();
  }

  // 發起專案售價申請 (連動商機與底價)
  function applySpecialPrice() {
    const custName = document.getElementById("customerName")?.value.trim() || "";
    const total = document.getElementById("desktopGrandTotal")?.textContent || "$ 0";
    if (!custName) {
      alert("[請先挑選客戶] 發起專案售價前請先選取客戶名稱！");
      return;
    }
    const panelSpecials = document.getElementById("panelSpecials");
    const panelQuotes = document.getElementById("panelQuotes");
    if (panelSpecials && panelQuotes) {
      panelQuotes.classList.add("hidden");
      panelSpecials.classList.remove("hidden");
      alert(`[已轉入專案售價申請] 客戶：${custName} ｜ 預計總額：${total}\n請於專案售價中心進行底價折讓審核申請。`);
    }
  }

  // ============================================================================
  // 對話盒 (Modals) 控制與資料繫結
  // ============================================================================
  function setupQuotationModals() {
    // ----------------------------------------------------
    // 1. 客戶搜尋 Modal 與雙欄標頭連動
    // ----------------------------------------------------
    const custInput = document.getElementById("customerName");
    const custModal = document.getElementById("customerModal");
    const custClose = document.getElementById("btnCloseCustomerModal");
    const custSearch = document.getElementById("customerModalSearch");
    const custResults = document.getElementById("customerModalResults");
    const custCount = document.getElementById("customerModalCount");

    if (custInput && custModal) {
      custInput.addEventListener("click", () => {
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
               (c.tax_id && c.tax_id.toLowerCase().includes(q)) ||
               (c.phone && c.phone.toLowerCase().includes(q));
      });

      if (custCount) custCount.textContent = list.length;

      if (list.length === 0) {
        custResults.innerHTML = '<div style="padding:16px; text-align:center; color:var(--text-muted); font-size:0.9rem;">查無符合的客戶資料</div>';
        return;
      }

      custResults.innerHTML = list.map(c => {
        const sub = [c.tax_id ? `統編: ${c.tax_id}` : "", c.phone ? `電話: ${c.phone}` : ""].filter(Boolean).join(" ｜ ");
        return `
          <div class="customer-item" data-id="${c.id}" data-name="${escapeHtml(c.name)}">
            <div class="customer-item-name">${escapeHtml(c.name)}</div>
            ${sub ? `<div class="customer-item-sub">${escapeHtml(sub)}</div>` : ''}
          </div>
        `;
      }).join("");

      custResults.querySelectorAll('.customer-item').forEach(el => {
        el.addEventListener("click", () => {
          const cid = parseInt(el.dataset.id, 10);
          const found = customersCache.find(c => c.id === cid);
          if (found) {
            currentSelectedCustomer = found;
            if (custInput) custInput.value = found.name;

            // 連動雙欄式折疊抬頭欄位
            const taxInp = document.getElementById("quoteTaxId");
            if (taxInp) taxInp.value = found.tax_id || "";
            const phoneInp = document.getElementById("quotePhone");
            if (phoneInp) phoneInp.value = found.phone || "";
            const faxInp = document.getElementById("quoteFax");
            if (faxInp) faxInp.value = found.fax || "";
            const termsInp = document.getElementById("quotePaymentTerms");
            if (termsInp) termsInp.value = found.payment_terms || "月結 30 天";
            const addrInp = document.getElementById("quoteDeliveryAddress");
            if (addrInp) addrInp.value = found.address || "";
            const contInp = document.getElementById("quoteContact");
            if (contInp) contInp.value = found.default_contact || "";

            // 更新簡要資訊卡與預覽單號
            const detailCard = document.getElementById("customerDetailCard");
            if (detailCard) detailCard.classList.remove("hidden");
            const sumText = document.getElementById("customerDetailSummary");
            if (sumText) sumText.textContent = found.name + (found.default_contact ? ` - ${found.default_contact}` : "");

            const comp = document.getElementById("cdCompany");
            if (comp) comp.textContent = found.name;
            const tax = document.getElementById("cdTaxId");
            if (tax) tax.textContent = found.tax_id || "-";
            const cont = document.getElementById("cdContact");
            if (cont) cont.textContent = found.default_contact || "-";
            const phone = document.getElementById("cdPhone");
            if (phone) phone.textContent = found.phone || "-";
            const addr = document.getElementById("cdAddress");
            if (addr) addr.textContent = found.address || "-";

            updateQuoteNoPreview();
          }
          if (custModal) custModal.classList.add("hidden");
        });
      });
    }

    // ----------------------------------------------------
    // 2. 產品搜尋 Modal
    // ----------------------------------------------------
    const prodModal = document.getElementById("productModal");
    const prodClose = document.getElementById("btnCloseModal");
    const prodSearch = document.getElementById("productSearch");
    const prodResults = document.getElementById("productResults");

    if (prodClose && prodModal) {
      prodClose.addEventListener("click", () => prodModal.classList.add("hidden"));
    }

    if (prodSearch) {
      prodSearch.addEventListener("input", (e) => {
        renderProductModalList(e.target.value.trim());
      });
    }

    // ----------------------------------------------------
    // 3. 交期 / 備註 Modal
    // ----------------------------------------------------
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

    // ----------------------------------------------------
    // 4. LINE 報價 Modal
    // ----------------------------------------------------
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

    // ----------------------------------------------------
    // 5. 料號切換 Modal
    // ----------------------------------------------------
    const skuModal = document.getElementById("skuModal");
    const skuClose = document.getElementById("btnCloseSkuModal");
    const skuCancel = document.getElementById("btnCancelSkuModal");

    if (skuClose && skuModal) {
      skuClose.addEventListener("click", () => skuModal.classList.add("hidden"));
    }
    if (skuCancel && skuModal) {
      skuCancel.addEventListener("click", () => skuModal.classList.add("hidden"));
    }

    // ----------------------------------------------------
    // 6. 智慧 BOM 匯入對話盒
    // ----------------------------------------------------
    const bomModal = document.getElementById("smartBomModal");
    const btnCloseBom = document.getElementById("btnCloseBomModal");
    const btnCancelBom = document.getElementById("btnCancelBomModal");
    const btnConfirmBom = document.getElementById("btnImportBomConfirm");
    const bomInput = document.getElementById("smartBomInput");

    if (btnCloseBom && bomModal) btnCloseBom.addEventListener("click", () => bomModal.classList.add("hidden"));
    if (btnCancelBom && bomModal) btnCancelBom.addEventListener("click", () => bomModal.classList.add("hidden"));
    if (btnConfirmBom && bomInput) {
      btnConfirmBom.addEventListener("click", () => {
        const text = bomInput.value.trim();
        if (!text) {
          alert("[請貼入 BOM 內容] 內容不得為空！");
          return;
        }
        const lines = text.split("\n");
        let addedCount = 0;
        lines.forEach(line => {
          const trimmed = line.trim();
          if (!trimmed) return;
          const parts = trimmed.split(/[\t,， ]+/);
          const model = parts[0];
          let qty = 1;
          for (let p of parts.slice(1)) {
            const num = parseInt(p, 10);
            if (!isNaN(num) && num > 0) {
              qty = num;
              break;
            }
          }
          if (model) {
            addBlankItem({
              model_no: model,
              quantity: qty,
              spec: "智慧 BOM 批次匯入",
              dealer_price: 0,
              discount_pct: 100,
              final_unit_price: 0,
              note_lead_time: "待確認"
            });
            addedCount++;
          }
        });
        bomInput.value = "";
        if (bomModal) bomModal.classList.add("hidden");
        calculateTotal();
        alert(`[BOM 匯入成功] 已成功匯入 ${addedCount} 筆品項！可點擊「重整經銷價」批次載入定價。`);
      });
    }

    // ----------------------------------------------------
    // 7. 從舊報價單匯入對話盒
    // ----------------------------------------------------
    const oldQuoteModal = document.getElementById("importOldQuoteModal");
    const btnCloseOld = document.getElementById("btnCloseImportOldQuoteModal");
    const btnCancelOld = document.getElementById("btnCancelImportOldQuoteModal");
    const btnConfirmOld = document.getElementById("btnImportOldQuoteConfirm");
    const oldQuoteInput = document.getElementById("importOldQuoteInput");

    if (btnCloseOld && oldQuoteModal) btnCloseOld.addEventListener("click", () => oldQuoteModal.classList.add("hidden"));
    if (btnCancelOld && oldQuoteModal) btnCancelOld.addEventListener("click", () => oldQuoteModal.classList.add("hidden"));
    if (btnConfirmOld && oldQuoteInput) {
      btnConfirmOld.addEventListener("click", () => {
        const text = oldQuoteInput.value.trim();
        if (!text) {
          alert("[請貼入舊單內容] 內容不得為空！");
          return;
        }
        const lines = text.split("\n");
        let addedCount = 0;
        lines.forEach(line => {
          const trimmed = line.trim();
          if (!trimmed) return;
          const parts = trimmed.split(/[\t,，]+/);
          if (parts.length >= 1) {
            const model = parts[0].trim();
            const qty = parts.length >= 2 ? (parseInt(parts[1], 10) || 1) : 1;
            const price = parts.length >= 3 ? (parseFloat(parts[2].replace(/[^0-9.]/g, '')) || 0) : 0;
            if (model) {
              addBlankItem({
                model_no: model,
                quantity: qty,
                spec: "舊單匯入品項",
                dealer_price: price,
                discount_pct: 100,
                final_unit_price: price,
                note_lead_time: "舊單原規格"
              });
              addedCount++;
            }
          }
        });
        oldQuoteInput.value = "";
        if (oldQuoteModal) oldQuoteModal.classList.add("hidden");
        calculateTotal();
        alert(`[舊單匯入成功] 已成功轉入 ${addedCount} 筆品項！`);
      });
    }

    // ----------------------------------------------------
    // 8. 清空品項二次確認對話盒
    // ----------------------------------------------------
    const clearModal = document.getElementById("clearConfirmModal");
    const btnCloseClear = document.getElementById("btnCloseClearModal");
    const btnCancelClear = document.getElementById("btnCancelClearItems");
    const btnConfirmClear = document.getElementById("btnConfirmClearItems");

    if (btnCloseClear && clearModal) btnCloseClear.addEventListener("click", () => clearModal.classList.add("hidden"));
    if (btnCancelClear && clearModal) btnCancelClear.addEventListener("click", () => clearModal.classList.add("hidden"));
    if (btnConfirmClear && clearModal) {
      btnConfirmClear.addEventListener("click", () => {
        const itemsContainer = document.getElementById("itemsContainer");
        if (itemsContainer) itemsContainer.innerHTML = "";
        const desktopTableBody = document.getElementById("quoteDesktopTableBody");
        if (desktopTableBody) desktopTableBody.innerHTML = "";
        quoteItemCount = 0;
        calculateTotal();
        clearModal.classList.add("hidden");
        alert("[已清空] 所有報價品項已清空，可重新建立新清單。");
      });
    }

    // ----------------------------------------------------
    // 9. 軟刪除二次確認對話盒
    // ----------------------------------------------------
    const delModal = document.getElementById("deleteConfirmModal");
    const btnCloseDel = document.getElementById("btnCloseDeleteModal");
    const btnCancelDel = document.getElementById("btnCancelDeleteQuote");
    const btnConfirmDel = document.getElementById("btnConfirmDeleteQuote");

    if (btnCloseDel && delModal) btnCloseDel.addEventListener("click", () => delModal.classList.add("hidden"));
    if (btnCancelDel && delModal) btnCancelDel.addEventListener("click", () => delModal.classList.add("hidden"));
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
          alert(`[已移至垃圾桶] 報價單【${pendingDeleteQuoteNo}】已安全移至垃圾桶！\n若需還原，可點擊「回收站 / 垃圾桶」隨時一鍵復原。`);
          pendingDeleteQuoteNo = null;
          loadQuotationHistory("", false);
        } catch (err) {
          alert("[軟刪除失敗] " + err.message);
        } finally {
          btnConfirmDel.disabled = false;
          btnConfirmDel.textContent = "確定移至垃圾桶";
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

    if (quoteProductsCache.length === 0) {
      try {
        const resp = await fetch("/api/products");
        if (resp.ok) {
          quoteProductsCache = await resp.json();
        }
      } catch (err) {
        console.warn("載入產品庫失敗:", err);
      }
    }
    renderProductModalList("");
  }

  function renderProductModalList(query = "") {
    const listEl = document.getElementById("productResults");
    if (!listEl) return;

    const filtered = (quoteProductsCache || []).filter(p => {
      if (!query) return true;
      const q = query.toLowerCase();
      return (p.model_no && p.model_no.toLowerCase().includes(q)) ||
             (p.spec && p.spec.toLowerCase().includes(q)) ||
             (p.series && p.series.toLowerCase().includes(q));
    });

    if (filtered.length === 0) {
      listEl.innerHTML = '<div style="padding:16px; text-align:center; color:var(--text-muted); font-size:0.9rem;">查無符合的士林產品型號</div>';
      return;
    }

    listEl.innerHTML = filtered.map(p => {
      const price = p.dealer_price ? `$ ${p.dealer_price.toLocaleString()}` : "-";
      return `
        <div class="product-item" data-model="${escapeHtml(p.model_no)}" data-price="${p.dealer_price || 0}" data-spec="${escapeHtml(p.spec || p.series || '')}">
          <div style="display:flex; justify-content:space-between; align-items:center;">
            <div>
              <div class="product-item-name">${escapeHtml(p.model_no)}</div>
              <div class="product-item-sub">${escapeHtml(p.spec || p.series || "士林標準系列")}</div>
            </div>
            <div style="text-align:right;">
              <div style="font-weight:700; color:var(--primary-color, #2563eb); font-size:0.95rem;">${price}</div>
              <span class="stock-badge stock-in" style="margin-top:2px;">在庫確認</span>
            </div>
          </div>
        </div>
      `;
    }).join("");

    listEl.querySelectorAll('.product-item').forEach(el => {
      el.addEventListener("click", () => {
        const model = el.dataset.model;
        const price = parseFloat(el.dataset.price) || 0;
        const spec = el.dataset.spec || "";

        if (currentEditingItemId) {
          // 1. 手機端元件更新
          const titleBtn = document.getElementById(`${currentEditingItemId}-title`);
          const codeInput = document.getElementById(`${currentEditingItemId}-code`);
          const dealerPriceInput = document.getElementById(`${currentEditingItemId}-dealer-price`);
          const copyBtn = document.getElementById(`${currentEditingItemId}-copy-btn`);
          const sugDiscountInput = document.getElementById(`${currentEditingItemId}-sug-discount`);
          const sugPriceInput = document.getElementById(`${currentEditingItemId}-sug-price`);
          const qtyInput = document.getElementById(`${currentEditingItemId}-qty`);
          const finalPriceInput = document.getElementById(`${currentEditingItemId}-final-price`);

          // 2. 電腦端表格元件更新
          const dModelBtn = document.getElementById(`d-${currentEditingItemId}-model`);
          const dSpec = document.getElementById(`d-${currentEditingItemId}-spec`);
          const dDP = document.getElementById(`d-${currentEditingItemId}-dealer-price`);
          const dCopy = document.getElementById(`d-${currentEditingItemId}-copy-btn`);
          const dPrice = document.getElementById(`d-${currentEditingItemId}-price`);
          const dSub = document.getElementById(`d-${currentEditingItemId}-subtotal`);

          if (titleBtn) {
            titleBtn.textContent = model;
            titleBtn.classList.add("selected");
          }
          if (codeInput) codeInput.value = model;
          if (dealerPriceInput) dealerPriceInput.value = `$ ${price.toLocaleString()}`;
          if (copyBtn) copyBtn.classList.remove("hidden");

          if (dModelBtn) {
            dModelBtn.textContent = model;
            dModelBtn.classList.add("selected");
          }
          if (dSpec) dSpec.textContent = spec || "-";
          if (dDP) dDP.textContent = `$ ${price.toLocaleString()}`;
          if (dCopy) dCopy.classList.remove("hidden");

          const disc = parseFloat(sugDiscountInput?.value) || 100;
          const sp = Math.round(price * (disc / 100));
          if (sugPriceInput) sugPriceInput.value = sp;
          if (dPrice) dPrice.value = sp;

          const q = parseInt(qtyInput?.value) || 1;
          const sub = Math.round(sp * q);
          if (finalPriceInput) finalPriceInput.value = sub > 0 ? sub.toLocaleString() : "";
          if (dSub) dSub.textContent = sub > 0 ? `$ ${sub.toLocaleString()}` : "$ 0";

          const mRow = document.getElementById(currentEditingItemId);
          if (mRow) mRow.dataset.spec = spec;

          calculateTotal();
        }

        const modal = document.getElementById("productModal");
        if (modal) modal.classList.add("hidden");
      });
    });
  }

  function openDeliveryModal() {
    const modal = document.getElementById("deliveryModal");
    if (modal) modal.classList.remove("hidden");
  }

  function openSkuModal(itemId) {
    const modal = document.getElementById("skuModal");
    const title = document.getElementById("skuModalModelTitle");
    const list = document.getElementById("skuModalList");
    const model = document.getElementById(`${itemId}-code`)?.value || "未指定";

    if (title) title.textContent = `型號：${model}`;
    if (list) {
      list.innerHTML = `
        <div style="padding:10px 12px; border:2px solid var(--primary-color, #2563eb); background:#eff6ff; border-radius:6px; cursor:pointer;">
          <div style="font-weight:700; color:#1e293b;">${escapeHtml(model)} (標準料號)</div>
          <div style="font-size:0.8rem; color:#059669; margin-top:2px;">現貨充足 ｜ 預設主要料號</div>
        </div>
      `;
    }
    if (modal) modal.classList.remove("hidden");
  }

  function openLineQuoteModal() {
    const custName = document.getElementById("customerName")?.value.trim() || "貴客戶";
    const projName = document.getElementById("quoteProjectName")?.value.trim() || "設備報價";
    const qDate = document.getElementById("quoteDate")?.value || new Date().toISOString().split("T")[0];
    const activeSales = localStorage.getItem("shihlin_pwa_active_sales") || "曾維崧";

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
        const model = r.querySelector('[id$="-code"]')?.value || "未指定型號";
        const sp = parseFloat(r.querySelector('[id$="-sug-price"]')?.value) || 0;
        const q = parseInt(r.querySelector('[id$="-qty"]')?.value) || 1;
        const rowSub = Math.round(sp * q);
        subtotal += rowSub;
        const note = r.querySelector('[id$="-delivery"]')?.value || "";
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

  async function saveQuotationDraft() {
    const custName = document.getElementById("customerName")?.value.trim();
    if (!custName) {
      alert("[請填寫客戶名稱] 請點擊客戶名稱欄位進行選取！");
      return;
    }

    const itemsContainer = document.getElementById("itemsContainer");
    const rows = itemsContainer ? itemsContainer.querySelectorAll('.item-row:not(.subtotal-divider-row)') : [];
    if (rows.length === 0) {
      alert("[無品項] 報價單必須包含至少一項產品！");
      return;
    }

    const itemsData = [];
    let subtotal = 0;
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      const model = r.querySelector('[id$="-code"]')?.value;
      if (!model) {
        alert(`[品項未指定] 第 ${i + 1} 項尚未指定產品型號！`);
        return;
      }
      const dpStr = r.querySelector('[id$="-dealer-price"]')?.value.replace(/[^0-9.]/g, '') || "0";
      const dealerPrice = parseFloat(dpStr) || 0;
      const discount = parseFloat(r.querySelector('[id$="-sug-discount"]')?.value) || 100;
      const sp = parseFloat(r.querySelector('[id$="-sug-price"]')?.value) || 0;
      const q = parseInt(r.querySelector('[id$="-qty"]')?.value) || 1;
      const rowSub = Math.round(sp * q);
      subtotal += rowSub;
      const leadTime = r.querySelector('[id$="-delivery"]')?.value || "";

      itemsData.push({
        model_no: model,
        spec: r.dataset.spec || "",
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
    let cleanCust = custName.replace(/有限公司|股份有限公司|企業|實業/g, "").trim();
    let prefixName = cleanCust.slice(0, 4) || "士電";
    let prefixInput = document.getElementById("quoteNoPrefix")?.value.trim();
    let prefix = prefixInput || `${prefixName}${yymm}`;

    const projName = document.getElementById("quoteProjectName")?.value.trim() || "標準設備報價";
    const projType = document.getElementById("quoteProjectType")?.value || "一般報價";
    const qDate = document.getElementById("quoteDate")?.value || now.toISOString().split("T")[0];
    const vDate = document.getElementById("quoteValidDate")?.value || qDate;
    const tax = Math.round(subtotal * 0.05);
    const grandTotal = subtotal + tax;

    const payload = {
      prefix: prefix,
      quote_date: qDate,
      valid_date: vDate,
      customer_name: custName,
      tax_id: document.getElementById("quoteTaxId")?.value.trim() || currentSelectedCustomer?.tax_id || "",
      contact_name: document.getElementById("quoteContact")?.value.trim() || currentSelectedCustomer?.default_contact || "",
      phone: document.getElementById("quotePhone")?.value.trim() || currentSelectedCustomer?.phone || "",
      delivery_address: document.getElementById("quoteDeliveryAddress")?.value.trim() || currentSelectedCustomer?.address || "",
      payment_terms: document.getElementById("quotePaymentTerms")?.value.trim() || "月結 30 天",
      project_name: projName,
      project_type: projType,
      subtotal: subtotal,
      tax_rate: 5.0,
      tax_amount: tax,
      grand_total: grandTotal,
      status: "追蹤中",
      items: itemsData
    };

    const btnSubmit = document.getElementById("btnSubmitDraft");
    if (btnSubmit) {
      btnSubmit.disabled = true;
      btnSubmit.textContent = "正在安全儲存中...";
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
      alert(`[儲存成功] 報價單已成功寫入資料庫！\n指派單號：【${result.quote_no}】\n遵循 Append-Only 保全規範，未覆寫任何舊紀錄。`);

      // 切換至歷史清單檢視
      switchQuoteSubTab("history");
    } catch (err) {
      console.error("儲存報價單失敗:", err);
      alert("[儲存失敗] " + err.message);
    } finally {
      if (btnSubmit) {
        btnSubmit.disabled = false;
        btnSubmit.textContent = "儲存報價單 (Append-Only)";
      }
    }
  }

  async function loadQuotationHistory(keyword = "", isTrash = false) {
    const resultsContainer = document.getElementById("quoteHistoryResults");
    if (!resultsContainer) return;
    resultsContainer.innerHTML = '<div style="padding:24px; text-align:center; color:var(--text-muted); font-size:0.88rem;">正在載入報價紀錄...</div>';

    try {
      const url = `/api/quotations?include_deleted=true${keyword ? `&keyword=${encodeURIComponent(keyword)}` : ''}`;
      const resp = await fetch(url);
      if (!resp.ok) throw new Error("讀取報價單資料失敗");
      const allQuotes = await resp.json();

      const filtered = (allQuotes || []).filter(q => isTrash ? (q.is_deleted === 1) : (!q.is_deleted || q.is_deleted === 0));

      if (filtered.length === 0) {
        resultsContainer.innerHTML = `
          <div style="padding:32px; text-align:center; color:var(--text-muted); font-size:0.9rem;">
            ${isTrash ? "目前回收站中沒有已刪除的報價單。" : "查無符合條件的歷史報價單紀錄。"}
          </div>
        `;
        return;
      }

      resultsContainer.innerHTML = filtered.map(q => `
        <div class="quote-history-card" data-quoteno="${escapeHtml(q.quote_no)}">
          <div style="display:flex; justify-content:space-between; align-items:flex-start; margin-bottom:6px;">
            <div>
              <span style="font-family:monospace; font-weight:700; color:var(--primary-color, #2563eb); font-size:1rem;">${escapeHtml(q.quote_no)}</span>
              ${q.version > 1 ? `<span style="display:inline-block; margin-left:6px; font-size:0.75rem; padding:1px 6px; background:#eff6ff; color:#2563eb; border-radius:4px; font-weight:600;">Rev.${q.version}</span>` : ''}
            </div>
            <div style="display:flex; align-items:center; gap:8px;">
              <span class="status-pill ${q.status === '已成交' ? 'success' : 'info'}">${escapeHtml(q.status || "追蹤中")}</span>
              ${isTrash ? `
                <button type="button" class="btn-restore-quote secondary" data-quoteno="${escapeHtml(q.quote_no)}" style="padding:3px 10px; font-size:0.8rem; background:#10b981; color:#fff; border:none; border-radius:4px; cursor:pointer;">一鍵復原</button>
              ` : `
                <button type="button" class="btn-trash-quote danger" data-quoteno="${escapeHtml(q.quote_no)}" style="padding:3px 8px; font-size:0.8rem; background:transparent; color:#ef4444; border:1px solid #fecaca; border-radius:4px; cursor:pointer;">移至垃圾桶</button>
              `}
            </div>
          </div>
          <div style="font-size:0.92rem; color:var(--text-main, #1e293b); font-weight:600; margin-bottom:4px;">
            ${escapeHtml(q.customer_name)} ｜ ${escapeHtml(q.project_name || "標準設備報價")}
          </div>
          <div style="display:flex; justify-content:space-between; align-items:center; font-size:0.82rem; color:var(--text-muted, #64748b);">
            <span>報價日期：${escapeHtml(q.quote_date)}</span>
            <span style="font-weight:700; color:var(--primary-color, #2563eb); font-size:1.05rem;">$ ${(q.grand_total || 0).toLocaleString()}</span>
          </div>
        </div>
      `).join("");

      // 點擊卡片載入編輯
      resultsContainer.querySelectorAll('.quote-history-card').forEach(card => {
        card.addEventListener("click", (e) => {
          if (e.target.closest('.btn-trash-quote') || e.target.closest('.btn-restore-quote')) return;
          const qNo = card.dataset.quoteno;
          loadQuotationToForm(qNo);
        });
      });

      // 移至垃圾桶按鈕事件
      resultsContainer.querySelectorAll('.btn-trash-quote').forEach(btn => {
        btn.addEventListener("click", (e) => {
          e.stopPropagation();
          pendingDeleteQuoteNo = btn.dataset.quoteno;
          const targetEl = document.getElementById("deleteTargetQuoteNo");
          if (targetEl) targetEl.textContent = pendingDeleteQuoteNo;
          const delModal = document.getElementById("deleteConfirmModal");
          if (delModal) delModal.classList.remove("hidden");
        });
      });

      // 一鍵復原按鈕事件
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

      const header = data.header;
      const items = data.items || [];
      currentBaseQuoteNo = header.quote_no;

      const custInput = document.getElementById("customerName");
      if (custInput) custInput.value = header.customer_name || "";

      const projInput = document.getElementById("quoteProjectName");
      if (projInput) projInput.value = header.project_name || "";

      const projTypeInput = document.getElementById("quoteProjectType");
      if (projTypeInput) projTypeInput.value = header.project_type || "一般報價";

      const dateInput = document.getElementById("quoteDate");
      if (dateInput) dateInput.value = header.quote_date || new Date().toISOString().split("T")[0];

      const validDateInput = document.getElementById("quoteValidDate");
      if (validDateInput) validDateInput.value = header.valid_date || header.quote_date || "";

      const taxIdInput = document.getElementById("quoteTaxId");
      if (taxIdInput) taxIdInput.value = header.tax_id || "";

      const phoneInput = document.getElementById("quotePhone");
      if (phoneInput) phoneInput.value = header.phone || "";

      const faxInput = document.getElementById("quoteFax");
      if (faxInput) faxInput.value = header.fax || "";

      const termsInput = document.getElementById("quotePaymentTerms");
      if (termsInput) termsInput.value = header.payment_terms || "月結 30 天";

      const addrInput = document.getElementById("quoteDeliveryAddress");
      if (addrInput) addrInput.value = header.delivery_address || "";

      const contInput = document.getElementById("quoteContact");
      if (contInput) contInput.value = header.contact_name || "";

      // 提取編號前綴以利版次遞增
      if (header.quote_no) {
        const parts = header.quote_no.split("-");
        const prefixInput = document.getElementById("quoteNoPrefix");
        if (prefixInput && parts.length > 0) prefixInput.value = parts[0];
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

      items.forEach(itm => {
        addBlankItem(itm);
      });

      calculateTotal();
      updateQuoteNoPreview();

      // 切換至編輯子分頁
      switchQuoteSubTab("create");

      alert(`[已載入報價單] 已成功載入單號【${header.quote_no}】！\n若您修改後再次點擊儲存，系統將遵循 Append-Only 規則自動生成新版次 (Rev+1)。`);
    } catch (err) {
      alert("[載入失敗] " + err.message);
    }
  }

  function loadQuotationDraftsInbox() {
    const draftsContainer = document.getElementById("quoteDraftsInboxResults");
    if (!draftsContainer) return;

    let drafts = [];
    try {
      drafts = JSON.parse(localStorage.getItem("shihlin_mobile_quote_drafts") || "[]");
    } catch (e) {
      drafts = [];
    }

    if (drafts.length === 0) {
      draftsContainer.innerHTML = `
        <div style="padding:32px; text-align:center; color:var(--text-muted); font-size:0.9rem;">
          目前草稿匣內無任何暫存草稿。<br>
          <span style="font-size:0.8rem; color:var(--text-muted); display:inline-block; margin-top:6px;">
            在手機上外出訪客時輸入的離線草稿，將自動歸納於此處方便一鍵載入。
          </span>
        </div>
      `;
      return;
    }

    draftsContainer.innerHTML = drafts.map((d, idx) => `
      <div class="quote-history-card" style="margin-bottom:12px;">
        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:6px;">
          <span style="font-weight:700; color:var(--primary-color, #2563eb); font-size:0.95rem;">草稿 #${idx + 1} ｜ ${escapeHtml(d.customer_name || '未命名客戶')}</span>
          <span style="font-size:0.8rem; color:var(--text-muted);">${escapeHtml(d.saved_at || '')}</span>
        </div>
        <div style="font-size:0.85rem; color:var(--text-main); margin-bottom:8px;">
          案場：${escapeHtml(d.project_name || '設備報價')} ｜ 品項數：${d.items ? d.items.length : 0} 項
        </div>
        <div style="display:flex; justify-content:flex-end; gap:8px;">
          <button type="button" class="btn-delete-draft" data-index="${idx}" style="padding:4px 12px; font-size:0.82rem; background:transparent; color:#ef4444; border:1px solid #fecaca; border-radius:4px; cursor:pointer;">刪除草稿</button>
          <button type="button" class="btn-load-draft" data-index="${idx}" style="padding:4px 14px; font-size:0.82rem; background:var(--primary-color, #2563eb); color:#fff; border:none; border-radius:4px; cursor:pointer;">載入至編輯區</button>
        </div>
      </div>
    `).join("");

    draftsContainer.querySelectorAll('.btn-load-draft').forEach(btn => {
      btn.addEventListener("click", () => {
        const idx = parseInt(btn.dataset.index, 10);
        const draft = drafts[idx];
        if (!draft) return;

        const custInput = document.getElementById("customerName");
        if (custInput) custInput.value = draft.customer_name || "";
        const projInput = document.getElementById("quoteProjectName");
        if (projInput) projInput.value = draft.project_name || "";

        const itemsContainer = document.getElementById("itemsContainer");
        if (itemsContainer) itemsContainer.innerHTML = "";
        const desktopTableBody = document.getElementById("quoteDesktopTableBody");
        if (desktopTableBody) desktopTableBody.innerHTML = "";
        quoteItemCount = 0;

        (draft.items || []).forEach(itm => addBlankItem(itm));
        calculateTotal();
        switchQuoteSubTab("create");
        alert("[已載入草稿] 手機草稿已成功轉入主工作區！");
      });
    });

    draftsContainer.querySelectorAll('.btn-delete-draft').forEach(btn => {
      btn.addEventListener("click", () => {
        const idx = parseInt(btn.dataset.index, 10);
        if (!confirm("確定要刪除這筆手機暫存草稿嗎？")) return;
        drafts.splice(idx, 1);
        localStorage.setItem("shihlin_mobile_quote_drafts", JSON.stringify(drafts));
        loadQuotationDraftsInbox();
      });
    });
  }

  async function runPriceAnalysis() {
    const resultsContainer = document.getElementById("priceAnalysisResults");
    if (!resultsContainer) return;

    const modelQuery = document.getElementById("analysisModelSearch")?.value.trim().toLowerCase() || "";
    const custQuery = document.getElementById("analysisCustomerSearch")?.value.trim().toLowerCase() || "";

    resultsContainer.innerHTML = '<div style="padding:24px; text-align:center; color:var(--text-muted); font-size:0.88rem;">正在分析歷史單價與折讓區間...</div>';

    try {
      const resp = await fetch("/api/quotations?include_deleted=false");
      if (!resp.ok) throw new Error("讀取報價紀錄失敗");
      const quotes = await resp.json();

      let matchedItems = [];
      for (let q of (quotes || [])) {
        if (custQuery && !q.customer_name?.toLowerCase().includes(custQuery)) continue;
        const detailResp = await fetch(`/api/quotations/${encodeURIComponent(q.quote_no)}`);
        if (detailResp.ok) {
          const detail = await detailResp.json();
          for (let itm of (detail.items || [])) {
            if (modelQuery && !itm.model_no?.toLowerCase().includes(modelQuery)) continue;
            matchedItems.push({
              quote_no: q.quote_no,
              customer_name: q.customer_name,
              quote_date: q.quote_date,
              model_no: itm.model_no,
              dealer_price: itm.dealer_price || 0,
              discount_pct: itm.discount_pct || 100,
              final_unit_price: itm.final_unit_price || 0,
              quantity: itm.quantity || 1
            });
          }
        }
      }

      if (matchedItems.length === 0) {
        resultsContainer.innerHTML = `
          <div style="padding:32px; text-align:center; color:var(--text-muted); font-size:0.9rem;">
            查無相符的歷史單價紀錄。請嘗試調整型號或客戶關鍵字。
          </div>
        `;
        return;
      }

      // 計算統計分析指標
      const prices = matchedItems.map(m => m.final_unit_price).filter(p => p > 0);
      const discounts = matchedItems.map(m => m.discount_pct).filter(d => d > 0);
      const maxPrice = prices.length ? Math.max(...prices) : 0;
      const minPrice = prices.length ? Math.min(...prices) : 0;
      const avgDisc = discounts.length ? (discounts.reduce((a, b) => a + b, 0) / discounts.length).toFixed(1) : 100;

      resultsContainer.innerHTML = `
        <div style="display:grid; grid-template-columns:repeat(auto-fit, minmax(140px, 1fr)); gap:12px; margin-bottom:16px;">
          <div style="background:var(--bg-card); border:1px solid var(--border-color); border-radius:8px; padding:12px; text-align:center;">
            <div style="font-size:0.75rem; color:var(--text-muted);">歷史最高單價</div>
            <div style="font-size:1.15rem; font-weight:700; color:#ef4444; margin-top:2px;">$ ${maxPrice.toLocaleString()}</div>
          </div>
          <div style="background:var(--bg-card); border:1px solid var(--border-color); border-radius:8px; padding:12px; text-align:center;">
            <div style="font-size:0.75rem; color:var(--text-muted);">歷史最低單價</div>
            <div style="font-size:1.15rem; font-weight:700; color:#10b981; margin-top:2px;">$ ${minPrice.toLocaleString()}</div>
          </div>
          <div style="background:var(--bg-card); border:1px solid var(--border-color); border-radius:8px; padding:12px; text-align:center;">
            <div style="font-size:0.75rem; color:var(--text-muted);">平均折讓率</div>
            <div style="font-size:1.15rem; font-weight:700; color:var(--primary-color, #2563eb); margin-top:2px;">${avgDisc} %</div>
          </div>
          <div style="background:var(--bg-card); border:1px solid var(--border-color); border-radius:8px; padding:12px; text-align:center;">
            <div style="font-size:0.75rem; color:var(--text-muted);">累計成交/報價筆數</div>
            <div style="font-size:1.15rem; font-weight:700; color:var(--text-main); margin-top:2px;">${matchedItems.length} 筆</div>
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
                <th style="padding:8px 10px; text-align:center;">折讓</th>
                <th style="padding:8px 10px; text-align:right;">成交單價</th>
                <th style="padding:8px 10px; text-align:center;">日期</th>
              </tr>
            </thead>
            <tbody>
              ${matchedItems.map(m => `
                <tr style="border-bottom:1px solid var(--border-color);">
                  <td style="padding:8px 10px; font-family:monospace; font-weight:600; color:var(--primary-color);">${escapeHtml(m.quote_no)}</td>
                  <td style="padding:8px 10px;">${escapeHtml(m.customer_name)}</td>
                  <td style="padding:8px 10px; font-weight:600;">${escapeHtml(m.model_no)}</td>
                  <td style="padding:8px 10px; text-align:right; color:var(--text-muted);">$ ${m.dealer_price.toLocaleString()}</td>
                  <td style="padding:8px 10px; text-align:center; font-weight:600;">${m.discount_pct}%</td>
                  <td style="padding:8px 10px; text-align:right; font-weight:700; color:#10b981;">$ ${m.final_unit_price.toLocaleString()}</td>
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
});

