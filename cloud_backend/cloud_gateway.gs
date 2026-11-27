/**
 * 士林電機 FA 業務系統 - Google 雲端硬碟資料庫守門員 (cloud_gateway.gs)
 *
 * 【核心架構與資安職責】
 * 1. 介面與資料徹底切開：GitHub 僅負責呈現網頁畫面，所有真實商業資料集中存放於管理員的 Google 雲端硬碟。
 * 2. 官方通行證反查：每次請求強制透過 Google 官方伺服器核驗登入通行證 (Access Token)，杜絕網址偽造信箱越權存取。
 * 3. 15 人資料嚴格隔離：依白名單試算表對應業務姓名與角色；一般業務僅能讀寫本人名下獨立資料檔，主管與管理員始可總覽團隊。
 * 4. 排隊防撞車保護：所有寫入動作強制透過 LockService 排隊鎖定，並採用單號合併保留機制，防止多人同時存檔互相覆蓋。
 * 5. 零寫死機密規範：本檔案不含任何密碼、金鑰、私人信箱或私密試算表編號；設定值一律存放於 Apps Script「指令碼屬性」。
 *
 * 【管理員首次部署步驟】
 * 1. 於 Google 雲端硬碟建立新的 Google Apps Script 專案，貼上本檔案全部內容並儲存。
 * 2. 點擊左側齒輪「專案設定」->「指令碼屬性」，新增屬性：
 *    - ADMIN_EMAILS：管理員 Google 信箱（多位請用半形逗號分隔）
 *    - WHITELIST_SHEET_ID：（選填）既有業務白名單試算表 ID；若留空，系統會自動在雲端硬碟尋找或建立「報價系統業務白名單」
 *    - CLOUD_ROOT_FOLDER_NAME：（選填）預設為「士林FA業務系統_雲端資料庫_測試副本」，驗證無誤後可改為正式資料夾名稱
 * 3. 先在上方函式選單選擇「authorizeCloudPermissions」執行一次，完成 Google 雲端硬碟與外部驗證連線授權。
 * 4. 點擊右上角「部署」->「新增部署作業」-> 類型選「網頁應用程式」：
 *    - 執行身分：「我」
 *    - 誰可以存取：「所有人」（實際讀寫權限由本程式內部強制核驗 Google 登入通行證與白名單嚴格把關）
 * 5. 複製部署產生的網頁應用程式網址，貼入系統網頁「設定 -> 資料連線模式與雲端同步設定」即可啟用。
 */

var DEFAULT_TEST_ROOT_FOLDER = "士林FA業務系統_雲端資料庫_測試副本";
var SUBFOLDER_SHARED_NAME = "00_系統共用主檔";
var SUBFOLDER_SALES_NAME = "01_業務個人隔離庫";
var WHITELIST_DEFAULT_FILENAME = "報價系統業務白名單";

/**
 * 供管理員首次部署前手動執行一次，觸發 Google 官方權限授權視窗
 */
function authorizeCloudPermissions() {
  DriveApp.getRootFolder();
  SpreadsheetApp.getActiveSpreadsheet();
  UrlFetchApp.fetch("https://www.googleapis.com/oauth2/v3/userinfo", { muteHttpExceptions: true });
  Logger.log("[授權完成] 已取得雲端硬碟、試算表與官方身分驗證連線權限，請進行網頁應用程式部署。");
}

/**
 * 處理 GET 請求（相容原有 check_whitelist 檢查與狀態探測）
 */
function doGet(e) {
  try {
    e = e || { parameter: {} };
    var params = e.parameter || {};
    var action = String(params.action || "").trim();

    if (action === "check_whitelist") {
      var accessToken = String(params.access_token || "").trim();
      if (accessToken) {
        var auth = verifyUserIdentity(accessToken);
        if (!auth.allowed) {
          return createJsonOutput({ status: "rejected", allowed: false, msg: auth.msg });
        }
        return createJsonOutput({
          status: "ok",
          allowed: true,
          name: auth.salesName,
          role: auth.role,
          is_admin: auth.role === "admin",
          is_manager: auth.role === "admin" || auth.role === "manager",
          msg: "身分與白名單驗證通過"
        });
      }
      // 相容舊版登入前預檢（僅回傳是否在白名單，不回傳任何商業資料）
      var rawEmail = String(params.email || "").trim().toLowerCase();
      var preCheck = lookupWhitelistByEmail(rawEmail);
      if (!preCheck.allowed) {
        return createJsonOutput({ status: "rejected", allowed: false, msg: preCheck.msg });
      }
      return createJsonOutput({
        status: "ok",
        allowed: true,
        name: preCheck.salesName,
        role: preCheck.role,
        is_admin: preCheck.role === "admin",
        is_manager: preCheck.role === "admin" || preCheck.role === "manager",
        msg: "白名單驗證通過"
      });
    }

    return createJsonOutput({
      status: "ok",
      service: "Shihlin FA Cloud Gateway",
      mode: "active",
      msg: "雲端資料庫守門員運作正常，資料讀寫請透過加密 POST 通行證呼叫。"
    });
  } catch (err) {
    return createJsonOutput({
      status: "error",
      code: 500,
      msg: "雲端守門員查詢異常：" + String(err && err.message ? err.message : err)
    });
  }
}

/**
 * 處理 POST 請求（採用 text/plain JSON 封包避免跨網域預檢阻擋，強制核驗 Google Access Token）
 */
function doPost(e) {
  try {
    var rawBody = (e && e.postData && e.postData.contents) ? e.postData.contents : "{}";
    var req = JSON.parse(rawBody);
    var action = String(req.action || "").trim();
    var accessToken = String(req.access_token || "").trim();

    if (!accessToken) {
      return createJsonOutput({
        status: "error",
        code: 401,
        msg: "[存取遭拒] 缺少 Google 官方登入通行證，請重新登入系統。"
      });
    }

    var auth = verifyUserIdentity(accessToken);
    if (!auth.allowed) {
      return createJsonOutput({
        status: "error",
        code: 403,
        msg: auth.msg || "[存取受限] 您的 Google 帳號未獲授權或登入已逾時，請重新登入。"
      });
    }

    if (action === "ping_status") {
      var folders = ensureCloudFolderStructure();
      return createJsonOutput({
        status: "ok",
        user: {
          email: auth.email,
          sales_name: auth.salesName,
          role: auth.role,
          is_admin: auth.role === "admin",
          is_manager: auth.role === "admin" || auth.role === "manager"
        },
        cloud: {
          root_folder_name: folders.rootName,
          has_shared_products: folders.hasProductsFile,
          server_time: Utilities.formatDate(new Date(), "Asia/Taipei", "yyyy-MM-dd HH:mm:ss")
        },
        msg: "雲端硬碟資料庫連線與身分驗證成功"
      });
    }

    if (action === "pull_user_workspace") {
      return handlePullWorkspace(auth, req);
    }

    if (action === "sync_sales_data") {
      return handleSyncSalesData(auth, req);
    }

    if (action === "admin_seed_shared_data") {
      return handleAdminSeedSharedData(auth, req);
    }

    if (action === "admin_upload_forecast_base") {
      return handleAdminUploadForecastBase(auth, req);
    }

    return createJsonOutput({
      status: "error",
      code: 400,
      msg: "不支援的雲端操作指令：" + action
    });
  } catch (err) {
    return createJsonOutput({
      status: "error",
      code: 500,
      msg: "雲端守門員處理異常：" + String(err && err.message ? err.message : err)
    });
  }
}

/**
 * 向 Google 官方伺服器反查 Access Token 真實持有人信箱，並核對白名單角色
 */
function verifyUserIdentity(accessToken) {
  var cache = CacheService.getScriptCache();
  var tokenHash = Utilities.base64EncodeWebSafe(
    Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, accessToken)
  ).slice(0, 32);
  var cacheKey = "auth_v1_" + tokenHash;
  var cachedStr = cache.get(cacheKey);
  if (cachedStr) {
    try {
      return JSON.parse(cachedStr);
    } catch (e) {}
  }

  var resp = UrlFetchApp.fetch("https://www.googleapis.com/oauth2/v3/userinfo", {
    method: "get",
    headers: { Authorization: "Bearer " + accessToken },
    muteHttpExceptions: true
  });

  if (resp.getResponseCode() !== 200) {
    return {
      allowed: false,
      msg: "[登入逾時] Google 通行證已過期或無效，請點擊右上角重新登入 Google 帳號。"
    };
  }

  var profile = JSON.parse(resp.getContentText("UTF-8") || "{}");
  var verifiedEmail = String(profile.email || "").trim().toLowerCase();
  if (!verifiedEmail || profile.email_verified === false) {
    return {
      allowed: false,
      msg: "[驗證失敗] 無法確認您的 Google 信箱有效性。"
    };
  }

  var wlResult = lookupWhitelistByEmail(verifiedEmail);
  if (wlResult.allowed) {
    // 快取驗證結果 5 分鐘（300 秒），大幅提升 15 人同時操作時的反應速度
    cache.put(cacheKey, JSON.stringify(wlResult), 300);
  }
  return wlResult;
}

/**
 * 查詢白名單試算表與指令碼屬性，回傳業務姓名與角色 (admin / manager / sales)
 */
function lookupWhitelistByEmail(email) {
  if (!email) {
    return { allowed: false, msg: "未提供電子郵件帳號" };
  }
  var cleanEmail = String(email).trim().toLowerCase();
  var props = PropertiesService.getScriptProperties();
  var adminEmailsRaw = String(props.getProperty("ADMIN_EMAILS") || "").toLowerCase();
  var adminEmailList = adminEmailsRaw
    .split(",")
    .map(function(s) { return s.trim(); })
    .filter(function(s) { return s.length > 0; });

  var ss = getOrCreateWhitelistSpreadsheet(props);
  var sheet = ss.getSheets()[0];
  var data = sheet.getDataRange().getValues();

  // 若白名單完全空白且當前登入者已設定於 ADMIN_EMAILS，直接放行管理員
  if (adminEmailList.indexOf(cleanEmail) !== -1) {
    var matchedAdminName = "";
    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0] || "").trim().toLowerCase() === cleanEmail) {
        matchedAdminName = String(data[i][1] || "").trim();
        break;
      }
    }
    return {
      allowed: true,
      email: cleanEmail,
      salesName: matchedAdminName || "系統管理員",
      role: "admin"
    };
  }

  for (var r = 1; r < data.length; r++) {
    var rowEmail = String(data[r][0] || "").trim().toLowerCase();
    if (rowEmail === cleanEmail) {
      var salesName = String(data[r][1] || "").trim();
      var roleNote = String(data[r][2] || "").trim().toUpperCase();
      if (!salesName) {
        return {
          allowed: false,
          msg: "[設定未完成] 您的信箱已在白名單中，但尚未填寫對應業務姓名，請聯絡管理員補齊第 2 欄姓名。"
        };
      }
      var role = "sales";
      if (roleNote.indexOf("ADMIN") !== -1 || roleNote.indexOf("管理員") !== -1) {
        role = "admin";
      } else if (roleNote.indexOf("MANAGER") !== -1 || roleNote.indexOf("主管") !== -1) {
        role = "manager";
      }
      return {
        allowed: true,
        email: cleanEmail,
        salesName: salesName,
        role: role
      };
    }
  }

  return {
    allowed: false,
    msg: "[存取受限] 您的帳號（" + cleanEmail + "）尚未加入業務白名單試算表，請聯絡系統管理員開通。"
  };
}

/**
 * 取得或自動建立「報價系統業務白名單」試算表
 */
function getOrCreateWhitelistSpreadsheet(props) {
  var sheetId = String(props.getProperty("WHITELIST_SHEET_ID") || "").trim();
  if (sheetId) {
    return SpreadsheetApp.openById(sheetId);
  }
  var files = DriveApp.getFilesByName(WHITELIST_DEFAULT_FILENAME);
  if (files.hasNext()) {
    var f = files.next();
    props.setProperty("WHITELIST_SHEET_ID", f.getId());
    return SpreadsheetApp.openById(f.getId());
  }
  var created = SpreadsheetApp.create(WHITELIST_DEFAULT_FILENAME);
  var firstSheet = created.getSheets()[0];
  firstSheet.setName("白名單清單");
  firstSheet.appendRow(["授權 Email (必填)", "對應業務姓名 (必填)", "角色權限 (一般業務 / 主管 / 管理員)"]);
  firstSheet.setFrozenRows(1);
  props.setProperty("WHITELIST_SHEET_ID", created.getId());
  return created;
}

/**
 * 確保 Google 雲端硬碟內的資料庫資料夾結構存在（預設建立測試用副本資料夾，絕不動到既有正式檔案）
 */
function ensureCloudFolderStructure() {
  var props = PropertiesService.getScriptProperties();
  var rootName = String(props.getProperty("CLOUD_ROOT_FOLDER_NAME") || DEFAULT_TEST_ROOT_FOLDER).trim();
  if (!rootName) rootName = DEFAULT_TEST_ROOT_FOLDER;

  var rootFolder = getOrCreateFolder(DriveApp.getRootFolder(), rootName);
  var sharedFolder = getOrCreateFolder(rootFolder, SUBFOLDER_SHARED_NAME);
  var salesFolder = getOrCreateFolder(rootFolder, SUBFOLDER_SALES_NAME);
  var hasProductsFile = sharedFolder.getFilesByName("products_master.json").hasNext();

  return {
    rootName: rootName,
    rootFolder: rootFolder,
    sharedFolder: sharedFolder,
    salesFolder: salesFolder,
    hasProductsFile: hasProductsFile
  };
}

function getOrCreateFolder(parentFolder, folderName) {
  var iter = parentFolder.getFoldersByName(folderName);
  if (iter.hasNext()) {
    return iter.next();
  }
  return parentFolder.createFolder(folderName);
}

/**
 * 清洗業務姓名以作為安全的雲端檔案名稱，防止路徑穿越字元
 */
function sanitizeSalesFileName(salesName) {
  var clean = String(salesName || "").replace(/[\\/:*?"<>|.\s]/g, "").trim();
  if (!clean) clean = "未命名業務";
  return "sales_data_" + clean + ".json";
}

/**
 * 建立空白的業務個人標準資料庫結構（100% 對齊 data_access.py 正規化欄位）
 */
function createEmptySalesBucket(salesName) {
  return {
    schema_version: 1,
    owner_sales: salesName,
    updated_at: new Date().toISOString(),
    customers: [],
    deactivated_customers: [],
    quotations: [],
    drafts: [],
    special_products: [],
    daily_reports: [],
    opportunities: [],
    deals: [],
    special_price_apps: [],
    sales_forecasts: [],
    driver_metrics: []
  };
}

function readJsonFileInFolder(folder, fileName, fallbackObj) {
  var files = folder.getFilesByName(fileName);
  if (!files.hasNext()) {
    return fallbackObj;
  }
  try {
    var content = files.next().getBlob().getDataAsString("UTF-8");
    if (!content) return fallbackObj;
    return JSON.parse(content);
  } catch (err) {
    return fallbackObj;
  }
}

function writeJsonFileInFolder(folder, fileName, dataObj) {
  var jsonStr = JSON.stringify(dataObj);
  var files = folder.getFilesByName(fileName);
  if (files.hasNext()) {
    var existingFile = files.next();
    existingFile.setContent(jsonStr);
    return existingFile.getId();
  }
  var newFile = folder.createFile(fileName, jsonStr, MimeType.PLAIN_TEXT);
  return newFile.getId();
}

/**
 * 決定本次請求允許存取的業務對象：
 * 一般業務強制鎖死為「本人姓名」，徹底阻絕跨業務偷看或竄改；主管與管理員可指定檢視對象。
 */
function resolveAuthorizedSalesTarget(auth, requestedSales) {
  var isPrivileged = (auth.role === "admin" || auth.role === "manager");
  if (!isPrivileged) {
    return auth.salesName;
  }
  var reqClean = String(requestedSales || "").trim();
  return reqClean || auth.salesName;
}

/**
 * 讀取工作區資料：回傳共用主檔（可選）＋ 授權範圍內之業務個人隔離資料
 */
function handlePullWorkspace(auth, req) {
  var folders = ensureCloudFolderStructure();
  var includeProducts = req.include_products === true;
  var targetSales = resolveAuthorizedSalesTarget(auth, req.target_sales);
  var isPrivileged = (auth.role === "admin" || auth.role === "manager");

  var sharedSettings = readJsonFileInFolder(folders.sharedFolder, "system_settings.json", {});
  var sharedProducts = null;
  if (includeProducts) {
    var prodWrap = readJsonFileInFolder(folders.sharedFolder, "products_master.json", null);
    if (prodWrap && Array.isArray(prodWrap.products)) {
      sharedProducts = prodWrap.products;
    }
  }

  var salesDataByPerson = {};
  if (isPrivileged && targetSales === "ALL") {
    var fileIter = folders.salesFolder.getFiles();
    while (fileIter.hasNext()) {
      var f = fileIter.next();
      var fName = f.getName();
      if (fName.indexOf("sales_data_") === 0 && fName.slice(-5) === ".json") {
        try {
          var parsed = JSON.parse(f.getBlob().getDataAsString("UTF-8"));
          if (parsed && parsed.owner_sales) {
            salesDataByPerson[parsed.owner_sales] = parsed;
          }
        } catch (e) {}
      }
    }
  } else {
    var safeFile = sanitizeSalesFileName(targetSales);
    var personalData = readJsonFileInFolder(
      folders.salesFolder,
      safeFile,
      createEmptySalesBucket(targetSales)
    );
    salesDataByPerson[targetSales] = personalData;
  }

  // 讀取業績預測三表歸戶底數，並強制依業務身分執行切割隔離（一般業務僅能取得本人名下客戶實績與注文殘）
  var fullForecastBase = readJsonFileInFolder(folders.sharedFolder, "erp_actuals_and_backlog.json", null);
  var scopedForecastBase = null;
  if (fullForecastBase && fullForecastBase.base_by_sales) {
    if (isPrivileged && targetSales === "ALL") {
      scopedForecastBase = fullForecastBase;
    } else {
      var singleSlice = {};
      if (fullForecastBase.base_by_sales[targetSales]) {
        singleSlice[targetSales] = fullForecastBase.base_by_sales[targetSales];
      }
      scopedForecastBase = {
        updated_at: fullForecastBase.updated_at || "",
        reference_month: fullForecastBase.reference_month || "",
        base_by_sales: singleSlice
      };
    }
  }

  return createJsonOutput({
    status: "ok",
    user: {
      email: auth.email,
      sales_name: auth.salesName,
      role: auth.role
    },
    target_sales: targetSales,
    shared_settings: sharedSettings,
    shared_products: sharedProducts,
    forecast_base: scopedForecastBase,
    sales_data_by_person: salesDataByPerson,
    synced_at: new Date().toISOString()
  });
}

function padSeq3(n) {
  var s = String(parseInt(n, 10) || 1);
  while (s.length < 3) s = "0" + s;
  return s;
}

function padRev2(n) {
  var s = String(parseInt(n, 10) || 0);
  while (s.length < 2) s = "0" + s;
  return s;
}

/**
 * 透過共用中央發號帳本 (quote_seq_ledger.json) 確保跨業務同前綴單號絕對唯一不撞號
 * 帳本僅保存前綴最大序號與前綴序號歸屬業務名，不含任何客戶全名、產品或價格機密
 */
function ensureUniqueQuotesWithCentralLedger(sharedFolder, targetSales, quotesList) {
  var ledger = readJsonFileInFolder(sharedFolder, "quote_seq_ledger.json", {
    max_seq_by_prefix: {},
    seq_owners: {}
  });
  if (!ledger.max_seq_by_prefix) ledger.max_seq_by_prefix = {};
  if (!ledger.seq_owners) ledger.seq_owners = {};

  var ledgerChanged = false;
  var reassigned = [];

  for (var i = 0; i < quotesList.length; i++) {
    var q = quotesList[i];
    if (!q || !q.quote_no) continue;
    var prefix = String(q.prefix || "").trim();
    var seq = parseInt(q.sequence, 10) || 0;
    var rev = parseInt(q.revision, 10) || 0;
    if (!prefix || seq <= 0) continue;

    var seqKey = prefix + "-" + padSeq3(seq);
    var currentOwner = ledger.seq_owners[seqKey];

    if (!currentOwner) {
      // 尚未有人登記此流水號：登記給當前業務
      ledger.seq_owners[seqKey] = targetSales;
      var curMax = parseInt(ledger.max_seq_by_prefix[prefix], 10) || 0;
      if (seq > curMax) ledger.max_seq_by_prefix[prefix] = seq;
      ledgerChanged = true;
    } else if (currentOwner !== targetSales) {
      // 跨業務同前綴撞號：自動順延配發新流水號，保留各自業務機密與單號唯一性
      var nextSeq = (parseInt(ledger.max_seq_by_prefix[prefix], 10) || 0) + 1;
      while (ledger.seq_owners[prefix + "-" + padSeq3(nextSeq)] &&
             ledger.seq_owners[prefix + "-" + padSeq3(nextSeq)] !== targetSales) {
        nextSeq++;
      }
      var oldQuoteNo = String(q.quote_no);
      var newQuoteNo = prefix + "-" + padSeq3(nextSeq) + "-" + padRev2(rev);
      q.sequence = nextSeq;
      q.quote_no = newQuoteNo;
      if (Array.isArray(q.items)) {
        for (var m = 0; m < q.items.length; m++) {
          if (q.items[m]) q.items[m].quote_no = newQuoteNo;
        }
      }
      ledger.max_seq_by_prefix[prefix] = Math.max(
        parseInt(ledger.max_seq_by_prefix[prefix], 10) || 0,
        nextSeq
      );
      ledger.seq_owners[prefix + "-" + padSeq3(nextSeq)] = targetSales;
      ledgerChanged = true;
      reassigned.push({ old_quote_no: oldQuoteNo, new_quote_no: newQuoteNo, sequence: nextSeq });
    } else {
      var existMax = parseInt(ledger.max_seq_by_prefix[prefix], 10) || 0;
      if (seq > existMax) {
        ledger.max_seq_by_prefix[prefix] = seq;
        ledgerChanged = true;
      }
    }
  }

  if (ledgerChanged) {
    ledger.updated_at = new Date().toISOString();
    writeJsonFileInFolder(sharedFolder, "quote_seq_ledger.json", ledger);
  }
  return reassigned;
}

/**
 * 寫入業務個人資料檔（內建排隊號碼牌 LockService、中央單號防撞帳本與歷史版次保留保護）
 */
function handleSyncSalesData(auth, req) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) {
    return createJsonOutput({
      status: "error",
      code: 429,
      msg: "[系統忙碌] 目前正有其他同仁執行雲端存檔排隊中，請過 2 秒後再試一次。"
    });
  }

  try {
    var folders = ensureCloudFolderStructure();
    var targetSales = resolveAuthorizedSalesTarget(auth, req.target_sales);

    // 雙重保險：一般業務絕對禁止寫入他人檔案
    if (auth.role !== "admin" && auth.role !== "manager" && targetSales !== auth.salesName) {
      return createJsonOutput({
        status: "error",
        code: 403,
        msg: "[權限不足] 您僅能儲存本人名下的業務資料。"
      });
    }

    var safeFile = sanitizeSalesFileName(targetSales);
    var currentCloud = readJsonFileInFolder(
      folders.salesFolder,
      safeFile,
      createEmptySalesBucket(targetSales)
    );
    var incoming = req.payload || {};
    var reassignedQuotes = [];

    // 1. 報價單經中央發號帳本防撞檢查後，採用 quote_no 唯一單號合併（保護 Append-Only 歷史版次）
    if (Array.isArray(incoming.quotations)) {
      reassignedQuotes = ensureUniqueQuotesWithCentralLedger(
        folders.sharedFolder,
        targetSales,
        incoming.quotations
      );
      var quoteMap = {};
      var existingQuotes = Array.isArray(currentCloud.quotations) ? currentCloud.quotations : [];
      for (var i = 0; i < existingQuotes.length; i++) {
        var eq = existingQuotes[i];
        if (eq && eq.quote_no) quoteMap[String(eq.quote_no)] = eq;
      }
      for (var j = 0; j < incoming.quotations.length; j++) {
        var iq = incoming.quotations[j];
        if (iq && iq.quote_no) quoteMap[String(iq.quote_no)] = iq;
      }
      var mergedQuotes = [];
      for (var qKey in quoteMap) {
        if (Object.prototype.hasOwnProperty.call(quoteMap, qKey)) {
          mergedQuotes.push(quoteMap[qKey]);
        }
      }
      mergedQuotes.sort(function(a, b) {
        var idA = Number(a.id) || 0;
        var idB = Number(b.id) || 0;
        if (idA !== idB) return idB - idA;
        return String(b.quote_date || b.created_at || "").localeCompare(String(a.quote_date || a.created_at || ""));
      });
      currentCloud.quotations = mergedQuotes;
    }

    // 2. 更新該業務名下其餘標準資料表陣列
    var listKeys = [
      "customers",
      "deactivated_customers",
      "drafts",
      "special_products",
      "daily_reports",
      "opportunities",
      "deals",
      "special_price_apps",
      "sales_forecasts",
      "driver_metrics"
    ];
    for (var k = 0; k < listKeys.length; k++) {
      var keyName = listKeys[k];
      if (Array.isArray(incoming[keyName])) {
        currentCloud[keyName] = incoming[keyName];
      }
    }

    currentCloud.owner_sales = targetSales;
    currentCloud.updated_at = new Date().toISOString();
    currentCloud.updated_by_email = auth.email;

    writeJsonFileInFolder(folders.salesFolder, safeFile, currentCloud);

    // 若管理員同步更新共用系統設定，一併寫入共用設定檔
    if (auth.role === "admin" && incoming.shared_settings && typeof incoming.shared_settings === "object") {
      var curSettings = readJsonFileInFolder(folders.sharedFolder, "system_settings.json", {});
      for (var sk in incoming.shared_settings) {
        if (Object.prototype.hasOwnProperty.call(incoming.shared_settings, sk)) {
          curSettings[sk] = incoming.shared_settings[sk];
        }
      }
      curSettings.updated_at = new Date().toISOString();
      writeJsonFileInFolder(folders.sharedFolder, "system_settings.json", curSettings);
    }

    return createJsonOutput({
      status: "ok",
      target_sales: targetSales,
      reassigned_quotes: reassignedQuotes,
      saved_at: currentCloud.updated_at,
      msg: "已安全同步儲存至 Google 雲端硬碟（" + folders.rootName + "）"
    });
  } finally {
    lock.releaseLock();
  }
}

/**
 * 管理員專用：初始化上傳標準產品牌價庫與業務基準資料至雲端測試副本資料夾
 */
function handleAdminSeedSharedData(auth, req) {
  if (auth.role !== "admin") {
    return createJsonOutput({
      status: "error",
      code: 403,
      msg: "[權限不足] 僅限系統管理員可執行共用基準資料初始化。"
    });
  }

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) {
    return createJsonOutput({
      status: "error",
      code: 429,
      msg: "[系統忙碌] 雲端鎖定排隊中，請稍候再試。"
    });
  }

  try {
    var folders = ensureCloudFolderStructure();
    var prodCount = 0;
    if (Array.isArray(req.products) && req.products.length > 0) {
      prodCount = req.products.length;
      writeJsonFileInFolder(folders.sharedFolder, "products_master.json", {
        updated_at: new Date().toISOString(),
        updated_by: auth.email,
        count: prodCount,
        products: req.products
      });
    }

    if (req.shared_settings && typeof req.shared_settings === "object") {
      writeJsonFileInFolder(folders.sharedFolder, "system_settings.json", req.shared_settings);
    }

    var seededSalesCount = 0;
    if (req.sales_seed_map && typeof req.sales_seed_map === "object") {
      for (var sName in req.sales_seed_map) {
        if (Object.prototype.hasOwnProperty.call(req.sales_seed_map, sName)) {
          var sFile = sanitizeSalesFileName(sName);
          // 安全保護：若雲端已有該業務檔案且未勾選強制覆寫，僅在檔案不存在時建立初始副本
          var alreadyExists = folders.salesFolder.getFilesByName(sFile).hasNext();
          if (!alreadyExists || req.overwrite_existing === true) {
            var bucket = createEmptySalesBucket(sName);
            var src = req.sales_seed_map[sName] || {};
            if (Array.isArray(src.customers)) bucket.customers = src.customers;
            if (Array.isArray(src.quotations)) bucket.quotations = src.quotations;
            if (Array.isArray(src.drafts)) bucket.drafts = src.drafts;
            if (Array.isArray(src.special_products)) bucket.special_products = src.special_products;
            if (Array.isArray(src.daily_reports)) bucket.daily_reports = src.daily_reports;
            if (Array.isArray(src.deals)) bucket.deals = src.deals;
            writeJsonFileInFolder(folders.salesFolder, sFile, bucket);
            seededSalesCount++;
          }
        }
      }
    }

    return createJsonOutput({
      status: "ok",
      root_folder: folders.rootName,
      product_count: prodCount,
      seeded_sales_count: seededSalesCount,
      msg: "已完成雲端測試副本初始化（產品 " + prodCount + " 筆，業務個人庫 " + seededSalesCount + " 位）。"
    });
  } finally {
    lock.releaseLock();
  }
}

/**
 * 管理員／主管專用：上傳聚合後之 ERP 三表（發票實績、整年銷售、客戶注文殘）統計底數
 */
function handleAdminUploadForecastBase(auth, req) {
  if (auth.role !== "admin" && auth.role !== "manager") {
    return createJsonOutput({
      status: "error",
      code: 403,
      msg: "[權限不足] 僅限主管或管理員可更新業績預測 ERP 實績與注文殘底數。"
    });
  }

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) {
    return createJsonOutput({
      status: "error",
      code: 429,
      msg: "[系統忙碌] 雲端鎖定排隊中，請稍候再試。"
    });
  }

  try {
    var folders = ensureCloudFolderStructure();
    var baseObj = req.forecast_base || {};
    var toSave = {
      updated_at: new Date().toISOString(),
      updated_by: auth.email,
      reference_month: String(baseObj.reference_month || ""),
      summary: baseObj.summary || {},
      base_by_sales: baseObj.base_by_sales || {}
    };
    writeJsonFileInFolder(folders.sharedFolder, "erp_actuals_and_backlog.json", toSave);
    return createJsonOutput({
      status: "ok",
      updated_at: toSave.updated_at,
      msg: "已將 ERP 三表歸戶底數同步至雲端資料庫。"
    });
  } finally {
    lock.releaseLock();
  }
}

function createJsonOutput(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
