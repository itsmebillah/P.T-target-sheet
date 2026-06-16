function doGet(e) {
  if (e && e.parameter && e.parameter.asset) {
    return AssetService.getAsset(e.parameter.asset);
  }

  var template = HtmlService.createTemplateFromFile("Index");
  template.pwaAssetBaseUrl = ScriptApp.getService().getUrl() || "";
  return template.evaluate()
    .setTitle("PT CPI - Target & Analytics Portal")
    .addMetaTag("viewport", "width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no");
}

var SHEET_NAMES = {
  USERS: "User_Master",
  HIERARCHY: "Hierarchy_Master",
  PRODUCTS: "Product_Master",
  LEDGER: "Monthly_Ledger",
  SALES: "Sales_Value"
};

var LEDGER_LAYOUT = {
  HEADER_ROWS: 2,
  DATA_START_ROW: 3,
  FIRST_PRODUCT_COLUMN: 10,
  STATIC_HEADERS: [
    "Month",
    "RSM",
    "TSM_Name",
    "SR_Name",
    "TSM_ID",
    "Dealer_Code",
    "Dealer_Name",
    "Territory",
    "Dealer_Target"
  ]
};

function login(userId, password) {
  try {
    var user = AuthenticationService.authenticate(userId, password);
    var sessionToken = AuthenticationService.createSession(user);
    var monthOptions = DateService.getMonthOptions();
    var defaultMonth = DateService.getPreferredMonth();
    var response = {
      ok: true,
      sessionToken: sessionToken,
      user: {
        tsmId: user.tsmId,
        name: user.name,
        area: user.area,
        role: user.role
      },
      monthOptions: monthOptions,
      defaultMonth: defaultMonth
    };

    if (user.role === "tsm") {
      response.matrixData = MatrixService.getTsmMatrix(user, defaultMonth);
    } else if (user.role === "admin") {
      response.adminData = AdminDashboardService.getDashboardData(defaultMonth);
    }

    return response;
  } catch (error) {
    return { ok: false, message: error.message };
  }
}

function getTsmMatrixData(sessionToken, monthKey) {
  try {
    var user = AuthenticationService.requireRole(sessionToken, ["tsm"]);
    return {
      ok: true,
      matrixData: MatrixService.getTsmMatrix(user, monthKey)
    };
  } catch (error) {
    return { ok: false, message: error.message };
  }
}

function saveMonthlyTargets(sessionToken, payloadString) {
  try {
    var user = AuthenticationService.requireRole(sessionToken, ["tsm"]);
    var payload = JSON.parse(payloadString || "{}");
    return LedgerService.saveMonthlyTargets(user, payload);
  } catch (error) {
    return { ok: false, message: error.message };
  }
}

function getAdminDashboardData(sessionToken, monthKey) {
  try {
    AuthenticationService.requireRole(sessionToken, ["admin"]);
    return {
      ok: true,
      adminData: AdminDashboardService.getDashboardData(monthKey)
    };
  } catch (error) {
    return { ok: false, message: error.message };
  }
}

function getSubmissionPrintHtml(sessionToken, contextString) {
  return getSubmissionReportHtml(sessionToken, contextString);
}

function getSubmissionReportHtml(sessionToken, contextString) {
  try {
    var user = AuthenticationService.requireRole(sessionToken, ["tsm"]);
    var context = JSON.parse(contextString || "{}");
    var payload = PrintService.buildSubmissionPayload(user, context);
    return {
      ok: true,
      html: PrintService.renderPrintHtml(payload)
    };
  } catch (error) {
    return { ok: false, message: error.message };
  }
}

function getSubmissionHistory(sessionToken) {
  try {
    var user = AuthenticationService.requireRole(sessionToken, ["tsm"]);
    return {
      ok: true,
      history: LedgerService.getSubmissionHistory(user)
    };
  } catch (error) {
    return { ok: false, message: error.message };
  }
}

var AppContext = (function() {
  function getSpreadsheet() {
    return SpreadsheetApp.getActiveSpreadsheet();
  }

  function getSheetOrThrow(sheetName) {
    var sheet = getSpreadsheet().getSheetByName(sheetName);
    if (!sheet) {
      throw new Error("Missing sheet: " + sheetName);
    }
    return sheet;
  }

  function getSheet(sheetName) {
    return getSpreadsheet().getSheetByName(sheetName);
  }

  function getDataRows(sheetName) {
    var sheet = getSheetOrThrow(sheetName);
    var values = sheet.getDataRange().getValues();
    return values.length > 1 ? values.slice(1) : [];
  }

  function ensureColumns(sheet, requiredColumns) {
    var currentColumns = sheet.getMaxColumns();
    if (currentColumns < requiredColumns) {
      sheet.insertColumnsAfter(currentColumns, requiredColumns - currentColumns);
    }
  }

  function ensureRows(sheet, requiredRows) {
    var currentRows = sheet.getMaxRows();
    if (currentRows < requiredRows) {
      sheet.insertRowsAfter(currentRows, requiredRows - currentRows);
    }
  }

  return {
    getSpreadsheet: getSpreadsheet,
    getSheetOrThrow: getSheetOrThrow,
    getSheet: getSheet,
    getDataRows: getDataRows,
    ensureColumns: ensureColumns,
    ensureRows: ensureRows
  };
})();

var DateService = (function() {
  function getTimeZone() {
    return Session.getScriptTimeZone() || AppContext.getSpreadsheet().getSpreadsheetTimeZone() || "Asia/Dhaka";
  }

  function normalizeMonthKey(value) {
    if (!value) {
      return formatMonth(new Date());
    }

    if (Object.prototype.toString.call(value) === "[object Date]" && !isNaN(value.getTime())) {
      return formatMonth(value);
    }

    var text = String(value).trim();
    if (!text) {
      return formatMonth(new Date());
    }

    if (/^[A-Za-z]{3}-\d{4}$/.test(text)) {
      return text;
    }

    if (/^[A-Za-z]+ \d{4}$/.test(text)) {
      var monthDate = new Date("1 " + text);
      if (!isNaN(monthDate.getTime())) {
        return formatMonth(monthDate);
      }
    }

    if (/^\d{4}-\d{2}$/.test(text)) {
      return formatMonth(new Date(text + "-01T00:00:00"));
    }

    var fallback = new Date(text);
    if (!isNaN(fallback.getTime())) {
      return formatMonth(fallback);
    }

    throw new Error("Invalid month format: " + text);
  }

  function formatMonth(dateObj) {
    return Utilities.formatDate(dateObj, getTimeZone(), "MMM-yyyy");
  }

  function formatTimestamp(dateObj) {
    return Utilities.formatDate(dateObj || new Date(), getTimeZone(), "dd MMM yyyy, hh:mm a");
  }

  function getMonthOptions() {
    var ledgerMonths = LedgerService.getAvailableMonths();
    var options = ledgerMonths.slice();
    var seen = {};
    var now = new Date();
    now.setDate(1);

    for (var i = 0; i < options.length; i++) {
      seen[options[i]] = true;
    }

    for (var current = 0; current <= 5; current++) {
      pushMonth(options, seen, new Date(now.getFullYear(), now.getMonth() + current, 1));
    }

    for (var past = 1; past <= 6; past++) {
      pushMonth(options, seen, new Date(now.getFullYear(), now.getMonth() - past, 1));
    }

    return options;
  }

  function getPreferredMonth() {
    var ledgerMonths = LedgerService.getAvailableMonths();
    if (ledgerMonths.length) {
      return ledgerMonths[0];
    }
    return formatMonth(new Date());
  }

  function monthKeyToDate(monthKey) {
    var normalized = normalizeMonthKey(monthKey);
    var parts = normalized.split("-");
    var monthMap = {
      Jan: 0,
      Feb: 1,
      Mar: 2,
      Apr: 3,
      May: 4,
      Jun: 5,
      Jul: 6,
      Aug: 7,
      Sep: 8,
      Oct: 9,
      Nov: 10,
      Dec: 11
    };
    return new Date(Number(parts[1]), monthMap[parts[0]], 1);
  }

  function compareMonthKeysDesc(left, right) {
    return monthKeyToDate(right).getTime() - monthKeyToDate(left).getTime();
  }

  function pushMonth(target, seen, dateObj) {
    var monthKey = formatMonth(dateObj);
    if (!seen[monthKey]) {
      target.push(monthKey);
      seen[monthKey] = true;
    }
  }

  function monthKeyToReferencePart(monthKey) {
    var parts = normalizeMonthKey(monthKey).split("-");
    var monthMap = {
      Jan: "01",
      Feb: "02",
      Mar: "03",
      Apr: "04",
      May: "05",
      Jun: "06",
      Jul: "07",
      Aug: "08",
      Sep: "09",
      Oct: "10",
      Nov: "11",
      Dec: "12"
    };
    return parts[1] + monthMap[parts[0]];
  }

  return {
    getTimeZone: getTimeZone,
    normalizeMonthKey: normalizeMonthKey,
    formatMonth: formatMonth,
    formatTimestamp: formatTimestamp,
    getMonthOptions: getMonthOptions,
    getPreferredMonth: getPreferredMonth,
    compareMonthKeysDesc: compareMonthKeysDesc,
    monthKeyToReferencePart: monthKeyToReferencePart
  };
})();

var AuthenticationService = (function() {
  var CACHE_PREFIX = "target-sheet-session:";
  var SESSION_TTL_SECONDS = 21600;

  function authenticate(userId, password) {
    var inputId = normalize(userId);
    var inputPassword = normalize(password);

    if (!inputId || !inputPassword) {
      throw new Error("TSM ID and password are required.");
    }

    var users = UserRepository.getUsers();
    for (var i = 0; i < users.length; i++) {
      if (normalize(users[i].tsmId) === inputId && normalize(users[i].password) === inputPassword) {
        if (["admin", "tsm"].indexOf(users[i].role) === -1) {
          throw new Error("Unsupported role: " + users[i].role);
        }
        return users[i];
      }
    }

    throw new Error("Invalid TSM ID or password.");
  }

  function createSession(user) {
    var token = Utilities.getUuid();
    CacheService.getScriptCache().put(CACHE_PREFIX + token, JSON.stringify({
      tsmId: user.tsmId,
      name: user.name,
      area: user.area,
      role: user.role
    }), SESSION_TTL_SECONDS);
    return token;
  }

  function requireRole(token, roles) {
    var user = requireSession(token);
    if (roles.indexOf(user.role) === -1) {
      throw new Error("You are not authorized to perform this action.");
    }
    return user;
  }

  function requireSession(token) {
    if (!token) {
      throw new Error("Session expired. Please log in again.");
    }

    var raw = CacheService.getScriptCache().get(CACHE_PREFIX + token);
    if (!raw) {
      throw new Error("Session expired. Please log in again.");
    }

    return JSON.parse(raw);
  }

  function normalize(value) {
    return value === undefined || value === null ? "" : String(value).trim().toLowerCase();
  }

  return {
    authenticate: authenticate,
    createSession: createSession,
    requireRole: requireRole
  };
})();

var UserRepository = (function() {
  function getUsers() {
    var rows = AppContext.getDataRows(SHEET_NAMES.USERS);
    var users = [];

    for (var i = 0; i < rows.length; i++) {
      if (!rows[i][0]) {
        continue;
      }

      users.push({
        tsmId: cleanString(rows[i][0]),
        password: cleanString(rows[i][1] || rows[i][0]),
        name: cleanString(rows[i][2]),
        area: cleanString(rows[i][3]),
        role: cleanString(rows[i][4]).toLowerCase()
      });
    }

    return users;
  }

  function getActiveTsms() {
    var users = getUsers();
    var hierarchyRows = HierarchyService.getAllRows();
    var rsmByTsm = {};
    var uniqueTsms = {};
    var result = [];

    for (var i = 0; i < hierarchyRows.length; i++) {
      if (!rsmByTsm[hierarchyRows[i].tsmId]) {
        rsmByTsm[hierarchyRows[i].tsmId] = hierarchyRows[i].rsmName;
      }
    }

    users.forEach(function(user) {
      var tsmKey = cleanString(user.tsmId).toLowerCase();
      if (user.role !== "tsm" || !tsmKey || uniqueTsms[tsmKey]) {
        return;
      }

      uniqueTsms[tsmKey] = true;
      result.push({
        tsmId: user.tsmId,
        name: user.name,
        area: user.area,
        rsmName: rsmByTsm[user.tsmId] || ""
      });
    });

    return result;
  }

  function cleanString(value) {
    return value === undefined || value === null ? "" : String(value).trim();
  }

  return {
    getUsers: getUsers,
    getActiveTsms: getActiveTsms
  };
})();

var HierarchyService = (function() {
  function getAllRows() {
    var rows = AppContext.getDataRows(SHEET_NAMES.HIERARCHY);
    var result = [];

    for (var i = 0; i < rows.length; i++) {
      if (!rows[i][1] || !rows[i][5]) {
        continue;
      }

      result.push({
        rsmName: cleanString(rows[i][0]),
        tsmId: cleanString(rows[i][1]),
        dealerName: cleanString(rows[i][2]),
        srName: cleanString(rows[i][3]),
        territory: cleanString(rows[i][4]),
        dealerCode: cleanString(rows[i][5]),
        assignmentKey: buildAssignmentKey(rows[i][5], rows[i][3])
      });
    }

    return result;
  }

  function getRowsForTsm(tsmId) {
    var allRows = getAllRows();
    var seen = {};
    var result = [];
    var normalizedTsm = cleanString(tsmId).toLowerCase();

    for (var i = 0; i < allRows.length; i++) {
      if (allRows[i].tsmId.toLowerCase() !== normalizedTsm) {
        continue;
      }

      var key = buildAssignmentKey(allRows[i].dealerCode, allRows[i].srName);
      if (!seen[key]) {
        seen[key] = true;
        result.push(allRows[i]);
      }
    }

    return result;
  }

  function getUniqueAssignments() {
    var allRows = getAllRows();
    var seen = {};
    var result = [];

    for (var i = 0; i < allRows.length; i++) {
      var key = [
        cleanString(allRows[i].tsmId).toLowerCase(),
        buildAssignmentKey(allRows[i].dealerCode, allRows[i].srName)
      ].join("|");
      if (!seen[key]) {
        seen[key] = true;
        result.push(allRows[i]);
      }
    }

    return result;
  }

  function buildDealerKey(dealerCode) {
    return cleanString(dealerCode).toLowerCase();
  }

  function buildAssignmentKey(dealerCode, srName) {
    return buildDealerKey(dealerCode) + "|" + cleanString(srName).toLowerCase();
  }

  function cleanString(value) {
    return value === undefined || value === null ? "" : String(value).trim();
  }

  return {
    getAllRows: getAllRows,
    getUniqueAssignments: getUniqueAssignments,
    getRowsForTsm: getRowsForTsm,
    buildDealerKey: buildDealerKey,
    buildAssignmentKey: buildAssignmentKey
  };
})();

var ProductService = (function() {
  function getProducts() {
    var rows = AppContext.getDataRows(SHEET_NAMES.PRODUCTS);
    var products = [];

    for (var i = 0; i < rows.length; i++) {
      if (!rows[i][0] || !rows[i][1]) {
        continue;
      }

      var product = {
        productId: cleanString(rows[i][0]),
        productName: cleanString(rows[i][1]),
        size: cleanString(rows[i][2]),
        dpRate: parseNumber(rows[i][3]),
        tpRate: parseNumber(rows[i][4])
      };
      product.productKey = buildProductKey(product.productName, product.size);
      products.push(product);
    }

    if (!products.length) {
      throw new Error("Product_Master has no products.");
    }

    return products;
  }

  function buildProductKey(productName, size) {
    return cleanString(productName) + "|" + cleanString(size);
  }

  function cleanString(value) {
    return value === undefined || value === null ? "" : String(value).trim();
  }

  function parseNumber(value) {
    var parsed = parseFloat(value);
    return isNaN(parsed) ? 0 : parsed;
  }

  return {
    getProducts: getProducts,
    buildProductKey: buildProductKey
  };
})();

var LedgerService = (function() {
  function ensureStructure() {
    var sheet = AppContext.getSheetOrThrow(SHEET_NAMES.LEDGER);
    var products = ProductService.getProducts();
    var requiredColumns = LEDGER_LAYOUT.FIRST_PRODUCT_COLUMN - 1 + products.length;

    AppContext.ensureColumns(sheet, requiredColumns);
    AppContext.ensureRows(sheet, LEDGER_LAYOUT.DATA_START_ROW);

    var row1 = createBlankRow(requiredColumns);
    var row2 = createBlankRow(requiredColumns);

    for (var i = 0; i < LEDGER_LAYOUT.STATIC_HEADERS.length; i++) {
      row2[i] = LEDGER_LAYOUT.STATIC_HEADERS[i];
    }

    for (var p = 0; p < products.length; p++) {
      var columnIndex = LEDGER_LAYOUT.FIRST_PRODUCT_COLUMN + p - 1;
      row1[columnIndex] = products[p].productName;
      row2[columnIndex] = products[p].size;
    }

    sheet.getRange(1, 1, LEDGER_LAYOUT.HEADER_ROWS, requiredColumns).setValues([row1, row2]);
  }

  function getProductColumnMap() {
    var products = ProductService.getProducts();
    var map = {};

    for (var i = 0; i < products.length; i++) {
      map[products[i].productKey] = LEDGER_LAYOUT.FIRST_PRODUCT_COLUMN + i;
    }

    return map;
  }

  function getLedgerRows() {
    ensureStructure();

    var sheet = AppContext.getSheetOrThrow(SHEET_NAMES.LEDGER);
    var products = ProductService.getProducts();
    var requiredColumns = LEDGER_LAYOUT.FIRST_PRODUCT_COLUMN - 1 + products.length;
    var lastRow = sheet.getLastRow();
    var rows = [];
    var dedupedByMonthDealer = {};
    var orderedKeys = [];

    if (lastRow < LEDGER_LAYOUT.DATA_START_ROW) {
      return rows;
    }

    var values = sheet.getRange(LEDGER_LAYOUT.DATA_START_ROW, 1, lastRow - LEDGER_LAYOUT.DATA_START_ROW + 1, requiredColumns).getValues();
    for (var i = 0; i < values.length; i++) {
      if (!values[i][0] || !values[i][5]) {
        continue;
      }

      var productQuantities = {};
      for (var p = 0; p < products.length; p++) {
        productQuantities[products[p].productKey] = parseNumber(values[i][LEDGER_LAYOUT.FIRST_PRODUCT_COLUMN - 1 + p]);
      }

      var normalizedMonth = safeMonth(values[i][0]);
      if (!normalizedMonth) {
        continue;
      }

      var rowModel = {
        month: normalizedMonth,
        rsmName: cleanString(values[i][1]),
        tsmName: cleanString(values[i][2]),
        srName: cleanString(values[i][3]),
        tsmId: cleanString(values[i][4]),
        dealerCode: cleanString(values[i][5]),
        dealerName: cleanString(values[i][6]),
        territory: cleanString(values[i][7]),
        dealerTarget: parseNumber(values[i][8]),
        productQuantities: productQuantities,
        assignmentKey: HierarchyService.buildAssignmentKey(values[i][5], values[i][3]),
        values: values[i]
      };

      var dedupeKey = buildMonthDealerKey(rowModel.month, rowModel.dealerCode, rowModel.srName);
      if (!dedupedByMonthDealer[dedupeKey]) {
        orderedKeys.push(dedupeKey);
      }
      dedupedByMonthDealer[dedupeKey] = rowModel;
    }

    for (var k = 0; k < orderedKeys.length; k++) {
      rows.push(dedupedByMonthDealer[orderedKeys[k]]);
    }

    return rows;
  }

  function getAvailableMonths() {
    var rows = getLedgerRows();
    var seen = {};
    var months = [];

    for (var i = 0; i < rows.length; i++) {
      if (rows[i].month && !seen[rows[i].month]) {
        months.push(rows[i].month);
        seen[rows[i].month] = true;
      }
    }

    months.sort(DateService.compareMonthKeysDesc);
    return months;
  }

  function getRowsForTsmAndMonth(tsmId, monthKey) {
    var rows = getLedgerRows();
    var normalizedMonth = DateService.normalizeMonthKey(monthKey);
    var normalizedTsm = cleanString(tsmId).toLowerCase();
    var map = {};

    for (var i = 0; i < rows.length; i++) {
      if (rows[i].month === normalizedMonth && rows[i].tsmId.toLowerCase() === normalizedTsm) {
        map[HierarchyService.buildAssignmentKey(rows[i].dealerCode, rows[i].srName)] = rows[i];
      }
    }

    return map;
  }

  function getSubmissionRows(tsmId, monthKey) {
    var rows = getLedgerRows();
    var normalizedMonth = DateService.normalizeMonthKey(monthKey);
    var normalizedTsm = cleanString(tsmId).toLowerCase();
    var results = [];

    for (var i = 0; i < rows.length; i++) {
      if (rows[i].month === normalizedMonth && rows[i].tsmId.toLowerCase() === normalizedTsm) {
        results.push(rows[i]);
      }
    }

    results.sort(function(left, right) {
      return cleanString(left.dealerName).localeCompare(cleanString(right.dealerName));
    });
    return results;
  }

  function saveMonthlyTargets(user, payload) {
    ensureStructure();

    var monthKey = DateService.normalizeMonthKey(payload.month);
    var hierarchyRows = HierarchyService.getRowsForTsm(user.tsmId);
    if (!hierarchyRows.length) {
      throw new Error("No hierarchy rows found for this TSM.");
    }

    var hierarchyByDealer = {};
    for (var h = 0; h < hierarchyRows.length; h++) {
      hierarchyByDealer[HierarchyService.buildAssignmentKey(hierarchyRows[h].dealerCode, hierarchyRows[h].srName)] = hierarchyRows[h];
    }

    var products = ProductService.getProducts();
    var productMap = getProductColumnMap();
    var rows = getLedgerRows();
    var uniqueOrder = [];
    var uniqueMap = {};

    for (var i = 0; i < rows.length; i++) {
      var existingKey = buildMonthDealerKey(rows[i].month, rows[i].dealerCode, rows[i].srName);
      if (!uniqueMap[existingKey]) {
        uniqueOrder.push(existingKey);
      }
      uniqueMap[existingKey] = rows[i].values.slice();
    }

    var submittedDealers = payload.dealers || [];
    for (var d = 0; d < submittedDealers.length; d++) {
      var dealerPayload = submittedDealers[d];
      var dealerKey = HierarchyService.buildAssignmentKey(dealerPayload.dealerCode, dealerPayload.srName);
      var hierarchy = hierarchyByDealer[dealerKey];

      if (!hierarchy) {
        throw new Error("Unauthorized dealer submission for dealer code " + dealerPayload.dealerCode + ".");
      }

      var rowValues = createLedgerRow(monthKey, user, hierarchy, dealerPayload.quantities || {}, products, productMap);
      var mapKey = buildMonthDealerKey(monthKey, hierarchy.dealerCode, hierarchy.srName);

      if (!uniqueMap[mapKey]) {
        uniqueOrder.push(mapKey);
      }
      uniqueMap[mapKey] = rowValues;
    }

    var outputRows = [];
    for (var o = 0; o < uniqueOrder.length; o++) {
      outputRows.push(uniqueMap[uniqueOrder[o]]);
    }

    writeLedgerRows(outputRows);

    var submittedAt = DateService.formatTimestamp(new Date());
    var totalTargetQty = 0;
    for (var s = 0; s < submittedDealers.length; s++) {
      var submissionHierarchy = hierarchyByDealer[HierarchyService.buildAssignmentKey(submittedDealers[s].dealerCode, submittedDealers[s].srName)];
      if (!submissionHierarchy) {
        continue;
      }
      totalTargetQty += createLedgerRow(monthKey, user, submissionHierarchy, submittedDealers[s].quantities || {}, products, productMap)[8];
    }

    var summary = {
      month: monthKey,
      tsmName: user.name,
      area: user.area,
      totalDealersSubmitted: submittedDealers.length,
      totalTargetQty: totalTargetQty,
      submittedAt: submittedAt,
      referenceNumber: buildReferenceNumber(monthKey, user.tsmId)
    };
    SubmissionMetadataService.recordSubmission(user, summary);

    return {
      ok: true,
      message: "Target submitted successfully.",
      summary: summary
    };
  }

  function getSubmissionHistory(user) {
    var rows = getLedgerRows();
    var normalizedTsm = cleanString(user.tsmId).toLowerCase();
    var grouped = {};

    for (var i = 0; i < rows.length; i++) {
      if (rows[i].tsmId.toLowerCase() !== normalizedTsm) {
        continue;
      }

      if (!grouped[rows[i].month]) {
        grouped[rows[i].month] = {
          month: rows[i].month,
          totalDealers: 0,
          totalTargetQty: 0
        };
      }

      grouped[rows[i].month].totalDealers += 1;
      grouped[rows[i].month].totalTargetQty += parseNumber(rows[i].dealerTarget);
    }

    return Object.keys(grouped).sort(DateService.compareMonthKeysDesc).map(function(monthKey) {
      var meta = SubmissionMetadataService.getSubmissionMeta(user.tsmId, monthKey) || {};
      return {
        month: monthKey,
        referenceNumber: meta.referenceNumber || buildReferenceNumber(monthKey, user.tsmId),
        submittedAt: meta.submittedAt || "",
        totalDealers: grouped[monthKey].totalDealers,
        totalTargetQty: grouped[monthKey].totalTargetQty
      };
    });
  }

  function createLedgerRow(monthKey, user, hierarchy, quantities, products, productMap) {
    var width = LEDGER_LAYOUT.FIRST_PRODUCT_COLUMN - 1 + products.length;
    var row = createBlankRow(width);
    var dealerTarget = 0;

    row[0] = monthKey;
    row[1] = hierarchy.rsmName;
    row[2] = user.name;
    row[3] = hierarchy.srName;
    row[4] = user.tsmId;
    row[5] = hierarchy.dealerCode;
    row[6] = hierarchy.dealerName;
    row[7] = hierarchy.territory;

    for (var i = 0; i < products.length; i++) {
      var productKey = products[i].productKey;
      var value = parseNumber(quantities[productKey]);
      dealerTarget += value;
      row[productMap[productKey] - 1] = value;
    }

    row[8] = dealerTarget;
    return row;
  }

  function writeLedgerRows(outputRows) {
    var sheet = AppContext.getSheetOrThrow(SHEET_NAMES.LEDGER);
    var currentDataRows = Math.max(sheet.getLastRow() - LEDGER_LAYOUT.DATA_START_ROW + 1, 0);
    var columnCount = LEDGER_LAYOUT.FIRST_PRODUCT_COLUMN - 1 + ProductService.getProducts().length;
    var requiredLastRow = LEDGER_LAYOUT.DATA_START_ROW - 1 + Math.max(outputRows.length, 1);

    AppContext.ensureRows(sheet, requiredLastRow);

    if (outputRows.length) {
      sheet.getRange(LEDGER_LAYOUT.DATA_START_ROW, 1, outputRows.length, columnCount).setValues(outputRows);
    }

    if (currentDataRows > outputRows.length) {
      sheet.getRange(LEDGER_LAYOUT.DATA_START_ROW + outputRows.length, 1, currentDataRows - outputRows.length, columnCount).clearContent();
    }
  }

  function buildReferenceNumber(monthKey, tsmId) {
    return "TRG-" + DateService.monthKeyToReferencePart(monthKey) + "-" + cleanString(tsmId).toUpperCase();
  }

  function buildMonthDealerKey(monthKey, dealerCode, srName) {
    return DateService.normalizeMonthKey(monthKey) + "|" + HierarchyService.buildAssignmentKey(dealerCode, srName);
  }

  function createBlankRow(length) {
    var row = [];
    for (var i = 0; i < length; i++) {
      row.push("");
    }
    return row;
  }

  function parseNumber(value) {
    var parsed = parseFloat(value);
    return isNaN(parsed) ? 0 : parsed;
  }

  function cleanString(value) {
    return value === undefined || value === null ? "" : String(value).trim();
  }

  function safeMonth(value) {
    try {
      return DateService.normalizeMonthKey(value);
    } catch (error) {
      return "";
    }
  }

  return {
    ensureStructure: ensureStructure,
    getProductColumnMap: getProductColumnMap,
    getLedgerRows: getLedgerRows,
    getAvailableMonths: getAvailableMonths,
    getRowsForTsmAndMonth: getRowsForTsmAndMonth,
    getSubmissionRows: getSubmissionRows,
    saveMonthlyTargets: saveMonthlyTargets,
    getSubmissionHistory: getSubmissionHistory,
    buildReferenceNumber: buildReferenceNumber
  };
})();

var MatrixService = (function() {
  function getTsmMatrix(user, monthKey) {
    var normalizedMonth = DateService.normalizeMonthKey(monthKey);
    var products = ProductService.getProducts();
    var allRows = HierarchyService.getAllRows().filter(function(row) {
      return row.tsmId.toLowerCase() === String(user.tsmId || "").toLowerCase();
    });
    var dealerColumns = HierarchyService.getRowsForTsm(user.tsmId);
    var savedRows = LedgerService.getRowsForTsmAndMonth(user.tsmId, normalizedMonth);
    var previousEntries = {};
    var dealerTargets = {};
    var alreadySubmitted = false;
    var uniqueSrMap = {};
    var uniqueDealerMap = {};

    for (var i = 0; i < dealerColumns.length; i++) {
      var dealerKey = dealerColumns[i].assignmentKey || HierarchyService.buildAssignmentKey(dealerColumns[i].dealerCode, dealerColumns[i].srName);
      var savedRow = savedRows[dealerKey];
      previousEntries[dealerKey] = {};
      dealerTargets[dealerKey] = 0;
      uniqueSrMap[String(dealerColumns[i].srName || "").toLowerCase()] = true;
      uniqueDealerMap[String(dealerColumns[i].dealerCode || "").toLowerCase()] = true;

      if (savedRow) {
        alreadySubmitted = true;
      }

      for (var p = 0; p < products.length; p++) {
        var value = savedRow ? parseNumber(savedRow.productQuantities[products[p].productKey]) : 0;
        previousEntries[dealerKey][products[p].productKey] = value;
        dealerTargets[dealerKey] += value;
      }
    }

    return {
      month: normalizedMonth,
      products: products,
      dealerColumns: dealerColumns,
      previousEntries: previousEntries,
      dealerTargets: dealerTargets,
      alreadySubmitted: alreadySubmitted,
      debugCounts: {
        totalHierarchyRecordsLoaded: allRows.length,
        totalSrCountLoaded: Object.keys(uniqueSrMap).length,
        totalDealerCountLoaded: Object.keys(uniqueDealerMap).length,
        totalColumnsRendered: dealerColumns.length,
        totalDealersRendered: dealerColumns.length,
        totalSavedRowsForMonth: Object.keys(savedRows).length
      }
    };
  }

  function parseNumber(value) {
    var parsed = parseFloat(value);
    return isNaN(parsed) ? 0 : parsed;
  }

  return {
    getTsmMatrix: getTsmMatrix
  };
})();

var SalesService = (function() {
  function getActualMaps() {
    var sheet = AppContext.getSheet(SHEET_NAMES.SALES);
    if (!sheet) {
      return { bySrMonth: {}, bySrMonthProduct: {}, isAvailable: false, productSupport: false };
    }

    var values = sheet.getDataRange().getValues();
    if (values.length < 2) {
      return { bySrMonth: {}, bySrMonthProduct: {}, isAvailable: false, productSupport: false };
    }

    var headers = values[0].map(function(header) {
      return normalizeHeader(header);
    });
    var columnMap = {
      date: findHeaderIndex(headers, ["date", "sales_date", "invoice_date", "month"]),
      sr: findHeaderIndex(headers, ["sr_name", "sr", "sr_representative", "sales_rep"]),
      qty: findHeaderIndex(headers, ["qty", "quantity", "sales_qty", "actual_sales", "actual_qty", "ctn"]),
      productName: findHeaderIndex(headers, ["product_name", "product", "sku", "item_name", "item"]),
      size: findHeaderIndex(headers, ["size", "pack_size", "pack", "variant"])
    };

    if (columnMap.date === -1 || columnMap.sr === -1 || columnMap.qty === -1) {
      return { bySrMonth: {}, bySrMonthProduct: {}, isAvailable: false, productSupport: false };
    }

    var bySrMonth = {};
    var bySrMonthProduct = {};
    var hasProductSupport = columnMap.productName !== -1;

    for (var i = 1; i < values.length; i++) {
      var srValue = cleanString(values[i][columnMap.sr]);
      if (!srValue) {
        continue;
      }

      var monthKey = safeMonth(values[i][columnMap.date]);
      if (!monthKey) {
        continue;
      }

      var srKey = srValue.toLowerCase();
      var qty = parseNumber(values[i][columnMap.qty]);
      var baseKey = monthKey + "|" + srKey;
      bySrMonth[baseKey] = (bySrMonth[baseKey] || 0) + qty;

      if (hasProductSupport) {
        var productName = cleanString(values[i][columnMap.productName]);
        if (productName) {
          var size = columnMap.size === -1 ? "" : cleanString(values[i][columnMap.size]);
          var productKey = ProductService.buildProductKey(productName, size);
          bySrMonthProduct[baseKey + "|" + productKey] = (bySrMonthProduct[baseKey + "|" + productKey] || 0) + qty;
        }
      }
    }

    return {
      bySrMonth: bySrMonth,
      bySrMonthProduct: bySrMonthProduct,
      isAvailable: true,
      productSupport: hasProductSupport
    };
  }

  function parseNumber(value) {
    var parsed = parseFloat(value);
    return isNaN(parsed) ? 0 : parsed;
  }

  function cleanString(value) {
    return value === undefined || value === null ? "" : String(value).trim();
  }

  function normalizeHeader(value) {
    return cleanString(value).toLowerCase().replace(/[^a-z0-9]+/g, "_");
  }

  function findHeaderIndex(headers, aliases) {
    for (var i = 0; i < aliases.length; i++) {
      var normalizedAlias = aliases[i].toLowerCase();
      var found = headers.indexOf(normalizedAlias);
      if (found !== -1) {
        return found;
      }
    }
    return -1;
  }

  function safeMonth(value) {
    try {
      return DateService.normalizeMonthKey(value);
    } catch (error) {
      return "";
    }
  }

  return {
    getActualMaps: getActualMaps
  };
})();

var AdminDashboardService = (function() {
  function getDashboardData(monthKey) {
    var products = ProductService.getProducts();
    var ledgerRows = LedgerService.getLedgerRows();
    var months = DateService.getMonthOptions();
    var activeTsms = UserRepository.getActiveTsms();
    var hierarchyRows = HierarchyService.getUniqueAssignments();
    var actualMaps = SalesService.getActualMaps();

    return {
      selectedMonth: !monthKey ? DateService.getPreferredMonth() : (monthKey === "ALL" ? "ALL" : DateService.normalizeMonthKey(monthKey)),
      availableMonths: months,
      products: products,
      rawEntries: ledgerRows.map(function(row) {
        return {
          month: row.month,
          rsmName: row.rsmName,
          tsmName: row.tsmName,
          tsmId: row.tsmId,
          srName: row.srName,
          dealerCode: row.dealerCode,
          dealerName: row.dealerName,
          territory: row.territory,
          dealerTarget: row.dealerTarget,
          productQuantities: row.productQuantities
        };
      }),
      allActiveTsms: activeTsms,
      allHierarchyRows: hierarchyRows,
      actualBySrMonth: actualMaps.bySrMonth,
      actualBySrMonthProduct: actualMaps.bySrMonthProduct,
      salesDataAvailable: actualMaps.isAvailable,
      salesProductSupport: actualMaps.productSupport
    };
  }

  return {
    getDashboardData: getDashboardData
  };
})();

var SubmissionMetadataService = (function() {
  var KEY_PREFIX = "submission-meta:";

  function buildKey(tsmId, monthKey) {
    return KEY_PREFIX + cleanString(tsmId).toLowerCase() + ":" + DateService.normalizeMonthKey(monthKey);
  }

  function recordSubmission(user, summary) {
    var data = {
      month: DateService.normalizeMonthKey(summary.month),
      tsmId: user.tsmId,
      tsmName: user.name,
      area: user.area,
      referenceNumber: summary.referenceNumber,
      submittedAt: summary.submittedAt,
      totalDealersSubmitted: summary.totalDealersSubmitted,
      totalTargetQty: summary.totalTargetQty
    };
    PropertiesService.getDocumentProperties().setProperty(buildKey(user.tsmId, summary.month), JSON.stringify(data));
  }

  function getSubmissionMeta(tsmId, monthKey) {
    var raw = PropertiesService.getDocumentProperties().getProperty(buildKey(tsmId, monthKey));
    return raw ? JSON.parse(raw) : null;
  }

  function cleanString(value) {
    return value === undefined || value === null ? "" : String(value).trim();
  }

  return {
    recordSubmission: recordSubmission,
    getSubmissionMeta: getSubmissionMeta
  };
})();

var PrintService = (function() {
  function buildSubmissionPayload(user, context) {
    var monthKey = DateService.normalizeMonthKey(context.month);
    var submissionRows = LedgerService.getSubmissionRows(user.tsmId, monthKey);
    var products = ProductService.getProducts();
    var meta = SubmissionMetadataService.getSubmissionMeta(user.tsmId, monthKey) || {};
    var dealerSummary = [];
    var dealerColumns = [];
    var previousEntries = {};
    var totalTargetQty = 0;

    if (!submissionRows.length) {
      throw new Error("No submission found for " + monthKey + ".");
    }

    for (var i = 0; i < submissionRows.length; i++) {
      var dealer = submissionRows[i];
      dealerColumns.push({
        assignmentKey: dealer.assignmentKey || HierarchyService.buildAssignmentKey(dealer.dealerCode, dealer.srName),
        dealerCode: dealer.dealerCode,
        dealerName: dealer.dealerName,
        srName: dealer.srName,
        territory: dealer.territory
      });
      previousEntries[dealer.assignmentKey || HierarchyService.buildAssignmentKey(dealer.dealerCode, dealer.srName)] = {};
      for (var p = 0; p < products.length; p++) {
        previousEntries[dealer.assignmentKey || HierarchyService.buildAssignmentKey(dealer.dealerCode, dealer.srName)][products[p].productKey] = dealer.productQuantities[products[p].productKey] || 0;
      }
      totalTargetQty += Number(dealer.dealerTarget || 0);
      dealerSummary.push({
        dealerName: dealer.dealerName,
        dealerCode: dealer.dealerCode,
        srName: dealer.srName,
        dealerTarget: Number(dealer.dealerTarget || 0)
      });
    }

    return {
      type: "submission",
      companyName: "PT Consumer Products Industries",
      title: "Monthly Target Submission",
      summary: {
        month: monthKey,
        tsmName: context.tsmName || meta.tsmName || user.name,
        area: context.area || meta.area || user.area,
        referenceNumber: context.referenceNumber || meta.referenceNumber || LedgerService.buildReferenceNumber(monthKey, user.tsmId),
        submittedAt: context.submittedAt || meta.submittedAt || "",
        totalDealersSubmitted: context.totalDealersSubmitted !== undefined ? context.totalDealersSubmitted : dealerSummary.length,
        totalTargetQty: context.totalTargetQty !== undefined ? context.totalTargetQty : totalTargetQty
      },
      dealerSummary: dealerSummary,
      detailedMatrix: {
        products: products,
        dealerColumns: dealerColumns,
        previousEntries: previousEntries
      },
      options: {
        autoPrint: !!context.autoPrint,
        pdfMode: !!context.pdfMode
      }
    };
  }

  function renderPrintHtml(payload) {
    var template = HtmlService.createTemplateFromFile("PrintView");
    template.printPayloadJson = JSON.stringify(payload);
    return template.evaluate().getContent();
  }

  return {
    buildSubmissionPayload: buildSubmissionPayload,
    renderPrintHtml: renderPrintHtml
  };
})();

var AssetService = (function() {
  function getAsset(assetName) {
    if (assetName === "manifest") {
      return ContentService.createTextOutput(JSON.stringify(buildManifest()))
        .setMimeType(ContentService.MimeType.JSON);
    }

    if (assetName === "sw") {
      return ContentService.createTextOutput(buildServiceWorkerScript())
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }

    if (assetName === "icon") {
      return ContentService.createTextOutput(buildIconSvg())
        .setMimeType(ContentService.MimeType.XML);
    }

    return ContentService.createTextOutput("Not found")
      .setMimeType(ContentService.MimeType.TEXT);
  }

  function buildManifest() {
    var baseUrl = ScriptApp.getService().getUrl() || "";
    var iconUrl = baseUrl + "?asset=icon";
    return {
      name: "PT CPI Target Management",
      short_name: "PT CPI Targets",
      description: "Mobile-first target entry and dashboard web app.",
      start_url: baseUrl,
      scope: baseUrl,
      display: "standalone",
      background_color: "#090d16",
      theme_color: "#111827",
      orientation: "portrait-primary",
      icons: [
        {
          src: iconUrl,
          sizes: "192x192",
          type: "image/svg+xml",
          purpose: "any maskable"
        },
        {
          src: iconUrl,
          sizes: "512x512",
          type: "image/svg+xml",
          purpose: "any maskable"
        }
      ]
    };
  }

  function buildServiceWorkerScript() {
    return [
      "const CACHE_NAME = 'pt-cpi-targets-v1';",
      "self.addEventListener('install', event => {",
      "  event.waitUntil((async () => {",
      "    const cache = await caches.open(CACHE_NAME);",
      "    try { await cache.add(self.registration.scope); } catch (error) {}",
      "    self.skipWaiting();",
      "  })());",
      "});",
      "self.addEventListener('activate', event => {",
      "  event.waitUntil((async () => {",
      "    const keys = await caches.keys();",
      "    await Promise.all(keys.filter(key => key !== CACHE_NAME).map(key => caches.delete(key)));",
      "    self.clients.claim();",
      "  })());",
      "});",
      "self.addEventListener('fetch', event => {",
      "  if (event.request.method !== 'GET') { return; }",
      "  event.respondWith((async () => {",
      "    const cache = await caches.open(CACHE_NAME);",
      "    const cached = await cache.match(event.request, { ignoreSearch: false });",
      "    if (cached) { return cached; }",
      "    try {",
      "      const response = await fetch(event.request);",
      "      if (response && response.ok && event.request.url.indexOf('script.google.com') !== -1) {",
      "        cache.put(event.request, response.clone());",
      "      }",
      "      return response;",
      "    } catch (error) {",
      "      const fallback = await cache.match(self.registration.scope);",
      "      if (fallback && event.request.mode === 'navigate') { return fallback; }",
      "      throw error;",
      "    }",
      "  })());",
      "});"
    ].join("\n");
  }

  function buildIconSvg() {
    return [
      '<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 512 512\">',
      '<defs><linearGradient id=\"g\" x1=\"0\" y1=\"0\" x2=\"1\" y2=\"1\">',
      '<stop offset=\"0%\" stop-color=\"#14b8a6\"/>',
      '<stop offset=\"100%\" stop-color=\"#2563eb\"/>',
      '</linearGradient></defs>',
      '<rect width=\"512\" height=\"512\" rx=\"112\" fill=\"#090d16\"/>',
      '<rect x=\"48\" y=\"48\" width=\"416\" height=\"416\" rx=\"96\" fill=\"url(#g)\" opacity=\"0.12\"/>',
      '<path d=\"M142 134h228c20 0 36 16 36 36v172c0 20-16 36-36 36H142c-20 0-36-16-36-36V170c0-20 16-36 36-36z\" fill=\"#0f172a\" stroke=\"#7dd3fc\" stroke-width=\"20\"/>',
      '<path d=\"M170 208h172M170 256h108M170 304h72\" stroke=\"#d1fae5\" stroke-width=\"24\" stroke-linecap=\"round\"/>',
      '<circle cx=\"358\" cy=\"298\" r=\"52\" fill=\"#22c55e\"/>',
      '<path d=\"M336 298l14 14 30-34\" stroke=\"#052e16\" stroke-width=\"18\" stroke-linecap=\"round\" stroke-linejoin=\"round\" fill=\"none\"/>',
      '</svg>'
    ].join("");
  }

  return {
    getAsset: getAsset
  };
})();
