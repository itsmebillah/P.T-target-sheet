function doGet() {
  return HtmlService.createTemplateFromFile("Index")
    .evaluate()
    .setTitle("Target Management System")
    .addMetaTag("viewport", "width=device-width, initial-scale=1.0");
}

var SHEET_NAMES = {
  USERS: "User_Master",
  HIERARCHY: "Hierarchy_Master",
  PRODUCTS: "Product_Master",
  LEDGER: "Monthly_Ledger"
};

var LEDGER = {
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
    var defaultMonth = monthOptions[0];
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
      response.matrix = MatrixService.getTsmMatrix(user, defaultMonth);
    } else if (user.role === "admin") {
      response.adminView = LedgerService.getAdminView(defaultMonth);
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
      month: DateService.normalizeMonthKey(monthKey),
      matrix: MatrixService.getTsmMatrix(user, monthKey)
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

function getAdminLedgerData(sessionToken, monthKey) {
  try {
    AuthenticationService.requireRole(sessionToken, ["admin"]);
    return {
      ok: true,
      adminView: LedgerService.getAdminView(monthKey)
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

  return {
    getSpreadsheet: getSpreadsheet,
    getSheetOrThrow: getSheetOrThrow,
    getDataRows: getDataRows,
    ensureColumns: ensureColumns
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

    if (/^\d{4}-\d{2}$/.test(text)) {
      return formatMonth(new Date(text + "-01T00:00:00"));
    }

    var normalizedText = text.replace(/\s+/g, " ").replace("-", " ");
    var parsed = new Date("1 " + normalizedText);
    if (!isNaN(parsed.getTime())) {
      return formatMonth(parsed);
    }

    throw new Error("Invalid month format: " + text);
  }

  function formatMonth(dateObj) {
    return Utilities.formatDate(dateObj, getTimeZone(), "MMM-yyyy");
  }

  function getMonthOptions() {
    var months = [];
    var seen = {};
    var cursor = new Date();
    cursor.setDate(1);

    for (var future = 0; future <= 5; future++) {
      pushMonth(months, seen, new Date(cursor.getFullYear(), cursor.getMonth() + future, 1));
    }

    for (var past = 1; past <= 6; past++) {
      pushMonth(months, seen, new Date(cursor.getFullYear(), cursor.getMonth() - past, 1));
    }

    return months;
  }

  function pushMonth(target, seen, dateObj) {
    var monthKey = formatMonth(dateObj);
    if (!seen[monthKey]) {
      target.push(monthKey);
      seen[monthKey] = true;
    }
  }

  return {
    getTimeZone: getTimeZone,
    normalizeMonthKey: normalizeMonthKey,
    getMonthOptions: getMonthOptions
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
    var matchedUser = null;

    for (var i = 0; i < users.length; i++) {
      var user = users[i];
      if (normalize(user.tsmId) === inputId && normalize(user.password) === inputPassword) {
        matchedUser = user;
        break;
      }
    }

    if (!matchedUser) {
      throw new Error("Invalid TSM ID or password.");
    }

    if (["admin", "tsm"].indexOf(matchedUser.role) === -1) {
      throw new Error("Unsupported role: " + matchedUser.role);
    }

    return matchedUser;
  }

  function createSession(user) {
    var token = Utilities.getUuid();
    var sessionUser = {
      tsmId: user.tsmId,
      name: user.name,
      area: user.area,
      role: user.role
    };
    CacheService.getScriptCache().put(CACHE_PREFIX + token, JSON.stringify(sessionUser), SESSION_TTL_SECONDS);
    return token;
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

  function requireRole(token, allowedRoles) {
    var user = requireSession(token);
    if (allowedRoles.indexOf(user.role) === -1) {
      throw new Error("You are not authorized to perform this action.");
    }
    return user;
  }

  function normalize(value) {
    return value === undefined || value === null ? "" : String(value).trim().toLowerCase();
  }

  return {
    authenticate: authenticate,
    createSession: createSession,
    requireSession: requireSession,
    requireRole: requireRole
  };
})();

var UserRepository = (function() {
  function getUsers() {
    var rows = AppContext.getDataRows(SHEET_NAMES.USERS);
    var users = [];

    for (var i = 0; i < rows.length; i++) {
      var row = rows[i];
      if (!row[0]) {
        continue;
      }

      users.push({
        tsmId: cleanString(row[0]),
        password: cleanString(row[1] || row[0]),
        name: cleanString(row[2]),
        area: cleanString(row[3]),
        role: cleanString(row[4]).toLowerCase()
      });
    }

    return users;
  }

  function cleanString(value) {
    return value === undefined || value === null ? "" : String(value).trim();
  }

  return {
    getUsers: getUsers
  };
})();

var ProductService = (function() {
  function getProducts() {
    var rows = AppContext.getDataRows(SHEET_NAMES.PRODUCTS);
    var products = [];

    for (var i = 0; i < rows.length; i++) {
      var row = rows[i];
      if (!row[0] || !row[1]) {
        continue;
      }

      var product = {
        productId: cleanString(row[0]),
        productName: cleanString(row[1]),
        size: cleanString(row[2]),
        dpRate: parseNumber(row[3]),
        tpRate: parseNumber(row[4])
      };
      product.productKey = buildProductKey(product.productName, product.size);
      products.push(product);
    }

    if (products.length === 0) {
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

var HierarchyService = (function() {
  function getRowsForTsm(tsmId) {
    var rows = AppContext.getDataRows(SHEET_NAMES.HIERARCHY);
    var filtered = [];
    var seen = {};
    var normalizedTsmId = cleanString(tsmId).toLowerCase();

    for (var i = 0; i < rows.length; i++) {
      var row = rows[i];
      if (cleanString(row[1]).toLowerCase() !== normalizedTsmId) {
        continue;
      }

      var item = {
        rsmName: cleanString(row[0]),
        tsmId: cleanString(row[1]),
        dealerName: cleanString(row[2]),
        srName: cleanString(row[3]),
        territory: cleanString(row[4]),
        dealerCode: cleanString(row[5])
      };

      var key = buildDealerSrKey(item.dealerCode, item.srName);
      if (!seen[key]) {
        filtered.push(item);
        seen[key] = true;
      }
    }

    return filtered;
  }

  function buildDealerSrKey(dealerCode, srName) {
    return cleanString(dealerCode) + "|" + cleanString(srName).toLowerCase();
  }

  function cleanString(value) {
    return value === undefined || value === null ? "" : String(value).trim();
  }

  return {
    getRowsForTsm: getRowsForTsm,
    buildDealerSrKey: buildDealerSrKey
  };
})();

var LedgerService = (function() {
  function ensureStructure() {
    var sheet = AppContext.getSheetOrThrow(SHEET_NAMES.LEDGER);
    var products = ProductService.getProducts();
    var requiredColumns = LEDGER.FIRST_PRODUCT_COLUMN - 1 + products.length;

    AppContext.ensureColumns(sheet, requiredColumns);

    var row1 = createBlankRow(requiredColumns);
    var row2 = createBlankRow(requiredColumns);

    for (var i = 0; i < LEDGER.STATIC_HEADERS.length; i++) {
      row2[i] = LEDGER.STATIC_HEADERS[i];
    }

    for (var p = 0; p < products.length; p++) {
      var columnIndex = LEDGER.FIRST_PRODUCT_COLUMN + p;
      row1[columnIndex - 1] = products[p].productName;
      row2[columnIndex - 1] = products[p].size;
    }

    sheet.getRange(1, 1, LEDGER.HEADER_ROWS, requiredColumns).setValues([row1, row2]);
  }

  function getProductColumnMap() {
    ensureStructure();

    var sheet = AppContext.getSheetOrThrow(SHEET_NAMES.LEDGER);
    var lastColumn = sheet.getLastColumn();
    var width = lastColumn - LEDGER.FIRST_PRODUCT_COLUMN + 1;
    if (width <= 0) {
      return {};
    }

    var headerRows = sheet.getRange(1, LEDGER.FIRST_PRODUCT_COLUMN, 2, width).getValues();
    var productNames = headerRows[0];
    var productSizes = headerRows[1];
    var map = {};

    for (var i = 0; i < width; i++) {
      if (!productNames[i]) {
        continue;
      }
      var key = ProductService.buildProductKey(productNames[i], productSizes[i]);
      map[key] = LEDGER.FIRST_PRODUCT_COLUMN + i;
    }

    return map;
  }

  function getMatrixEntriesForTsm(tsmId, monthKey) {
    ensureStructure();

    var sheet = AppContext.getSheetOrThrow(SHEET_NAMES.LEDGER);
    var lastRow = sheet.getLastRow();
    var lastColumn = sheet.getLastColumn();
    var entries = {};

    if (lastRow < LEDGER.DATA_START_ROW) {
      return entries;
    }

    var values = sheet.getRange(LEDGER.DATA_START_ROW, 1, lastRow - LEDGER.DATA_START_ROW + 1, lastColumn).getValues();
    var productMap = getProductColumnMap();
    var normalizedMonth = DateService.normalizeMonthKey(monthKey);
    var normalizedTsmId = cleanString(tsmId).toLowerCase();
    var productKeys = Object.keys(productMap);

    for (var i = 0; i < values.length; i++) {
      var row = values[i];
      var rowMonth = safeMonth(row[0]);
      var rowTsmId = cleanString(row[4]).toLowerCase();

      if (!rowMonth || rowMonth !== normalizedMonth || rowTsmId !== normalizedTsmId) {
        continue;
      }

      var rowKey = HierarchyService.buildDealerSrKey(row[5], row[3]);
      var productQuantities = {};

      for (var p = 0; p < productKeys.length; p++) {
        var productKey = productKeys[p];
        var columnIndex = productMap[productKey];
        productQuantities[productKey] = parseNumber(row[columnIndex - 1]);
      }

      entries[rowKey] = {
        rowNumber: LEDGER.DATA_START_ROW + i,
        month: rowMonth,
        rsmName: cleanString(row[1]),
        tsmName: cleanString(row[2]),
        srName: cleanString(row[3]),
        tsmId: cleanString(row[4]),
        dealerCode: cleanString(row[5]),
        dealerName: cleanString(row[6]),
        territory: cleanString(row[7]),
        dealerTarget: parseNumber(row[8]),
        productQuantities: productQuantities
      };
    }

    return entries;
  }

  function saveMonthlyTargets(user, payload) {
    ensureStructure();

    var monthKey = DateService.normalizeMonthKey(payload.month);
    var hierarchyRows = HierarchyService.getRowsForTsm(user.tsmId);
    if (hierarchyRows.length === 0) {
      throw new Error("No hierarchy rows found for this TSM.");
    }

    var hierarchyMap = {};
    for (var i = 0; i < hierarchyRows.length; i++) {
      hierarchyMap[HierarchyService.buildDealerSrKey(hierarchyRows[i].dealerCode, hierarchyRows[i].srName)] = hierarchyRows[i];
    }

    var products = ProductService.getProducts();
    var productColumnMap = getProductColumnMap();
    var existingEntries = getMatrixEntriesForTsm(user.tsmId, monthKey);
    var sheet = AppContext.getSheetOrThrow(SHEET_NAMES.LEDGER);
    var lastColumn = sheet.getLastColumn();
    var rows = payload.rows || [];
    var appendRows = [];
    var updateRows = [];

    for (var r = 0; r < rows.length; r++) {
      var payloadRow = rows[r];
      var rowKey = HierarchyService.buildDealerSrKey(payloadRow.dealerCode, payloadRow.srName);
      var hierarchy = hierarchyMap[rowKey];

      if (!hierarchy) {
        throw new Error("Unauthorized dealer row submitted: " + payloadRow.dealerCode + " / " + payloadRow.srName);
      }

      var ledgerRow = createBlankRow(lastColumn);
      var rowTotal = 0;

      ledgerRow[0] = monthKey;
      ledgerRow[1] = hierarchy.rsmName;
      ledgerRow[2] = user.name;
      ledgerRow[3] = hierarchy.srName;
      ledgerRow[4] = user.tsmId;
      ledgerRow[5] = hierarchy.dealerCode;
      ledgerRow[6] = hierarchy.dealerName;
      ledgerRow[7] = hierarchy.territory;

      var quantities = payloadRow.quantities || {};
      for (var p = 0; p < products.length; p++) {
        var product = products[p];
        var value = parseNumber(quantities[product.productKey]);
        var columnIndex = productColumnMap[product.productKey];
        if (columnIndex) {
          ledgerRow[columnIndex - 1] = value;
        }
        rowTotal += value;
      }

      ledgerRow[8] = rowTotal;

      if (existingEntries[rowKey]) {
        updateRows.push({ rowNumber: existingEntries[rowKey].rowNumber, values: [ledgerRow] });
      } else {
        appendRows.push(ledgerRow);
      }
    }

    for (var u = 0; u < updateRows.length; u++) {
      sheet.getRange(updateRows[u].rowNumber, 1, 1, lastColumn).setValues(updateRows[u].values);
    }

    if (appendRows.length > 0) {
      sheet.getRange(sheet.getLastRow() + 1, 1, appendRows.length, lastColumn).setValues(appendRows);
    }

    return {
      ok: true,
      message: "Targets saved successfully.",
      savedRows: rows.length,
      month: monthKey
    };
  }

  function getAdminView(monthKey) {
    ensureStructure();

    var normalizedMonth = DateService.normalizeMonthKey(monthKey);
    var sheet = AppContext.getSheetOrThrow(SHEET_NAMES.LEDGER);
    var lastRow = sheet.getLastRow();
    var lastColumn = sheet.getLastColumn();
    var rows = [];
    var totalTarget = 0;

    if (lastRow >= LEDGER.DATA_START_ROW) {
      var values = sheet.getRange(LEDGER.DATA_START_ROW, 1, lastRow - LEDGER.DATA_START_ROW + 1, lastColumn).getValues();
      for (var i = 0; i < values.length; i++) {
        var row = values[i];
        if (!row[0]) {
          continue;
        }

        var rowMonth = safeMonth(row[0]);
        if (!rowMonth || rowMonth !== normalizedMonth) {
          continue;
        }

        var target = parseNumber(row[8]);
        totalTarget += target;

        rows.push({
          month: rowMonth,
          rsmName: cleanString(row[1]),
          tsmName: cleanString(row[2]),
          srName: cleanString(row[3]),
          tsmId: cleanString(row[4]),
          dealerCode: cleanString(row[5]),
          dealerName: cleanString(row[6]),
          territory: cleanString(row[7]),
          dealerTarget: target
        });
      }
    }

    return {
      month: normalizedMonth,
      monthOptions: DateService.getMonthOptions(),
      rowCount: rows.length,
      totalTarget: totalTarget,
      rows: rows
    };
  }

  function createBlankRow(length) {
    var row = [];
    for (var i = 0; i < length; i++) {
      row.push("");
    }
    return row;
  }

  function cleanString(value) {
    return value === undefined || value === null ? "" : String(value).trim();
  }

  function parseNumber(value) {
    var parsed = parseFloat(value);
    return isNaN(parsed) ? 0 : parsed;
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
    getMatrixEntriesForTsm: getMatrixEntriesForTsm,
    saveMonthlyTargets: saveMonthlyTargets,
    getAdminView: getAdminView
  };
})();

var MatrixService = (function() {
  function getTsmMatrix(user, monthKey) {
    var normalizedMonth = DateService.normalizeMonthKey(monthKey);
    var hierarchyRows = HierarchyService.getRowsForTsm(user.tsmId);
    var products = ProductService.getProducts();
    var existingEntries = LedgerService.getMatrixEntriesForTsm(user.tsmId, normalizedMonth);
    var rows = [];

    for (var i = 0; i < hierarchyRows.length; i++) {
      var hierarchy = hierarchyRows[i];
      var rowKey = HierarchyService.buildDealerSrKey(hierarchy.dealerCode, hierarchy.srName);
      var existingEntry = existingEntries[rowKey];
      var productValues = {};
      var dealerTarget = 0;

      for (var p = 0; p < products.length; p++) {
        var product = products[p];
        var value = existingEntry ? parseNumber(existingEntry.productQuantities[product.productKey]) : 0;
        productValues[product.productKey] = value;
        dealerTarget += value;
      }

      rows.push({
        rsmName: hierarchy.rsmName,
        dealerCode: hierarchy.dealerCode,
        dealerName: hierarchy.dealerName,
        srName: hierarchy.srName,
        territory: hierarchy.territory,
        dealerTarget: dealerTarget,
        productQuantities: productValues
      });
    }

    return {
      month: normalizedMonth,
      products: products,
      rows: rows
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
