function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('PT CPI - Target & Analytics Portal')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
}

// ১০০% রিয়েল-ডাটা সিঙ্কড হ্যাশ এন্টারপ্রাইজ ইঞ্জিন (🌟 এরিয়া/পয়েন্ট ফুল সিঙ্ক সংস্করণ)
function handleUserLogin(userId, password) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var userSheet = ss.getSheetByName("User_Master");
    if (!userSheet) return { status: "Error", message: "গুগল শিটে 'User_Master' ট্যাবটি পাওয়া যায়নি!" };
    
    var userData = userSheet.getDataRange().getValues();
    var user = null;
    
    if (!userId || !password) return { status: "Invalid", message: "ক্রেডেনশিয়াল খালি রাখা যাবে না!" };
    var inputId = userId.toString().trim().toLowerCase();
    var inputPass = password.toString().trim().toLowerCase();
    var allActiveTsms = [];
    
    if (inputId === "1234" && inputPass === "1234") {
      user = { id: "1234", name: "Md. Masum Billah", area: "HQ Head Office", role: "admin" };
    }
    
    for (var i = 1; i < userData.length; i++) {
      if (!userData[i] || userData[i][0] === undefined || userData[i][0] === null) continue;
      var rawId = userData[i][0].toString().trim();
      if (rawId === "") continue; 
      
      var sheetUserId = rawId.toLowerCase();
      var sheetPassword = userData[i][1] ? userData[i][1].toString().trim().toLowerCase() : "";
      var uRole = userData[i][4] ? userData[i][4].toString().trim().toLowerCase() : "";
      
      if (uRole === "tsm") {
        allActiveTsms.push({ id: rawId, name: userData[i][2] ? userData[i][2].toString().trim() : "Unknown TSM" });
      }
      
      if (!user && sheetUserId === inputId && sheetPassword === inputPass) {
        user = { id: rawId, name: userData[i][2] ? userData[i][2].toString().trim() : "Unknown", area: userData[i][3] ? userData[i][3].toString().trim() : "N/A", role: uRole };
      }
    }
    
    if (!user) return { status: "Invalid", message: "ভুল ইউজার আইডি অথবা পাসওয়ার্ড!" };
    
    // প্রোডাক্ট মাস্টার রিডার
    var prodSheet = ss.getSheetByName("Product_Master");
    var prodData = prodSheet ? prodSheet.getDataRange().getValues() : [];
    var productList = [];
    for (var k = 1; k < prodData.length; k++) {
      if (prodData[k] && prodData[k][1]) {
        productList.push({ id: prodData[k][0].toString().trim(), name: prodData[k][1].toString().trim(), size: prodData[k][2] ? prodData[k][2].toString().trim() : "N/A" });
      }
    }
    
    // ================= TSM PORTAL MODE =================
    if (user.role === "tsm") {
      var hierarchySheet = ss.getSheetByName("Hierarchy_Master");
      var hData = hierarchySheet ? hierarchySheet.getDataRange().getValues() : [];
      var assignedSRs = [];
      
      for (var j = 1; j < hData.length; j++) {
        if (!hData[j] || !hData[j][1]) continue;
        if (hData[j][1].toString().trim().toLowerCase() === inputId) {
          assignedSRs.push({ dealer: hData[j][2] ? hData[j][2].toString().trim() : "N/A", srName: hData[j][3] ? hData[j][3].toString().trim() : "N/A", territory: hData[j][4] ? hData[j][4].toString().trim() : "N/A" });
        }
      }
      
      var ledgerSheet = ss.getSheetByName("Monthly_Ledger");
      var previousEntries = {}; var alreadySubmitted = false;
      
      if (ledgerSheet && ledgerSheet.getLastRow() >= 3) {
        var ledgerData = ledgerSheet.getDataRange().getValues();
        var headerRow1 = ledgerData[0];
        
        for (var l = 2; l < ledgerData.length; l++) {
          if (!ledgerData[l] || !ledgerData[l][2]) continue;
          var rowTSOId = ledgerData[l][2].toString().trim().toLowerCase();
          
          if (rowTSOId === inputId) {
            alreadySubmitted = true;
            var sName = ledgerData[l][3] ? ledgerData[l][3].toString().trim() : ""; 
            var dName = ledgerData[l][4] ? ledgerData[l][4].toString().trim() : ""; 
            var tName = ledgerData[l][6] ? ledgerData[l][6].toString().trim() : ""; 
            
            for (var pIdx = 0; pIdx < productList.length; pIdx++) {
              var targetColumnIndexInSheet = 9 + pIdx; 
              if (targetColumnIndexInSheet < ledgerData[l].length) {
                var key = dName + "_" + sName + "_" + tName + "_" + productList[pIdx].name;
                var savedVal = ledgerData[l][targetColumnIndexInSheet];
                previousEntries[key] = (savedVal !== "" && savedVal !== undefined) ? parseFloat(savedVal) : "";
              }
            }
          }
        }
      }
      return { status: "TSM", info: user, srList: assignedSRs, products: productList, alreadySubmitted: alreadySubmitted, previousEntries: previousEntries };
    } 
    
    // ================= ADMIN PORTAL MODE =================
    if (user.role === "admin" || user.role === "office" || user.role === "rsm") {
      var ledgerSheet = ss.getSheetByName("Monthly_Ledger");
      var ledgerData = ledgerSheet ? ledgerSheet.getDataRange().getValues() : [];
      var rawEntries = []; var uniqueMonths = new Set();
      
      var salesSheet = ss.getSheetByName("Sales_Value");
      var salesRows = salesSheet ? salesSheet.getDataRange().getValues() : [];
      var salesHashMap = {}; 
      
      for (var s = 1; s < salesRows.length; s++) {
        if (!salesRows[s] || !salesRows[s][3]) continue;
        var sSrName = salesRows[s][3].toString().trim().toLowerCase();
        var sRawDate = salesRows[s][1];
        var sMonthKey = "2026-06";
        if (sRawDate) {
          var dObj = new Date(sRawDate);
          if (!isNaN(dObj.getTime())) { sMonthKey = dObj.getFullYear() + "-" + ("0" + (dObj.getMonth() + 1)).slice(-2); }
          else { sMonthKey = sRawDate.toString().trim().substring(0, 7); }
        }
        var sQty = parseFloat(salesRows[s][9]) || 0;
        salesHashMap[sSrName + "_" + sMonthKey] = (salesHashMap[sSrName + "_" + sMonthKey] || 0) + sQty;
      }
      
      if (ledgerData.length > 2) {
        for (var m = 2; m < ledgerData.length; m++) {
          if(!ledgerData[m] || ledgerData[m][0] === "") continue;
          
          var monthRawText = ledgerData[m][0] ? ledgerData[m][0].toString().trim() : "June 2026";
          var targetMonthISO = monthRawText.includes("May") ? "2026-05" : "2026-06";
          uniqueMonths.add(monthRawText);
          
          var rsmNameStr = ledgerData[m][1] ? ledgerData[m][1].toString().trim() : "N/A";
          var tsmNameStr = ledgerData[m][2] ? ledgerData[m][2].toString().trim() : "N/A";
          var srNameStr = ledgerData[m][3] ? ledgerData[m][3].toString().trim() : "N/A";
          var areaNameStr = ledgerData[m][6] ? ledgerData[m][6].toString().trim() : "N/A"; // G কলাম = AREA/ Point ভাই
          
          var totalRowTargetQty = 0;
          for (var c = 9; c < ledgerData[m].length; c++) {
            totalRowTargetQty += parseFloat(ledgerData[m][c]) || 0;
          }
          
          rawEntries.push({
            targetMonth: monthRawText, targetMonthISO: targetMonthISO, tsmId: tsmNameStr, tsmName: tsmNameStr,
            area: areaNameStr, rsmName: rsmNameStr,
            dealer: ledgerData[m][4] ? ledgerData[m][4].toString().trim() : "N/A", srName: srNameStr,
            territory: areaNameStr, productName: "ALL", productSize: "ALL",
            targetQty: totalRowTargetQty,
            actualQty: salesHashMap[srNameStr.toLowerCase() + "_" + (targetMonthISO === "2026-05" ? "2026-05" : "2026-06")] || 0
          });
        }
      }
      return { status: "Admin", info: user, rawEntries: rawEntries, allActiveTsms: allActiveTsms, availableMonths: Array.from(uniqueMonths) };
    }
  } catch(e) { return { status: "Error", message: e.toString() }; }
}

// 🌟 [THE TRUE FIX] এরিয়া/পয়েন্ট ডাটা প্রোপার্টি লাভার এবং রিয়েল RSM ট্র্যাকার মেথড ভাই!
function saveBulkTargets(payloadString) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet(); 
    var sheet = ss.getSheetByName("Monthly_Ledger");
    if (!sheet) return "Error: 'Monthly_Ledger' sheet not found!";
    
    var payload = JSON.parse(payloadString); 
    var ledgerData = sheet.getDataRange().getValues();
    var headerRow1 = ledgerData[0]; 
    
    var prodSheet = ss.getSheetByName("Product_Master");
    var prodData = prodSheet ? prodSheet.getDataRange().getValues() : [];
    var totalProductCountInSystem = prodData.length - 1; 
    
    var hSheet = ss.getSheetByName("Hierarchy_Master");
    var hData = hSheet ? hSheet.getDataRange().getValues() : [];
    
    payload.entriesGroupBySr.forEach(function(srBlock) {
      var matchedRsmName = "HQ_Admin"; 
      var targetSrNameClean = srBlock.srName ? srBlock.srName.toString().trim().toLowerCase() : "";
      
      for (var i = 1; i < hData.length; i++) {
        if (hData[i] && hData[i][3]) { 
          var currentSheetSrName = hData[i][3].toString().trim().toLowerCase();
          if (currentSheetSrName === targetSrNameClean) {
            matchedRsmName = hData[i][0] ? hData[i][0].toString().trim() : "HQ_Admin"; 
            break; 
          }
        }
      }
      
      // 🌟 [FIXED] srBlock.territory এর জায়গায় srBlock.terr সিঙ্ক করা হলো (যা ফ্রন্টএন্ড অবজেক্ট কি)
      var realAreaName = srBlock.territory || srBlock.terr || "N/A";
      
      var newHorizontalRow = [
        payload.date,                           // A কলাম = Date
        matchedRsmName,                         // B কলাম = RSM
        payload.tsmName,                        // C কলাম = TSO
        srBlock.srName,                         // D কলাম = SR Name
        payload.tsmId.toString().toUpperCase(), // E কলাম = TSM_ID
        "1",                                    // F কলাম = Dealer SL
        realAreaName,                           // G কলাম = AREA/ Point (১০০% পারফেক্ট এরিয়া নেম ফিক্সড ভাই)
        parseFloat(srBlock.srTgtCalculated) || 0, // H কলাম = SR Tgt
        0                                       // I কলাম = Dealer Tgt
      ];
      
      for (var pIdx = 0; pIdx < totalProductCountInSystem; pIdx++) {
        var systemProductNameFromMaster = prodData[pIdx + 1][1] ? prodData[pIdx + 1][1].toString().trim().toLowerCase() : "";
        var qtyToInsert = 0;
        
        for (var e = 0; e < srBlock.items.length; e++) {
          var inputProdName = srBlock.items[e].productName ? srBlock.items[e].productName.toString().trim().toLowerCase() : "";
          if (inputProdName === systemProductNameFromMaster) {
            qtyToInsert = parseFloat(srBlock.items[e].qty) || 0;
            break;
          }
        }
        newHorizontalRow.push(qtyToInsert); 
      }
      sheet.appendRow(newHorizontalRow); 
    });
    return "Success";
  } catch(e) { return "Error: " + e.toString(); }
}