/**
 * Main.gs - Clean Baseline Architecture
 */

function doPost(e) {
  var lock = LockService.getScriptLock();
  var hasLock = false;
  try {
    hasLock = lock.tryLock(30000); // 30s lock
  } catch (lockErr) {
    console.warn("Could not acquire lock: " + lockErr.toString());
  }

  try {
    if (!e || !e.postData || !e.postData.contents) {
      return ContentService.createTextOutput(JSON.stringify({
        status: "ERROR",
        message: "No POST body received."
      })).setMimeType(ContentService.MimeType.JSON);
    }

    var data = JSON.parse(e.postData.contents);

    var woNumber = String(
      data.woNumber || 
      data["Work Order Number"] || 
      data["WO Number"] || 
      data["woNum"] || ""
    ).trim();

    var workOrderId = String(
      data.workOrderId || 
      data.woId || 
      data["Work Order ID"] || 
      data["WO ID"] || 
      data["parentWorkOrderId"] || ""
    ).trim();

    var templateId = String(
      data.templateId || 
      data["Template ID"] || 
      (typeof DEFAULT_TEMPLATE_ID !== "undefined" ? DEFAULT_TEMPLATE_ID : "")
    ).trim();

    // 1. Fetch parent and child records
    var sheetRowData = getWorkOrderRowDataFromSheet(woNumber, workOrderId);

    workOrderId = String(sheetRowData["Work Order ID"] || workOrderId || "").trim();
    woNumber = String(sheetRowData["Work Order Number"] || woNumber || "").trim();

    var cleanWoNum = String(woNumber || "").replace(/[^0-9]/g, "");
    var formattedWoStr = "WO-" + (cleanWoNum ? ("000000" + cleanWoNum).slice(-6) : "000000");

    // 2. Debounce window (4 seconds)
    var props = PropertiesService.getScriptProperties();
    var lockPropKey = "LAST_RUN_" + (cleanWoNum || workOrderId.toLowerCase());
    var lastRunTime = parseInt(props.getProperty(lockPropKey) || "0", 10);
    var currentTime = new Date().getTime();

    if (currentTime - lastRunTime < 4000) {
      return ContentService.createTextOutput(JSON.stringify({
        status: "SUCCESS",
        message: "Rapid re-trigger debounced."
      })).setMimeType(ContentService.MimeType.JSON);
    }
    props.setProperty(lockPropKey, String(currentTime));

    // 3. Resolve Revision
    var payloadRev = data.revision || data["Revision"] || data["Work Order Revision"] || "";
    var sheetRev = sheetRowData["Revision"] || sheetRowData["Revision Number"] || "";
    var currentRevVal = parseInt(payloadRev !== "" ? payloadRev : sheetRev, 10);
    if (isNaN(currentRevVal) || currentRevVal < 1) currentRevVal = 1;
    var revision = String(currentRevVal).trim();

    // 4. Resolve Product and Quantity
    var productNumber = String(data.productNumber || data["Product Number"] || sheetRowData["Product Number"] || "").trim();
    var productType = String(data.productType || data["Product Type"] || sheetRowData["Product Type"] || sheetRowData["Work Order Type"] || "").trim();

    var serialFormat = getSerialFormatFromProductsSheet(productNumber);

    var plannedQty = data.plannedQty || data["Planned Quantity"] || data["Actual Quantity"] || 
                     sheetRowData["Planned Quantity"] || sheetRowData["WO QTY (Planned)"] || 
                     sheetRowData["Actual Quantity"] || sheetRowData["WO QTY (Actual)"] || 1;
    plannedQty = parseInt(plannedQty, 10);
    if (isNaN(plannedQty) || plannedQty < 1) plannedQty = 1;

    var rawConsumed = data.consumedSerials || data["Consumed Serial Units"] || sheetRowData["Consumed Serial Units"] || [];
    var selectedConsumedSerials = [];
    if (Array.isArray(rawConsumed)) {
      selectedConsumedSerials = rawConsumed;
    } else if (typeof rawConsumed === "string" && rawConsumed.trim() !== "") {
      selectedConsumedSerials = rawConsumed.split(/[\s,]+/).filter(Boolean);
    }

    // 5. Reconcile Finished Goods Serials
    if (workOrderId && productNumber) {
      try {
        var serialSheet = getSerialSheet();
        var timestampStr = getFormattedTimestamp();

        reconcileFGSerials(
          workOrderId, 
          woNumber, 
          productNumber, 
          serialFormat, 
          plannedQty, 
          serialSheet, 
          timestampStr, 
          selectedConsumedSerials, 
          productType
        );

        sheetRowData = getWorkOrderRowDataFromSheet(woNumber, workOrderId);
      } catch (reconErr) {
        console.error("Reconciliation error: " + reconErr.toString());
      }
    }

    // 6. Timestamps
    var timeZone = Session.getScriptTimeZone() || "GMT-7";
    sheetRowData["Generated Timestamp"] = Utilities.formatDate(new Date(), timeZone, "MM/dd/yyyy HH:mm:ss");
    sheetRowData["Printed Date"] = Utilities.formatDate(new Date(), timeZone, "MM/dd/yyyy");
    sheetRowData["Revision"] = revision;

    // 7. Generate PDF
    var pdfInfo = generateTravelerPDF(woNumber, revision, templateId, sheetRowData);

    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      pdfUrl: pdfInfo ? pdfInfo.fileUrl : "",
      fileId: pdfInfo ? pdfInfo.fileId : "",
      revision: revision
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    console.error("CRITICAL ERROR in doPost: " + err.toString());
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    if (hasLock) {
      try { lock.releaseLock(); } catch (e) {}
    }
  }
}

function getSerialFormatFromProductsSheet(productNumber) {
  if (!productNumber) return "";
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var prodSheet = ss.getSheetByName("Products");
    if (!prodSheet) return "";

    var data = prodSheet.getDataRange().getDisplayValues();
    if (data.length < 2) return "";

    var headers = data[0].map(function(h) { return String(h).trim().toLowerCase(); });
    var pnCol = headers.indexOf("product number");
    var formatCol = headers.indexOf("serial format");

    if (formatCol === -1) formatCol = 7;
    if (pnCol === -1) pnCol = 0;

    var cleanTargetPn = String(productNumber).trim().toLowerCase();

    for (var i = 1; i < data.length; i++) {
      var rowPn = String(data[i][pnCol]).trim().toLowerCase();
      if (rowPn === cleanTargetPn) {
        return String(data[i][formatCol] || "").replace(/[^A-Za-z0-9]/g, "");
      }
    }
  } catch (err) {
    console.error("Error fetching serial format: " + err.toString());
  }
  return "";
}

function getWorkOrderRowDataFromSheet(woNumber, workOrderId) {
  if (!woNumber && !workOrderId) return {};

  try {
    var activeSs = SpreadsheetApp.getActiveSpreadsheet();
    var ss = SpreadsheetApp.openById(activeSs.getId());
    var sheet = ss.getSheetByName("Work Orders") || ss.getSheets()[0];
    if (!sheet) return {};

    var displayValues = sheet.getDataRange().getDisplayValues();
    if (displayValues.length < 2) return {};

    var headers = displayValues[0].map(function(h) { return String(h).trim().toLowerCase(); });
    var idColIdx = headers.indexOf("work order id");
    var woColIdx = headers.indexOf("work order number");

    var targetId = String(workOrderId || "").trim().toLowerCase();
    var cleanTargetWo = String(woNumber || "").replace(/[^0-9]/g, "");

    var rowObj = {};
    var isMatch = false;

    for (var r = 1; r < displayValues.length; r++) {
      var cellId = idColIdx !== -1 ? String(displayValues[r][idColIdx]).trim().toLowerCase() : "";
      var cellCleanWo = woColIdx !== -1 ? String(displayValues[r][woColIdx]).replace(/[^0-9]/g, "") : "";

      if ((targetId && cellId === targetId) || (!targetId && cellCleanWo === cleanTargetWo && cleanTargetWo !== "")) {
        isMatch = true;
        for (var col = 0; col < displayValues[0].length; col++) {
          rowObj[String(displayValues[0][col]).trim()] = displayValues[r][col];
        }
        break;
      }
    }

    if (isMatch) {
      var woId = rowObj["Work Order ID"] || rowObj["ID"] || workOrderId || woNumber;
      var effectiveWoNum = rowObj["Work Order Number"] || woNumber;

      rowObj["Related Serial Units"] = getRelatedSerialsFromSheet(ss, effectiveWoNum, woId, rowObj["Product Number"]);

      if (typeof PAGE2_TABLE_CONFIGS !== "undefined") {
        PAGE2_TABLE_CONFIGS.forEach(function(cfg) {
          var childRows = [];
          for (var s = 0; s < cfg.sheetNames.length; s++) {
            childRows = getGenericChildRowsFromSheet(ss, cfg.sheetNames[s], effectiveWoNum, woId);
            if (childRows.length > 0) break;
          }

          var baseKey = cfg.targetKey || cfg.sheetNames[0];
          rowObj[baseKey] = childRows;
          rowObj["Related " + baseKey] = childRows;
        });
      }
    }

    return rowObj;
  } catch (err) {
    console.error("Error in getWorkOrderRowDataFromSheet: " + err.toString());
  }

  return {};
}

function getRelatedSerialsFromSheet(ss, woNumber, woId, productNumber) {
  var serialList = [];
  var serialSheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!serialSheet) return serialList;

  var values = serialSheet.getDataRange().getDisplayValues();
  if (values.length < 2) return serialList;

  var searchNum = String(woNumber || "").replace(/[^0-9]/g, "");
  var cleanWoId = String(woId || "").trim().toLowerCase();

  var serialCol = CONFIG.COLS.SERIAL_NUMBER - 1;
  var buildWoCol = CONFIG.COLS.BUILD_WO_ID - 1;
  var parentCol = CONFIG.COLS.PARENT_SERIAL - 1;
  var statusCol = CONFIG.COLS.STATUS - 1;

  var childMap = {};
  for (var c = 1; c < values.length; c++) {
    var cStatus = String(values[c][statusCol] || "").trim();
    if (cStatus === "Cancelled") continue;

    var parentRef = String(values[c][parentCol] || "").trim();
    var thisSerial = String(values[c][serialCol] || "").trim();

    if (parentRef) {
      if (!childMap[parentRef]) childMap[parentRef] = [];
      if (childMap[parentRef].indexOf(thisSerial) === -1) {
        childMap[parentRef].push(thisSerial);
      }
    }
  }

  for (var r = 1; r < values.length; r++) {
    var rowStatus = String(values[r][statusCol] || "").trim();
    if (rowStatus === "Cancelled") continue;

    var rowBuildWo = String(values[r][buildWoCol] || "").trim().toLowerCase();
    var cleanRowBuild = rowBuildWo.replace(/[^0-9]/g, "");

    var isBuildMatch = false;
    if (cleanWoId && rowBuildWo === cleanWoId) {
      isBuildMatch = true;
    } else if (searchNum && (cleanRowBuild === searchNum || rowBuildWo.indexOf(searchNum) !== -1)) {
      isBuildMatch = true;
    }

    if (isBuildMatch) {
      var targetSerial = String(values[r][serialCol] || "").trim();
      var consumedStacked = "";

      if (childMap[targetSerial] && childMap[targetSerial].length > 0) {
        var uniqueChildren = [];
        childMap[targetSerial].forEach(function(s) {
          if (uniqueChildren.indexOf(s) === -1) {
            uniqueChildren.push(s);
          }
        });
        consumedStacked = uniqueChildren.join("\n");
      }

      serialList.push({
        "serialNumber": targetSerial,
        "parentUnitSerial": consumedStacked || "—"
      });
    }
  }

  serialList.sort(function(a, b) {
    return a.serialNumber.localeCompare(b.serialNumber);
  });

  return serialList;
}

function getGenericChildRowsFromSheet(ss, sheetName, woNumber, woId) {
  var rowsList = [];
  if (!ss) return rowsList;
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return rowsList;

  var values = sheet.getDataRange().getDisplayValues();
  if (values.length < 2) return rowsList;

  var headers = values[0];
  var cleanWoId = String(woId || "").trim().toLowerCase();
  var cleanWoNum = String(woNumber || "").replace(/[^0-9]/g, "");

  var woColIdx = 1; // Column B default
  for (var h = 0; h < headers.length; h++) {
    var hName = String(headers[h] || "").trim().toLowerCase();
    if (hName === "work order id" || hName === "wo id" || hName === "work order") {
      woColIdx = h;
      break;
    }
  }

  for (var r = 1; r < values.length; r++) {
    var cellWoRef = String(values[r][woColIdx] || "").trim().toLowerCase();
    var cellWoCleanNum = cellWoRef.replace(/[^0-9]/g, "");

    var isMatch = false;
    if (cleanWoId && cellWoRef === cleanWoId) {
      isMatch = true;
    } else if (cleanWoNum && cellWoCleanNum === cleanWoNum && cleanWoNum.length >= 4) {
      isMatch = true;
    }

    if (isMatch) {
      var rowObj = {};
      for (var col = 0; col < headers.length; col++) {
        var hName = String(headers[col]).trim();
        rowObj[hName] = values[r][col];
      }
      rowsList.push(rowObj);
    }
  }

  return rowsList;
}