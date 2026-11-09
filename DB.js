/**
 * DB.gs - Database & Serial Table Operations
 */

function getSerialSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
  if (!sheet) {
    throw new Error("Sheet '" + CONFIG.SHEET_NAME + "' not found.");
  }
  return sheet;
}

function getFormattedTimestamp() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
}

function insertFGSerialRow(sheet, newSerial, productNumber, workOrderId, timestampStr) {
  var unitId = "SU-" + Utilities.getUuid().substring(0, 8);
  var newRow = [
    unitId,          // Col A: Serial Unit ID
    newSerial,       // Col B: Serial Number
    productNumber,   // Col C: Product Number
    workOrderId,     // Col D: Build WO ID
    "",              // Col E: Consumed WO ID
    "",              // Col F: Parent Unit Serial
    "Available",     // Col G: Status
    timestampStr,    // Col H: Created Date
    ""               // Col I: Consumed Date
  ];
  sheet.appendRow(newRow);
}

function reconcileConsumedUnits(sheet, currentSelectedList, workOrderId, primaryFGSerial, timestampStr) {
  var cleanWOId = String(workOrderId || "").trim().toLowerCase();
  var lastRow = sheet.getLastRow();

  if (lastRow <= 1) return;

  var rangeData = sheet.getRange(2, 1, lastRow - 1, 9).getValues();

  for (var i = 0; i < rangeData.length; i++) {
    var serial = String(rangeData[i][CONFIG.COLS.SERIAL_NUMBER - 1] || "").trim();
    var consumedWOId = String(rangeData[i][CONFIG.COLS.CONSUMED_WO_ID - 1] || "").trim().toLowerCase();
    var rowIndex = i + 2;

    var isCurrentlySelected = currentSelectedList.indexOf(serial) !== -1;
    var isAssignedToThisWO = (consumedWOId === cleanWOId && cleanWOId !== "");

    if (isCurrentlySelected) {
      sheet.getRange(rowIndex, CONFIG.COLS.CONSUMED_WO_ID).setValue(workOrderId);
      sheet.getRange(rowIndex, CONFIG.COLS.PARENT_SERIAL).setValue(primaryFGSerial);
      sheet.getRange(rowIndex, CONFIG.COLS.STATUS).setValue("Consumed");
      sheet.getRange(rowIndex, CONFIG.COLS.CONSUMED_DATE).setValue(timestampStr);
    } 
    else if (!isCurrentlySelected && isAssignedToThisWO) {
      sheet.getRange(rowIndex, CONFIG.COLS.CONSUMED_WO_ID).setValue("");
      sheet.getRange(rowIndex, CONFIG.COLS.PARENT_SERIAL).setValue("");
      sheet.getRange(rowIndex, CONFIG.COLS.STATUS).setValue("Available");
      sheet.getRange(rowIndex, CONFIG.COLS.CONSUMED_DATE).setValue("");
    }
  }
  SpreadsheetApp.flush();
}