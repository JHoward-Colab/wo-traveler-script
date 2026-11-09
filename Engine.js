/**
 * Engine.gs - Serial Number Generation, Reconciliation & Rollback Logic
 */

function reconcileFGSerials(
  workOrderId,
  woNumber,
  productNumber,
  serialFormat,
  plannedQty,
  serialSheet,
  timestampStr,
  selectedConsumedSerials,
  productType
) {
  if (!serialSheet) {
    throw new Error("reconcileFGSerials: serialSheet is null or undefined.");
  }

  var values = serialSheet.getDataRange().getDisplayValues();
  var cleanWoId = String(workOrderId || "").trim().toLowerCase();
  var cleanWoNum = String(woNumber || "").replace(/[^0-9]/g, "");

  var colSerial = 1;
  var colBuildWo = 3;
  var colConsumedWo = 4;
  var colParentSerial = 5;
  var colStatus = 6;
  var colConsumedDate = 8;

  var cleanType = String(productType || "").toLowerCase().trim();
  var isShock = cleanType.indexOf("shock") !== -1 || String(productNumber || "").indexOf("246-99-") === 0;

  // 1. Collect existing active FG units built by this WO
  var existingFG = [];
  if (values.length >= 2) {
    for (var r = 1; r < values.length; r++) {
      var rowStatus = String(values[r][colStatus] || "").trim();
      if (rowStatus === "Cancelled") continue;

      var rowBuildWo = String(values[r][colBuildWo] || "").trim().toLowerCase();
      var cleanRowBuild = rowBuildWo.replace(/[^0-9]/g, "");

      var isMatch = false;
      if (cleanWoId && rowBuildWo === cleanWoId) {
        isMatch = true;
      } else if (cleanWoNum && cleanRowBuild === cleanWoNum) {
        isMatch = true;
      }

      if (isMatch) {
        existingFG.push({
          rowIndex: r + 1,
          serialNumber: String(values[r][colSerial] || "").trim(),
          status: rowStatus
        });
      }
    }
  }

  // =========================================================================
  // CASE A: SHOCKS (1:1 Damper Inheritance)
  // =========================================================================
  if (isShock && selectedConsumedSerials && selectedConsumedSerials.length > 0) {
    var targetDamperSerials = selectedConsumedSerials.slice(0, plannedQty);

    for (var i = 0; i < existingFG.length; i++) {
      var fgUnit = existingFG[i];
      if (targetDamperSerials.indexOf(fgUnit.serialNumber) === -1) {
        serialSheet.getRange(fgUnit.rowIndex, colStatus + 1).setValue("Cancelled");
        serialSheet.getRange(fgUnit.rowIndex, colBuildWo + 1).setValue("");
        serialSheet.getRange(fgUnit.rowIndex, colConsumedWo + 1).setValue("");
        serialSheet.getRange(fgUnit.rowIndex, colParentSerial + 1).setValue("");
      }
    }

    existingFG = existingFG.filter(function(u) {
      return targetDamperSerials.indexOf(u.serialNumber) !== -1;
    });

    var existingSerialsMap = existingFG.map(function(u) { return u.serialNumber; });
    for (var d = 0; d < targetDamperSerials.length; d++) {
      var neededSerial = targetDamperSerials[d];
      if (existingSerialsMap.indexOf(neededSerial) === -1) {
        var newRowIndex = serialSheet.getLastRow() + 1;
        var newUnitId = "SU-" + Utilities.getUuid().slice(0, 8);

        serialSheet.appendRow([
          newUnitId,          // Col A: Serial Unit ID
          neededSerial,       // Col B: Serial Number
          productNumber,      // Col C: Product Number
          workOrderId,        // Col D: Build WO ID
          "",                 // Col E: Consumed WO ID
          "",                 // Col F: Parent Unit Serial
          "Available",        // Col G: Status
          timestampStr,       // Col H: Created Date
          ""                  // Col I: Consumed Date
        ]);

        existingFG.push({
          rowIndex: newRowIndex,
          serialNumber: neededSerial,
          status: "Available"
        });
      }
    }
    SpreadsheetApp.flush();

  // =========================================================================
  // CASE B: STANDARD BUILDS, SUB-ASSEMBLIES & DAMPERS
  // =========================================================================
  } else {
    existingFG.sort(function(a, b) {
      return a.serialNumber.localeCompare(b.serialNumber);
    });

    // Quantity Reduction: Cancel excess from the tail
    if (existingFG.length > plannedQty) {
      for (var x = plannedQty; x < existingFG.length; x++) {
        var excessRow = existingFG[x].rowIndex;
        var excessSerial = existingFG[x].serialNumber;

        serialSheet.getRange(excessRow, colStatus + 1).setValue("Cancelled");
        serialSheet.getRange(excessRow, colBuildWo + 1).setValue("");
        serialSheet.getRange(excessRow, colConsumedWo + 1).setValue("");
        serialSheet.getRange(excessRow, colParentSerial + 1).setValue("");

        if (values.length >= 2) {
          for (var c = 1; c < values.length; c++) {
            var childParentRef = String(values[c][colParentSerial] || "").trim();
            var childStatus = String(values[c][colStatus] || "").trim();

            if (childParentRef === excessSerial && childStatus === "Consumed") {
              var childRow = c + 1;
              serialSheet.getRange(childRow, colStatus + 1).setValue("Available");
              serialSheet.getRange(childRow, colConsumedWo + 1).setValue("");
              serialSheet.getRange(childRow, colParentSerial + 1).setValue("");
              serialSheet.getRange(childRow, colConsumedDate + 1).setValue("");
            }
          }
        }
      }
      SpreadsheetApp.flush();
      existingFG = existingFG.slice(0, plannedQty);
    }

    // Quantity Increase: Generate sequential serials
    var neededCount = plannedQty - existingFG.length;
    if (neededCount > 0) {
      for (var n = 0; n < neededCount; n++) {
        var nextIndex = existingFG.length + 1;
        var newSerial = generateNextSerial(productNumber, woNumber, nextIndex, serialFormat);
        var newRowIndex = serialSheet.getLastRow() + 1;
        var newUnitId = "SU-" + Utilities.getUuid().slice(0, 8);

        serialSheet.appendRow([
          newUnitId,          // Col A: Serial Unit ID
          newSerial,          // Col B: Serial Number
          productNumber,      // Col C: Product Number
          workOrderId,        // Col D: Build WO ID
          "",                 // Col E: Consumed WO ID
          "",                 // Col F: Parent Unit Serial
          "Available",        // Col G: Status
          timestampStr,       // Col H: Created Date
          ""                  // Col I: Consumed Date
        ]);

        existingFG.push({
          rowIndex: newRowIndex,
          serialNumber: newSerial,
          status: "Available"
        });
      }
      SpreadsheetApp.flush();
    }
  }

  // =========================================================================
  // RECONCILE CONSUMED SERIAL ASSIGNMENTS & DESELECTION ROLLBACKS
  // =========================================================================
  values = serialSheet.getDataRange().getDisplayValues();

  // 1. Deselected Rollback: Release any unit previously consumed that is no longer selected
  if (values.length >= 2) {
    for (var r2 = 1; r2 < values.length; r2++) {
      var cWo = String(values[r2][colConsumedWo] || "").trim().toLowerCase();
      var cStatus = String(values[r2][colStatus] || "").trim();
      var cSerial = String(values[r2][colSerial] || "").trim();

      var isConsumedByThisWo = (cleanWoId && cWo === cleanWoId) || 
                               (cleanWoNum && cWo.replace(/[^0-9]/g, "") === cleanWoNum);

      if (isConsumedByThisWo && cStatus === "Consumed") {
        if (!selectedConsumedSerials || selectedConsumedSerials.indexOf(cSerial) === -1) {
          var releaseRow = r2 + 1;
          serialSheet.getRange(releaseRow, colStatus + 1).setValue("Available");
          serialSheet.getRange(releaseRow, colConsumedWo + 1).setValue("");
          serialSheet.getRange(releaseRow, colParentSerial + 1).setValue("");
          serialSheet.getRange(releaseRow, colConsumedDate + 1).setValue("");
        }
      }
    }
  }

  // 2. Assign Consumed Serials to Target Parents
  if (selectedConsumedSerials && selectedConsumedSerials.length > 0 && existingFG.length > 0 && values.length >= 2) {
    for (var s = 0; s < selectedConsumedSerials.length; s++) {
      var currentConsumedSerial = selectedConsumedSerials[s];
      var targetFgParent = isShock ? currentConsumedSerial : existingFG[s % existingFG.length].serialNumber;

      for (var r3 = 1; r3 < values.length; r3++) {
        var rowSer = String(values[r3][colSerial] || "").trim();
        var rowBuildWo = String(values[r3][colBuildWo] || "").trim().toLowerCase();

        if (rowSer === currentConsumedSerial && rowBuildWo !== cleanWoId && rowBuildWo.replace(/[^0-9]/g, "") !== cleanWoNum) {
          var assignRow = r3 + 1;
          serialSheet.getRange(assignRow, colStatus + 1).setValue("Consumed");
          serialSheet.getRange(assignRow, colConsumedWo + 1).setValue(workOrderId);
          serialSheet.getRange(assignRow, colParentSerial + 1).setValue(targetFgParent);
          serialSheet.getRange(assignRow, colConsumedDate + 1).setValue(timestampStr);
          break;
        }
      }
    }
  }

  SpreadsheetApp.flush();
}

function generateNextSerial(productNumber, woNumber, index, serialFormat) {
  var cleanWo = String(woNumber || "").replace(/[^0-9]/g, "");
  var woPadded = ("000000" + cleanWo).slice(-6);
  var idxPadded = ("000" + index).slice(-3);

  var prefix = String(serialFormat || "").replace(/[^A-Za-z0-9]/g, "").trim();

  if (!prefix || prefix.length < 4) {
    var cleanPn = String(productNumber || "").replace(/[^A-Za-z0-9]/g, "");
    prefix = cleanPn.slice(0, 8) || "43081002";
  }

  return prefix + woPadded + idxPadded;
}