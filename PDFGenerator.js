/**
 * PDFGenerator.gs - Traveler PDF Output & Cleanup
 */

function generateTravelerPDF(woNumber, revision, templateId, rowData) {
  console.log("Starting PDF generation for WO: " + woNumber);

  var effectiveTemplateId = String(templateId || DEFAULT_TEMPLATE_ID || "").trim();
  if (!effectiveTemplateId) {
    throw new Error("Template ID is missing or invalid.");
  }

  var templateFile = DriveApp.getFileById(effectiveTemplateId);
  var targetFolder = DriveApp.getFolderById(WORK_ORDER_PDF_FOLDER_ID);
  var archiveFolder = DriveApp.getFolderById(ARCHIVE_FOLDER_ID);

  var cleanWo = String(woNumber || "").replace(/[^0-9]/g, "");
  var formattedWoStr = "WO-" + (cleanWo ? ("000000" + cleanWo).slice(-6) : "000000");
  
  var currentRev = String(revision !== undefined && revision !== null && revision !== "" ? revision : "1").trim();
  var pdfName = formattedWoStr + " Ver " + currentRev + ".pdf";

  // Clean up existing duplicate in active folder or archive previous revisions
  try {
    var existingActiveFiles = targetFolder.getFiles();
    var searchPrefix = formattedWoStr + " Ver ";
    while (existingActiveFiles.hasNext()) {
      var file = existingActiveFiles.next();
      var fName = file.getName();
      if (fName.indexOf(searchPrefix) === 0 && fName.toLowerCase().indexOf(".pdf") !== -1) {
        if (fName === pdfName) {
          file.setTrashed(true);
        } else {
          file.moveTo(archiveFolder);
          console.log("Archived older revision: " + fName);
        }
      }
    }

    var existingArchiveFiles = archiveFolder.getFiles();
    while (existingArchiveFiles.hasNext()) {
      var aFile = existingArchiveFiles.next();
      if (aFile.getName() === pdfName) {
        aFile.setTrashed(true);
      }
    }
  } catch (archErr) {
    console.warn("Archive step note: " + archErr.toString());
  }

  var copyDoc = templateFile.makeCopy("TEMP_WO_" + cleanWo, targetFolder);
  var copyId = copyDoc.getId();

  try {
    var doc = DocumentApp.openById(copyId);
    var body = doc.getBody();
    var header = doc.getHeader();
    var footer = doc.getFooter();

    populateDocPlaceholders(body, header, footer, rowData, formattedWoStr, currentRev);
    insertQRCodeImage(body, formattedWoStr);
    populateChildTables(body, rowData);

    doc.saveAndClose();

    var pdfBlob = DriveApp.getFileById(copyId).getAs("application/pdf").setName(pdfName);
    var finalPdfFile = targetFolder.createFile(pdfBlob);
    
    try {
      finalPdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (sharingErr) {
      console.warn("Sharing note: " + sharingErr.toString());
    }

    console.log("Successfully created PDF: " + finalPdfFile.getUrl());

    return {
      fileName: pdfName,
      fileId: finalPdfFile.getId(),
      fileUrl: finalPdfFile.getUrl()
    };

  } catch (pdfErr) {
    console.error("Error inside generateTravelerPDF: " + pdfErr.toString());
    throw pdfErr;
  } finally {
    try {
      DriveApp.getFileById(copyId).setTrashed(true);
    } catch (e) {
      console.warn("Could not trash temp file: " + e.toString());
    }
  }
}

function insertQRCodeImage(body, woNumberStr) {
  if (!body || !woNumberStr) return;

  var qrTokens = ["<<[WO_QR_Code]>>", "*<<[Work Order Number]>>*", "*<<Work Order Number>>*"];
  var qrApiUrl = "https://quickchart.io/qr?size=150&text=" + encodeURIComponent(woNumberStr);
  
  var qrBlob;
  try {
    qrBlob = UrlFetchApp.fetch(qrApiUrl, { muteHttpExceptions: true }).getBlob();
  } catch (e) {
    console.warn("Could not fetch QR code image blob: " + e.toString());
    return;
  }

  for (var i = 0; i < qrTokens.length; i++) {
    var token = qrTokens[i];
    var rangeElement = body.findText(token.replace(/([.*+?^${}()|[\]\/\\])/g, "\\$1"));

    while (rangeElement) {
      var element = rangeElement.getElement();
      if (element.getType() === DocumentApp.ElementType.TEXT) {
        var textNode = element.asText();
        var parent = textNode.getParent();
        
        textNode.setText("");
        if (parent.getType() === DocumentApp.ElementType.PARAGRAPH) {
          var img = parent.asParagraph().appendInlineImage(qrBlob);
          img.setWidth(75);
          img.setHeight(75);
        }
      }
      rangeElement = body.findText(token.replace(/([.*+?^${}()|[\]\/\\])/g, "\\$1"), rangeElement);
    }
  }
}

function safeReplace(container, token, value) {
  if (!container || !token) return;
  var valStr = (value === undefined || value === null || String(value).trim() === "") ? " " : String(value).trim();
  var escapedToken = String(token).replace(/[\-\[\]\/\{\}\(\)\*\+\?\.\\\^\$\|]/g, "\\$&");
  var escapedValue = valStr.replace(/\\/g, "\\\\").replace(/\$/g, "\\$");

  try {
    container.replaceText(escapedToken, escapedValue);
  } catch (e) {
    console.warn("Replace text warning for token '" + token + "': " + e.toString());
  }
}

function populateDocPlaceholders(body, header, footer, rowData, woNumber, revision) {
  rowData = rowData || {};
  var cleanWoStr = String(woNumber || "").trim();
  var revStr = String(revision !== undefined && revision !== null && revision !== "" ? revision : "1").trim();

  var sections = [body];
  if (header) sections.push(header);
  if (footer) sections.push(footer);

  function replaceEverywhere(token, value) {
    for (var s = 0; s < sections.length; s++) {
      safeReplace(sections[s], token, value);
    }
  }

  replaceEverywhere("<<[Work Order Number]>>", cleanWoStr);
  replaceEverywhere("<<[Revision]>>", revStr);
  replaceEverywhere("<<[Status]>>", rowData["Status"] || rowData["WO Status"]);
  replaceEverywhere("<<[Closed Date]>>", rowData["Closed Date"] || rowData["WO Close Date"]);
  replaceEverywhere("<<[Work Order Type]>>", rowData["Work Order Type"] || rowData["WO Type"] || rowData["Product Type"]);
  replaceEverywhere("<<[Product Number]>>", rowData["Product Number"]);
  replaceEverywhere("<<[BOM Revision]>>", rowData["BOM Revision"]);
  replaceEverywhere("<<[Product Description]>>", rowData["Product Description"]);
  replaceEverywhere("<<[PO or SO Number]>>", rowData["PO or SO Number"] || rowData["PO #"]);
  replaceEverywhere("<<[Customer]>>", rowData["Customer"]);
  replaceEverywhere("<<[Planned Quantity]>>", rowData["Planned Quantity"] || rowData["WO QTY (Planned)"] || "1");
  replaceEverywhere("<<[Actual Quantity]>>", rowData["Actual Quantity"] || rowData["WO QTY (Actual)"]);
  replaceEverywhere("<<[Planned Start Date]>>", rowData["Planned Start Date"] || rowData["WO Start Date (Planned)"]);
  replaceEverywhere("<<[Actual Start Date]>>", rowData["Actual Start Date"] || rowData["WO Start Date (Actual)"]);
  replaceEverywhere("<<[Comments]>>", rowData["Comments"]);

  // Timestamps
  replaceEverywhere("<<Generated Timestamp>>", rowData["Generated Timestamp"]);
  replaceEverywhere("<<[Generated Timestamp]>>", rowData["Generated Timestamp"]);
  replaceEverywhere("<<Printed Date>>", rowData["Printed Date"]);
  replaceEverywhere("<<[Printed Date]>>", rowData["Printed Date"]);

  // Standard fallbacks
  replaceEverywhere("<<Work Order Number>>", cleanWoStr);
  replaceEverywhere("<<Revision>>", revStr);
  replaceEverywhere("<<Product Number>>", rowData["Product Number"]);
  replaceEverywhere("<<Planned Quantity>>", rowData["Planned Quantity"] || rowData["WO QTY (Planned)"] || "1");
}

function populateChildTables(body, rowData) {
  var tables = body.getTables();

  for (var t = 0; t < tables.length; t++) {
    var table = tables[t];
    var tableText = table.getText();

    // 1. Serial Units Table
    if (tableText.indexOf("Related Serial Units") !== -1 || tableText.indexOf("Target Serial Number") !== -1 || tableText.indexOf("Serial Number") !== -1) {
      renderRepeatRows(table, rowData["Related Serial Units"], function(row, item) {
        row.getCell(0).setText(item["serialNumber"] || item["Serial Number"] || " ");
        if (row.getNumCells() > 1) {
          row.getCell(1).setText(item["parentUnitSerial"] || item["Parent Unit Serial"] || item["consumedSerial"] || "—");
        }
      });
    }

    // 2. Deviations Table
    if (tableText.indexOf("Deviations") !== -1) {
      var devList = rowData["WO Deviations"] || rowData["Related WO Deviations"] || [];
      renderRepeatRows(table, devList, function(row, item) {
        row.getCell(0).setText(item["Deviation Date"] || item["Date"] || " ");
        if (row.getNumCells() > 1) row.getCell(1).setText(item["Current Specification"] || item["Current Spec"] || " ");
        if (row.getNumCells() > 2) row.getCell(2).setText(item["Requested Deviation"] || " ");
        if (row.getNumCells() > 3) row.getCell(3).setText(item["Reason for Deviation"] || " ");
      });
    }

    // 3. BOM Changes Table
    if (tableText.indexOf("BOM Changes") !== -1) {
      var bomList = rowData["WO BOM Changes"] || rowData["Related WO BOM Changes"] || [];
      renderRepeatRows(table, bomList, function(row, item) {
        row.getCell(0).setText(item["Primary Part Number"] || item["Primary Part #"] || " ");
        if (row.getNumCells() > 1) row.getCell(1).setText(item["Primary Part Description"] || " ");
        if (row.getNumCells() > 2) row.getCell(2).setText(item["Primary Quantity"] || item["Primary Qty"] || " ");
        if (row.getNumCells() > 3) row.getCell(3).setText(item["Sub Part Number"] || " ");
        if (row.getNumCells() > 4) row.getCell(4).setText(item["Sub Part Description"] || item["Sub Description"] || " ");
        if (row.getNumCells() > 5) row.getCell(5).setText(item["Sub Quantity"] || item["Sub Qty"] || " ");
      });
    }

    // 4. Critical Assembly Checks Table
    if (tableText.indexOf("Critical Assembly Checks") !== -1 || tableText.indexOf("Traveler Checks") !== -1 || tableText.indexOf("Checks") !== -1) {
      var checkList = rowData["WO Traveler Checks"] || rowData["Related WO Traveler Checks"] || [];
      renderRepeatRows(table, checkList, function(row, item) {
        row.getCell(0).setText(item["Step Number"] || item["Step #"] || " ");
        if (row.getNumCells() > 1) row.getCell(1).setText(item["Action / Check Description"] || " ");
        if (row.getNumCells() > 2) row.getCell(2).setText(item["Requirement / Spec"] || " ");
        if (row.getNumCells() > 3) row.getCell(3).setText(item["Status"] || " ");
        if (row.getNumCells() > 4) {
          var initial = item["Operator Initial"] || item["Operator Verification"] || "";
          row.getCell(4).setText(initial ? "[ " + initial + " ]" : "[   ]");
        }
      });
    }
  }

  safeReplace(body, "<<Start:", "");
  safeReplace(body, "<<End>>", "");
}

function renderRepeatRows(table, itemList, populateCallback) {
  itemList = itemList || [];
  var numRows = table.getNumRows();
  var templateRowIdx = -1;

  for (var r = 0; r < numRows; r++) {
    if (table.getRow(r).getText().indexOf("<<Start:") !== -1) {
      templateRowIdx = r;
      break;
    }
  }

  if (templateRowIdx === -1) return;

  var templateRow = table.getRow(templateRowIdx);

  if (itemList.length === 0) {
    for (var c = 0; c < templateRow.getNumCells(); c++) {
      templateRow.getCell(c).setText(" ");
    }
    return;
  }

  for (var i = 0; i < itemList.length; i++) {
    var newRow = table.appendTableRow(templateRow.copy());
    populateCallback(newRow, itemList[i]);

    for (var c = 0; c < newRow.getNumCells(); c++) {
      var cell = newRow.getCell(c);
      cell.setPaddingTop(1.5);
      cell.setPaddingBottom(1.5);
      cell.setPaddingLeft(3);
      cell.setPaddingRight(3);

      for (var p = 0; p < cell.getNumChildren(); p++) {
        var child = cell.getChild(p);
        if (child.getType() === DocumentApp.ElementType.PARAGRAPH) {
          var para = child.asParagraph();
          para.setSpacingBefore(0);
          para.setSpacingAfter(0);
          para.setLineSpacing(1.0);
          
          var textObj = para.editAsText();
          textObj.setFontSize(8.5);
        }
      }
    }
  }

  table.removeRow(templateRowIdx);
}