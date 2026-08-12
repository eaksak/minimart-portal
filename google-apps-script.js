/**
 * Google Apps Script for Silom POS - Sales Consolidation & Read-only API
 * 
 * Instructions:
 * 1. Create a new Google Sheet.
 * 2. Click Extensions > Apps Script.
 * 3. Delete any default code and paste this script.
 * 4. In the Apps Script editor, click on "Services" (left sidebar, + button), select "Drive API", and add it.
 * 5. Reload the spreadsheet, open the "Silom POS" menu, and run:
 *    - "ตั้งค่าระบบและดูโฟลเดอร์อัปโหลด"
 *    - "ติดตั้งซิงค์อัตโนมัติทุก 1 ชั่วโมง"
 * 6. Upload sales files only to the protected Drive upload folder.
 * 7. Deploy as Web App:
 *    - Click Deploy > New deployment.
 *    - Select type: Web App.
 *    - Set "Execute as": Me.
 *    - Set "Who has access": Anyone.
 *    - Click Deploy, authorize permissions, and copy the Web App URL.
 * 8. Paste the Web App URL into the dashboard settings under the Import tab.
 */

// --- CONFIGURATION ---
var CONFIG_SHEET_NAME = "Config";
var MASTER_SHEET_NAME = "MasterSales";
var DEFAULT_UPLOAD_FOLDER = "Silom Sales Upload";
var DEFAULT_ARCHIVE_FOLDER = "Silom Sales Archive";
var SPREADSHEET_ID_PROPERTY = "SILOM_SPREADSHEET_ID";

/**
 * Public read-only endpoint. It never imports files or writes to Drive/Sheets.
 */
function doGet(e) {
  var response = {};
  
  try {
    var action = (e && e.parameter && e.parameter.action) || "salesData";

    if (action === "dailySummary") {
      var requestedDate = normalizeDateKey_(e.parameter.date);
      if (!requestedDate) {
        var dateError = new Error("Missing or invalid 'date' parameter. Use YYYY-MM-DD.");
        dateError.code = "INVALID_DATE";
        throw dateError;
      }
      response = getDailySummaryResponse_(requestedDate);
    } else {
      var dataObj = getMasterSalesData();
      response = {
        status: "success",
        mode: "read-only",
        data: dataObj.data,
        dateColumns: dataObj.dateColumns
      };
    }
  } catch (err) {
    response = {
      status: "error",
      code: err.code || "REQUEST_FAILED",
      message: err.toString()
    };
  }
  
  return ContentService.createTextOutput(JSON.stringify(response))
    .setMimeType(ContentService.MimeType.JSON);
}

function getDailySummaryResponse_(requestedDate) {
  var dataObj = getMasterSalesData();
  var dateColumn = "";

  for (var i = 0; i < dataObj.dateColumns.length; i++) {
    if (normalizeDateKey_(dataObj.dateColumns[i]) === requestedDate) {
      dateColumn = dataObj.dateColumns[i];
      break;
    }
  }

  if (!dateColumn) {
    var notFoundError = new Error("No Silom sales data found for " + requestedDate);
    notFoundError.code = "DATE_NOT_FOUND";
    throw notFoundError;
  }

  var revenueData = {
    REV_COOP_SALES: 0,
    REV_CANTEEN_RICE: 0,
    REV_BAKERY_CONSIGN: 0,
    REV_CONSIGNMENT: 0,
    REV_UNIFORM: 0
  };

  dataObj.data.forEach(function(item) {
    var qty = Number(item[dateColumn]) || 0;
    if (qty === 0) return;

    var amount = qty * (Number(item["ราคา/หน่วย"]) || 0);
    var categoryNumber = getSilomCategoryNumber_(item["หมวดหมู่"]);

    if (categoryNumber === 4) {
      revenueData.REV_CANTEEN_RICE += amount;
    } else if (categoryNumber === 8) {
      revenueData.REV_BAKERY_CONSIGN += amount;
    } else if (categoryNumber === 9) {
      revenueData.REV_CONSIGNMENT += amount;
    } else if (categoryNumber === 10) {
      revenueData.REV_UNIFORM += amount;
    } else {
      revenueData.REV_COOP_SALES += amount;
    }
  });

  Object.keys(revenueData).forEach(function(key) {
    revenueData[key] = Math.round(revenueData[key] * 100) / 100;
  });

  var total = Object.keys(revenueData).reduce(function(sum, key) {
    return sum + revenueData[key];
  }, 0);

  return {
    status: "success",
    mode: "read-only",
    date: requestedDate,
    sourceDateColumn: dateColumn,
    data: revenueData,
    total: Math.round(total * 100) / 100,
    generatedAt: new Date().toISOString()
  };
}

function normalizeDateKey_(value) {
  if (!value) return "";

  if (value instanceof Date) {
    if (isNaN(value.getTime())) return "";
    return Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }

  var text = String(value).trim();
  var isoMatch = text.match(/^(\d{4})[/-](\d{1,2})[/-](\d{1,2})$/);
  var day;
  var month;
  var year;

  if (isoMatch) {
    year = Number(isoMatch[1]);
    month = Number(isoMatch[2]);
    day = Number(isoMatch[3]);
  } else {
    var dmyMatch = text.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})$/);
    if (!dmyMatch) return "";
    day = Number(dmyMatch[1]);
    month = Number(dmyMatch[2]);
    year = Number(dmyMatch[3]);
    if (year < 100) year += 2000;
    if (year > 2400) year -= 543;
  }

  var parsed = new Date(year, month - 1, day);
  if (
    parsed.getFullYear() !== year ||
    parsed.getMonth() !== month - 1 ||
    parsed.getDate() !== day
  ) return "";

  return String(year) + "-" + pad2_(month) + "-" + pad2_(day);
}

function pad2_(value) {
  return value < 10 ? "0" + value : String(value);
}

function getSilomCategoryNumber_(category) {
  var value = String(category || "").trim();
  if (/^A[_\s-]*เครื่องแบบเครื่องหมาย/i.test(value)) return 10;
  if (/^Uncategory$/i.test(value) || value === "ไม่มีหมวดหมู่") return 11;
  var numericPrefix = value.match(/^(\d+)/);
  return numericPrefix ? Number(numericPrefix[1]) : 11;
}


/**
 * Reject public write requests. Uploads must go through the protected Drive folder.
 */
function doPost() {
  return ContentService.createTextOutput(JSON.stringify({
    status: "error",
    code: "METHOD_NOT_ALLOWED",
    message: "Direct uploads are disabled. Upload the Excel file to the authorized Google Drive folder."
  }))
    .setMimeType(ContentService.MimeType.JSON);
}

/** Add private operational controls to the bound spreadsheet. */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Silom POS")
    .addItem("ตั้งค่าระบบและดูโฟลเดอร์อัปโหลด", "initializeSilomSystem")
    .addItem("ซิงค์ไฟล์ยอดขายตอนนี้", "syncUploadedSalesFilesFromMenu")
    .addSeparator()
    .addItem("ติดตั้งซิงค์อัตโนมัติทุก 1 ชั่วโมง", "installHourlySyncTrigger")
    .addItem("ยกเลิกซิงค์อัตโนมัติ", "removeScheduledSyncTriggersFromMenu")
    .addToUi();
}

function initializeSilomSystem() {
  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!activeSpreadsheet) {
    throw new Error("Open the bound Google Sheet before initializing the system.");
  }
  PropertiesService.getScriptProperties()
    .setProperty(SPREADSHEET_ID_PROPERTY, activeSpreadsheet.getId());

  var config = getOrSetupEnvironment();
  var uploadFolder = DriveApp.getFolderById(config.uploadFolderId);
  SpreadsheetApp.getUi().alert(
    "ตั้งค่าระบบเรียบร้อย\n\nอัปโหลดไฟล์ยอดขายที่โฟลเดอร์นี้เท่านั้น:\n" + uploadFolder.getUrl()
  );
}

/** Private sync function used by the time trigger. */
function syncUploadedSalesFiles() {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    throw new Error("Another sales sync is already running.");
  }

  try {
    var config = getOrSetupEnvironment();
    return importNewExcelFiles(config);
  } finally {
    lock.releaseLock();
  }
}

function syncUploadedSalesFilesFromMenu() {
  var result = syncUploadedSalesFiles();
  SpreadsheetApp.getUi().alert(
    "ซิงค์เสร็จเรียบร้อย\nประมวลผล " + result.filesProcessed.length + " ไฟล์"
  );
}

function installHourlySyncTrigger() {
  initializeSpreadsheetReference_();
  removeScheduledSyncTriggers_();
  ScriptApp.newTrigger("syncUploadedSalesFiles")
    .timeBased()
    .everyHours(1)
    .create();
  SpreadsheetApp.getUi().alert("ติดตั้งการซิงค์อัตโนมัติทุก 1 ชั่วโมงเรียบร้อย");
}

function removeScheduledSyncTriggersFromMenu() {
  var removed = removeScheduledSyncTriggers_();
  SpreadsheetApp.getUi().alert("ยกเลิก Trigger จำนวน " + removed + " รายการเรียบร้อย");
}

function removeScheduledSyncTriggers_() {
  var removed = 0;
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === "syncUploadedSalesFiles") {
      ScriptApp.deleteTrigger(trigger);
      removed++;
    }
  });
  return removed;
}

function initializeSpreadsheetReference_() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) {
    throw new Error("Open the bound Google Sheet before installing the trigger.");
  }
  PropertiesService.getScriptProperties()
    .setProperty(SPREADSHEET_ID_PROPERTY, spreadsheet.getId());
}

function getSilomSpreadsheet_() {
  var spreadsheetId = PropertiesService.getScriptProperties()
    .getProperty(SPREADSHEET_ID_PROPERTY);
  if (spreadsheetId) {
    return SpreadsheetApp.openById(spreadsheetId);
  }

  var activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (activeSpreadsheet) return activeSpreadsheet;

  throw new Error("Silom spreadsheet is not initialized. Run initializeSilomSystem from the spreadsheet menu.");
}

/**
 * Setup sheets and Google Drive folders if they don't exist, and return folder IDs
 */
function getOrSetupEnvironment() {
  var ss = getSilomSpreadsheet_();
  
  // 1. Setup Config Sheet
  var configSheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  if (!configSheet) {
    configSheet = ss.insertSheet(CONFIG_SHEET_NAME);
    configSheet.appendRow(["Key", "Value", "Description"]);
    configSheet.getRange("A1:C1").setFontWeight("bold");
    configSheet.appendRow(["UPLOAD_FOLDER_ID", "", "Folder ID for raw Excel uploads"]);
    configSheet.appendRow(["ARCHIVE_FOLDER_ID", "", "Folder ID for processed files"]);
  }
  
  // 2. Setup Master Sheet
  var masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);
  if (!masterSheet) {
    masterSheet = ss.insertSheet(MASTER_SHEET_NAME);
    // Initialize headers
    var headers = ["รหัสสินค้า", "ชื่อสินค้า", "หมวดหมู่", "ราคา/หน่วย", "ต้นทุน", "กำไร/ขาดทุน", "ประเภทภาษี"];
    masterSheet.appendRow(headers);
    masterSheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
  }
  
  // 3. Read/Generate folder configurations
  var uploadFolderId = "";
  var archiveFolderId = "";
  
  var data = configSheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === "UPLOAD_FOLDER_ID") {
      uploadFolderId = data[i][1];
    } else if (data[i][0] === "ARCHIVE_FOLDER_ID") {
      archiveFolderId = data[i][1];
    }
  }
  
  // Create folders if IDs are blank or invalid
  var parentFolder = DriveApp.getFileById(ss.getId()).getParents();
  var rootFolder = parentFolder.hasNext() ? parentFolder.next() : DriveApp.getRootFolder();
  
  if (!uploadFolderId || !isValidFolder(uploadFolderId)) {
    var uFolder = rootFolder.createFolder(DEFAULT_UPLOAD_FOLDER);
    uploadFolderId = uFolder.getId();
    updateConfigValue("UPLOAD_FOLDER_ID", uploadFolderId);
  }
  
  if (!archiveFolderId || !isValidFolder(archiveFolderId)) {
    var aFolder = rootFolder.createFolder(DEFAULT_ARCHIVE_FOLDER);
    archiveFolderId = aFolder.getId();
    updateConfigValue("ARCHIVE_FOLDER_ID", archiveFolderId);
  }
  
  return {
    uploadFolderId: uploadFolderId,
    archiveFolderId: archiveFolderId
  };
}

function isValidFolder(id) {
  try {
    var folder = DriveApp.getFolderById(id);
    return folder && !folder.isTrashed();
  } catch (e) {
    return false;
  }
}

function updateConfigValue(key, value) {
  var ss = getSilomSpreadsheet_();
  var sheet = ss.getSheetByName(CONFIG_SHEET_NAME);
  var range = sheet.getDataRange();
  var values = range.getValues();
  for (var i = 1; i < values.length; i++) {
    if (values[i][0] === key) {
      sheet.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
}

/**
 * Scan UPLOAD_FOLDER for .xlsx files, convert to Google Sheets, parse sales data,
 * merge into MASTER_SHEET, and archive the original file.
 */
function importNewExcelFiles(config) {
  var uploadFolder = DriveApp.getFolderById(config.uploadFolderId);
  var archiveFolder = DriveApp.getFolderById(config.archiveFolderId);
  
  // Find xlsx and xls files in upload folder
  var files = uploadFolder.getFiles();
  var filesProcessed = [];
  var processedDates = [];
  
  while (files.hasNext()) {
    var file = files.next();
    var filename = file.getName();
    
    // Process only Excel files
    if (filename.indexOf(".xlsx") !== -1 || filename.indexOf(".xls") !== -1) {
      try {
        var tempSheetId = convertExcelToGoogleSheet(file.getId(), config.uploadFolderId);
        var dateStr = processSingleSheet(tempSheetId);
        
        // Delete the temporary sheet
        DriveApp.getFileById(tempSheetId).setTrashed(true);
        
        // Move raw file to archive
        file.moveTo(archiveFolder);
        
        filesProcessed.push(filename);
        if (dateStr) processedDates.push(dateStr);
      } catch (err) {
        Logger.log("Failed to process " + filename + ": " + err.toString());
        throw new Error("Failed to process " + filename + ": " + err.message);
      }
    }
  }
  
  // Return list of dates and process results
  return {
    filesProcessed: filesProcessed,
    dateColumns: processedDates
  };
}

/**
 * Convert Excel to Google Sheets format using Advanced Drive Service (supports v2 and v3)
 */
function convertExcelToGoogleSheet(fileId, folderId) {
  var file = DriveApp.getFileById(fileId);
  var blob = file.getBlob();
  
  if (typeof Drive !== 'undefined' && Drive.Files) {
    // Check Drive API v2
    if (typeof Drive.Files.insert === 'function') {
      var resourceV2 = {
        title: "Temp_" + file.getName(),
        mimeType: MimeType.GOOGLE_SHEETS,
        parents: [{id: folderId}]
      };
      var newFileV2 = Drive.Files.insert(resourceV2, blob);
      return newFileV2.id;
    }
    // Check Drive API v3
    else if (typeof Drive.Files.create === 'function') {
      var resourceV3 = {
        name: "Temp_" + file.getName(),
        mimeType: MimeType.GOOGLE_SHEETS,
        parents: [folderId]
      };
      var newFileV3 = Drive.Files.create(resourceV3, blob);
      return newFileV3.id;
    }
  }
  
  throw new Error("Drive API Advanced Service is not enabled. Please click Services (+) in Apps Script, select Drive API, and add it.");
}

/**
 * Process a temporary Google Spreadsheet containing daily sales report
 * and merge into the MasterSales spreadsheet.
 */
function processSingleSheet(tempSheetId) {
  var tempSs = SpreadsheetApp.openById(tempSheetId);
  var tempSheet = tempSs.getSheets()[0];
  var rawData = tempSheet.getDataRange().getValues();
  
  // Find header row and identify the date
  var headerRowIndex = -1;
  var dateStr = "";
  
  for (var i = 0; i < Math.min(50, rawData.length); i++) {
    var row = rawData[i];
    var formattedRow = row.map(formatHeaderValue);
    var rowStr = formattedRow.join("");
    
    if (rowStr.indexOf("รหัสสินค้า") !== -1 && rowStr.indexOf("ชื่อสินค้า") !== -1) {
      headerRowIndex = i;
      
      // Look for a date column (DD/MM/YYYY) in the headers
      for (var colIdx = 0; colIdx < formattedRow.length; colIdx++) {
        var cellVal = formattedRow[colIdx];
        var dateMatch = cellVal.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})$/);
        if (dateMatch) {
          dateStr = cellVal; // e.g. "01/08/2026"
          break;
        }
      }
      break;
    }
  }
  
  // Fallbacks for date identification
  if (!dateStr) {
    // Fallback 1: Extract from cell B5 (From date)
    if (rawData[4] && String(rawData[4][0]).trim() === "From" && rawData[4][1]) {
      var fromDate = new Date(rawData[4][1]);
      if (fromDate && !isNaN(fromDate.getTime())) {
        dateStr = Utilities.formatDate(fromDate, Session.getScriptTimeZone(), "dd/MM/yyyy");
      }
    }
  }
  
  if (headerRowIndex === -1 || !dateStr) {
    throw new Error("Could not identify header row or report date in file.");
  }
  
  var headers = rawData[headerRowIndex].map(formatHeaderValue);
  
  // Map index of columns of interest in raw file
  var skuColIdx = headers.indexOf("รหัสสินค้า");
  var nameColIdx = headers.indexOf("ชื่อสินค้า");
  var categoryColIdx = headers.indexOf("หมวดหมู่");
  var priceColIdx = headers.indexOf("ราคา/หน่วย");
  var costColIdx = headers.indexOf("ต้นทุน");
  var profitColIdx = headers.indexOf("กำไร/ขาดทุน");
  var taxColIdx = headers.indexOf("ประเภทภาษี");
  var qtyColIdx = headers.indexOf(dateStr); // Quantity sold on this specific day
  
  if (qtyColIdx === -1) {
    // Try column "จำนวน"
    qtyColIdx = headers.indexOf("จำนวน");
  }
  
  if (skuColIdx === -1 || nameColIdx === -1) {
    throw new Error("Missing required columns in report file.");
  }
  
  // Load Master sheet
  var ss = getSilomSpreadsheet_();
  var masterSheet = ss.getSheetByName(MASTER_SHEET_NAME);
  var masterData = masterSheet.getDataRange().getValues();
  var masterHeaders = masterData[0].map(function(h) { return String(h).trim(); });
  
  // Check if date column exists in Master sheet. If not, append it.
  var masterDateColIdx = masterHeaders.indexOf(dateStr);
  if (masterDateColIdx === -1) {
    masterHeaders.push(dateStr);
    masterSheet.getRange(1, masterHeaders.length).setValue(dateStr).setFontWeight("bold");
    masterDateColIdx = masterHeaders.length - 1;
  }
  
  // Create mapping of Master SKU -> Master Row Index (0-indexed)
  var masterSkuMap = {};
  for (var rIdx = 1; rIdx < masterData.length; rIdx++) {
    var sku = String(masterData[rIdx][0]).trim();
    if (sku) masterSkuMap[sku] = rIdx;
  }
  
  // Read and process product rows from import sheet
  var itemsToInsert = [];
  var updates = []; // Array of objects {row, col, value}
  
  for (var rIdx = headerRowIndex + 1; rIdx < rawData.length; rIdx++) {
    var row = rawData[rIdx];
    var sku = String(row[skuColIdx]).trim();
    var name = String(row[nameColIdx]).trim();
    
    if (!sku || !name) continue; // Skip invalid rows
    
    var category = categoryColIdx !== -1 ? String(row[categoryColIdx]).trim() : "Uncategory";
    var price = priceColIdx !== -1 ? Number(row[priceColIdx]) || 0 : 0;
    var cost = costColIdx !== -1 ? Number(row[costColIdx]) || 0 : 0;
    var profit = profitColIdx !== -1 ? Number(row[profitColIdx]) || 0 : 0;
    var tax = taxColIdx !== -1 ? String(row[taxColIdx]).trim() : "N";
    var qty = qtyColIdx !== -1 ? Number(row[qtyColIdx]) || 0 : 0;
    
    if (sku === "รวมทั้งสิ้น" || sku === "Total") continue; // Skip total row
    
    if (sku in masterSkuMap) {
      // Update existing item values
      var targetRow = masterSkuMap[sku] + 1; // 1-indexed row in sheet
      
      // Update product info (to latest values)
      masterSheet.getRange(targetRow, 2).setValue(name);
      masterSheet.getRange(targetRow, 3).setValue(category);
      masterSheet.getRange(targetRow, 4).setValue(price);
      masterSheet.getRange(targetRow, 5).setValue(cost);
      masterSheet.getRange(targetRow, 6).setValue(profit);
      masterSheet.getRange(targetRow, 7).setValue(tax);
      
      // Write/Overwrite date quantity
      masterSheet.getRange(targetRow, masterDateColIdx + 1).setValue(qty);
    } else {
      // Prepare new item row
      var newRow = [sku, name, category, price, cost, profit, tax];
      
      // Pad row to match target date column index
      while (newRow.length < masterHeaders.length) {
        newRow.push("");
      }
      newRow[masterDateColIdx] = qty;
      itemsToInsert.push(newRow);
    }
  }
  
  // Insert any new items to Master sheet
  if (itemsToInsert.length > 0) {
    masterSheet.getRange(masterSheet.getLastRow() + 1, 1, itemsToInsert.length, itemsToInsert[0].length).setValues(itemsToInsert);
  }
  
  return dateStr;
}

function formatHeaderValue(cellValue) {
  if (cellValue instanceof Date) {
    return Utilities.formatDate(cellValue, Session.getScriptTimeZone(), "dd/MM/yyyy");
  }
  var str = String(cellValue || '').trim();
  if (str.length > 10 && (str.indexOf("GMT") !== -1 || str.indexOf("T") !== -1)) {
    var d = new Date(str);
    if (!isNaN(d.getTime())) {
      return Utilities.formatDate(d, Session.getScriptTimeZone(), "dd/MM/yyyy");
    }
  }
  return str;
}

/**
 * Fetch and format data from MasterSales sheet to return as JSON object
 */
function getMasterSalesData() {
  var ss = getSilomSpreadsheet_();
  var sheet = ss.getSheetByName(MASTER_SHEET_NAME);
  var values = sheet.getDataRange().getValues();
  
  if (values.length <= 1) {
    return { data: [], dateColumns: [] };
  }
  
  var headers = values[0].map(formatHeaderValue);
  
  // Identify which columns are date columns (columns index 7 and onwards)
  var dateColumns = [];
  for (var colIdx = 7; colIdx < headers.length; colIdx++) {
    var match = headers[colIdx].match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})$/);
    if (match) {
      dateColumns.push(headers[colIdx]);
    }
  }
  
  var products = [];
  for (var rIdx = 1; rIdx < values.length; rIdx++) {
    var row = values[rIdx];
    var sku = String(row[0]).trim();
    if (!sku) continue;
    
    var item = {
      "รหัสสินค้า": sku,
      "ชื่อสินค้า": String(row[1]).trim(),
      "หมวดหมู่": String(row[2]).trim(),
      "ราคา/หน่วย": Number(row[3]) || 0,
      "ต้นทุน": Number(row[4]) || 0,
      "กำไร/ขาดทุน": Number(row[5]) || 0,
      "ประเภทภาษี": String(row[6]).trim(),
      "รวมทั้งสิ้น": 0,
      "จำนวน": 0
    };
    
    // Accumulate dates values
    var totalQty = 0;
    for (var colIdx = 7; colIdx < headers.length; colIdx++) {
      var colHeader = headers[colIdx];
      var qty = Number(row[colIdx]) || 0;
      item[colHeader] = qty;
      totalQty += qty;
    }
    
    item["จำนวน"] = totalQty;
    item["รวมทั้งสิ้น"] = totalQty * item["ราคา/หน่วย"];
    
    products.push(item);
  }
  
  return {
    data: products,
    dateColumns: dateColumns
  };
}
