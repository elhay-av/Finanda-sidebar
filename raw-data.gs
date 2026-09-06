/**
 * Raw Data Management for Finanda Transactions
 * 
 * Manages the dedicated 'RawData' tab in the Google Spreadsheet,
 * storing all raw Finanda transaction fields and persisting
 * categorization and comment metadata across syncs.
 */

const RAW_DATA_SHEET_NAME = "RawData";

const RAW_DATA_HEADERS = [
  "TransID",
  "AccountType",
  "BillingCurrency",
  "BizExpense",
  "CreateDate",
  "CreateSource",
  "Credit",
  "Debit",
  "Description",
  "InstituteTransID",
  "LoadType",
  "OpposingAccount",
  "OpposingTransID",
  "ReferenceID",
  "ReferenceID2",
  "ReportedBalance",
  "TransCurrency",
  "TransDate",
  "TransType",
  "TransValueDate",
  "TotalPayments",
  "AccID",
  "Amount",
  "AccountID",
  "category",
  "CatDesc",
  "CatGroupID",
  "CatGroup",
  "CategoryType",
  "AccountNumber",
  "institute_code",
  "instituteDesc",
  "hashtags",
  "MonthKey",
  "Category",
  "Comment",
];

/**
 * Safely formats any value for writing into Google Sheets.
 */
function formatRawCellValue(val) {
  if (val === undefined || val === null) {
    return "";
  }
  if (typeof val === "object") {
    try {
      return JSON.stringify(val);
    } catch (e) {
      return String(val);
    }
  }
  return val;
}

/**
 * Ensures the 'RawData' tab exists in the active spreadsheet.
 * If not, creates it with header row, formats headers, and hides the tab.
 * 
 * @returns {GoogleAppsScript.Spreadsheet.Sheet} The RawData sheet
 */
function ensureRawDataTab() {
  const spreadsheet = getProtectedActiveSpreadsheet();
  let sheet = spreadsheet.getSheetByName(RAW_DATA_SHEET_NAME);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(RAW_DATA_SHEET_NAME);
    sheet.getRange(1, 1, 1, RAW_DATA_HEADERS.length).setValues([RAW_DATA_HEADERS]);
    sheet.getRange(1, 1, 1, RAW_DATA_HEADERS.length).setFontWeight("bold").setBackground("#f0f4f8");
    sheet.setFrozenRows(1);
    try {
      sheet.hideSheet();
    } catch (e) {
      Logger.log("Notice: could not hide RawData sheet: " + e.message);
    }
  } else {
    // If sheet exists but has no rows
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, RAW_DATA_HEADERS.length).setValues([RAW_DATA_HEADERS]);
      sheet.getRange(1, 1, 1, RAW_DATA_HEADERS.length).setFontWeight("bold").setBackground("#f0f4f8");
      sheet.setFrozenRows(1);
    }
    try {
      if (!sheet.isSheetHidden()) {
        sheet.hideSheet();
      }
    } catch (e) {
      Logger.log("Notice: could not verify RawData hidden status: " + e.message);
    }
  }
  return sheet;
}

/**
 * Appends only genuinely new transactions to the RawData tab.
 * Deduplicates based on TransID (or _id fallback).
 * 
 * @param {Array<Object>} transactions List of Finanda transaction objects
 * @param {string} monthKey The year-month string (e.g. "2026-09")
 * @returns {number} Count of new transactions appended
 */
function appendNewRawTransactions(transactions, monthKey) {
  if (!transactions || !transactions.length) {
    return 0;
  }
  const sheet = ensureRawDataTab();
  const lastRow = sheet.getLastRow();

  // Build a Set of existing TransIDs
  const existingIds = new Set();
  if (lastRow > 1) {
    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      const val = ids[i][0];
      if (val !== "" && val !== null && val !== undefined) {
        existingIds.add(String(val).trim());
      }
    }
  }

  const newRows = [];
  for (let i = 0; i < transactions.length; i++) {
    const item = transactions[i];
    const transId = String(
      item.TransID != null ? item.TransID : item._id != null ? item._id : "",
    ).trim();

    if (transId && existingIds.has(transId)) {
      continue; // Skip existing record
    }
    if (transId) {
      existingIds.add(transId);
    }

    const row = RAW_DATA_HEADERS.map((header) => {
      if (header === "TransID") {
        return transId;
      }
      if (header === "MonthKey") {
        return monthKey;
      }
      if (header === "Category") {
        return "";
      }
      if (header === "Comment") {
        return "";
      }
      return formatRawCellValue(item[header]);
    });
    newRows.push(row);
  }

  if (newRows.length > 0) {
    sheet
      .getRange(lastRow + 1, 1, newRows.length, RAW_DATA_HEADERS.length)
      .setValues(newRows);
    Logger.log(`[RawData] Appended ${newRows.length} new transactions for ${monthKey}`);
  } else {
    Logger.log(`[RawData] All ${transactions.length} transactions already exist in RawData`);
  }

  return newRows.length;
}

/**
 * Retrieves all transactions for a given monthKey from RawData.
 * 
 * @param {string} monthKey The year-month string (e.g. "2026-09")
 * @returns {Array<Object>} List of records containing row metadata and item fields
 */
function getRawTransactionsForMonth(monthKey) {
  const sheet = ensureRawDataTab();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return [];
  }

  const lastCol = sheet.getLastColumn();
  const allValues = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const headers = allValues[0];
  const colIndexMap = {};
  headers.forEach((h, idx) => {
    colIndexMap[h] = idx;
  });

  const monthColIdx = colIndexMap["MonthKey"];
  const categoryColIdx = colIndexMap["Category"];
  const commentColIdx = colIndexMap["Comment"];
  const transIdColIdx = colIndexMap["TransID"];

  const results = [];
  for (let r = 1; r < allValues.length; r++) {
    const rowValues = allValues[r];
    const rowMonth = rowValues[monthColIdx];
    if (String(rowMonth).trim() === monthKey) {
      const item = {};
      headers.forEach((h, idx) => {
        item[h] = rowValues[idx];
      });

      item.Amount = Number(item.Amount) || 0;
      item.Credit = Number(item.Credit) || 0;
      item.Debit = Number(item.Debit) || 0;
      item._id = item.TransID; // Compatibility

      results.push({
        rowIndex: r + 1, // 1-based row index in sheet
        transId: String(rowValues[transIdColIdx] || "").trim(),
        category: String(rowValues[categoryColIdx] || "").trim(),
        comment: String(rowValues[commentColIdx] || "").trim(),
        item: item,
      });
    }
  }

  return results;
}

/**
 * Updates Category and Comment for a single transaction in RawData by TransID.
 * 
 * @param {string} transId The transaction ID
 * @param {string} category The mapped budget group ID or 'SKIP'
 * @param {string} comment Reasoning / documentation note
 */
function updateRawDataCategory(transId, category, comment) {
  if (!transId) return;
  const sheet = ensureRawDataTab();
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const transIdColIdx = headers.indexOf("TransID") + 1;
  const categoryColIdx = headers.indexOf("Category") + 1;
  const commentColIdx = headers.indexOf("Comment") + 1;

  if (categoryColIdx === 0) return;

  const ids = sheet.getRange(2, transIdColIdx, lastRow - 1, 1).getValues();
  const searchId = String(transId).trim();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]).trim() === searchId) {
      const rowIndex = i + 2;
      sheet.getRange(rowIndex, categoryColIdx).setValue(category);
      if (commentColIdx > 0 && comment !== undefined) {
        sheet.getRange(rowIndex, commentColIdx).setValue(comment);
      }
      Logger.log(
        `[RawData] Updated transId ${transId} at row ${rowIndex}: category=${category}, comment=${comment}`,
      );
      return;
    }
  }
  Logger.log(`[RawData] TransID ${transId} not found in RawData tab`);
}

/**
 * Batch updates Category and Comment for multiple rows in RawData.
 * 
 * @param {Array<{rowIndex: number, category: string, comment?: string}>} updates
 */
function batchUpdateRawDataCategories(updates) {
  if (!updates || !updates.length) return;
  const sheet = ensureRawDataTab();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const categoryColIdx = headers.indexOf("Category") + 1;
  const commentColIdx = headers.indexOf("Comment") + 1;

  updates.forEach((up) => {
    if (up.rowIndex) {
      sheet.getRange(up.rowIndex, categoryColIdx).setValue(up.category);
      if (commentColIdx > 0 && up.comment !== undefined) {
        sheet.getRange(up.rowIndex, commentColIdx).setValue(up.comment);
      }
    }
  });
}
