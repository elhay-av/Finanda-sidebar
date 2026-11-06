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
 * Normalizes any date value or string into a "YYYY-MM" month key.
 * Handles Google Sheets auto-coerced Date objects, ISO strings, etc.
 * 
 * @param {*} val Date object, string, or any cell value
 * @returns {string} Normalized "YYYY-MM" string or empty string
 */
function parseMonthKey(val) {
  if (!val) return "";
  if (val instanceof Date) {
    const y = val.getFullYear();
    const m = (val.getMonth() + 1).toString().padStart(2, "0");
    return `${y}-${m}`;
  }
  const str = String(val).trim();
  if (/^\d{4}-\d{2}/.test(str)) {
    return str.substring(0, 7);
  }
  const d = new Date(str);
  if (!isNaN(d.getTime())) {
    const y = d.getFullYear();
    const m = (d.getMonth() + 1).toString().padStart(2, "0");
    return `${y}-${m}`;
  }
  return str;
}

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
 * Also configures plain text format for identifier and date key columns.
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

  // Ensure MonthKey and TransID columns are formatted as Plain Text
  try {
    const monthCol = RAW_DATA_HEADERS.indexOf("MonthKey") + 1;
    const transIdCol = RAW_DATA_HEADERS.indexOf("TransID") + 1;
    const numRows = Math.max(sheet.getMaxRows(), 100);
    if (monthCol > 0) {
      sheet.getRange(1, monthCol, numRows, 1).setNumberFormat("@");
    }
    if (transIdCol > 0) {
      sheet.getRange(1, transIdCol, numRows, 1).setNumberFormat("@");
    }
  } catch (fmtErr) {
    Logger.log("Notice: Could not set plain text format: " + fmtErr.message);
  }

  return sheet;
}

/**
 * Normalizes a date cell (Date object or "YYYY-MM-DD..." string) to "YYYY-MM-DD".
 */
function normalizeDayKey(val, tz) {
  if (!val) return "";
  if (val instanceof Date) {
    return Utilities.formatDate(val, tz, "yyyy-MM-dd");
  }
  return String(val).trim().substring(0, 10);
}

/**
 * Builds a fingerprint that identifies "the same real-world transaction"
 * even when Finanda re-issues it under a new TransID.
 */
function rawFingerprint(item, tz, withDescription) {
  const parts = [
    String(item.AccID != null ? item.AccID : "").trim(),
    normalizeDayKey(item.TransDate, tz),
    (Number(item.Debit) || 0).toFixed(2),
    (Number(item.Credit) || 0).toFixed(2),
  ];
  if (withDescription) {
    parts.push(String(item.Description != null ? item.Description : "").replace(/\s+/g, " ").trim());
  }
  return parts.join("|");
}

function getRawTransId(item) {
  return String(
    item.TransID != null ? item.TransID : item._id != null ? item._id : "",
  ).trim();
}

/**
 * Syncs pulled Finanda transactions into the RawData tab.
 *
 * Finanda sometimes replaces a transaction with a new record (new TransID),
 * e.g. when a card charge is first loaded and later re-loaded as invoiced.
 * The old record disappears from Finanda, but RawData used to keep it, so the
 * same expense showed up twice. To fix this, a RawData row for this month whose
 * TransID is no longer returned by Finanda is treated as "stale" and is matched
 * (same account, date, debit, credit and preferably description) to its
 * replacement:
 *   - replacement is new          -> the stale row is overwritten in place
 *                                    (Category/Comment are kept)
 *   - replacement already in sheet -> the stale row's Category/Comment move to it
 *                                    (if it has none) and the stale row is removed
 *   - no replacement, uncategorized -> the stale row is removed
 *   - no replacement, categorized  -> kept, logged (already counted in the year sheet)
 * Two real identical purchases are never merged, because one side must have
 * disappeared from Finanda first.
 *
 * @param {Array<Object>} transactions Pulled transactions for the month
 * @param {string} monthKey e.g. "2026-10"
 * @param {Set<string>} [fetchedIds] All TransIDs Finanda currently returns (all months)
 * @param {string} [fetchedFromDate] Earliest TransDate Finanda returned ("YYYY-MM-DD")
 * @returns {{added:number, replaced:number, merged:number, removed:number, flagged:number}}
 */
function appendNewRawTransactions(transactions, monthKey, fetchedIds, fetchedFromDate) {
  const stats = { added: 0, replaced: 0, merged: 0, removed: 0, flagged: 0 };
  transactions = transactions || [];
  const sheet = ensureRawDataTab();
  const tz = sheet.getParent().getSpreadsheetTimeZone();
  const lastRow = sheet.getLastRow();
  const lastCol = Math.max(sheet.getLastColumn(), RAW_DATA_HEADERS.length);
  const headers = sheet
    .getRange(1, 1, 1, lastCol)
    .getValues()[0]
    .map((h) => String(h).trim());
  const colOf = (name) => {
    const idx = headers.indexOf(name);
    return idx !== -1 ? idx : RAW_DATA_HEADERS.indexOf(name);
  };
  const transIdIdx = colOf("TransID");
  const monthIdx = colOf("MonthKey");
  const categoryIdx = colOf("Category");
  const commentIdx = colOf("Comment");

  const toRowValues = (item, transId, category, comment) =>
    headers.map((header, idx) => {
      if (idx === transIdIdx) return transId;
      if (idx === monthIdx) return "'" + monthKey; // force literal text
      if (idx === categoryIdx) return category || "";
      if (idx === commentIdx) return comment || "";
      return header ? formatRawCellValue(item[header]) : "";
    });

  // Read existing rows
  const values = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, lastCol).getValues() : [];
  const rows = values.map((rowValues, i) => {
    const item = {};
    headers.forEach((h, idx) => {
      if (h) item[h] = rowValues[idx];
    });
    return {
      rowIndex: i + 2,
      transId: String(rowValues[transIdIdx] || "").trim(),
      monthKey: parseMonthKey(rowValues[monthIdx]),
      category: String(rowValues[categoryIdx] || "").trim(),
      comment: String(rowValues[commentIdx] || "").trim(),
      item: item,
    };
  });
  const existingIds = new Set(rows.map((r) => r.transId).filter(Boolean));

  // Genuinely new transactions (by TransID)
  const newItems = [];
  transactions.forEach((item) => {
    const transId = getRawTransId(item);
    if (transId && existingIds.has(transId)) return;
    if (transId) existingIds.add(transId);
    newItems.push({ item: item, transId: transId, claimed: false });
  });

  // Reconcile stale rows - only when Finanda's data covers the whole month
  const canReconcile =
    fetchedIds && fetchedIds.size > 0 && fetchedFromDate && fetchedFromDate <= monthKey + "-01";
  const rowsToDelete = [];
  if (canReconcile) {
    const monthRows = rows.filter((r) => r.monthKey === monthKey && r.transId);
    const stale = monthRows.filter((r) => !fetchedIds.has(r.transId));
    const liveRows = monthRows
      .filter((r) => fetchedIds.has(r.transId))
      .map((r) => ({ row: r, claimed: false }));

    const findMatch = (pool, getItem, fp, withDesc) =>
      pool.find((c) => !c.claimed && rawFingerprint(getItem(c), tz, withDesc) === fp);

    stale.forEach((staleRow) => {
      let newMatch = null;
      let liveMatch = null;
      for (const withDesc of [true, false]) {
        const fp = rawFingerprint(staleRow.item, tz, withDesc);
        newMatch = findMatch(newItems, (c) => c.item, fp, withDesc);
        if (newMatch) break;
        liveMatch = findMatch(liveRows, (c) => c.row.item, fp, withDesc);
        if (liveMatch) break;
      }

      if (newMatch) {
        newMatch.claimed = true;
        sheet
          .getRange(staleRow.rowIndex, 1, 1, headers.length)
          .setValues([toRowValues(newMatch.item, newMatch.transId, staleRow.category, staleRow.comment)]);
        stats.replaced++;
        Logger.log(`[RawData] Replaced stale ${staleRow.transId} with ${newMatch.transId} (row ${staleRow.rowIndex})`);
      } else if (liveMatch) {
        liveMatch.claimed = true;
        const live = liveMatch.row;
        if (!live.category && staleRow.category) {
          sheet.getRange(live.rowIndex, categoryIdx + 1).setValue(staleRow.category);
          sheet.getRange(live.rowIndex, commentIdx + 1).setValue(staleRow.comment);
          rowsToDelete.push(staleRow.rowIndex);
          stats.merged++;
          Logger.log(`[RawData] Merged stale ${staleRow.transId} into ${live.transId}, removed row ${staleRow.rowIndex}`);
        } else if (!staleRow.category) {
          rowsToDelete.push(staleRow.rowIndex);
          stats.removed++;
          Logger.log(`[RawData] Removed uncategorized stale duplicate ${staleRow.transId} of ${live.transId}`);
        } else {
          // Both were categorized -> the amount was already counted twice in the year sheet
          sheet
            .getRange(staleRow.rowIndex, commentIdx + 1)
            .setValue(`DUPLICATE of ${live.transId} - fix the year sheet, then delete this row`);
          stats.flagged++;
          Logger.log(`[RawData] WARNING: stale ${staleRow.transId} and ${live.transId} are both categorized`);
        }
      } else if (!staleRow.category) {
        rowsToDelete.push(staleRow.rowIndex);
        stats.removed++;
        Logger.log(`[RawData] Removed uncategorized row ${staleRow.transId} - no longer exists in Finanda`);
      } else {
        Logger.log(`[RawData] Notice: categorized row ${staleRow.transId} no longer exists in Finanda, kept`);
      }
    });
  } else {
    Logger.log(`[RawData] Skipping stale-row reconciliation for ${monthKey} (fetched data does not cover the month)`);
  }

  // Delete bottom-up so row indexes stay valid
  rowsToDelete
    .sort((a, b) => b - a)
    .forEach((rowIndex) => sheet.deleteRow(rowIndex));

  const toAppend = newItems
    .filter((c) => !c.claimed)
    .map((c) => toRowValues(c.item, c.transId, "", ""));
  if (toAppend.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, toAppend.length, headers.length).setValues(toAppend);
  }
  stats.added = toAppend.length;
  Logger.log(`[RawData] Sync ${monthKey}: ${JSON.stringify(stats)}`);
  return stats;
}

/**
 * Retrieves all transactions for a given monthKey from RawData.
 * Uses robust month matching that handles Date objects, strings,
 * and falls back to transaction dates (TransDate, TransValueDate).
 * 
 * @param {string} monthKey The year-month string (e.g. "2026-08")
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

  let transIdColIdx = -1;
  let monthColIdx = -1;
  let categoryColIdx = -1;
  let commentColIdx = -1;

  // Search backwards to accurately locate the trailing Category/Comment columns
  for (let i = headers.length - 1; i >= 0; i--) {
    const h = String(headers[i]).trim().toLowerCase();
    if (commentColIdx === -1 && h === "comment") {
      commentColIdx = i;
    }
    if (categoryColIdx === -1 && (h === "category" || h === "categorie")) {
      categoryColIdx = i;
    }
    if (monthColIdx === -1 && h === "monthkey") {
      monthColIdx = i;
    }
    if (transIdColIdx === -1 && h === "transid") {
      transIdColIdx = i;
    }
  }

  // Fallbacks if not located by header name
  if (transIdColIdx === -1) transIdColIdx = 0;
  if (categoryColIdx === -1) categoryColIdx = headers.length - 2;
  if (commentColIdx === -1) commentColIdx = headers.length - 1;
  if (monthColIdx === -1) monthColIdx = headers.length - 3;

  const results = [];
  for (let r = 1; r < allValues.length; r++) {
    const rowValues = allValues[r];
    const item = {};
    headers.forEach((h, idx) => {
      item[h] = rowValues[idx];
    });

    item.Amount = Number(item.Amount) || 0;
    item.Credit = Number(item.Credit) || 0;
    item.Debit = Number(item.Debit) || 0;
    item._id = item.TransID; // Compatibility

    // Check MonthKey with multiple fallbacks
    const rowMonth = rowValues[monthColIdx];
    const parsedRowMonth = parseMonthKey(rowMonth);
    const transDateMonth = item.TransDate ? parseMonthKey(item.TransDate) : "";
    const transValueDateMonth = item.TransValueDate ? parseMonthKey(item.TransValueDate) : "";

    const isMatch =
      parsedRowMonth === monthKey ||
      (item.TotalPayments ? transValueDateMonth === monthKey : transDateMonth === monthKey) ||
      transDateMonth === monthKey ||
      transValueDateMonth === monthKey;

    if (isMatch) {
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
  let transIdColIdx = -1;
  let categoryColIdx = -1;
  let commentColIdx = -1;

  for (let i = headers.length - 1; i >= 0; i--) {
    const h = String(headers[i]).trim().toLowerCase();
    if (commentColIdx === -1 && h === "comment") {
      commentColIdx = i + 1;
    }
    if (categoryColIdx === -1 && (h === "category" || h === "categorie")) {
      categoryColIdx = i + 1;
    }
    if (transIdColIdx === -1 && h === "transid") {
      transIdColIdx = i + 1;
    }
  }

  if (categoryColIdx === -1) categoryColIdx = headers.length - 1;
  if (transIdColIdx === -1) transIdColIdx = 1;

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
  let categoryColIdx = -1;
  let commentColIdx = -1;

  for (let i = headers.length - 1; i >= 0; i--) {
    const h = String(headers[i]).trim().toLowerCase();
    if (commentColIdx === -1 && h === "comment") {
      commentColIdx = i + 1;
    }
    if (categoryColIdx === -1 && (h === "category" || h === "categorie")) {
      categoryColIdx = i + 1;
    }
  }

  if (categoryColIdx === -1) categoryColIdx = headers.length - 1;

  updates.forEach((up) => {
    if (up.rowIndex) {
      sheet.getRange(up.rowIndex, categoryColIdx).setValue(up.category);
      if (commentColIdx > 0 && up.comment !== undefined) {
        sheet.getRange(up.rowIndex, commentColIdx).setValue(up.comment);
      }
    }
  });
}
