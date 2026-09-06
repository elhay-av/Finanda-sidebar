function filterThisMonth(data, monthString) {
  Logger.log(`Month string is ${monthString}`);
  if (!data?.accounts?.CheckingAccounts?.length) {
    throw new Error("Accounts must be greater than 0");
  }
  const allAccounts = data.accounts.CheckingAccounts.find(
    (account) => account.AccountID === "unified_checking",
  );
  if (!allAccounts) {
    throw new Error("unified_checking not found");
  }
  const transactions = allAccounts?.Transactions;
  if (!transactions) {
    throw new Error("allAccounts.Transactions, not found");
  }
  const filteredData = transactions.filter((item) => {
    if (item.TotalPayments) {
      return item.TransValueDate.startsWith(monthString);
    } else {
      return item.TransDate.startsWith(monthString);
    }
  });
  return filteredData;
}

function filterTransactionsWithoutDebit(transactions) {
  if (!transactions.length) {
    throw new Error("Transactions must be greater than 0");
  }
  return transactions.filter((item) => item.Debit !== 0 || item.Credit > 0);
}

function splitByType(transactions) {
  return transactions.reduce(
    (acc, item) => {
      if (item.Credit > 0 || item.Debit < 0) {
        if (item.CatGroup !== "העברות פנימיות") {
          acc.income.push(item);
        }
      } else {
        acc.expanses.push(item);
      }

      return acc;
    },
    {
      expanses: [],
      income: [],
    },
  );
}

function groupByBalanceCategory(transactions, defaultGroup) {
  const groupsToCategory = getGroupsMapping();
  const res = transactions.reduce((acc, item) => {
    let selectedGroup = groupsToCategory[item.category];
    if (selectedGroup === "SKIP") {
      return acc;
    }

    if (Array.isArray(selectedGroup)) {
      selectedGroup = selectedGroup.find((condition) => {
        if (typeof condition.value === "string") {
          return item[condition.key]?.includes(condition.value);
        } else {
          return item[condition.key] === condition.value;
        }
      })?.group;
    }
    const groupKey = selectedGroup || defaultGroup;
    if (!acc[groupKey]) {
      Logger.log(`Warning: groupKey '${groupKey}' not found in acc. Item category: ${item.category}`);
      acc[groupKey] = [];
    }
    // console.log('PUSH', groupKey, acc[groupKey])
    acc[groupKey].push(item);
    return acc;
  }, getEmptyGroups());

  const filteredGroups = Object.entries(res).reduce((acc, [key, val]) => {
    if (val.length) {
      acc[key] = val;
    }
    return acc;
  }, {});
  return filteredGroups;
}

function openFinandaSideBar(authUrl) {
  var template = HtmlService.createTemplateFromFile("FinandaSideBar.html");
  template.authRequired = !!authUrl;
  template.authUrl = authUrl || "";

  var htmlOutput = template
    .evaluate()
    .setWidth(400)
    .setTitle("Elhay's Family budget");

  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}



function processFinandaData(data, year, month) {
  checkAuthorization();
  if (year == null || month == null) {
    throw new Error("שגיאה בעיבוד הנתונים - שנה או חודש חסרים");
  }

  const monthString = `${year}-${(month + 1).toString().padStart(2, "0")}`;
  const monthlyTransactions = filterThisMonth(data, monthString);
  Logger.log(`Monthly transactions fetched from Finanda: ${monthlyTransactions.length}`);
  const transactionsWithDebit =
    filterTransactionsWithoutDebit(monthlyTransactions);

  // 1. Ensure RawData tab exists and append newly pulled transactions (deduplicating by TransID)
  ensureRawDataTab();
  const addedCount = appendNewRawTransactions(transactionsWithDebit, monthString);
  Logger.log(`Appended ${addedCount} new transactions to RawData for ${monthString}`);

  // 2. Read back all transactions for this month from RawData
  const monthRecords = getRawTransactionsForMonth(monthString);
  Logger.log(`Total transactions in RawData for ${monthString}: ${monthRecords.length}`);

  // 3. Find uncategorized records (Category is empty)
  const uncategorizedRecords = monthRecords.filter(
    (rec) => !rec.category || rec.category.trim() === "",
  );

  Logger.log(`Uncategorized transactions in RawData: ${uncategorizedRecords.length}`);

  const accountsMap =
    data?.accounts?.CheckingAccounts?.reduce((acc, account) => {
      acc[account.AccountNum] = account.AccountDesc;
      return acc;
    }, {}) || {};

  if (uncategorizedRecords.length === 0) {
    getProtectedUi().alert("כל הפעולות לחודש זה כבר סווגו בגיליון.");
    return;
  }

  // 4. Categorize uncategorized transactions using groups mapping rules
  const groupsToCategory = getGroupsMapping();
  const rawUpdates = [];
  const newlyCategorized = {};
  const uncategorizedIncome = [];
  const uncategorizedExpanses = [];

  for (let i = 0; i < uncategorizedRecords.length; i++) {
    const rec = uncategorizedRecords[i];
    const item = rec.item;

    const isIncome = item.Credit > 0 || item.Debit < 0;

    // Skip internal transfers
    if (isIncome && item.CatGroup === "העברות פנימיות") {
      rawUpdates.push({
        rowIndex: rec.rowIndex,
        category: "SKIP",
        comment: "skipped - internal transfer (העברות פנימיות)",
      });
      continue;
    }

    let selectedGroup = groupsToCategory[item.category];
    if (selectedGroup === "SKIP") {
      rawUpdates.push({
        rowIndex: rec.rowIndex,
        category: "SKIP",
        comment: "skipped by group-settings rule",
      });
      continue;
    }

    if (Array.isArray(selectedGroup)) {
      selectedGroup = selectedGroup.find((condition) => {
        if (typeof condition.value === "string") {
          return item[condition.key]?.includes(condition.value);
        } else {
          return item[condition.key] === condition.value;
        }
      })?.group;
    }

    if (selectedGroup) {
      const groupKey = selectedGroup.toString();
      rawUpdates.push({
        rowIndex: rec.rowIndex,
        category: groupKey,
        comment: "categorized by group-settings",
      });
      if (!newlyCategorized[groupKey]) {
        newlyCategorized[groupKey] = [];
      }
      newlyCategorized[groupKey].push(item);
    } else {
      // Keep as uncategorized for user to decide in sidebar
      if (isIncome) {
        uncategorizedIncome.push(item);
      } else {
        uncategorizedExpanses.push(item);
      }
    }
  }

  // 5. Batch update RawData with newly categorized and skipped entries
  batchUpdateRawDataCategories(rawUpdates);

  // 6. Update the year sheet with newly categorized items and open sidebar for uncategorized
  updateSheetWithNewTransactions(
    newlyCategorized,
    uncategorizedIncome,
    uncategorizedExpanses,
    accountsMap,
    year,
    month,
  );
}
