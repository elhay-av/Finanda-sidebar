function filterThisMonth(data, monthString) {
  Logger.log(`Month string is ${monthString} : ${data.toString()}`);
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
    // console.log('PUSH', selectedGroup || defaultGroup, acc[selectedGroup || defaultGroup])
    acc[selectedGroup || defaultGroup].push(item);
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

function getDateRange() {
  const activeSheet = getProtectedActiveSpreadsheet().getActiveSheet();

  const activeRange = activeSheet.getActiveRange();
  const columnIndex = activeRange.getColumn();
  const row = activeRange.getRow();

  const rangeDateHeader = activeSheet
    .getRange(row - 1, columnIndex - 1, 1, 1)
    .getValue();
  const selectedDate = new Date(rangeDateHeader);
  Logger.log(`Selected month by range ${rangeDateHeader} ${selectedDate}`);

  return { year: selectedDate.getFullYear(), month: selectedDate.getMonth() };
}

function processFinandaData(data, year, month) {
  checkAuthorization();
  if (year === undefined || month === undefined) {
    const range = getDateRange();
    year = range.year;
    month = range.month;
  }

  const monthString = `${year}-${(month + 1).toString().padStart(2, "0")}`;
  const monthlyTransactions = filterThisMonth(data, monthString);
  const transactionsWithDebit =
    filterTransactionsWithoutDebit(monthlyTransactions);
  const transactionsByType = splitByType(transactionsWithDebit);

  const expanses = groupByBalanceCategory(
    transactionsByType.expanses,
    DEFAULT_GROUPS.expanses,
  );
  const income = groupByBalanceCategory(
    transactionsByType.income,
    DEFAULT_GROUPS.income,
  );

  const accountsMap = data?.accounts?.CheckingAccounts.reduce(
    (acc, account) => {
      acc[account.AccountNum] = account.AccountDesc;
      return acc;
    },
    {},
  );

  // TODO: (Elhay) Find מידע על תשלומים כמו ״בר מים״
  updateSheetData(income, expanses, accountsMap, year, month);
}
