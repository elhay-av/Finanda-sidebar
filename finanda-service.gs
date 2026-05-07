const DEVICE_ID = '1253';

const groupsToCategory = getGroupsMapping();
const globalHeaders = {
  // accept: "application/json, text/plain, */*",
  // "accept-language": "en-US,en;q=0.9,he;q=0.8",
  // origin: "https://premium.finanda.co.il",
  // priority: "u=1, i",
  // "sec-ch-ua": `"Chromium";v="142", "Google Chrome";v="142", "Not_A Brand";v="99"`,
  // "sec-ch-ua-mobile": "?0",
  // "sec-ch-ua-platform": "macOS",
  // "sec-fetch-dest": "empty",
  // "sec-fetch-mode": "cors",
  // "sec-fetch-site": "same-site",
  // "user-agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/142.0.0.0 Safari/537.36"
};

function getURLSearchParams(params) {
  return Object.entries(params).reduce((acc, [key, val]) => {
    if (acc) {
      acc += '&'
    }
    acc += `${key}=${encodeURIComponent(val).replace('!', '%21')}`
    // acc += `${key}=${val}`
    return acc;
  }, '');
}

function login () {
  const session = getFinandaSession();
  if (session) {
    console.log('Session restored from cache, no login required');
    return session;
  }

  const body = getURLSearchParams({
			password: getFinandaPassword(),
			device: `Chrome-${DEVICE_ID}`,
			version: '1.63',
			appVersion: '1.73',
			// tzOffset: `${(new Date()).getTimezoneOffset()}`,
			// model: 'desktop-Apple',
			checkSubscription: 'true',
			caller: 'web',
			userid: 'elhayav@gmail.com',
		});

console.log('bodyPramsList', body)
	let response = UrlFetchApp.fetch('https://cloud.finanda.co.il/login', {
    muteHttpExceptions: true,
		method: 'POST',
    ...globalHeaders,
		contentType: 'application/x-www-form-urlencoded',
		payload: body
	});

  console.log('Part-1 auth with user and password:', response.getContentText())
  let data = {};
  try {
    data = JSON.parse(response.getContentText());
  } catch {
    throw new Error(`login with user and password failed: ${response.getResponseCode()}`);
  }
  

  if (response.getResponseCode() === 400 && data.errorMessage === "authentication-required") {
    const mfaBody = getURLSearchParams({
      encSession: data.session,
      mfaMethod: data.authenticationMethods[0].type,
      encMfaInput: data.authenticationMethods[0].value,
      validationTyp: 'user'
    });
    const mfaResponse = UrlFetchApp.fetch('https://cloud.finanda.co.il/requestMFA', {
      muteHttpExceptions: true,
      method: 'POST',
      ...globalHeaders,
      contentType: 'application/x-www-form-urlencoded',
      payload: mfaBody
    });
    console.log('Part-2 generate mfa code:', mfaResponse.getContentText())
    
    if (mfaResponse.getResponseCode() === 200) {
      const mfaCode = SpreadsheetApp.getUi().prompt('הכנס קוד שנשלח בSMS');

      const mfaAuthBody = getURLSearchParams({
        encSession: data.session,
        device: `Chrome-${DEVICE_ID}`,
        mfaCode: mfaCode.getResponseText(),
        validationType: 'user'
      });
      const mfaAuthResponse = UrlFetchApp.fetch('https://cloud.finanda.co.il/authenticateMFA', {
        muteHttpExceptions: true,
        method: 'POST',
        ...globalHeaders,
        contentType: 'application/x-www-form-urlencoded',
        payload: mfaAuthBody
      });
      console.log('Part-3 auth with mfa code:', mfaAuthResponse.getContentText())

      if (mfaAuthResponse.getResponseCode() === 200) {
        response = UrlFetchApp.fetch('https://cloud.finanda.co.il/login', {
          muteHttpExceptions: true,
          method: 'POST',
          ...globalHeaders,
          contentType: 'application/x-www-form-urlencoded',
          payload: body
        });

        console.log('Part-4 RE-auth user and password:', mfaAuthResponse.getContentText())
        data = JSON.parse(response.getContentText());
      } else {
        throw new Error(`wrong MFA code. status code (${mfaCode.getResponseText()}): ${mfaAuthResponse.getResponseCode()}, ${mfaResponse.getContentText()}`);
      }
    } else {
      throw new Error(`MFA Login failed. status code: ${mfaResponse.getResponseCode()}, ${mfaResponse.getContentText()}`);
    }
  }

	if (!data.success) {
		throw new Error('login data is undefined (data.success)');
	}
	if (data.userStatus !== 'verified') {
		throw new Error('userStatus is not verified:', JSON.stringify(data || 'No Data provided'));
	}

	if (data.session) {
    setFinandaSession(data.session);
		return data.session;
	} else {
		throw new Error('Session is empty', JSON.stringify(data || 'No Data provided'));
	}
}

function getProfileInitiation (sessionId) {
	const response = UrlFetchApp.fetch('https://cloud.finanda.co.il/profileInitiation', {
		method: 'POST',
    ...globalHeaders,
		contentType: 'application/x-www-form-urlencoded',
		payload: getURLSearchParams({
			caller: 'web',
			session: sessionId,
		})
	});
	
  if (response.getResponseCode() !== 200) {
    throw new Error(`profileInitiation failed. status code: ${response.getResponseCode()}, ${response.getContentText()}`);
  }

	const data = JSON.parse(response.getContentText());

	if (!data.success) {
		throw new Error(data);
	}

	return data;
}

function filterThisMonth (data, monthString) {
	if (!data?.accounts?.CheckingAccounts?.length) {
		throw new Error('Accounts must be greater than 0');
	}
	const allAccounts = data.accounts.CheckingAccounts.find(account => account.AccountID === 'unified_checking')
	if (!allAccounts) {
		throw new Error('unified_checking not found');
	}
	const transactions = allAccounts?.Transactions
	if (!transactions) {
		throw new Error('allAccounts.Transactions, not found');
	}
	const filteredData = transactions.filter(item => {
    if (item.TotalPayments) {
      return item.TransValueDate.startsWith(monthString);
    } else {
      return item.TransDate.startsWith(monthString);
    }
  });
	return filteredData;
}

function filterTransactionsWithoutDebit (transactions) {
	if (!transactions.length) {
		throw new Error('Transactions must be greater than 0');
	}
	return transactions.filter(item => item.Debit !== 0 || item.Credit > 0);
}

function splitByType (transactions) {
	return transactions.reduce((acc, item) => {
		if (item.Credit > 0 || item.Debit < 0) {
			if (item.CatGroup !== 'העברות פנימיות') {
				acc.income.push(item);
			}
		} else {
			acc.expanses.push(item);
		}

		return acc;
	}, {
		expanses: [], income: []
	});
}

function groupByBalanceCategory (transactions, defaultGroup) {
	const res = transactions.reduce((acc, item) => {
		let selectedGroup = groupsToCategory[item.category];
		if (selectedGroup === 'SKIP') {
			return acc;
		}

		if (Array.isArray(selectedGroup)) {
			selectedGroup = selectedGroup.find(condition => {
				if (typeof condition.value === 'string') {
					return item[condition.key]?.includes(condition.value);
				} else {
					return item[condition.key] === condition.value;
				}
			})?.group;
		}
    console.log('PUSH', selectedGroup || defaultGroup, acc[selectedGroup || defaultGroup])
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

function openFinandaSideBar () {
  var htmlOutput = HtmlService
    .createHtmlOutputFromFile('FinandaSideBar.html')
    .setWidth(400)
    .setTitle('טעינת מידע...');
    
  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}

function getDateRange() {
  const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  const activeRange = activeSheet.getActiveRange();
  const columnIndex = activeRange.getColumn();
  const row = activeRange.getRow();

  const rangeDateHeader = activeSheet.getRange(row - 1, columnIndex - 1, 1, 1).getValue();
  const selectedDate = new Date(rangeDateHeader);
  Logger.log(`Selected month by range ${rangeDateHeader} ${selectedDate}`);

  return {year: selectedDate.getFullYear(), month: selectedDate.getMonth()};
}
function processFinandaData(data) {
	const {year, month} = getDateRange();

	const monthString = `${year}-${(month + 1).toString().padStart(2, '0')}`;
	const monthlyTransactions = filterThisMonth(data, monthString);
	const transactionsWithDebit = filterTransactionsWithoutDebit(monthlyTransactions);
	const transactionsByType = splitByType(transactionsWithDebit);
  
	const expanses = groupByBalanceCategory(transactionsByType.expanses, DEFAULT_GROUPS.expanses)
	const income = groupByBalanceCategory(transactionsByType.income, DEFAULT_GROUPS.income);

  const accountsMap = data?.accounts?.CheckingAccounts.reduce((acc, account) => {
    acc[account.AccountNum] = account.AccountDesc;
    return acc;
  }, {});

	// TODO: (Elhay) Find מידע על תשלומים כמו ״בר מים״
	updateSheetData(income, expanses, accountsMap);
}