function doGet(e) {
  return ContentService.createTextOutput("Hello World!");
}

function onInstall(e) {
  // This adds the menu items immediately upon installation
  onOpen(e); 
}

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  
  ui.createMenu('משיכת מידע')
      .addItem('משיכת מידע', 'UpdateByRange')
      .addToUi();
}

// Function to call to programmatically create the installable trigger
function setupTrigger() {
  var triggers = ScriptApp.getProjectTriggers();
  var triggerExists = false;
  
  // Check if it's already there to avoid duplicates
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'openFinandaSideBar') {
      triggerExists = true;
      break;
    }
  }
  
  if (!triggerExists) {
    ScriptApp.newTrigger('openFinandaSideBar')
      .forSpreadsheet(SpreadsheetApp.getActive())
      .onOpen()
      .create();
  }
}

function itemTemplate (item, defaultGroup, groups, columnIndex, accountsMap, year) {
  const amount = Math.abs(item.Amount);
  const date = item.TotalPayments ? new Date(item.TransValueDate) : new Date(item.TransDate);

  Logger.log(`item.Description: ${item.Description}, defaultGroup: ${defaultGroup}, groups: ${groups.length}, columnIndex: ${columnIndex}, year: ${year}`)
  return `
      <div id="${item._id}" class="balance__item">
        <div>
          <b>${accountsMap[item.AccountNumber] || item.instituteDesc} - ${item.Description}</b>
          <div>${amount}${item.TransCurrency === 'ILS' ? '₪' : '$'}</div>
          <div>${date.getDate()} ${months[date.getMonth()]}</div>
        </div>
        <div class="balance__actions">
          <select class="select-menu" name="select_id" id="select_id" onchange="onTransactionChange(this.value, '${amount}', '${escape(item.Description)}', '${columnIndex}', '${item._id}', '${year}')">
              <option value="">שייך תנועה</option>
              ${groups.map(group => `<option value="${group.id}">${group.name}</option>`)}
          </select>
          <button class="standard-button" onclick="onTransactionChange('${defaultGroup}', '${item.Amount}', '${escape(item.Description)}', '${columnIndex}', '${item._id}', '${year}')">חריגה</button>
          <button class="standard-button" onclick="document.getElementById('${item._id}').remove()">דלג</button>
        </div>
      </div>
  `
}

function openSideBar (income, expanses, columnIndex, accountsMap, year) {
  Logger.log('Open sidebar')
  var htmlOutput = HtmlService
    .createHtmlOutputFromFile('TransactionsSideBar.html')
    .setWidth(400)
    .setTitle('סיכום חלוקת פעולות');

    const groups = getGroups();

    htmlOutput.append('<h2>הכנסות חריגות</h2>');
    const incomGroups = groups.filter(item => item.type === 'הכנסה');
    (income || []).forEach(item => {
      htmlOutput.append(itemTemplate(item, DEFAULT_GROUPS.income, incomGroups, columnIndex, accountsMap, year));
    });

    htmlOutput.append('<h2>הוצאות חריגות</h2>');
    const expansesGroups = groups.filter(item => item.type !== 'הכנסה' && item.type !== 'הוצאה משתנה');
    (expanses || []).forEach(item => {
      htmlOutput.append(itemTemplate(item, DEFAULT_GROUPS.expanses, expansesGroups, columnIndex, accountsMap, year));
    });
    
    Logger.log('sidebar html is ready')
  SpreadsheetApp.getUi().showSidebar(htmlOutput);
}

function showExpansesFormDialog() {
  var html = HtmlService.createHtmlOutput('<iframe height="100%" width="100%" src="https://docs.google.com/spreadsheets/d/e/2PACX-1vRunJmjpGvKAZnQQdKhP9T9Wn4hvnhd2DkLJDQAbB68aQSlwT3KnP3bg-nqFVCdDXR_pfCNHPK7mSB_/pubhtml?gid=818273642&amp;single=true&amp;widget=true&amp;headers=false"></iframe>')
      .setWidth(400)
      .setHeight(600);
  SpreadsheetApp.getUi()
      .showModalDialog(html, 'הוצאות חריגות');
}

/**
 * Retrieves the Finanda credentials to be used by the sidebar.
 * This function is called from the client-side using google.script.run.getCredentials()
 * 
 * @returns {Object} An object containing the user and pass
 */
function getCredentials() {
  try {
    return {
      user: getFinandaUser(),
      pass: getFinandaPassword()
    };
  } catch (error) {
    Logger.log("Failed to retrieve credentials:", error);
    throw new Error("Unable to retrieve Finanda credentials.");
  }
}

function updateCellFromSideBar(Amount, Description, groupId, columnIndex, year) {
  const targetSheet = year ? SpreadsheetApp.getActiveSpreadsheet().getSheetByName(year.toString()) : SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (!targetSheet) {
    throw new Error(`לא נמצא גיליון עבור השנה ${year}`);
  }

  const activeRange = targetSheet.getRange(4, 1, 100, 1);
  const values = activeRange.getValues().map(i => i[0]);
  const rowIndex = values.indexOf(parseInt(groupId)) + 1;

  // console.log('groupId', groupId, columnIndex, rowIndex, values)
  updateCell(targetSheet.getRange(rowIndex + 3, columnIndex), [{ Amount, Description: unescape(Description) }]);
  SpreadsheetApp.getActive().toast(`עודכן`);
}

function findColumnForMonth(sheet, year, month) {
  const lastColumn = sheet.getLastColumn();
  if (lastColumn < 2) return null;
  
  // Row 3 contains the month headers
  const headerValues = sheet.getRange(3, 2, 1, lastColumn - 1).getValues()[0];
  
  for (let i = 0; i < headerValues.length; i++) {
    const cellValue = headerValues[i];
    if (cellValue instanceof Date || (typeof cellValue === 'string' && cellValue !== '')) {
      const date = new Date(cellValue);
      if (!isNaN(date.getTime())) {
        if (date.getFullYear() === year && date.getMonth() === month) {
          return i + 2; // +2 offset (starts at column B)
        }
      }
    }
  }
  return null;
}

function updateCell (range, transactions) {
  if (transactions && transactions.length) {
    const sum = transactions.map(i => Math.abs(i.Amount));
    let notes = transactions.map(i => `${i.Description}: ${Math.abs(i.Amount)}`);
    
    Logger.log(`[updateCell] Row: ${range.getRowIndex()}, Column: ${range.getColumn()}: =${sum.join('+')} note: ${notes.join('\n')}`);

    const currentValues = range.getFormula().replace('=', '').split('+').map(Number);
    const mergedValues = [...new Set([...currentValues, ...sum])].filter(Boolean);
    
    range.setFormula(`=${mergedValues.join('+')}`);

    const currentNote = range.getNote() || '';
    if (currentNote) {
      const filteredNotes = notes.filter(i => !currentNote.includes(i));
      const notesString = [currentNote, ...filteredNotes].filter(Boolean).join(' \n');
      range.setNote(notesString);
    } else {
      range.setNote(notes.join(' \n'));
    }
  }
}

async function UpdateByRange() {
  openFinandaSideBar()
}


function updateSheetData(income, expanses, accountsMap, year, month) {  
  // If year or month are undefined, try fallback to getDateRange
  if (year === undefined || month === undefined) {
    const range = getDateRange();
    year = range.year;
    month = range.month;
  }

  const targetSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(year.toString());
  if (!targetSheet) {
    SpreadsheetApp.getUi().alert(`שגיאה: לא נמצא גיליון בשם ${year}`);
    return;
  }

  // Activate the target sheet tab
  targetSheet.activate();

  const baseColumnIndex = findColumnForMonth(targetSheet, year, month);
  if (!baseColumnIndex) {
    SpreadsheetApp.getUi().alert(`שגיאה: לא נמצאה עמודה מתאימה לחודש ${month + 1} בגיליון ${year}`);
    return;
  }

  // The actual values column is the second column of the month's pair (baseColumnIndex + 1)
  const actualColumnIndex = baseColumnIndex + 1;

  const startRow = 4;
  const lastRow = targetSheet.getLastRow();
  if (lastRow < startRow) {
    SpreadsheetApp.getUi().alert(`שגיאה: אין שורות נתונים בגיליון ${year}`);
    return;
  }

  Logger.log(`[updateSheetData] Starting update for sheet: ${year}, column: ${actualColumnIndex}, rows: ${startRow} to ${lastRow}`);

  for (let rowIndex = startRow; rowIndex <= lastRow; rowIndex++) {
    const groupId = targetSheet.getRange(rowIndex, 1).getValue();
    if (!groupId) continue;

    const groupIdStr = groupId.toString().trim();
    if (groupIdStr === "" || groupIdStr === DEFAULT_GROUPS.income || groupIdStr === DEFAULT_GROUPS.expanses) {
      continue;
    }

    const transactions = income[groupIdStr] || expanses[groupIdStr];
    if (transactions && transactions.length) {
      const range = targetSheet.getRange(rowIndex, actualColumnIndex);
      updateCell(range, transactions);
    }
  }

  Logger.log(`[updateSheetData] End of update loop`);

  if (income[DEFAULT_GROUPS.income]?.length || expanses[DEFAULT_GROUPS.expanses]?.length) {
    openSideBar(income[DEFAULT_GROUPS.income], expanses[DEFAULT_GROUPS.expanses], actualColumnIndex, accountsMap, year);
  }

  SpreadsheetApp.getUi().alert('העדכון הסתיים בהצלחה');
}

const months = [
			'ינואר',
			'פברואר',
			'מרץ',
			'אפריל',
			'מאי',
			'יוני',
			'יולי',
			'אוגוסט',
			'ספטמבר',
			'אוקטובר',
			'נובמבר',
			'דצמבר'
		];
