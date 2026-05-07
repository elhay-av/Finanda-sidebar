function onOpen() {
  var ui = SpreadsheetApp.getUi();
  
  ui.createMenu('משיכת מידע')
      .addItem('פיננדה', 'UpdateByRange')
      .addItem('סיכום הוצאות חריגות', 'showExpansesFormDialog')
      .addToUi();
}

function itemTemplate (item, defaultGroup, groups, columnIndex, accountsMap) {
  const amount = Math.abs(item.Amount);
  const date = item.TotalPayments ? new Date(item.TransValueDate) : new Date(item.TransDate);

  Logger.log(`item.Description: ${item.Description}, defaultGroup: ${defaultGroup}, groups: ${groups.length}, columnIndex: ${columnIndex}`)
  return `
      <div id="${item._id}" class="balance__item">
        <div>
          <b>${accountsMap[item.AccountNumber] || item.instituteDesc} - ${item.Description}</b>
          <div>${amount}${item.TransCurrency === 'ILS' ? '₪' : '$'}</div>
          <div>${date.getDate()} ${months[date.getMonth()]}</div>
        </div>
        <div class="balance__actions">
          <select class="select-menu" name="select_id" id="select_id" onchange="onTransactionChange(this.value, '${amount}', '${escape(item.Description)}', '${columnIndex}', '${item._id}')">
              <option value="">שייך תנועה</option>
              ${groups.map(group => `<option value="${group.id}">${group.name}</option>`)}
          </select>
          <button class="standard-button" onclick="onTransactionChange('${defaultGroup}', '${item.Amount}', '${escape(item.Description)}', '${columnIndex}', '${item._id}')">חריגה</button>
          <button class="standard-button" onclick="document.getElementById('${item._id}').remove()">דלג</button>
        </div>
      </div>
  `
}

function openSideBar (income, expanses, columnIndex, accountsMap) {
  var htmlOutput = HtmlService
    .createHtmlOutputFromFile('TransactionsSideBar.html')
    .setWidth(400)
    .setTitle('סיכום חלוקת פעולות');

    const groups = getGroups();

    htmlOutput.append('<h2>הכנסות חריגות</h2>');
    const incomGroups = groups.filter(item => item.type === 'הכנסה');
    (income || []).forEach(item => {
      htmlOutput.append(itemTemplate(item, DEFAULT_GROUPS.income, incomGroups, columnIndex, accountsMap));
    });

    htmlOutput.append('<h2>הוצאות חריגות</h2>');
    const expansesGroups = groups.filter(item => item.type !== 'הכנסה' && item.type !== 'הוצאה משתנה');
    (expanses || []).forEach(item => {
      htmlOutput.append(itemTemplate(item, DEFAULT_GROUPS.expanses, expansesGroups, columnIndex, accountsMap));
    });
    
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
    console.error("Failed to retrieve credentials:", error);
    throw new Error("Unable to retrieve Finanda credentials.");
  }
}

function updateCellFromSideBar(Amount, Description, groupId, columnIndex) {
  const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const activeRange = activeSheet.getRange(4, 1, 100, 1);
  const values = activeRange.getValues().map(i => i[0]);
  const rowIndex = values.indexOf(parseInt(groupId)) + 1;

  // console.log('groupId', groupId, columnIndex, rowIndex, values)
  updateCell(activeSheet.getRange(rowIndex + 3, columnIndex), [{ Amount, Description: unescape(Description) }]);
  SpreadsheetApp.getActive().toast(`עודכן`);
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

function updateSheetData(income, expanses, accountsMap) {  
  const activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const activeRange = activeSheet.getActiveRange();
  const row = activeRange.getRow();
  const columnIndex = activeRange.getColumn();
  
  new Array(activeRange.getNumRows()).fill(true).forEach((_, index) => {
    const rowIndex = index + row;
    const groupId = activeSheet.getRange(rowIndex, 1).getValue();
    
    if (groupId.toString() === DEFAULT_GROUPS.income || groupId.toString() === DEFAULT_GROUPS.expanses) {
      return;
    }

    const transactions = income[groupId] || expanses[groupId];
    const range = activeSheet.getRange(rowIndex, columnIndex);

    updateCell(range, transactions);
  });

  if (income[DEFAULT_GROUPS.income]?.length || expanses[DEFAULT_GROUPS.expanses]?.length) {
    openSideBar(income[DEFAULT_GROUPS.income], expanses[DEFAULT_GROUPS.expanses], columnIndex, accountsMap);
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
