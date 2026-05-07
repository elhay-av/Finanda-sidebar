const GROUPS_SHEET_NAME = 'settings-groups'

let grouppingData = null;

function getGrouppingDataRange() {
  if (grouppingData) {
    return grouppingData;
  }

  const settingsSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(GROUPS_SHEET_NAME);
  ranges = settingsSheet.getNamedRanges().reduce((acc, range) => {
    acc[range.getName()] = range.getRange().getValues();
    return acc;
  }, {});

  grouppingData = ranges['grouppingData'];
  return grouppingData;
}

const DEFAULT_GROUPS = {
  income: "5",
  expanses: "30",
};

function getGroups () {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('מאזן')
  const range = sheet.getRange(6, 1, 50, 4);
  const values = range.getValues();

  const idColIndex = 0;
  const typeColIndex = 2;
  const nameColIndex = 3;

  return values.filter(i => i[typeColIndex] !== 'הוצאה משתנה').map((row, index) => {
    Logger.log(`Row: ${index + 1} -> groupId: ${row[idColIndex]}, type: ${row[typeColIndex]}, name: ${row[nameColIndex]}`);
    return {
      name: row[nameColIndex],
      type: row[typeColIndex],
      id: row[idColIndex]
    }
  });
}

function getEmptyGroups() {
  return getGroups().reduce((acc, i) => {
    acc[i.id] = [];
    return acc;
  }, {})
}

function getGroupsMapping() {
  // sourceCatId	key	value	destinationGroup	קבוצה במאזן
  const columnIndexes = {
    sourceCatId: 0,
    key: 1,
    value: 2,
    destinationGroup: 3
  };

  const data = getGrouppingDataRange();
  const returnData = data.reduce((acc, row) => {
    const groupId = row[columnIndexes.sourceCatId];
    let newGroup;
    
    if (row[columnIndexes.key]) {
      newGroup = {
        "key": row[columnIndexes.key],
        "value": row[columnIndexes.value],
        "group": row[columnIndexes.destinationGroup] + ''
      };
    } else {
      newGroup = row[columnIndexes.destinationGroup] + '';
    }
     

    if (!groupId) {
      return acc;
    }

    // if group not exist yet
    if (!acc[groupId]) {
      if (typeof newGroup === 'object') {
        acc[groupId] = [newGroup];  
      } else {
        acc[groupId] = newGroup;
      }

      return acc;
    }
    // if group exist as array
    if (Array.isArray(acc[groupId])) {
      acc[groupId].push(newGroup);
      return acc;
    }

    // if group exist as object
    acc[groupId] = [acc[groupId], newGroup];

    return acc;
  }, {});

  // console.log('Groupping data:', JSON.stringify(returnData, null, 2));
  return returnData;
}