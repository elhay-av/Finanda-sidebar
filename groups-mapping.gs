const GROUPS_SHEET_NAME = "settings-groups";

var grouppingData = null;

function getGrouppingDataRange() {
  if (grouppingData) {
    return grouppingData;
  }

  const settingsSheet =
    getProtectedActiveSpreadsheet().getSheetByName(GROUPS_SHEET_NAME);
  ranges = settingsSheet.getNamedRanges().reduce((acc, range) => {
    acc[range.getName()] = range.getRange().getValues();
    return acc;
  }, {});

  grouppingData = ranges["grouppingData"];
  return grouppingData;
}

function getDefaultGroups() {
  return {
    income: getSettings()["DEFAULT_INCOOME_GROUP_ID"][0][0] + "",
    expanses: getSettings()["DEFAULT_EXPANSE_GROUP_ID"][0][0] + "",
  };
}

function getGroups() {
  const sheet = getProtectedActiveSpreadsheet().getSheetByName("מאזן");
  const range = sheet.getRange(7, 1, 50, 4);
  const values = range.getValues();

  const idColIndex = 0;
  const subTypeColIndex = 1;
  const typeColIndex = 2;
  const nameColIndex = 3;

  return values
    .filter((i) => !["פייבוקס", "מזומן"].includes(i[subTypeColIndex]))
    .map((row, index) => {
      // Logger.log(`Row: ${index + 1} -> groupId: ${row[idColIndex]}, type: ${row[typeColIndex]}, name: ${row[nameColIndex]}`);
      return {
        name: row[nameColIndex],
        type: row[typeColIndex],
        id: row[idColIndex],
      };
    });
}

function getEmptyGroups() {
  return getGroups().reduce((acc, i) => {
    acc[i.id] = [];
    return acc;
  }, {});
}

function getGroupsMapping() {
  // sourceCatId	key	value	destinationGroup	קבוצה במאזן
  const columnIndexes = {
    sourceCatId: 0,
    key: 1,
    value: 2,
    destinationGroup: 3,
  };

  const data = getGrouppingDataRange();
  const returnData = data.reduce((acc, row) => {
    const groupId = row[columnIndexes.sourceCatId];
    let newGroup;

    if (row[columnIndexes.key]) {
      newGroup = {
        key: row[columnIndexes.key],
        value: row[columnIndexes.value],
        group: row[columnIndexes.destinationGroup] + "",
      };
    } else {
      newGroup = row[columnIndexes.destinationGroup] + "";
    }

    if (!groupId) {
      return acc;
    }

    // if group not exist yet
    if (!acc[groupId]) {
      if (typeof newGroup === "object") {
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

/**
 * Adds a new auto-categorization rule row to the 'group-settings' sheet.
 *
 * Column layout in group-settings:
 * Col A (1): sourceCatId (item.category or rule.sourceCatId)
 * Col B (2): key / Search property (rule.searchProperty)
 * Col C (3): value / Search value (rule.searchValue)
 * Col D (4): destinationGroup (rule.destinationGroupId or "SKIP" if empty)
 * Col E (5): Source category / CatDesc (rule.sourceCategory)
 * Col F (6): Dest category / Table1[תיאור] (rule.destCategoryName or "")
 * Col G (7): Notes (rule.notes)
 *
 * @param {Object} rule The rule configuration from the sidebar modal
 * @returns {Object} { success: boolean, message: string }
 */
function addAutoCategoryRule(rule) {
  checkAuthorization();
  if (!rule) {
    throw new Error("נתוני החוק ריקים");
  }

  const spreadsheet = getProtectedActiveSpreadsheet();
  const sheet =
    spreadsheet.getSheetByName("group-settings") ||
    spreadsheet.getSheetByName(GROUPS_SHEET_NAME);

  if (!sheet) {
    throw new Error("גיליון group-settings לא נמצא");
  }

  const destGroupId =
    rule.destinationGroupId != null && rule.destinationGroupId !== ""
      ? rule.destinationGroupId
      : "SKIP";

  const rowData = [
    rule.sourceCatId != null ? rule.sourceCatId : "",
    rule.searchProperty || "",
    rule.searchValue || "",
    destGroupId,
    rule.sourceCategory || "",
    rule.destCategoryName || "",
    rule.notes || "",
  ];

  sheet.appendRow(rowData);
  getProtectedActive().toast("חוק סיווג חדש נוסף ל-group-settings");

  return { success: true };
}
