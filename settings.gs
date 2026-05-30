const SETTINGS_SHEET_NAME = 'settings'

const MINUTE = 1000*60;
const MAX_SESSION_DURATION = MINUTE * 5;

let settings = null;

function getSettings() {
  if (settings) {
    return settings;
  }

  const settingsSheet = getProtectedActiveSpreadsheet().getSheetByName(SETTINGS_SHEET_NAME);
  settings = settingsSheet.getNamedRanges().reduce((acc, range) => {
    acc[range.getName()] = range.getRange().getValues();
    return acc;
  }, {});
  return settings;
}

function setSettingsRangeValue(rangeName, value) {
  const settingsSheet = getProtectedActiveSpreadsheet().getSheetByName(SETTINGS_SHEET_NAME);
  settingsSheet.getNamedRanges().forEach((range) => {
    if (range.getName() === rangeName) {
      range.getRange().setValue(value);
    }
  })
}

function getFinandaPassword () {
  return getSettings()['FinandaPassword'][0][0];
}

function getFinandaUser () {
  return getSettings()['FinandaUser'][0][0];
}

function getFinandaSession () {
  const session = getSettings()['FinandaSession'][0][0];
  const sessionTime = getSettings()['FinandaSessionTime'][0][0];

  if (!session) {
    return '';
  }

  if (new Date().getTime() - parseInt(sessionTime) < MAX_SESSION_DURATION) {
    return session;
  }

  return '';
}

function setFinandaSession (session) {
  setSettingsRangeValue('FinandaSession', session);
  setSettingsRangeValue('FinandaSessionTime', new Date().getTime());
}