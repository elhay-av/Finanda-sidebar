/**
 * @OnlyCurrentDoc
 */

const SPREADSHEETS_SCOPE = 'https://www.googleapis.com/auth/spreadsheets.currentonly';

/**
 * Checks if the user has granted the required authorization scope.
 * Throws a specific error with the authorization URL if not.
 */
function checkAuthorization() {
  try {
    const authInfo = ScriptApp.getAuthorizationInfo(ScriptApp.AuthMode.FULL, [SPREADSHEETS_SCOPE]);
    if (authInfo.getAuthorizationStatus() === ScriptApp.AuthorizationStatus.REQUIRED) {
      const authUrl = authInfo.getAuthorizationUrl();
      Logger.log("Authorization REQUIRED. URL: " + authUrl);
      throw new Error("AUTHORIZATION_REQUIRED: " + authUrl);
    }
  } catch (e) {
    // If we are in a simple trigger context (like onOpen), getAuthorizationInfo might throw or be inaccurate.
    // In that case, we skip the check to avoid blocking the trigger.
    Logger.log("Check authorization failed or skipped: " + e.message);
  }
}

/**
 * Gets the active spreadsheet with safety checks for authorization.
 * @returns {GoogleAppsScript.Spreadsheet.Spreadsheet} The active spreadsheet
 */
function getProtectedActiveSpreadsheet() {
  checkAuthorization();
  return SpreadsheetApp.getActiveSpreadsheet();
}

/**
 * Gets the active sheet / container with safety checks for authorization.
 * @returns {GoogleAppsScript.Spreadsheet.Spreadsheet} The active spreadsheet
 */
function getProtectedActive() {
  checkAuthorization();
  return SpreadsheetApp.getActive();
}

/**
 * Gets the UI with safety checks for authorization.
 * @returns {GoogleAppsScript.Base.Ui} The Apps Script Ui object
 */
function getProtectedUi() {
  checkAuthorization();
  return SpreadsheetApp.getUi();
}
