/*************************************************************
 * Triggers
 * Functions to install and manage triggers
 *************************************************************/

/**
 * Installs all necessary triggers for the app.
 * Includes time-based and edit triggers.
 */
function installTriggers() {
  Logger.log("Installing triggers...");

  // Time-based trigger to refresh global variables every 5 hours
  ScriptApp.newTrigger("refreshGlobalVariablesCache")
    .timeBased()
    .everyHours(5)
    .create();

  // Edit trigger to handle seen by edits
  ScriptApp.newTrigger("handleSeenByEdit")
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onEdit()
    .create();

  Logger.log("✅ All triggers installed.");
}

/**
 * Function called by onEdit trigger.
 * Delegates to handleSeenByEdit in ContactSheetsFunctions.
 * @param {Object} e Edit event object.
 */
function handleSeenByEdit(e) {
  if (!e) return;

  // Just call the function from ContactSheetsFunctions, which now uses ensureGlobalVariables
  ContactSheetsFunctions_handleSeenByEdit(e);
}
