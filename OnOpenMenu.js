/*************************************************************
 * onOpen()
 * Builds "Automation Settings" menu
 *************************************************************/
function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ensureAutomationInfoSheetURL();

  if (AUTOMATION_INFO_SHEET_URL) {
    loadAutomationInfoSheet();
    getGlobalVariables(); 
  }

  ui.createMenu("Automation Settings")
    .addItem("Set Automation Info Sheet URL", "updateAutomationInfoSheetURL")
    .addItem("Copy Automation Info Sheet URL", "copyAutomationInfoSheetURLToClipboard")
    .addItem("Backup Automaiton Info Sheet URL", "updateAutomationInfoSheetBackupRow")
    .addItem("Select Completed Months", "showCompletedMonthsDialog")
    .addSeparator()
    .addItem("Refresh Global Variables", "refreshGlobalVariablesWithPopup")
    .addSeparator()
    .addItem("Refresh All Employee Data", "refreshEmployeeData")
    .addItem("Refresh Trigger", "ensureOnURLEditTrigger")
    .addItem("Show Cached Variables", "showVariablesPopup")
    .addItem("Update Seen By Dropdowns", "updateSeenByDropdownsManual")
    .addToUi();

  ensureOnURLEditTrigger();
  ensureHandleSeenByEditTrigger();
}
