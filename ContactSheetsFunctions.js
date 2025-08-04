/*************************************************************
 * Contact Sheets Functions
 * Functions related to reading/writing data in contact sheets,
 * including workers, children, dropdowns, and reminders.
 *************************************************************/

/**
 * Retrieves an array of child objects from the "Hearing Tracker" tab.
 * Each object contains child's name and case number.
 * @returns {Array} Array of objects {name, caseNumber}
 */
function getChildrenArray() {
  ensureAutomationInfoSheetURL();
  const caseTrackerSS = SpreadsheetApp.openByUrl(CASE_TRACKER_URL);
  const hearingTracker = caseTrackerSS.getSheetByName("Hearing Tracker");
  if (!hearingTracker) throw new Error("Hearing Tracker tab not found in Case Tracker sheet.");

  const lastRow = hearingTracker.getLastRow();
  const values = hearingTracker.getRange(2, 2, lastRow - 1, 2).getValues(); // Skip header row

  const children = [];

  values.forEach(row => {
    const caseNumber = row[0];
    const childNames = row[1];

    if (caseNumber && childNames) {
      const names = childNames.split(",").map(n => n.trim()).filter(Boolean);
      names.forEach(name => {
        children.push({
          name: name,
          caseNumber: caseNumber
        });
      });
    }
  });

  Logger.log(`✅ Parsed ${children.length} children.`);
  return children;
}

/**
 * Outputs children names and case numbers to all contact sheets.
 * Writes child name in column 1 and case number in column 2.
 * Stops if no children found.
 */
function outputChildrenToSheets() {
  const children = getChildrenArray();
  if (children.length === 0) throw new Error("No children found to output.");

  const contactSheets = getContactSheets();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let totalWritten = 0;

  contactSheets.forEach(({ name }) => {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return;

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return;

    const writeRows = Math.min(children.length, lastRow - 1);
    const nameValues = [];
    const caseValues = [];

    for (let i = 0; i < writeRows; i++) {
      nameValues.push([children[i].name]);
      caseValues.push([children[i].caseNumber]);
    }

    sheet.getRange(2, 1, writeRows, 1).setValues(nameValues);
    sheet.getRange(2, 2, writeRows, 1).setValues(caseValues);

    Logger.log(`✅ Wrote ${writeRows} children + case numbers to ${sheet.getName()}`);
    totalWritten += writeRows;
  });

  SpreadsheetApp.getUi().alert(`✅ Output ${totalWritten} children across all sheets.`);
}

/**
 * Updates the "Seen By" dropdowns in all contact sheets.
 * Pulls from CPSEmployeeInfo and Additional Workers Info tabs.
 * @param {boolean} showUI Whether to show a popup when complete.
 */
function updateSeenByDropdowns(showUI = false) {
  ensureAutomationInfoSheetURL();
  const automationSS = SpreadsheetApp.openByUrl(AUTOMATION_INFO_SHEET_URL);

  const employeeSheet = automationSS.getSheetByName("CPSEmployeeInfo");
  const employees = employeeSheet
    .getRange(2, 1, employeeSheet.getLastRow() - 1, 1)
    .getValues()
    .flat()
    .filter(name => !!name && name.trim() !== "");

  const additionalSheet = automationSS.getSheetByName("Additional Workers Info");
  const additionalWorkers = additionalSheet
    ? additionalSheet.getRange(2, 1, additionalSheet.getLastRow() - 1, 1)
      .getValues()
      .flat()
      .filter(name => !!name && name.trim() !== "")
    : [];

  const allWorkers = [...new Set([...employees, ...additionalWorkers])];

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(allWorkers, true)
    .setAllowInvalid(true)
    .build();

  const contactSheets = getContactSheets();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  contactSheets.forEach(cs => {
    const sheet = ss.getSheetByName(cs.name);
    if (!sheet) return;

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    for (let i = 2; i <= lastRow; i++) {
      const cellValue = sheet.getRange(i, 1).getValue();
      if (cellValue) {
        sheet.getRange(i, cs.seenByCol).setDataValidation(rule);
      }
    }
  });

  if (showUI) {
    SpreadsheetApp.getUi().alert(`✅ Dropdowns updated for ${allWorkers.length} workers.`);
  }
}

/**
 * Retrieves worker info by name from CPSEmployeeInfo and Additional Workers Info tabs.
 * @param {string} name Worker name to search for.
 * @returns {Object|null} Object with worker info or null if not found.
 */
function getWorkerInfoByName(name) {
  if (!name) return null;

  const automationInfoSS = SpreadsheetApp.openByUrl(AUTOMATION_INFO_SHEET_URL);

  const tabsToSearch = ['CPSEmployeeInfo', 'Additional Workers Info'];

  for (let tabName of tabsToSearch) {
    const sheet = automationInfoSS.getSheetByName(tabName);
    if (!sheet) continue;

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const nameIdx = headers.indexOf('workerName');
    const emailIdx = headers.indexOf('workerEmail');
    const supNameIdx = headers.indexOf('supervisorName');
    const supEmailIdx = headers.indexOf('supervisorEmail');
    const countyIdx = headers.indexOf('workerCounty');

    for (let i = 1; i < data.length; i++) {
      if (data[i][nameIdx] && data[i][nameIdx].toString().trim() === name.trim()) {
        return {
          workerName: data[i][nameIdx],
          workerEmail: data[i][emailIdx],
          supervisorName: data[i][supNameIdx],
          supervisorEmail: data[i][supEmailIdx],
          workerCounty: countyIdx >= 0 ? data[i][countyIdx] : null
        };
      }
    }
  }

  return null;
}

/**
 * Retrieves contact sheets info, using cache, properties, or scanning the spreadsheet.
 * @param {boolean} forceRefresh - If true, forces a fresh scan.
 * @returns {Array} Array of contact sheet objects {name, seenByCol}
 */
function getContactSheets(forceRefresh = false) {
  const cache = CacheService.getScriptCache();

  if (!forceRefresh) {
    const cached = cache.get("contactSheets");
    if (cached) {
      Logger.log("✅ Loaded contactSheets from CACHE.");
      return JSON.parse(cached);
    }
  }

  if (!forceRefresh) {
    const props = PropertiesService.getScriptProperties().getProperty("contactSheets");
    if (props) {
      Logger.log("✅ Loaded contactSheets from PROPERTIES.");
      const parsed = JSON.parse(props);
      cache.put("contactSheets", props, 18000);
      return parsed;
    }
  }

  const backup = getContactSheetsBackupRow();
  if (!forceRefresh && backup) {
    Logger.log("✅ Loaded contactSheets from BACKUP row.");
    cache.put("contactSheets", JSON.stringify(backup), 18000);
    return backup;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allSheets = ss.getSheets();
  const contactSheets = [];

  allSheets.forEach(sheet => {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const seenByColIndex = headers.findIndex(h => h && h.toString().trim().toLowerCase() === 'seen by');
    if (seenByColIndex !== -1) {
      contactSheets.push({
        name: sheet.getName(),
        seenByCol: seenByColIndex + 1
      });
    }
  });

  Logger.log("🔄 Freshly discovered contactSheets.");

  cache.put("contactSheets", JSON.stringify(contactSheets), 18000);
  PropertiesService.getScriptProperties().setProperty("contactSheets", JSON.stringify(contactSheets));
  updateContactSheetsBackupRow(contactSheets);

  return contactSheets;
}

/**
 * Retrieves the backup row of contact sheets stored in the Variables tab.
 * @returns {Array|null} Parsed contact sheets array or null.
 */
function getContactSheetsBackupRow() {
  try {
    ensureAutomationInfoSheetURL();
    const sheet = automationInfoSheet.getSheetByName("Variables");
    if (!sheet) {
      Logger.log("❌ Variables sheet missing for backup retrieval.");
      return null;
    }

    const values = sheet.getDataRange().getValues();
    for (let i = 1; i < values.length; i++) {
      if (values[i][0] === "contactSheets") {
        return JSON.parse(values[i][1]);
      }
    }
    return null;
  } catch (e) {
    Logger.log("❌ Error retrieving contactSheets backup: " + e.message);
    return null;
  }
}

/**
 * Updates the backup row in the Variables tab with contact sheets info.
 * @param {Array} contactSheets Array of contact sheet objects.
 */
function updateContactSheetsBackupRow(contactSheets) {
  try {
    ensureAutomationInfoSheetURL();
    const sheet = automationInfoSheet.getSheetByName("Variables");
    if (!sheet) return;

    const values = sheet.getDataRange().getValues();
    for (let i = 1; i < values.length; i++) {
      if (values[i][0] === "contactSheets") {
        sheet.getRange(i + 1, 2).setValue(JSON.stringify(contactSheets));
        return;
      }
    }
    // If backup key not found, append it
    sheet.appendRow(["contactSheets", JSON.stringify(contactSheets)]);
  } catch (e) {
    Logger.log("❌ Error updating contactSheets backup: " + e.message);
  }
}
