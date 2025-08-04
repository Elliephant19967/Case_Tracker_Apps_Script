/*************************************************************
 * Utilities & Global Variable Handling
 *************************************************************/

// List of required global keys
const REQUIRED_KEYS = [
  "MAIN_WORKER_EMAIL",
  "MAIN_WORKER_NAME",
  "MAIN_SUPERVISOR_EMAIL",
  "MAIN_SUPERVISOR_NAME",
  "SSM_NAME",
  "SSM_EMAIL"
];

/**
 * Loads the Automation Info Sheet and retrieves global variables from the "Variables" tab.
 * @returns {Object} Key-value pairs of variables.
 */
function loadAutomationInfoSheet() {
  try {
    ensureAutomationInfoSheetURL();
    Logger.log("Opening Automation Info Sheet URL: " + AUTOMATION_INFO_SHEET_URL);

    automationInfoSheet = SpreadsheetApp.openByUrl(AUTOMATION_INFO_SHEET_URL);
  } catch (e) {
    throw new Error("Failed to open Automation Info Sheet: " + e.message);
  }

  const sheet = automationInfoSheet.getSheetByName("Variables");
  if (!sheet) throw new Error("❌ 'Variables' sheet not found in Automation Info Sheet.");

  const values = sheet.getDataRange().getValues();
  const vars = {};

  Logger.log("📥 Loading Global Variables from 'Variables' sheet...");
  for (let i = 1; i < values.length; i++) {
    const key = values[i][0]?.toString().trim();
    const value = values[i][1]?.toString().trim();
    if (!key || !value) {
      Logger.log(`⚠️ Skipping blank key or value at row ${i + 1}`);
      continue;
    }
    vars[key] = value;
    Logger.log(`✅ Loaded: ${key} = ${value}`);
  }

  if (Object.keys(vars).length === 0) {
    Logger.log("⚠️ No valid variables loaded from sheet.");
    return {};
  }

  return vars;
}

/**
 * Ensures the specified global variables are loaded and cached.
 * Tries Cache → Properties → Sheet fallback.
 * Optionally shows a popup to display the loaded variables.
 * @param {string[]} keys - Array of required variable keys.
 * @param {boolean} showPopup - Whether to display the variables in a popup dialog.
 * @returns {Object} Object containing the required global variables.
 */
function ensureGlobalVariables(keys = [], showPopup = false) {
  ensureAutomationInfoSheetURL();
  const cache = CacheService.getScriptCache();

  // Try cache
  let cachedVars = cache.get("globalVariables");
  if (cachedVars) {
    const vars = JSON.parse(cachedVars);
    Logger.log("✅ Loaded global variables from CACHE.");
    const filtered = filterVars(vars, keys);
    if (showPopup) showVariablesPopup(filtered);
    return filtered;
  } else {
    Logger.log("⚠️ Cache empty.");
  }

  // Try PropertiesService
  let propsVars = loadVarsFromProperties();
  if (propsVars && Object.keys(propsVars).length > 0) {
    cache.put("globalVariables", JSON.stringify(propsVars), 18000); // 5 hrs
    Logger.log("✅ Loaded global variables from PROPERTIES.");
    const filtered = filterVars(propsVars, keys);
    if (showPopup) showVariablesPopup(filtered);
    return filtered;
  } else {
    Logger.log("⚠️ Properties empty.");
  }

  // Fallback to sheet refresh
  Logger.log("🔄 Falling back to SHEET (refreshGlobalVariablesCache)");
  const freshVars = refreshGlobalVariablesCache();
  const filtered = filterVars(freshVars, keys);
  if (showPopup) showVariablesPopup(filtered);
  return filtered;
}

/**
 * Filters the variables object to include only specified keys.
 * @param {Object} vars - The full variables object.
 * @param {string[]} keys - The keys to include.
 * @returns {Object} Filtered variables object.
 */
function filterVars(vars, keys) {
  if (!keys || keys.length === 0) return vars;
  const filtered = {};
  keys.forEach(key => {
    if (vars.hasOwnProperty(key)) filtered[key] = vars[key];
  });
  return filtered;
}

/**
 * @deprecated Replaced by ensureGlobalVariables().
 * This function is kept only for backward compatibility and will be removed later.
 * It now calls ensureGlobalVariables(REQUIRED_KEYS) directly and logs a warning.
 */
function getGlobalVariables(showPopup = false) {
  Logger.log("⚠️ getGlobalVariables() is deprecated. Use ensureGlobalVariables(REQUIRED_KEYS) instead.");
  return ensureGlobalVariables(REQUIRED_KEYS, showPopup);
}

/**
 * Refreshes the global variables cache from the sheet.
 * Updates temporary globals and caches the variables.
 * @returns {Object} The refreshed variables.
 */
function refreshGlobalVariablesCache() {
  try {
    Logger.log("🔎 Attempting to load variables from SHEET...");
    const vars = loadAutomationInfoSheet(); // Loads from Variables tab

    if (!vars || Object.keys(vars).length === 0) {
      throw new Error("No variables loaded from 'Variables' sheet! Check tab name & key names.");
    }

    // Update script-level temporary globals if needed
    CASE_TRACKER_URL = vars["caseTrackerUrl"] || "";
    SSM_NAME = vars["ssmName"] || "";
    // Add other globals as needed

    // Save permanently
    CacheService.getScriptCache().put("globalVariables", JSON.stringify(vars), 18000);
    PropertiesService.getScriptProperties().setProperty("globalVariablesJSON", JSON.stringify(vars));

    Logger.log("✅ Refreshed variables from SHEET.");
    return vars;

  } catch (e) {
    Logger.log("❌ Failed to refresh variables from SHEET: " + e.message);
    return {};
  }
}

/**
 * Loads variables JSON from PropertiesService.
 * @returns {Object|null} Parsed variables object or null on failure.
 */
function loadVarsFromProperties() {
  try {
    const jsonStr = PropertiesService.getScriptProperties().getProperty("globalVariablesJSON");
    return jsonStr ? JSON.parse(jsonStr) : null;
  } catch (e) {
    Logger.log("❌ Failed to load vars from PropertiesService: " + e.message);
    return null;
  }
}

/**
 * Saves variables JSON to PropertiesService.
 * @param {Object} vars - Variables to save.
 */
function saveVarsToProperties(vars) {
  try {
    PropertiesService.getScriptProperties().setProperty("globalVariablesJSON", JSON.stringify(vars));
  } catch (e) {
    Logger.log("❌ Failed to save vars to PropertiesService: " + e.message);
  }
}

/**
 * Shows a popup with global variables info.
 * @param {Object} vars - Variables to display.
 */
function showVariablesPopup(vars) {
  if (!vars || Object.keys(vars).length === 0) {
    SpreadsheetApp.getUi().alert("⚠️ No variables loaded!\n\nDebugging:\n- Cache empty?\n- Properties empty?\n- Sheet not reachable?");
    return;
  }

  let message = "Global Variables:\n\n";
  for (let key in vars) {
    message += `${key}: ${vars[key]}\n`;
  }
  SpreadsheetApp.getUi().alert(message);
}

/**
 * Schedules a time-based trigger to refresh global variables cache every 5 hours.
 */
function scheduleCacheRefresh() {
  ScriptApp.newTrigger("refreshGlobalVariablesCache")
    .timeBased()
    .everyHours(5)
    .create();
}
