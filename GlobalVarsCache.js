/*************************************************************
 * Global Variable Loading & Caching
 * 
 * This module is responsible for ensuring global variables
 * (like URLs, emails, and names) are reliably available in 
 * memory. It uses three layers of redundancy:
 * 
 *  1. CacheService - short-term (fast access)
 *  2. PropertiesService - persistent storage
 *  3. Variables Sheet - ultimate source of truth
 * 
 * Functions in this file should be called before running any
 * automation logic that depends on global variables.
 *************************************************************/

/**
 * ensureGlobalVariables()
 * 
 * Verifies that required global variables are available.
 * If missing, it refreshes the cache and throws an error if
 * required variables are still not found.
 * 
 * @param {string[]} requiredKeys (optional)
 *   - List of variable keys that must be present.
 *   - If empty, it just ensures something is loaded.
 * 
 * @returns {Object} Loaded variables (key/value pairs)
 */
function ensureGlobalVariables(requiredKeys = []) {
  Logger.log("🔎 Ensuring global variables are loaded...");

  // First, try loading from existing cache/properties/sheet
  let vars = getGlobalVariables(false);

  // If specific keys are required, verify they all exist
  if (requiredKeys.length > 0) {
    const missingKeys = requiredKeys.filter(k => !vars.hasOwnProperty(k) || !vars[k]);

    // If any keys are missing, force a refresh from the sheet
    if (missingKeys.length > 0) {
      Logger.log(`⚠️ Missing required variables: ${missingKeys.join(", ")}. Refreshing...`);
      vars = refreshGlobalVariablesCache();

      // Double-check after refresh
      const stillMissing = requiredKeys.filter(k => !vars.hasOwnProperty(k) || !vars[k]);
      if (stillMissing.length > 0) {
        throw new Error(`❌ Critical missing variables: ${stillMissing.join(", ")}. Check Variables sheet.`);
      }
    }
  }

  Logger.log("✅ Global variables confirmed loaded.");
  return vars;
}

/**
 * loadAutomationInfoSheet()
 * 
 * Opens the Automation Info Sheet using its URL and loads all
 * key/value pairs from the "Variables" tab.
 * 
 * @returns {Object} Loaded variables (key/value pairs)
 */
function loadAutomationInfoSheet() {
  try {
    ensureAutomationInfoSheetURL();
    Logger.log("Opening Automation Info Sheet URL: " + AUTOMATION_INFO_SHEET_URL);

    automationInfoSheet = SpreadsheetApp.openByUrl(AUTOMATION_INFO_SHEET_URL);
  } catch (e) {
    throw new Error("Failed to open Automation Info Sheet: " + e.message);
  }

  // Load Variables tab
  const sheet = automationInfoSheet.getSheetByName("Variables");
  if (!sheet) throw new Error("❌ 'Variables' sheet not found in Automation Info Sheet.");

  // Read all rows
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
 * getGlobalVariables()
 * 
 * Attempts to load global variables from cache or persistent 
 * storage. If none found, falls back to refreshing from sheet.
 * 
 * @param {boolean} showPopup 
 *   - If true, shows a popup with loaded variables.
 * 
 * @returns {Object} Loaded variables
 */
function getGlobalVariables(showPopup = false) {
  ensureAutomationInfoSheetURL();
  const cache = CacheService.getScriptCache();
  let cachedVars = cache.get("globalVariables");

  // 1) Try cache first
  if (cachedVars) {
    const vars = JSON.parse(cachedVars);
    Logger.log("✅ Loaded global variables from CACHE: " + JSON.stringify(vars));
    if (showPopup) showVariablesPopup(vars);
    return vars;
  } else {
    Logger.log("⚠️ Cache empty.");
  }

  // 2) Try PropertiesService next
  let propsVars = loadVarsFromProperties();
  if (propsVars && Object.keys(propsVars).length > 0) {
    cache.put("globalVariables", JSON.stringify(propsVars), 18000); // cache for 5 hrs
    Logger.log("✅ Loaded global variables from PROPERTIES: " + JSON.stringify(propsVars));
    if (showPopup) showVariablesPopup(propsVars);
    return propsVars;
  } else {
    Logger.log("⚠️ Properties empty.");
  }

  // 3) Fallback: refresh from sheet
  Logger.log("🔄 Falling back to SHEET (refreshGlobalVariablesCache)");
  const freshVars = refreshGlobalVariablesCache();
  if (showPopup) showVariablesPopup(freshVars);
  return freshVars;
}

/**
 * refreshGlobalVariablesCache()
 * 
 * Forces a reload of variables from the sheet and updates 
 * both CacheService and PropertiesService for redundancy.
 * 
 * @returns {Object} Refreshed variables
 */
function refreshGlobalVariablesCache() {
  try {
    Logger.log("🔎 Attempting to load variables from SHEET...");
    const vars = loadAutomationInfoSheet(); // Loads from Variables tab

    if (!vars || Object.keys(vars).length === 0) {
      throw new Error("No variables loaded from 'Variables' sheet! Check tab name & key names.");
    }

    // Optionally update script-level variables for immediate use
    CASE_TRACKER_URL = vars["caseTrackerUrl"] || "";
    SSM_NAME = vars["ssmName"] || "";
    // ... repeat for other globals as needed

    // Save redundantly
    CacheService.getScriptCache().put("globalVariables", JSON.stringify(vars), 18000); // cache 5 hrs
    PropertiesService.getScriptProperties().setProperty("globalVariablesJSON", JSON.stringify(vars));

    Logger.log("✅ Refreshed variables from SHEET: " + JSON.stringify(vars));
    return vars;

  } catch (e) {
    Logger.log("❌ Failed to refresh variables from SHEET: " + e.message);
    return {};
  }
}

/**
 * saveVarsToProperties()
 * 
 * Saves global variables to PropertiesService for persistent storage.
 */
function saveVarsToProperties(vars) {
  try {
    PropertiesService.getScriptProperties().setProperty("globalVariablesJSON", JSON.stringify(vars));
  } catch (e) {
    Logger.log("❌ Failed to save vars to PropertiesService: " + e.message);
  }
}

/**
 * loadVarsFromProperties()
 * 
 * Loads global variables from PropertiesService (if any).
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
 * scheduleCacheRefresh()
 * 
 * Installs a time-based trigger to refresh global variables cache 
 * automatically every 5 hours.
 */
function scheduleCacheRefresh() {
  ScriptApp.newTrigger("refreshGlobalVariablesCache")
    .timeBased()
    .everyHours(5)
    .create();
}
