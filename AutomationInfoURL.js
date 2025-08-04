/*************************************************************
 * Automation Info Sheet URL Management
 *************************************************************/
function ensureAutomationInfoSheetURL() {
  if (AUTOMATION_INFO_SHEET_URL && AUTOMATION_INFO_SHEET_URL.trim() !== "") {
    return AUTOMATION_INFO_SHEET_URL;
  }

  const cachedUrl = getCachedBackupAutomationInfoSheetURL();
  if (cachedUrl) {
    AUTOMATION_INFO_SHEET_URL = cachedUrl;
    updateAutomationInfoSheetBackupRow(AUTOMATION_INFO_SHEET_URL);
    getGlobalVariables();
    return AUTOMATION_INFO_SHEET_URL;
  }

  const propUrl = PropertiesService.getUserProperties().getProperty("AUTOMATION_INFO_SHEET_URL");
  if (propUrl && propUrl.trim() !== "") {
    AUTOMATION_INFO_SHEET_URL = propUrl;
    updateAutomationInfoSheetBackupRow(AUTOMATION_INFO_SHEET_URL);
    getGlobalVariables();
    return AUTOMATION_INFO_SHEET_URL;
  }

  return getAutomationInfoURL();
}

function getCachedBackupAutomationInfoSheetURL() {
  try {
    if (!automationInfoSheet) {
      if (AUTOMATION_INFO_SHEET_URL) {
        automationInfoSheet = SpreadsheetApp.openByUrl(AUTOMATION_INFO_SHEET_URL);
      } else {
        return null;
      }
    }

    const sheet = automationInfoSheet.getSheetByName("Variables");
    if (!sheet) {
      Logger.log("❌ 'Variables' sheet not found in automationInfoSheet.");
      return null;
    }

    const values = sheet.getDataRange().getValues();
    for (let i = 1; i < values.length; i++) {
      if (values[i][0] === "automationInfoSheetURL") {
        return values[i][1];
      }
    }
  } catch (e) {
    Logger.log("Error retrieving cached backup URL: " + e.message);
  }
  return null;
}
function getAutomationInfoURL() {
  const ui = SpreadsheetApp.getUi();
  const userProps = PropertiesService.getUserProperties();

  const response = ui.prompt("Enter the URL of your Automation Info Sheet:");
  if (response.getSelectedButton() !== ui.Button.OK) {
    ui.alert("You must provide the Automation Info Sheet URL to continue.");
    throw new Error("Automation Info Sheet URL is required.");
  }

  const newUrl = response.getResponseText().trim();
  if (!newUrl) {
    ui.alert("Invalid URL provided. Please try again.");
    throw new Error("Invalid Automation Info Sheet URL.");
  }

  // Save and open
  userProps.setProperty("AUTOMATION_INFO_SHEET_URL", newUrl);
  AUTOMATION_INFO_SHEET_URL = newUrl;
  automationInfoSheet = SpreadsheetApp.openByUrl(newUrl);

  updateAutomationInfoSheetBackupRow(newUrl);
  getGlobalVariables();

  ui.alert("Automation Info Sheet URL saved and variables refreshed.");
  return newUrl;
}
function updateAutomationInfoSheetBackupRow(url) {
  try {
    if (!automationInfoSheet) {
      automationInfoSheet = SpreadsheetApp.openByUrl(url);
    }

    const sheet = automationInfoSheet.getSheetByName("Variables");
    if (!sheet) {
      Logger.log("❌ 'Variables' sheet not found in automationInfoSheet.");
      return;
    }

    const values = sheet.getDataRange().getValues();
    let foundRow = -1;

    // Search for existing row
    for (let i = 1; i < values.length; i++) {
      if (values[i][0] === "automationInfoSheetURL") {
        foundRow = i + 1;
        break;
      }
    }

    if (foundRow === -1) {
      sheet.appendRow([
        "automationInfoSheetURL",
        url,
        "This is autofilled, do not type in this box"
      ]);
      Logger.log("Backup row for automationInfoSheetURL appended.");
    } else {
      sheet.getRange(foundRow, 2).setValue(url);
      sheet.getRange(foundRow, 3).setValue("This is autofilled, do not type in this box");
      Logger.log(`Backup row for automationInfoSheetURL updated at row ${foundRow}.`);
    }
  } catch (e) {
    Logger.log("Error updating backup row: " + e.message);
  }
}

function updateAutomationInfoSheetURL() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.prompt("Enter the new Automation Info Sheet URL:");

  if (response.getSelectedButton() === ui.Button.OK) {
    const newUrl = response.getResponseText().trim();
    if (!newUrl) {
      ui.alert("Invalid URL provided. Update cancelled.");
      return;
    }

    AUTOMATION_INFO_SHEET_URL = newUrl;
    PropertiesService.getUserProperties().setProperty("AUTOMATION_INFO_SHEET_URL", newUrl);
    automationInfoSheet = SpreadsheetApp.openByUrl(newUrl);

    updateAutomationInfoSheetBackupRow(newUrl);
    getGlobalVariables();

    ui.alert("Automation Info Sheet URL updated and variables refreshed.");
  }
}
function copyAutomationInfoSheetURLToClipboard() {
  const ui = SpreadsheetApp.getUi();
  if (!AUTOMATION_INFO_SHEET_URL) {
    ui.alert("Automation Info Sheet URL is not set.");
    return;
  }

  const html = `
    <html>
      <body>
        <input type="text" value="${AUTOMATION_INFO_SHEET_URL}" id="urlInput" readonly style="width: 100%;">
        <button onclick="copyUrl()">Copy URL to Clipboard</button>
        <script>
          function copyUrl() {
            const urlInput = document.getElementById('urlInput');
            urlInput.select();
            document.execCommand('copy');
            alert('URL copied to clipboard!');
            google.script.run.withSuccessHandler(() => google.script.host.close()).postCopyUpdate();
          }
        </script>
      </body>
    </html>
  `;
  ui.showModalDialog(HtmlService.createHtmlOutput(html).setWidth(400).setHeight(150), "Copy Automation Info Sheet URL");
}
function postCopyUpdate() {
  updateAutomationInfoSheetBackupRow(AUTOMATION_INFO_SHEET_URL);
  getGlobalVariables();
}