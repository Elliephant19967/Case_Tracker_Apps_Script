/*************************************************************
 * sendMonthlyContactReminders()
 * 
 * Main entry point:
 * - Processes ALL "MonthName Contacts" sheets for months ≤ current month.
 * - Skips sheets marked as completed (CONTACT_COMPLETE_MONTHS).
 * - Sends reminders for missing contact entries.
 * - Prior-month reminders only go out on Mondays.
 * - If no reminders are sent for a sheet (and it's Monday), that month is
 *   marked complete and added to CONTACT_COMPLETE_MONTHS in global variables.
 *************************************************************/
function sendMonthlyContactReminders() {
  try {
    // Load global variables (from cache or Variables tab)
    getGlobalVariables(false);
  } catch (error) {
    Logger.log(`❌ Failed to load global variables: ${error.message}`);
    return;
  }

  const today = new Date();
  const currentMonth = today.getMonth(); // 0-based month index
  const isMonday = today.getDay() === 1; // 1 = Monday

  // Parse CONTACT_COMPLETE_MONTHS (comma-separated string) into an array
  const completeMonths = (CONTACT_COMPLETE_MONTHS || "")
    .split(",")
    .map(m => m.trim())
    .filter(Boolean); // remove empty strings

  Logger.log(`Running Contact Reminders on ${today}`);
  Logger.log(`Complete Months currently marked: ${completeMonths}`);

  // Open Case Tracker and iterate over all sheets
  const ss = SpreadsheetApp.openByUrl(CASE_TRACKER_URL);
  const allSheets = ss.getSheets();

  for (let sheet of allSheets) {
    const sheetName = sheet.getName();

    // Only process sheets with names like "January Contacts"
    const monthMatch = sheetName.match(/^([A-Za-z]+) Contacts$/);
    if (!monthMatch) continue;

    const monthName = monthMatch[1];
    const sheetMonth = getMonthNumberFromName(monthName);

    // Skip if this sheet is for a future month
    if (sheetMonth > currentMonth) continue;

    // Skip if this month is already marked as complete
    if (completeMonths.includes(monthName)) {
      Logger.log(`Skipping ${sheetName}: already marked complete.`);
      continue;
    }

    // Process the sheet and track how many reminders were sent
    const remindersSent = processContactSheet(sheet, today, currentMonth, isMonday);

    // If it's Monday AND no reminders were sent, mark the month as complete
    if (isMonday && remindersSent === 0) {
      Logger.log(`No reminders sent for ${sheetName}. Marking as complete.`);
      completeMonths.push(monthName);

      // Update CONTACT_COMPLETE_MONTHS in global variables
      updateGlobalVariable("CONTACT_COMPLETE_MONTHS", completeMonths.join(", "));
    }
  }
}

/*************************************************************
 * processContactSheet(sheet, today, currentMonth, isMonday)
 * 
 * Processes a single contact sheet row by row.
 * - Skips contacts already entered or missing key info.
 * - Applies Monday-only rule for prior months.
 * - Calls sendContactReminderRow() to handle actual reminder sending.
 * 
 * Returns: Number of reminders sent for this sheet.
 *************************************************************/
function processContactSheet(sheet, today, currentMonth, isMonday) {
  const data = sheet.getDataRange().getValues();
  let remindersSent = 0;

  // Skip the header row (row 1)
  for (let i = 1; i < data.length; i++) {
    const [
      childName,          // A: Child Name
      caseID,             // B: Case ID
      dateSeenRaw,        // C: Date Seen
      seenBy,             // D: Seen By (dropdown name)
      dateContactEntered, // E: Date Contact Entered
      lastReminderSent,   // F: Last Reminder Sent
      missed,             // G: Missed (optional)
      reasonMissed        // H: Reason Missed (optional)
    ] = data[i];

    // Skip rows missing required data
    if (!childName || !dateSeenRaw || !seenBy) continue;

    // Skip rows where contact is already entered
    if (dateContactEntered && dateContactEntered !== "") continue;

    const dateSeen = new Date(dateSeenRaw);
    if (isNaN(dateSeen)) continue;

    // Determine if this is from a prior month
    const isPriorMonth = dateSeen.getMonth() < today.getMonth() || 
                         dateSeen.getFullYear() < today.getFullYear();

    // Apply Monday-only rule for prior months
    if (isPriorMonth && !isMonday) {
      Logger.log(`Skipping ${childName}: prior month and today is not Monday.`);
      continue;
    }

    // Send reminder and increment count if successful
    const sent = sendContactReminderRow(data[i], today, sheet, i + 1);
    if (sent) remindersSent++;
  }

  return remindersSent;
}

/*************************************************************
 * sendContactReminderRow(rowData, today, sheet, rowIndex)
 * 
 * Handles all reminder sending logic for a single row:
 * - Looks up worker info from Automation Info (CPSEmployeeInfo).
 * - Chooses correct reminder template based on timing (standard, reprimanding, post-month).
 * - Sends email and updates "Last Reminder Sent" column.
 * 
 * Returns: true if a reminder was sent, false otherwise.
 *************************************************************/
function sendContactReminderRow(rowData, today, sheet, rowIndex) {
  const [
    childName, caseID, dateSeenRaw, seenBy,
    dateContactEntered, lastReminderSent, missed, reasonMissed
  ] = rowData;

  const dateSeen = new Date(dateSeenRaw);
  const daysLeftInMonth = daysRemainingInMonth(today);
  const daysSinceSeen = Math.floor((today - dateSeen) / (1000 * 60 * 60 * 24));

  // Lookup worker info
  const empInfo = getWorkerInfoByName(seenBy);

  // If worker name matches the user (MAIN_WORKER_NAME)
  let recipients = [];
  let bccList = [];
  let body = "";

  if (MAIN_WORKER_NAME === seenBy) {
    if (dateSeen.getMonth() < today.getMonth()) {
      body = getPostMonthContactReminderHtml(MAIN_WORKER_NAME, MAIN_SUPERVISOR_NAME,
        childName, caseID, Utilities.formatDate(dateSeen, GLOBAL_TIMEZONE, "MM/dd/yyyy"));
    } else if (daysLeftInMonth <= 7) {
      body = getReprimandingContactReminderHtml(MAIN_WORKER_NAME, MAIN_SUPERVISOR_NAME,
        childName, caseID, Utilities.formatDate(dateSeen, GLOBAL_TIMEZONE, "MM/dd/yyyy"),
        daysSinceSeen, daysLeftInMonth);
    } else {
      body = getStandardContactReminderHtml(MAIN_WORKER_NAME, MAIN_SUPERVISOR_NAME,
        childName, caseID, Utilities.formatDate(dateSeen, GLOBAL_TIMEZONE, "MM/dd/yyyy"),
        daysSinceSeen, daysLeftInMonth);
    }
    recipients = [MAIN_WORKER_EMAIL];
    bccList = [MAIN_SUPERVISOR_EMAIL];
    if (daysLeftInMonth <= 7) bccList.push(SSM_EMAIL);

  } else if (empInfo) {
    // External worker logic
    if (dateSeen.getMonth() < today.getMonth()) {
      body = getPostMonthContactReminderHtml(empInfo.workerName, empInfo.supervisorName,
        childName, caseID, Utilities.formatDate(dateSeen, GLOBAL_TIMEZONE, "MM/dd/yyyy"));
    } else if (daysLeftInMonth <= 7) {
      body = getReprimandingContactReminderHtml(empInfo.workerName, empInfo.supervisorName,
        childName, caseID, Utilities.formatDate(dateSeen, GLOBAL_TIMEZONE, "MM/dd/yyyy"),
        daysSinceSeen, daysLeftInMonth);
    } else {
      body = getStandardContactReminderHtml(empInfo.workerName, empInfo.supervisorName,
        childName, caseID, Utilities.formatDate(dateSeen, GLOBAL_TIMEZONE, "MM/dd/yyyy"),
        daysSinceSeen, daysLeftInMonth);
    }
    recipients = [empInfo.workerEmail];
    bccList = [empInfo.supervisorEmail];
    if (daysLeftInMonth <= 7) bccList.push(SSM_EMAIL);

  } else {
    Logger.log(`Could not find worker info for ${seenBy}. Skipping.`);
    return false;
  }

  try {
    GmailApp.sendEmail(recipients.join(","), `Contact Entry Reminder – ${childName}`, "", {
      bcc: bccList.join(","),
      htmlBody: body
    });
    sheet.getRange(rowIndex, 6).setValue(
      Utilities.formatDate(today, GLOBAL_TIMEZONE, "MM/dd/yyyy")
    );
    return true;
  } catch (error) {
    Logger.log(`❌ Failed to send reminder for ${childName}: ${error.message}`);
    return false;
  }
}


/*************************************************************
 * getMonthNumberFromName(name)
 * 
 * Converts month name (January, February, etc.) into 0-based month index.
 *************************************************************/
function getMonthNumberFromName(name) {
  const months = [
    "January","February","March","April","May","June",
    "July","August","September","October","November","December"
  ];
  return months.findIndex(m => m.toLowerCase() === name.toLowerCase());
}

/*************************************************************
 * daysRemainingInMonth(date)
 * 
 * Returns number of days left in the month for the given date.
 *************************************************************/
function daysRemainingInMonth(date) {
  const lastDay = new Date(date.getFullYear(), date.getMonth() + 1, 0);
  return Math.ceil((lastDay - date) / (1000 * 60 * 60 * 24));
}
