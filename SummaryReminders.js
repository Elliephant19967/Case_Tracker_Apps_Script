/*************************************************************
 * sendSummaryReminders()
 *
 * Sends reminder emails based on summary due dates and submission statuses.
 * 
 * Combines:
 * - Robust global variable handling and URL loading.
 * - Original email-type logic for different reminder tiers.
 *************************************************************/
function sendSummaryReminders() {
  // Ensure global variables are loaded (no popup)
  getGlobalVariables(false);

  // Open the Case Tracker sheet by URL
  const ss = SpreadsheetApp.openByUrl(CASE_TRACKER_URL);
  const sheet = ss.getSheets()[0]; // First sheet assumed
  const data = sheet.getDataRange().getValues();
  const today = new Date();

  Logger.log(`Running sendSummaryReminders() on ${today}`);

  // Loop through rows (skip header)
  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    // Column mapping based on layout
    const fullName = row[0];                // Case Name (LastName, FirstName)
    const courtDateRaw = row[4];            // Next Court Date
    const dueDateRaw = row[7];              // Summary Due Date
    const submitted = row[9] === true;      // Submitted checkbox
    const summaryLink = row[12];            // Link to Summary

    // Parse last name from "LastName, FirstName"
    let lastName = "";
    if (fullName && fullName.includes(",")) {
      lastName = fullName.split(",")[0].trim();
    } else {
      lastName = fullName ? fullName.trim() : "";
      if (!fullName) {
        Logger.log(`Row ${i + 1}: Missing Case Name`);
      } else {
        Logger.log(`Row ${i + 1}: Unexpected Case Name format: "${fullName}"`);
      }
    }

    // Convert date columns to Date objects
    const dueDate = dueDateRaw ? new Date(dueDateRaw) : null;
    const courtDate = courtDateRaw ? new Date(courtDateRaw) : null;

    // Skip if already submitted or missing required fields
    if (submitted || !dueDate || !lastName) {
      Logger.log(`Row ${i + 1}: Skipped (submitted=${submitted}, dueDate=${dueDate}, lastName=${lastName})`);
      continue;
    }

    // Date calculations
    const daysLate = Math.floor((today - dueDate) / (1000 * 60 * 60 * 24));
    const daysUntilHearing = courtDate
      ? Math.ceil((courtDate - today) / (1000 * 60 * 60 * 24))
      : 0;

    // Reminder dates
    const firstReminderDate = new Date(dueDate);
    firstReminderDate.setDate(dueDate.getDate() - 1);

    const followUpDate = new Date(dueDate);
    followUpDate.setDate(dueDate.getDate() + 6);
    const formattedFollowUpDate = followUpDate.toLocaleDateString();

    // Counts for severe overdue logic
    const remindersSent = Math.max(0, daysLate);
    const supervisorRemindersSent = Math.max(0, daysLate - 1);

    // Variables for email construction
    let emailSubject = "";
    let emailHtmlBody = "";
    let recipients = [];

    /**************** Determine which reminder email to send ****************/

    if (today.toDateString() === firstReminderDate.toDateString()) {
      // Day before due date
      emailSubject = `Summary Due Tomorrow`;
      emailHtmlBody = getStandardSummaryReminderHtml(lastName, summaryLink, false);
      recipients = [MAIN_WORKER_EMAIL];

      Logger.log(`Row ${i + 1}: Sending "Due Tomorrow" email for ${lastName}`);

    } else if (today.toDateString() === dueDate.toDateString()) {
      // Due date
      emailSubject = `Summary Due Today`;
      emailHtmlBody = getStandardSummaryReminderHtml(lastName, summaryLink, true);
      recipients = [MAIN_WORKER_EMAIL];

      Logger.log(`Row ${i + 1}: Sending "Due Today" email for ${lastName}`);

    } else if (daysLate > 0 && daysLate < 7) {
      // 1-6 days late (supervisor included)
      emailSubject = `Summary Overdue: ${lastName}`;
      emailHtmlBody = getSupervisorIncludedSummaryReminderHtml(
        lastName,
        daysLate,
        formattedFollowUpDate,
        summaryLink
      );
      recipients = [MAIN_WORKER_EMAIL, MAIN_SUPERVISOR_EMAIL];

      Logger.log(`Row ${i + 1}: Sending "Overdue 1-6 days" email for ${lastName}`);

    } else if (daysLate >= 7) {
      // 7+ days late (SSM included)
      emailSubject = `Urgent: ${lastName} Summary Severely Overdue`;
      emailHtmlBody = getReprimandingSummaryReminderHtml(
        lastName,
        remindersSent,
        daysUntilHearing,
        supervisorRemindersSent,
        dueDate.toLocaleDateString(),
        summaryLink
      );
      recipients = [MAIN_WORKER_EMAIL, MAIN_SUPERVISOR_EMAIL, SSM_EMAIL];

      Logger.log(`Row ${i + 1}: Sending "Severely Overdue" email for ${lastName}`);
    }

    /**************** Send Email ****************/

    // Skip if no recipients or email content determined
    if (recipients.length === 0 || !emailSubject || !emailHtmlBody) {
      Logger.log(`Row ${i + 1}: ⚠️ No email sent (no recipients or content)`);
      continue;
    }

    // Send email with robust error handling
    try {
      sendEmailWithHtml(
        recipients.filter(email => email && email.trim() !== ""), // remove blanks
        emailSubject,
        emailHtmlBody
      );
      Logger.log(`Row ${i + 1}: Email sent successfully to ${recipients.join(", ")}`);
    } catch (e) {
      Logger.log(`Row ${i + 1}: ❌ Failed to send email: ${e.message}`);
    }
  }
}



/*************************************************************
 * sendEmailWithHtml(to, subject, htmlBody)
 *
 * Sends an HTML email.
 * Accepts single or array of recipients.
 * Logs recipient, subject, and function call for debugging.
 *************************************************************/
function sendEmailWithHtml(to, subject, htmlBody) {
  const recipients = Array.isArray(to) ? to.join(",") : to;

  // Log details before sending
  Logger.log(`sendEmailWithHtml() invoked`);
  Logger.log(`Recipients: ${recipients}`);
  Logger.log(`Subject: ${subject}`);
  Logger.log(`Email body preview: ${htmlBody.substring(0, 100)}...`);

  // Send the email
  GmailApp.sendEmail(recipients, subject, "", { htmlBody: htmlBody });
}
