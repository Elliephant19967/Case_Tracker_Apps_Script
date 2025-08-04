/*************************************************************
 * Emails.gs
 * 
 * Functions for generating email content and signatures.
 * Uses global variables loaded from Automation Info Sheet.
 *************************************************************/

/**
 * Returns the HTML string for the email signature.
 * Uses global variables for phone numbers and contact info.
 * @returns {string} HTML string for email signature
 */
function getEmailSignatureHtml() {
  // Signature image hosted on Google Drive
  const signatureImageUrl = "https://drive.google.com/uc?export=view&id=14hutBSg3irYox5g8rMQJYUVYY1A885L1";

  // Use globals for phone numbers if available, fallback to hardcoded defaults
  const officeExt = WORKER_OFFICE_EXTENSION || "20657";
  const cellNumber = WORKER_CELL_NUMBER || "681-3833-4019";

  return `
    <br/><br/>
    <img src="${signatureImageUrl}" alt="Department of Human Services Logo" style="width:200px; height:auto; margin-bottom:10px;" />
    
    <p><strong>${MAIN_WORKER_NAME || "Ellie Brewer"}</strong><br/>
    Child Protective Service Worker<br/>
    Bureau for Social Services<br/>
    Kanawha County Department of Health and Human Resources<br/>
    4190 Washington St. W<br/>
    Charleston, WV 25313</p>
    
    <p>P: 304-746-2360 ext. ${officeExt} (Office)<br/>
    P: ${cellNumber} (Cell)<br/>
    F: 304-558-0798</p>
  `;
}

/**
 * Standard Summary Reminder (Day before due date and due date)
 * @param {string} lastName - Last name for the case summary
 * @param {string} summaryLink - Link to the summary
 * @param {boolean} isDueToday - true if due today, false if due tomorrow
 * @returns {string} HTML string for email body
 */
function getStandardSummaryReminderHtml(lastName, summaryLink, isDueToday) {
  const whenText = isDueToday ? "due today" : "due tomorrow";
  return `
    <p>Hey ${MAIN_WORKER_NAME},</p>
    <p>The ${lastName} summary is ${whenText}. A link to the summary is included below to remove a barrier to completing this task and help avoid distraction: </p>
    <p><a href="${summaryLink}">${summaryLink}</a></p>
    <br></br>
    ${getEmailSignatureHtml()}
  `;
}

/**
 * Supervisor Included Summary Reminder (1-6 days late)
 * @param {string} lastName - Last name for the case summary
 * @param {number} daysLate - Number of days the summary is late
 * @param {string} followUpDate - Date by which follow-up is required
 * @param {string} summaryLink - Link to the summary
 * @returns {string} HTML string for email body
 */
function getSupervisorIncludedSummaryReminderHtml(lastName, daysLate, followUpDate, summaryLink) {
  return `
    <p>Hey ${MAIN_WORKER_NAME},</p>
    <p>The ${lastName} summary is ${daysLate} days late. You should get it submitted soon. If the summary isn't submitted by ${followUpDate}, ${MAIN_SUPERVISOR_NAME} will be required to go with you, and they don't want that.</p>
    <p>A link to the summary is included below to remove a barrier to completing this task and help avoid distraction:</p>
    <p><a href="${summaryLink}">${summaryLink}</a></p>
    <br></br>
    ${getEmailSignatureHtml()}
  `;
}

/**
 * Reprimanding Summary Reminder (7+ days late until court date)
 * @param {string} lastName - Last name for the case summary
 * @param {number} remindersSent - Number of reminders sent to worker
 * @param {number} daysUntilHearing - Days left until the hearing
 * @param {number} supervisorRemindersSent - Number of reminders sent to supervisor
 * @param {string} dueDateFormatted - Due date formatted as a string
 * @param {string} summaryLink - Link to the summary
 * @returns {string} HTML string for email body
 */
function getReprimandingSummaryReminderHtml(lastName, remindersSent, daysUntilHearing, supervisorRemindersSent, dueDateFormatted, summaryLink) {
  return `
    <p>${MAIN_WORKER_NAME},</p>
    <p>You have now received ${remindersSent} reminders about the ${lastName} summary being due on ${dueDateFormatted}.</p>
    <p>You only have ${daysUntilHearing} days until this hearing and ${MAIN_SUPERVISOR_NAME} is required to attend with you.</p>
    <p>${MAIN_SUPERVISOR_NAME} has been receiving these reminders for the past ${supervisorRemindersSent} days and now ${SSM_NAME} is included as well.</p>
    <p>You need to submit this ASAP so that you don't have lawyers calling ${MAIN_SUPERVISOR_NAME} or ${SSM_NAME}, which could result in disciplinary action.</p>
    <p>A link to the summary is included below to remove a barrier to completing this task and help avoid distraction:</p>
    <p><a href="${summaryLink}">${summaryLink}</a></p>
    <br></br>
    ${getEmailSignatureHtml()}
  `;
}

/**
 * Standard Contact Reminder Email (before last week of month)
 * @param {string} workerName - Name of the worker
 * @param {string} supervisorName - Name of the supervisor
 * @param {string} childName - Child's name
 * @param {string} caseID - Case ID
 * @param {string} dateSeen - Date child was seen (MM/dd/yyyy)
 * @param {number} daysSinceSeen - Days since the child was seen
 * @param {number} daysLeftInMonth - Days remaining in the month
 * @returns {string} HTML string for email body
 */
function getStandardContactReminderHtml(workerName, supervisorName, childName, caseID, dateSeen, daysSinceSeen, daysLeftInMonth) {
  return `
    <p>Hello ${workerName},</p>

    <p>This is a reminder that you last saw <strong>${childName}</strong> (Case ID: ${caseID}) ${daysSinceSeen} days ago on ${dateSeen}. 
    There are only ${daysLeftInMonth} days remaining in the month, and it is important that the contact is entered as soon as possible to ensure accurate and timely reporting.</p>

    <p>Failure to enter the contact by the end of the month will negatively impact my contact compliance percentage. 
    If the contact is not entered by the last week of the month, ${supervisorName}, your supervisor, and ${SSM_NAME} will be added onto these emails.</p>

    <p>This is an automated message and will be sent daily until the contact is entered and ${MAIN_WORKER_NAME} is notified. 
    ${MAIN_WORKER_NAME} is included in this email so you can reply to this message to let them know if a contact has been entered.</p>

    <p>Thank you for your prompt attention to this matter.</p>

    <br></br>
    ${getEmailSignatureHtml()}
  `;
}

/**
 * Reprimanding Contact Reminder Email (last week of month)
 * @param {string} workerName - Name of the worker
 * @param {string} supervisorName - Name of the supervisor
 * @param {string} childName - Child's name
 * @param {string} caseID - Case ID
 * @param {string} dateSeen - Date child was seen (MM/dd/yyyy)
 * @param {number} daysSinceSeen - Days since the child was seen
 * @param {number} daysRemaining - Days remaining in the month
 * @returns {string} HTML string for email body
 */
function getReprimandingContactReminderHtml(workerName, supervisorName, childName, caseID, dateSeen, daysSinceSeen, daysRemaining) {
  return `
    <p>Hello ${workerName},</p>

    <p>This is a reminder that you last saw <strong>${childName}</strong> (Case ID: ${caseID}) ${daysSinceSeen} days ago on ${dateSeen}. 
    The month is almost over, and it is critical that the contact is entered immediately to maintain compliance.</p>

    <p>There are only ${daysRemaining} remaining this month. Since it is the final week of the month, your supervisor ${supervisorName} and ${SSM_NAME} have been added to this email.</p>

    <p>This is an automated message and will be sent daily until the contact is entered and ${MAIN_WORKER_NAME} is notified. 
    ${MAIN_WORKER_NAME} is included in this email so you can reply to this message to let them know if a contact has been entered.</p>

    <p>Thank you for your immediate attention to this matter.</p>

    <br></br>
    ${getEmailSignatureHtml()}
  `;
}

/**
 * Post-Month Contact Reminder Email (after the month of contact)
 * @param {string} workerName - Name of the worker
 * @param {string} supervisorName - Name of the supervisor
 * @param {string} childName - Child's name
 * @param {string} caseID - Case ID
 * @param {string} dateSeen - Date child was seen (MM/dd/yyyy)
 * @returns {string} HTML string for email body
 */
function getPostMonthContactReminderHtml(workerName, supervisorName, childName, caseID, dateSeen) {
  return `
    <p>Hello ${workerName},</p>

    <p>This is an overdue reminder that you last saw <strong>${childName}</strong> (Case ID: ${caseID}), seen on ${dateSeen}. 
    The contact for this child is now past due for the previous month and must be entered immediately.</p>

    <p>Your supervisor ${supervisorName} and ${SSM_NAME} have been notified of this delay and may follow up with you directly.</p>

    <p>This is an automated message and will continue to be sent every Monday until the contact is entered and ${MAIN_WORKER_NAME} is notified. 
    ${MAIN_WORKER_NAME} is included in this email so you can reply to this message to let them know if a contact has been entered.</p>

    <p>Thank you for your urgent attention to this matter.</p>
    <br></br>
    ${getEmailSignatureHtml()}
  `;
}
