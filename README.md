# Case Management Automation Suite

This Google Apps Script project automates key workflows for case management, including:

- **Monthly Contact Reminders:** Automatically sends email reminders to workers for missing contact entries, with logic handling prior-month reminders, customized email templates, and completed month tracking.
- **Court Summary Reminders:** Sends timely reminders to staff for submitting court summaries, ensuring compliance with reporting deadlines.

The project also provides:

- Worker assignment dropdown management
- Global variables caching and syncing from a dedicated Variables sheet
- UI dialogs to manage completed months
- Trigger setup for scheduled and edit-based automation
- Integration with multiple worker info sources for dynamic email routing

---

## Features

### Monthly Contact Reminders

- Processes monthly contact sheets named like "January Contacts", etc.
- Sends reminders only for missing contact entries.
- Applies Monday-only rules for prior months.
- Marks months complete once all reminders have been sent.
- Supports multiple email templates based on timing and status.
- Tracks completion status with UI for manual adjustment.

### Court Summary Reminders

- Scans court summary tracker sheets.
- Sends automated email reminders for pending summaries.
- Supports configurable schedules and recipient info.
- Includes supervisor and special management bcc options.

### Additional Capabilities

- Automatically updates "Seen By" dropdowns based on worker info tabs.
- Retrieves and caches global variables efficiently with fallback strategies.
- Installs and manages time-based and edit triggers.
- Uses Spreadsheet and Gmail APIs for integrated workflow.

---

## Setup & Local Development

### Prerequisites

- [Node.js](https://nodejs.org/) (v14+ recommended)
- [npm](https://www.npmjs.com/get-npm)
- [Google Apps Script CLI (`clasp`)](https://github.com/google/clasp)

### Installation

1. Clone this repository:

   ```bash
   git clone https://github.com/Elliephant19967/Case_Tracker_Apps_Script.git
   cd https://github.com/Elliephant19967/Case_Tracker_Apps_Script.git
