# Career_Suite_AI_v3
# Automated Job Application Tracker & Dashboard v3 **REVISED NAME**

**Project:** Automated Job Application Tracker & Pipeline Manager v3.1 (Corrected Sections)
**Author:** Francis John LiButti (Original), AI Integration & Refinements by Assistant
**Project Name (Alt):** Automated Job Application Tracker & Dashboard
**Current Version:** v3.1 (Internal Milestone: "Robust Processing & Formula-Driven Analytics")
*(Reflects Gemini API fix, stable chart creation with formula-driven helper data, robust Peak Status, and reliable label management).*
**Last Updated:** May 2024 (Public Release)

## Overview

This Google Apps Script project automates the tedious process of tracking job applications. It leverages Gmail to parse application-related emails, extracts key information (company, job title, status), logs it into a Google Spreadsheet, and provides an insightful dashboard with key metrics and visualizations. This version (v3.1) significantly enhances parsing robustness with Gemini API integration (with regex fallback), improves dashboard stability through formula-driven helper data, and refines status and label management.

## Key Features

*   **Automated Email Processing:** Scans a specified Gmail label (e.g., "Job Application Tracker/To Process") for new application emails.
*   **AI-Powered Parsing (Gemini API):** Utilizes Google's Gemini API to intelligently extract Company Name, Job Title, and Application Status from email content.
*   **Regex Fallback Parsing:** Provides robust regex-based parsing if the Gemini API key is not configured or if API calls fail.
*   **Google Sheets Integration:**
    *   Logs all extracted application data into a structured Google Sheet.
    *   Automatically updates existing entries or creates new ones.
    *   Tracks "Peak Status" to see the furthest stage an application reached.
*   **Dynamic Dashboard:**
    *   Generates a dashboard in a separate Google Sheet tab.
    *   Displays key metrics (Total Apps, Active Apps, Interview Rate, Offer Rate, etc.).
    *   Visualizes data with charts:
        *   Platform Distribution (LinkedIn, Indeed, etc.)
        *   Applications Over Time (Weekly)
        *   Application Funnel (Peak Stages)
    *   Utilizes a hidden "Helper Sheet" with formulas for reliable chart data.
*   **Gmail Label Management:** Automatically moves processed emails to "Processed" or "Manual Review Needed" labels.
*   **Automated Stale Application Handling:** Marks applications as "Rejected" if there have been no updates for a configurable number of weeks (default: 7).
*   **Configuration Driven:** Most settings (Spreadsheet ID, sheet names, Gmail labels, keywords, API keys) are managed in a `Config.gs` file.
*   **Initial Setup Routine:** A one-time function (`initialSetup_LabelsAndSheet`) to create necessary Gmail labels, the Google Sheet, dashboard, helper sheet, and time-driven triggers.

## Technology Stack

*   Google Apps Script (JavaScript)
*   Google Sheets API
*   Gmail API
*   Google Drive API
*   Google Gemini API (Optional, but recommended for best parsing)

## Prerequisites

*   A Google Account.
*   Basic familiarity with Google Apps Script (to copy/paste code and run initial setup).
*   (Optional but Recommended) A Google Gemini API Key for AI-powered email parsing. You can obtain one from [Google AI Studio](https://aistudio.google.com/app/apikey).

## Setup and Installation

1.  **Create a New Google Apps Script Project:**
    *   Go to [script.google.com/create](https://script.google.com/create).
    *   Give your project a name (e.g., "Job Application Tracker").

2.  **Copy the Code:**
    *   The provided code is extensive. It's best practice to organize it into multiple `.gs` files within your Apps Script project for better maintainability. For example:
        *   `Config.gs` (for all the `const` declarations and configuration)
        *   `Main.gs` or `Code.gs` (for `processJobApplicationEmails`, `initialSetup_LabelsAndSheet`, parsing functions, etc.)
        *   `Dashboard.gs` (for dashboard creation, formatting, and chart update functions)
        *   `SheetHelpers.gs` (for `getOrCreateSpreadsheetAndSheet`, `setupSheetFormatting`, etc.)
        *   `TriggerSetup.gs` (for `createTimeDrivenTrigger`, `markStaleApplicationsAsRejected`)
    *   Alternatively, you can paste all the code into the default `Code.gs` file, but it will be very long.
    *   **Ensure the comment `// Paste after Section X` blocks are placed correctly if you are manually splitting the code.**

3.  **Save the Project:** Click the save icon.

4.  **Configure Gemini API Key (Optional but Recommended):**
    *   In the Apps Script editor, go to **Project Settings** (gear icon on the left).
    *   Scroll down to **Script Properties**.
    *   Click **Add script property**.
    *   **Property:** `GEMINI_API_KEY`
    *   **Value:** Paste your Gemini API Key.
    *   Click **Save script properties**.

5.  **Run Initial Setup:**
    *   In the Apps Script editor, select the function `initialSetup_LabelsAndSheet` from the function dropdown (next to the "Debug" and "Run" buttons).
    *   Click **Run**.
    *   **Authorization:** You'll be prompted to authorize the script.
        *   Click "Review permissions."
        *   Choose your Google account.
        *   You might see a "Google hasn’t verified this app" warning. Click "Advanced," then "Go to [Your Project Name] (unsafe)."
        *   Review the permissions the script needs (Gmail, Sheets, Drive) and click "Allow."
    *   The script will:
        *   Create necessary Gmail labels (`Job Application Tracker`, `/To Process`, `/Processed`, `/Manual Review Needed`).
        *   Create a new Google Spreadsheet named "Automated Job Application Tracker Data" (or open/find an existing one if `FIXED_SPREADSHEET_ID` is set and valid).
        *   Set up the "Applications" data sheet, "Dashboard" sheet, and a hidden "DashboardHelperData" sheet.
        *   Format these sheets and set up dashboard formulas.
        *   Attempt to create initial charts (it adds and then removes dummy data for this).
        *   Set up time-driven triggers for email processing and stale application rejection.
    *   You should see a UI alert in Google Sheets (if a sheet was active) or check the `Logger.log` output in the Apps Script editor under "Executions" for success messages or errors.

6.  **Configure `FIXED_SPREADSHEET_ID` (Recommended after first run):**
    *   After the `initialSetup_LabelsAndSheet` function successfully creates the spreadsheet, open it.
    *   Copy its ID from the URL (the long string of characters between `/d/` and `/edit`).
    *   Go back to your `Config.gs` file in the Apps Script editor.
    *   Paste this ID into the `FIXED_SPREADSHEET_ID` constant:
        ```javascript
        const FIXED_SPREADSHEET_ID = "YOUR_COPIED_SPREADSHEET_ID_HERE";
        ```
    *   Save the script. This makes sheet access more reliable.

7.  **Verify Triggers:**
    *   In the Apps Script editor, go to **Triggers** (clock icon on the left).
    *   You should see triggers for `processJobApplicationEmails` (e.g., every hour) and `markStaleApplicationsAsRejected` (e.g., daily).

## Configuration (`Config.gs`)

Modify the constants in `Config.gs` to customize the script's behavior:

*   `DEBUG_MODE`: Set to `true` for detailed logging, `false` for less verbose logs in production.
*   `FIXED_SPREADSHEET_ID`: Set this to your target spreadsheet's ID for direct access. If empty, the script will try to find/create `TARGET_SPREADSHEET_FILENAME`.
*   `TARGET_SPREADSHEET_FILENAME`: Name of the spreadsheet if `FIXED_SPREADSHEET_ID` is not used.
*   `SHEET_TAB_NAME`, `DASHBOARD_TAB_NAME`, `HELPER_SHEET_NAME`: Names for the different sheets.
*   `GMAIL_LABEL_PARENT`, `GMAIL_LABEL_TO_PROCESS`, etc.: Names for Gmail labels.
*   Column Indices (`PROCESSED_TIMESTAMP_COL`, etc.): 1-based indices for columns in the "Applications" sheet. **Change with caution as it affects all data operations.**
*   Status Values (`DEFAULT_STATUS`, `REJECTED_STATUS`, etc.): Standard status strings.
*   `STATUS_HIERARCHY`: Defines the progression and importance of statuses.
*   `WEEKS_THRESHOLD`: For `markStaleApplicationsAsRejected`.
*   Keywords (`REJECTION_KEYWORDS`, `OFFER_KEYWORDS`, etc.): Used by the regex fallback parser.
*   `PLATFORM_DOMAIN_KEYWORDS`, `IGNORED_DOMAINS`: For platform detection and company name cleaning.
*   `GEMINI_API_KEY_PROPERTY`: The name of the script property holding the Gemini API key.

## Usage

1.  **Label Emails in Gmail:**
    *   When you apply for a job or receive an update, apply the `Job Application Tracker/To Process` label to the relevant email thread in Gmail.
2.  **Automated Processing:**
    *   The `processJobApplicationEmails` function will run automatically based on its trigger (e.g., every hour).
    *   It will read emails from the "To Process" label, parse them, and update the Google Sheet.
    *   Processed emails will be moved to "Processed" or "Manual Review Needed."
3.  **View Data and Dashboard:**
    *   Open the "Automated Job Application Tracker Data" Google Sheet.
    *   The "Applications" tab contains the raw data.
    *   The "Dashboard" tab shows your application metrics and charts. The dashboard updates its data via formulas linked to the "Applications" and "DashboardHelperData" sheets. The charts should refresh automatically or when the sheet is opened. You can run `updateDashboardMetrics` manually if needed to force chart object recreation/verification.
4.  **Stale Application Updates:**
    *   The `markStaleApplicationsAsRejected` function runs daily (by default) to update old, non-finalized applications.

## Key Functions

*   `initialSetup_LabelsAndSheet()`: Run once to set up everything.
*   `processJobApplicationEmails()`: The main function that processes emails; triggered automatically.
*   `parseEmailWithGemini()`: Uses Gemini API for parsing.
*   `extractCompanyAndTitle()`: Regex-based fallback parsing.
*   `getOrCreateSpreadsheetAndSheet()`: Manages access to the target Google Sheet and data tab.
*   `setupSheetFormatting()`: Formats the "Applications" data sheet.
*   `getOrCreateDashboardSheet()`, `formatDashboardSheet()`: Creates and formats the dashboard.
*   `getOrCreateHelperSheet()`: Creates and manages the hidden helper sheet.
*   `updateDashboardMetrics()`, `updatePlatformDistributionChart()`, `updateApplicationsOverTimeChart()`, `updateApplicationFunnelChart()`: Manages dashboard data and chart creation/updates.
*   `markStaleApplicationsAsRejected()`: Marks old applications as rejected; triggered automatically.

## Future Enhancements / Ideas

*   More sophisticated handling of email threads with multiple updates.
*   UI for easier manual correction of parsed data directly from the spreadsheet.
*   Direct integration with calendar for scheduling interviews.
*   More granular analytics on the dashboard (e.g., source effectiveness, time-to-hire stages).
*   Support for attachments (e.g., resume versions sent).

## Contributing

Contributions, issues, and feature requests are welcome! Please feel free to fork the repository, make changes, and submit a pull request.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## Author

*   **Francis John LiButti** (Original Concept and Development)
*   AI Integration & Refinements by Assistant

---
*This script is provided as-is. Please test thoroughly before relying on it for critical data.*
