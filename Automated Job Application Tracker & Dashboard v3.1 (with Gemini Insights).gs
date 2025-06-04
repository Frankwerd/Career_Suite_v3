/**
 * Project: Automated Job Application Tracker & Pipeline Manager v3.1 (Corrected Sections)
 * Author: Francis John LiButti (Original), AI Integration & Refinements by Assistant
 * Project Name: Automated Job Application Tracker & Dashboard
 * Current Version: v3.1 (Internal Milestone: "Robust Processing & Formula-Driven Analytics")
 * (Reflects Gemini API fix, stable chart creation with formula-driven helper data, robust Peak Status, and reliable label management).
 * Date: May 18, 2025 (Based on last successful log)
 */

// File: Config.gs
// Description: Contains all global configuration constants for the Job Application Tracker project.
// Modifying values here will change the behavior of the script.const DEBUG_MODE = true;

// -- Spreadsheet Configuration --
const FIXED_SPREADSHEET_ID = "1L7Iy_YpVQj2eaDjkUDF6nFbqTlHhZJncei_S2W2Ygj4"; // Example: "1CnAIrjtTivPu_IUh4wrfwjJH0SmGnJNTGHfRJx5HwuE" or ""
const TARGET_SPREADSHEET_FILENAME = "Automated Job Application Tracker Data";
const SHEET_TAB_NAME = "Applications"; // Data sheet
const DASHBOARD_TAB_NAME = "Dashboard";   // Dashboard sheet
const HELPER_SHEET_NAME = "DashboardHelperData"; //Hidden Helper Sheet

// -- Gmail Label Configuration --
const GMAIL_LABEL_PARENT = "Job Application Tracker";
const GMAIL_LABEL_TO_PROCESS = GMAIL_LABEL_PARENT + "/To Process";
const GMAIL_LABEL_PROCESSED = GMAIL_LABEL_PARENT + "/Processed";
const GMAIL_LABEL_MANUAL_REVIEW = GMAIL_LABEL_PARENT + "/Manual Review Needed";

// --- Column Indices (1-based) ---
const PROCESSED_TIMESTAMP_COL = 1; const EMAIL_DATE_COL = 2; const PLATFORM_COL = 3; const COMPANY_COL = 4; const JOB_TITLE_COL = 5; const STATUS_COL = 6; const PEAK_STATUS_COL = 7; const LAST_UPDATE_DATE_COL = 8; const EMAIL_SUBJECT_COL = 9; const EMAIL_LINK_COL = 10; const EMAIL_ID_COL = 11;
const TOTAL_COLUMNS_IN_SHEET = 11; // Adjusted for Peak Status column

// --- Status Values ---
const DEFAULT_STATUS = "Applied";
const REJECTED_STATUS = "Rejected";
const OFFER_STATUS = "Offer Received";
const ACCEPTED_STATUS = "Offer Accepted"; // Primarily manual input
const INTERVIEW_STATUS = "Interview Scheduled";
const ASSESSMENT_STATUS = "Assessment/Screening";
const APPLICATION_VIEWED_STATUS = "Application Viewed";
const MANUAL_REVIEW_NEEDED = "N/A - Manual Review";
const DEFAULT_PLATFORM = "Other";

// --- Status Progression Order ---
const STATUS_HIERARCHY = {
  [MANUAL_REVIEW_NEEDED]: -1, "Update/Other": 0, [DEFAULT_STATUS]: 1, [APPLICATION_VIEWED_STATUS]: 2, [ASSESSMENT_STATUS]: 3, [INTERVIEW_STATUS]: 4, [OFFER_STATUS]: 5, [REJECTED_STATUS]: 5, [ACCEPTED_STATUS]: 6
};

// --- Config for Auto-Reject Stale Apps ---
const WEEKS_THRESHOLD = 7;
const FINAL_STATUSES_FOR_STALE_CHECK = new Set([REJECTED_STATUS, ACCEPTED_STATUS, "Withdrawn"]);

// Keywords
const REJECTION_KEYWORDS = ["unfortunately", "regret to inform", "not moving forward", "decided not to proceed", "other candidates", "filled the position", "thank you for your time but"];
const OFFER_KEYWORDS = ["pleased to offer", "offer of employment", "job offer", "formally offer you the position"];
const INTERVIEW_KEYWORDS = ["invitation to interview", "schedule an interview", "interview request", "like to speak with you", "next steps involve an interview", "interview availability"];
const ASSESSMENT_KEYWORDS = ["assessment", "coding challenge", "online test", "technical screen", "next step is a skill assessment", "take a short test"];
const APPLICATION_VIEWED_KEYWORDS = ["application was viewed", "your application was viewed by", "recruiter viewed your application", "company viewed your application", "viewed your profile for the role"];
const PLATFORM_DOMAIN_KEYWORDS = { "linkedin.com": "LinkedIn", "indeed.com": "Indeed", "wellfound.com": "Wellfound", "angel.co": "Wellfound" }; // Using full domains for better matching
const IGNORED_DOMAINS = new Set(['greenhouse.io', 'lever.co', 'myworkday.com', 'icims.com', 'ashbyhq.com', 'smartrecruiters.com', 'bamboohr.com', 'taleo.net', 'gmail.com', 'google.com', 'example.com']);

// --- Gemini API Configuration ---
const GEMINI_API_KEY_PROPERTY = 'GEMINI_API_KEY';

// --- Helper: Get or Create Gmail Label ---
function getOrCreateLabel(labelName) {
  if (!labelName || typeof labelName !== 'string' || labelName.trim() === "") {
    Logger.log(`[ERROR] Invalid labelName provided: "${labelName}"`);
    return null;
  }
  let label = null;
  try {
    label = GmailApp.getUserLabelByName(labelName);
  } catch (e) {
    Logger.log(`[ERROR] Error checking for label "${labelName}": ${e}`);
    return null; // Propagate error indication
  }
  if (!label) {
    if (DEBUG_MODE) Logger.log(`[DEBUG] Label "${labelName}" not found. Creating...`);
    try {
      label = GmailApp.createLabel(labelName);
      Logger.log(`[INFO] Successfully created label: "${labelName}"`);
    } catch (e) {
      Logger.log(`[ERROR] Failed to create label "${labelName}": ${e}\n${e.stack}`);
      return null; // Propagate error indication
    }
  } else {
    if (DEBUG_MODE) Logger.log(`[DEBUG] Label "${labelName}" already exists.`);
  }
  return label;
}

// --- Dashboard Helper: Column Index to Letter ---
function columnToLetter(column) {
  let temp, letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

// Paste after Section 1

// --- Helper: Setup Sheet Formatting (for Data Sheet "Applications") ---
function setupSheetFormatting(sheet) {
  // Robust check for valid sheet object at the very beginning
  if (!sheet || typeof sheet.getName !== 'function') {
    Logger.log(`[ERROR] SETUP_SHEET: Invalid sheet object passed. Parameter was: ${sheet}, Type: ${typeof sheet}`);
    // Consider throwing an error here or ensuring this is handled by the caller
    // For now, just exiting to prevent further errors on an invalid object.
    return;
  }
  Logger.log(`[DEBUG] SETUP_SHEET: Entered with sheet named: "${sheet.getName()}". Validating if it's the data sheet.`);


  // Do not attempt to format the dashboard or other unrelated sheets with this specific logic
  if (sheet.getName() !== SHEET_TAB_NAME) {
    if (DEBUG_MODE) Logger.log(`[DEBUG] SETUP_SHEET: Skipping data sheet formatting for a non-data-sheet tab: "${sheet.getName()}".`);
    return;
  }

  if (sheet.getLastRow() === 0 && sheet.getLastColumn() === 0) { // Only if truly empty
    Logger.log(`[INFO] SETUP_SHEET: Data sheet "${sheet.getName()}" is new/empty. Applying detailed formatting.`);
    let headers = new Array(TOTAL_COLUMNS_IN_SHEET).fill('');
    headers[PROCESSED_TIMESTAMP_COL - 1] = "Processed Timestamp"; headers[EMAIL_DATE_COL - 1] = "Email Date"; headers[PLATFORM_COL - 1] = "Platform"; headers[COMPANY_COL - 1] = "Company Name"; headers[JOB_TITLE_COL - 1] = "Job Title"; headers[STATUS_COL - 1] = "Status";
    headers[PEAK_STATUS_COL - 1] = "Peak Status";
    headers[LAST_UPDATE_DATE_COL - 1] = "Last Update Email Date"; headers[EMAIL_SUBJECT_COL - 1] = "Email Subject"; headers[EMAIL_LINK_COL - 1] = "Email Link"; headers[EMAIL_ID_COL - 1] = "Email ID";
    try { sheet.appendRow(headers); } catch(e) { Logger.log(`[ERROR] SETUP_SHEET: Failed to append header row: ${e}`); return; }


    const headerRange = sheet.getRange(1, 1, 1, TOTAL_COLUMNS_IN_SHEET);
    try {
        headerRange.setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle').setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
        sheet.setRowHeight(1, 40); sheet.setFrozenRows(1);
    } catch(e) { Logger.log(`[WARN] SETUP_SHEET: Error setting up header format: ${e}`);}

    const numDataRowsToFormat = sheet.getMaxRows() > 1 ? sheet.getMaxRows() - 1 : 1000;
    if (numDataRowsToFormat > 0) { // Only proceed if there are rows to format
        const allDataRange = sheet.getRange(2, 1, numDataRowsToFormat, TOTAL_COLUMNS_IN_SHEET);
        try { allDataRange.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP).setVerticalAlignment('top'); } catch(e) { Logger.log(`[WARN] SETUP_SHEET: Error setting data range wrap/align: ${e}`);}
        try { sheet.setRowHeights(2, numDataRowsToFormat, 30); }
        catch (e) { Logger.log(`[WARN] Setup: Could not set default row heights for new sheet: ${e}`); }
        if (EMAIL_LINK_COL > 0 && TOTAL_COLUMNS_IN_SHEET >= EMAIL_LINK_COL) { 
            try{const eLCR=sheet.getRange(2, EMAIL_LINK_COL, numDataRowsToFormat, 1);eLCR.setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);}
            catch(e){Logger.log(`[WARN] SETUP_SHEET: Col Link CLIP error for new sheet: ${e}`);}
        }
        const bR=sheet.getRange(2, 1, numDataRowsToFormat, TOTAL_COLUMNS_IN_SHEET);
        try{const b=bR.applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);b.setHeaderRowColor(null).setFirstRowColor("#E3F2FD").setSecondRowColor("#FFFFFF");}
        catch(e){Logger.log(`[WARN] Setup: Banding error for new sheet: ${e}`);}
    }
    
    try {
      sheet.setColumnWidth(PROCESSED_TIMESTAMP_COL,160); sheet.setColumnWidth(EMAIL_DATE_COL,120); sheet.setColumnWidth(PLATFORM_COL,100); sheet.setColumnWidth(COMPANY_COL,200); sheet.setColumnWidth(JOB_TITLE_COL,250); sheet.setColumnWidth(STATUS_COL,150);
      sheet.setColumnWidth(PEAK_STATUS_COL, 150);
      sheet.setColumnWidth(LAST_UPDATE_DATE_COL,160); sheet.setColumnWidth(EMAIL_SUBJECT_COL,300); sheet.setColumnWidth(EMAIL_LINK_COL,100); sheet.setColumnWidth(EMAIL_ID_COL,200);
    } catch (e) { Logger.log(`[WARN] Setup: Could not set col widths: ${e}`); }
    
    try { sheet.hideColumns(PEAK_STATUS_COL); Logger.log(`[INFO] SETUP_SHEET: Hid column ${PEAK_STATUS_COL} (Peak Status) for new sheet.`); }
    catch (e) { Logger.log(`[WARN] SETUP_SHEET: Could not hide Peak Status column for new sheet: ${e}`); }

    const lastUsedColumn = TOTAL_COLUMNS_IN_SHEET; 
    const maxColumnsInSheet = sheet.getMaxColumns(); 
    if (maxColumnsInSheet > lastUsedColumn) { 
        try { sheet.hideColumns(lastUsedColumn + 1, maxColumnsInSheet - lastUsedColumn); } 
        catch(e){ Logger.log(`[WARN] Setup: Hide unused cols error for new sheet: ${e}`); }
    }
    Logger.log(`[INFO] SETUP_SHEET: Detailed formatting attempt complete for new data sheet "${sheet.getName()}".`);

  } else { // Sheet already has content
    if(DEBUG_MODE)Logger.log(`[DEBUG] SETUP_SHEET: Sheet "${sheet.getName()}" has content. Ensuring key formats.`);
    try { if(sheet.getFrozenRows()<1) sheet.setFrozenRows(1); } catch(e){ Logger.log(`[WARN] SETUP_SHEET: Could not set frozen rows on existing sheet: ${e}`); }
    
    if(EMAIL_LINK_COL > 0 && TOTAL_COLUMNS_IN_SHEET >= EMAIL_LINK_COL && sheet.getLastRow() > 1){
        try{const ecl=sheet.getRange(2,EMAIL_LINK_COL,sheet.getLastRow()-1,1);if(ecl.getWrapStrategies()[0][0]!==SpreadsheetApp.WrapStrategy.CLIP){ecl.setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);}}
        catch(e){Logger.log(`[WARN] SETUP_SHEET: Error setting CLIP on existing email link column: ${e}`);}
    }
    
    if (PEAK_STATUS_COL <= sheet.getMaxColumns()) {
        if (sheet.isColumnHiddenByUser(PEAK_STATUS_COL) === false) { 
          try { sheet.hideColumns(PEAK_STATUS_COL); Logger.log(`[INFO] SETUP_SHEET: Ensured Peak Status column is hidden on existing sheet "${sheet.getName()}".`); } 
          catch(e) { Logger.log(`[WARN] Failed to hide Peak Status on existing sheet "${sheet.getName()}": ${e}`);}
        }
    } else if (DEBUG_MODE) {
        Logger.log(`[DEBUG] SETUP_SHEET: Peak status column (expected at index ${PEAK_STATUS_COL}) seems to be beyond max columns (${sheet.getMaxColumns()}) for existing sheet "${sheet.getName()}". Data structure might be from an older version.`);
    }
  }
}

// --- Helper: Get Sheet Access (Simplified: Fixed ID or Dynamic Find/Create without storing ID) ---
function getOrCreateSpreadsheetAndSheet() {
  let ss = null;
  let sheetTab = null; 

  if (FIXED_SPREADSHEET_ID && FIXED_SPREADSHEET_ID.trim() !== "" && FIXED_SPREADSHEET_ID !== "YOUR_SPREADSHEET_ID_HERE") {
    Logger.log(`[INFO] SPREADSHEET: Using Fixed ID: "${FIXED_SPREADSHEET_ID}"`);
    try {
      ss = SpreadsheetApp.openById(FIXED_SPREADSHEET_ID);
      Logger.log(`[INFO] SPREADSHEET: Opened "${ss.getName()}" (Fixed ID).`);
    } catch (e) {
      const msg = `FIXED ID FAIL: Could not open spreadsheet with ID "${FIXED_SPREADSHEET_ID}". ${e.message}. Please check the ID and your permissions. Script may stop or fail later.`;
      Logger.log(`[FATAL] ${msg}`);
      return { spreadsheet: null, sheet: null };
    }
  } else {
    Logger.log(`[INFO] SPREADSHEET: Fixed ID not set. Attempting to find/create sheet by name: "${TARGET_SPREADSHEET_FILENAME}"`);
    try {
      const files = DriveApp.getFilesByName(TARGET_SPREADSHEET_FILENAME);
      if (files.hasNext()) {
        const file = files.next();
        ss = SpreadsheetApp.open(file);
        Logger.log(`[INFO] SPREADSHEET: Found and opened existing spreadsheet by name: "${ss.getName()}" (ID: ${ss.getId()}).`);
        if (files.hasNext()) {
          Logger.log(`[WARN] SPREADSHEET: Multiple files found with the name "${TARGET_SPREADSHEET_FILENAME}". Used the first one.`);
        }
      } else {
        Logger.log(`[INFO] SPREADSHEET: No spreadsheet found by name "${TARGET_SPREADSHEET_FILENAME}". Attempting to create a new one.`);
        try {
          ss = SpreadsheetApp.create(TARGET_SPREADSHEET_FILENAME);
          Logger.log(`[INFO] SPREADSHEET: Successfully created new spreadsheet: "${ss.getName()}" (ID: ${ss.getId()}).`);
        } catch (eCreate) {
          const msg = `CREATE FAIL: Failed to create new spreadsheet named "${TARGET_SPREADSHEET_FILENAME}". ${eCreate.message}. Script may stop or fail later.`;
          Logger.log(`[FATAL] ${msg}`);
          return { spreadsheet: null, sheet: null };
        }
      }
    } catch (eDrive) {
      const msg = `DRIVE/OPEN FAIL: Error accessing Google Drive for "${TARGET_SPREADSHEET_FILENAME}". ${eDrive.message}. Script may stop or fail later.`;
      Logger.log(`[FATAL] ${msg}`);
      return { spreadsheet: null, sheet: null };
    }
  }

  if (ss) {
    sheetTab = ss.getSheetByName(SHEET_TAB_NAME);
    if (!sheetTab) {
      Logger.log(`[INFO] TAB: Data sheet "${SHEET_TAB_NAME}" not found in "${ss.getName()}". Creating...`);
      try {
        sheetTab = ss.insertSheet(SHEET_TAB_NAME, ss.getSheets().length); // Insert at the end initially
        ss.setActiveSheet(sheetTab); // Activate the new sheet
        // Move other sheets if "Applications" should be first (or near first after Dashboard)
        // This part can be complex if a dashboard sheet also needs specific ordering.
        // For now, just ensures data sheet exists. Dashboard function can move itself to front.

        Logger.log(`[INFO] TAB: Created data sheet "${SHEET_TAB_NAME}".`);
        const defaultSheet = ss.getSheetByName('Sheet1');
        if (defaultSheet && defaultSheet.getName().toLowerCase() === 'sheet1' && defaultSheet.getSheetId() !== sheetTab.getSheetId() && ss.getSheets().length > 1) {
          try { ss.deleteSheet(defaultSheet); Logger.log(`[INFO] TAB: Removed default 'Sheet1'.`); }
          catch (eDeleteDefault) { Logger.log(`[WARN] TAB: Failed to remove default 'Sheet1': ${eDeleteDefault.message}`); }
        }
      } catch (eTabCreate) {
        const msg = `TAB CREATE FAIL: Error creating data tab "${SHEET_TAB_NAME}" in "${ss.getName()}". ${eTabCreate.message}.`;
        Logger.log(`[FATAL] ${msg}`);
        return { spreadsheet: ss, sheet: null };
      }
    } else {
      Logger.log(`[INFO] TAB: Found existing data sheet "${SHEET_TAB_NAME}".`);
    }

    if (sheetTab && typeof sheetTab.getName === 'function') {
      Logger.log(`[DEBUG] GET_SHEET: Preparing to call setupSheetFormatting for sheet named: "${sheetTab.getName()}".`);
      try {
        setupSheetFormatting(sheetTab);
      } catch (e_format) {
        Logger.log(`[ERROR] GET_SHEET: Error occurred *during* call to setupSheetFormatting for sheet "${sheetTab.getName()}": ${e_format} \nStack: ${e_format.stack}`);
      }
    } else {
      Logger.log(`[ERROR] GET_SHEET: 'sheetTab' for data is NULL or not valid before calling setupSheetFormatting. Cannot format.`);
    }
  } else {
    Logger.log(`[FATAL] SPREADSHEET: Spreadsheet object ('ss') is null. Cannot proceed.`);
    return { spreadsheet: null, sheet: null };
  }
  return { spreadsheet: ss, sheet: sheetTab };
}

// Paste after Section 2

// --- Triggers ---
function createTimeDrivenTrigger(functionName = 'processJobApplicationEmails', hours = 1) {
  let exists = false;
  try {
    ScriptApp.getProjectTriggers().forEach(t => {
      if (t.getHandlerFunction() === functionName && t.getEventType() === ScriptApp.EventType.CLOCK) {
        exists = true;
      }
    });
    if (!exists) {
      ScriptApp.newTrigger(functionName).timeBased().everyHours(hours).create();
      Logger.log(`[INFO] TRIGGER: ${hours}-hourly trigger for "${functionName}" created successfully.`);
    } else {
      Logger.log(`[INFO] TRIGGER: ${hours}-hourly trigger for "${functionName}" already exists.`);
    }
  } catch (e) {
    Logger.log(`[ERROR] TRIGGER: Failed to create or verify ${hours}-hourly trigger for "${functionName}": ${e.message} (Stack: ${e.stack})`);
    // Optionally, inform user via UI if context allows and it's critical, but for setup, log is primary
  }
  return !exists; // Returns true if a new trigger was created in this call
}

function createOrVerifyStaleRejectTrigger(functionName = 'markStaleApplicationsAsRejected', hour = 2) { // Default to 2 AM
  let exists = false;
  try {
    ScriptApp.getProjectTriggers().forEach(t => {
      if (t.getHandlerFunction() === functionName && t.getEventType() === ScriptApp.EventType.CLOCK) {
        exists = true;
      }
    });
    if (!exists) {
      ScriptApp.newTrigger(functionName).timeBased().everyDays(1).atHour(hour).inTimezone(Session.getScriptTimeZone()).create();
      Logger.log(`[INFO] TRIGGER: Daily trigger for "${functionName}" (around ${hour}:00 script timezone) created successfully.`);
    } else {
      Logger.log(`[INFO] TRIGGER: Daily trigger for "${functionName}" already exists.`);
    }
  } catch (e) {
    Logger.log(`[ERROR] TRIGGER: Failed to create or verify daily trigger for "${functionName}": ${e.message} (Stack: ${e.stack})`);
  }
  return !exists;
}

// --- Initial Setup Function ---
function initialSetup_LabelsAndSheet() {
  Logger.log(`\n==== STARTING INITIAL SETUP (LABELS, SHEETS, TRIGGERS, DASHBOARD, HELPER) ====`);
  let messages = [];
  let overallSuccess = true;
  let dummyDataWasAdded = false; // <<<< DECLARED AT THE TOP OF THE FUNCTION

  // --- 1. Gmail Label Verification/Creation ---
  Logger.log("[INFO] SETUP: Verifying/Creating Labels...");
  const parentLabel = getOrCreateLabel(GMAIL_LABEL_PARENT); 
  Utilities.sleep(100); 
  const toProcessLabel = getOrCreateLabel(GMAIL_LABEL_TO_PROCESS); 
  Utilities.sleep(100);
  const processedLabel = getOrCreateLabel(GMAIL_LABEL_PROCESSED); 
  Utilities.sleep(100);
  const manualReviewLabel = getOrCreateLabel(GMAIL_LABEL_MANUAL_REVIEW);

  if (!parentLabel || !toProcessLabel || !processedLabel || !manualReviewLabel) {
    messages.push(`Labels: One or more labels FAILED.`);
    overallSuccess = false;
  } else {
    messages.push("Labels: OK.");
  }

  // --- 2. Spreadsheet & "Applications" Data Sheet Setup ---
  Logger.log("[INFO] SETUP: Verifying/Creating Data Sheet & Tab ('Applications')...");
  const { spreadsheet: ss, sheet: dataSh } = getOrCreateSpreadsheetAndSheet(); 

  if (!ss || !dataSh) {
    messages.push("Data Sheet/Tab ('Applications'): FAILED.");
    overallSuccess = false;
  } else {
    messages.push(`Data Sheet/Tab ('Applications'): OK. Using spreadsheet "${ss.getName()}" and tab "${dataSh.getName()}".`);

    // --- 3. Add Dummy Data (if sheet is new/empty for initial chart creation) ---
    // dummyDataWasAdded is already declared at the top of the function
    if (dataSh.getLastRow() <= 1) { 
      Logger.log("[INFO] SETUP: Applications sheet is empty/header only. Adding temporary dummy data.");
      try {
        const today = new Date();
        const weekAgo = new Date(today.getTime() - 7 * 24 * 60 * 60 * 1000);
        const twoWeeksAgo = new Date(today.getTime() - 14 * 24 * 60 * 60 * 1000);

        let dummyRowsData = [ // Renamed to avoid conflict with a possible global 'dummyRows'
          [new Date(), twoWeeksAgo, "LinkedIn", "Alpha Inc.", "Engineer I", DEFAULT_STATUS, DEFAULT_STATUS, twoWeeksAgo, "Applied to Alpha", "", ""],
          [new Date(), weekAgo, "Indeed", "Beta LLC", "Analyst Pro", APPLICATION_VIEWED_STATUS, APPLICATION_VIEWED_STATUS, weekAgo, "Viewed at Beta", "", ""],
          [new Date(), today, "Wellfound", "Gamma Solutions", "Manager X", INTERVIEW_STATUS, INTERVIEW_STATUS, today, "Interview at Gamma", "", ""]
        ];
        
        dummyRowsData = dummyRowsData.map(row => {
            while (row.length < TOTAL_COLUMNS_IN_SHEET) row.push("");
            return row.slice(0, TOTAL_COLUMNS_IN_SHEET);
        });

        dataSh.getRange(2, 1, dummyRowsData.length, TOTAL_COLUMNS_IN_SHEET).setValues(dummyRowsData);
        dummyDataWasAdded = true; // Set the flag defined at the top
        Logger.log(`[INFO] SETUP: Added ${dummyRowsData.length} dummy rows.`);
      } catch (e) {
        Logger.log(`[ERROR] SETUP: Failed to add dummy data: ${e.toString()} \nStack: ${e.stack}`);
      }
    }
  } // This closes the 'if (ss && dataSh)' for adding dummy data logic associated with dataSh


  // --- 4. Dashboard & Helper Sheet Setup (only if 'ss' is valid from previous block) ---
  if (ss) { // ss should be valid if we reached here without an early exit for !ss
    Logger.log("[INFO] SETUP: Verifying/Creating and Formatting Dashboard Sheet...");
    try {
      const dashboardSheet = getOrCreateDashboardSheet(ss);
      if (dashboardSheet) {
        formatDashboardSheet(dashboardSheet); 
        messages.push(`Dashboard Sheet: OK.`);
        
        Logger.log("[INFO] SETUP: Verifying/Creating Helper Sheet...");
        const helperSheet = getOrCreateHelperSheet(ss);
        if (helperSheet) { messages.push(`Helper Sheet: OK.`); } 
        else { messages.push(`Helper Sheet: FAILED.`); }
      } else { messages.push(`Dashboard Sheet: FAILED.`); overallSuccess = false; }
    } catch (e) {
      Logger.log(`[ERROR] SETUP: Error during dashboard/helper setup: ${e.toString()} \nStack: ${e.stack}`);
      messages.push(`Dashboard/Helper Sheet: FAILED - ${e.message}.`);
      overallSuccess = false;
    }

    // --- 5. Update Dashboard Metrics (will use dummy data if it was added) ---
    try {
      Logger.log("[INFO] SETUP: Attempting initial dashboard metrics (chart data) update...");
      updateDashboardMetrics(); 
      messages.push("Dashboard Chart Data: Update attempted.");
    } catch (e) {
      Logger.log(`[ERROR] SETUP: Failed during updateDashboardMetrics: ${e.toString()} \nStack: ${e.stack}`);
      messages.push(`Dashboard Chart Data: FAILED - ${e.message}.`);
      overallSuccess = false; 
    }

    // --- 6. Remove Dummy Data (if it was added AND dataSh is valid) ---
    // dummyDataWasAdded flag (defined at the function top) is checked here.
    // Also ensure dataSh is still valid.
    if (dummyDataWasAdded && dataSh && typeof dataSh.deleteRows === 'function') { 
      Logger.log("[INFO] SETUP: Removing temporary dummy data.");
      try {
        // We added 3 dummy rows
        dataSh.deleteRows(2, 3); 
        Logger.log("[INFO] SETUP: Dummy data removed.");
      } catch (e) {
        Logger.log(`[ERROR] SETUP: Failed to remove dummy data: ${e.toString()} \nStack: ${e.stack}.`);
      }
    }
  } else { // This 'else' corresponds to the 'if (ss)' for dashboard/helper setup and updateDashboardMetrics
    messages.push("Dashboard, Helper Sheets, and Metrics Update: SKIPPED (Spreadsheet object 'ss' was not available from getOrCreateSpreadsheetAndSheet).");
  }


  // --- 7. Trigger Verification/Creation ---
  Logger.log("[INFO] SETUP: Verifying/Creating Triggers...");
  // Assuming createTimeDrivenTrigger takes (functionName, hours) and createOrVerifyStaleRejectTrigger takes (functionName, hourOfDay)
  if (createTimeDrivenTrigger('processJobApplicationEmails', 1)) { messages.push("Email Processor Trigger: CREATED."); } 
  else { messages.push("Email Processor Trigger: Not newly created (check logs)."); }

  if (createOrVerifyStaleRejectTrigger('markStaleApplicationsAsRejected', 2)) { messages.push("Stale Reject Trigger: CREATED."); } 
  else { messages.push("Stale Reject Trigger: Not newly created (check logs)."); }

  // --- 8. Final Summary and UI Alert ---
  const finalMsg = `Initial setup process ${overallSuccess ? "completed." : "encountered ISSUES."}\n\nSummary:\n- ${messages.join('\n- ')}`;
  Logger.log(`\n==== INITIAL SETUP ${overallSuccess ? "OK" : "ISSUES"} ====\n${finalMsg.replace(/\n- /g,'\n  - ')}`);
  try {
    if (SpreadsheetApp.getActiveSpreadsheet() && SpreadsheetApp.getUi()) { SpreadsheetApp.getUi().alert( `Setup ${overallSuccess?"Complete":"Issues"}`, finalMsg, SpreadsheetApp.getUi().ButtonSet.OK); } 
    else { Logger.log("[INFO] UI Alert for initial setup skipped.");}
  } catch(e) { Logger.log(`[INFO] UI Alert skipped: ${e.message}.`); }
  Logger.log("==== END INITIAL SETUP ====");
}

// --- REGEX PARSING LOGIC (FALLBACK) ---
function parseCompanyFromDomain(sender) {
  const emailMatch = sender.match(/<([^>]+)>/); if (!emailMatch || !emailMatch[1]) return null;
  const emailAddress = emailMatch[1]; const domainParts = emailAddress.split('@'); if (domainParts.length !== 2) return null;
  let domain = domainParts[1].toLowerCase();
  if (IGNORED_DOMAINS.has(domain) && !domain.includes('wellfound.com') /*Exception for wellfound which can be an ATS domain*/ ) { return null; }
  domain = domain.replace(/^(?:careers|jobs|recruiting|apply|hr|talent|notification|notifications|team|hello|no-reply|noreply)[.-]?/i, '');
  domain = domain.replace(/\.(com|org|net|io|co|ai|dev|xyz|tech|ca|uk|de|fr|app|eu|us|info|biz|work|agency|careers|招聘|group|global|inc|llc|ltd|corp|gmbh)$/i, ''); // More comprehensive TLDs
  domain = domain.replace(/[^a-z0-9]+/gi, ' '); // Replace non-alphanumeric with space
  domain = domain.split(' ').map(word => word.charAt(0).toUpperCase() + word.slice(1)).join(' '); // Capitalize
  return domain.trim() || null;
}

function parseCompanyFromSenderName(sender) {
  const nameMatch = sender.match(/^"?(.*?)"?\s*</);
  let name = nameMatch ? nameMatch[1].trim() : sender.split('<')[0].trim();
  if (!name || name.includes('@') || name.length < 2) return null; // Basic validation

  // Remove common ATS/platform noise
  name = name.replace(/\|\s*(?:greenhouse|lever|wellfound|workday|ashby|icims|smartrecruiters|taleo|bamboohr|recruiterbox|jazzhr|workable|breezyhr|notion)\b/i, '');
  name = name.replace(/\s*(?:via Wellfound|via LinkedIn|via Indeed|from Greenhouse|from Lever|Careers at|Hiring at)\b/gi, '');
  // Remove common generic terms often appended
  name = name.replace(/\s*(?:Careers|Recruiting|Recruitment|Hiring Team|Hiring|Talent Acquisition|Talent|HR|Team|Notifications?|Jobs?|Updates?|Apply|Notification|Hello|No-?Reply|Support|Info|Admin|Department|Notifications)\b/gi, '');
  // Remove trailing legal entities, punctuation, and trim
  name = name.replace(/[|,_.\s]+(?:Inc\.?|LLC\.?|Ltd\.?|Corp\.?|GmbH|Solutions|Services|Group|Global|Technologies|Labs|Studio|Ventures)?$/i, '').trim();
  name = name.replace(/^(?:The|A)\s+/i, '').trim(); // Remove leading "The", "A"
  
  if (name.length > 1 && !/^(?:noreply|no-reply|jobs|careers|support|info|admin|hr|talent|recruiting|team|hello)$/i.test(name.toLowerCase())) {
    return name;
  }
  return null;
}

function extractCompanyAndTitle(message, platform, emailSubject, plainBody) {
  let company = MANUAL_REVIEW_NEEDED; let title = MANUAL_REVIEW_NEEDED;
  const sender = message.getFrom();
  if (DEBUG_MODE) Logger.log(`[DEBUG] RGX_PARSE: Fallback C/T for subj: "${emailSubject.substring(0,100)}"`);
  
  let tempCompanyFromDomain = parseCompanyFromDomain(sender);
  let tempCompanyFromName = parseCompanyFromSenderName(sender);
  if (DEBUG_MODE) Logger.log(`[DEBUG] RGX_PARSE: From Sender -> Name="${tempCompanyFromName}", Domain="${tempCompanyFromDomain}"`);

  // Platform-specific logic (example for Wellfound)
  if (platform === "Wellfound" && plainBody) {
    let wfCoSub = emailSubject.match(/update from (.*?)(?: \|| at |$)/i) || emailSubject.match(/application to (.*?)(?: successfully| at |$)/i) || emailSubject.match(/New introduction from (.*?)(?: for |$)/i);
    if (wfCoSub && wfCoSub[1]) company = wfCoSub[1].trim();
    // More specific Wellfound body parsing can be added here if needed
    if (title === MANUAL_REVIEW_NEEDED && plainBody && sender.toLowerCase().includes("team@hi.wellfound.com")) {
        const markerPhrase = "if there's a match, we will make an email introduction.";
        const markerIndex = plainBody.toLowerCase().indexOf(markerPhrase);
        if (markerIndex !== -1) {
            const relevantText = plainBody.substring(markerIndex + markerPhrase.length);
            const titleMatch = relevantText.match(/^\s*\*\s*([A-Za-z\s.,:&'\/-]+?)(?:\s*\(| at | \n|$)/m);
            if (titleMatch && titleMatch[1]) title = titleMatch[1].trim();
        }
    }
  }

  // Regex patterns for subject line parsing (ordered by specificity or commonness)
  const complexPatterns = [
    { r: /Application for(?: the)?\s+(.+?)\s+at\s+([^-:|–—]+)/i, tI: 1, cI: 2 }, // Title at Company
    { r: /Invite(?:.*?)(?:to|for)(?: an)? interview(?:.*?)\sfor\s+(?:the\s)?(.+?)(?:\s+at\s+([^-:|–—]+))?/i, tI: 1, cI: 2 }, // Interview for Title at Company
    { r: /Your application for(?: the)?\s+(.+?)\s+at\s+([^-:|–—]+)/i, tI: 1, cI: 2 },
    { r: /Regarding your application for\s+(.+?)(?:\s-\s(.*?))?(?:\s@\s(.*?))?$/i, tI: 1, cI: 3, cI2: 2}, // Greenhouse subjects
    { r: /^(?:Update on|Your Application to|Thank you for applying to)\s+([^-:|–—]+?)(?:\s*-\s*([^-:|–—]+))?$/i, cI: 1, tI: 2 }, // Lever style: Company - Title OR Title - Company
    { r: /applying to\s+(.+?)\s+at\s+([^-:|–—]+)/i, tI: 1, cI: 2 },
    { r: /interest in the\s+(.+?)\s+role(?:\s+at\s+([^-:|–—]+))?/i, tI: 1, cI: 2 },
    { r: /update on your\s+(.+?)\s+app(?:lication)?(?:\s+at\s+([^-:|–—]+))?/i, tI: 1, cI: 2 }
  ];

  for (const pI of complexPatterns) {
    let m = emailSubject.match(pI.r);
    if (m) {
      if (DEBUG_MODE) Logger.log(`[DEBUG] RGX_PARSE: Matched subject pattern: ${pI.r}`);
      let extractedTitle = pI.tI > 0 && m[pI.tI] ? m[pI.tI].trim() : null;
      let extractedCompany = pI.cI > 0 && m[pI.cI] ? m[pI.cI].trim() : null;
      if (!extractedCompany && pI.cI2 > 0 && m[pI.cI2]) extractedCompany = m[pI.cI2].trim();

      // Check if one part looks more like a company and the other a title, for Lever-style ambiguous subjects
      if (pI.cI === 1 && pI.tI === 2 && extractedCompany && extractedTitle) { // e.g. "Lever: Company - Title" or "Lever: Title - Company"
          // A simple heuristic: if one contains "Engineer", "Manager", "Analyst", etc. it's likely the title
          if (/\b(engineer|manager|analyst|developer|specialist|lead|director|coordinator|architect|consultant|designer|recruiter|associate|intern)\b/i.test(extractedCompany) &&
             !/\b(engineer|manager|analyst|developer|specialist|lead|director|coordinator|architect|consultant|designer|recruiter|associate|intern)\b/i.test(extractedTitle)) {
              // Swap them
              [extractedCompany, extractedTitle] = [extractedTitle, extractedCompany];
              if (DEBUG_MODE) Logger.log(`[DEBUG] RGX_PARSE: Swapped Company/Title based on keywords. C: ${extractedCompany}, T: ${extractedTitle}`);
          }
      }
      
      if (extractedTitle && (title === MANUAL_REVIEW_NEEDED || title === DEFAULT_STATUS)) title = extractedTitle;
      if (extractedCompany && (company === MANUAL_REVIEW_NEEDED || company === DEFAULT_PLATFORM)) company = extractedCompany;
      
      if (company !== MANUAL_REVIEW_NEEDED && title !== MANUAL_REVIEW_NEEDED && company !== DEFAULT_PLATFORM && title !== DEFAULT_STATUS) break;
    }
  }

  if (company === MANUAL_REVIEW_NEEDED && tempCompanyFromName) company = tempCompanyFromName; // Use parsed sender name if subject failed for company

  // Body Scan Fallback (if still needed)
  if ((company === MANUAL_REVIEW_NEEDED || title === MANUAL_REVIEW_NEEDED || company === DEFAULT_PLATFORM || title === DEFAULT_STATUS) && plainBody) {
    const bodyCleaned = plainBody.substring(0, 1000).replace(/<[^>]+>/g, ' '); // Process first 1000 chars
    if (company === MANUAL_REVIEW_NEEDED || company === DEFAULT_PLATFORM) {
      let bodyCompanyMatch = bodyCleaned.match(/(?:applying to|application with|interview with|position at|role at|opportunity at|Thank you for your interest in working at)\s+([A-Z][A-Za-z\s.&'-]+(?:LLC|Inc\.?|Ltd\.?|Corp\.?|GmbH|Group|Solutions|Technologies)?)(?:[.,\s\n\(]|$)/i);
      if (bodyCompanyMatch && bodyCompanyMatch[1]) company = bodyCompanyMatch[1].trim();
    }
    if (title === MANUAL_REVIEW_NEEDED || title === DEFAULT_STATUS) {
      let bodyTitleMatch = bodyCleaned.match(/(?:application for the|position of|role of|applying for the|interview for the|title:)\s+([A-Za-z][A-Za-z0-9\s.,:&'\/\(\)-]+?)(?:\s\(| at | with |[\s.,\n\(]|$)/i);
      if (bodyTitleMatch && bodyTitleMatch[1]) title = bodyTitleMatch[1].trim();
    }
  }

  if (company === MANUAL_REVIEW_NEEDED && tempCompanyFromDomain) company = tempCompanyFromDomain; // Last resort for company

  const cleanE = (entity, isTitle = false) => {
    if (!entity || entity === MANUAL_REVIEW_NEEDED || entity === DEFAULT_STATUS || entity === DEFAULT_PLATFORM || entity.toLowerCase() === "n/a") return MANUAL_REVIEW_NEEDED;
    let cl = entity.split(/[\n\r#(]| - /)[0]; // Take first line, remove text after # or some " - " patterns
    cl = cl.replace(/ (?:inc|llc|ltd|corp|gmbh)[\.,]?$/i, '').replace(/[,"']?$/, '');
    cl = cl.replace(/^(?:The|A)\s+/i, '');
    cl = cl.replace(/\s+/g, ' ').trim();
    if (isTitle) {
        cl = cl.replace(/JR\d+\s*[-–—]?\s*/i, '');
        cl = cl.replace(/\(Senior\)/i, 'Senior'); // Hoist (Senior)
        cl = cl.replace(/\(.*?(?:remote|hybrid|onsite|contract|part-time|full-time|intern|co-op|stipend|urgent|hiring|opening|various locations).*?\)/gi, '');
        cl = cl.replace(/[-–—:]\s*(?:remote|hybrid|onsite|contract|part-time|full-time|intern|co-op|various locations)\s*$/gi, '');
        cl = cl.replace(/^[-\s#*]+|[,\s]+$/g, ''); // Clean leading/trailing punctuation and spaces
    }
    cl = cl.replace(/[\u2018\u2019\u201A\u201B\u2032\u2035]/g, "'").replace(/[\u201C\u201D\u201E\u201F\u2033\u2036]/g, '"').replace(/&/gi, '&').replace(/ /gi, ' ');
    cl = cl.trim();
    return cl.length < 2 ? MANUAL_REVIEW_NEEDED : cl; // If cleaning results in too short, mark as N/A
  };
    
  company = cleanE(company);
  title = cleanE(title, true);

  if (DEBUG_MODE) Logger.log(`[DEBUG] RGX_PARSE: Final Fallback Result -> Company:"${company}", Title:"${title}"`);
  return {company: company, title: title};
}

function parseBodyForStatus(plainBody) {
  if (!plainBody || plainBody.length < 10) { if (DEBUG_MODE) Logger.log("[DEBUG] RGX_STATUS: Body too short/missing."); return null; }
  let bL = plainBody.toLowerCase().replace(/[.,!?;:()\[\]{}'"“”‘’\-–—]/g, ' ').replace(/\s+/g, ' ').trim();
  if (OFFER_KEYWORDS.some(k => bL.includes(k))) { Logger.log(`[DEBUG] RGX_STATUS: OFFER.`); return OFFER_STATUS; }
  if (INTERVIEW_KEYWORDS.some(k => bL.includes(k))) { Logger.log(`[DEBUG] RGX_STATUS: INTERVIEW.`); return INTERVIEW_STATUS; }
  if (ASSESSMENT_KEYWORDS.some(k => bL.includes(k))) { Logger.log(`[DEBUG] RGX_STATUS: ASSESSMENT.`); return ASSESSMENT_STATUS; }
  if (APPLICATION_VIEWED_KEYWORDS.some(k => bL.includes(k))) { Logger.log(`[DEBUG] RGX_STATUS: APP_VIEWED.`); return APPLICATION_VIEWED_STATUS; }
  if (REJECTION_KEYWORDS.some(k => bL.includes(k))) { Logger.log(`[DEBUG] RGX_STATUS: REJECTION.`); return REJECTED_STATUS; }
  if (DEBUG_MODE) Logger.log("[DEBUG] RGX_STATUS: No specific keywords found by regex.");
  return null;
}
// Paste after Section 3

// --- GEMINI API PARSING LOGIC ---
function parseEmailWithGemini(emailSubject, emailBody, apiKey) {
  if (!apiKey) {
    Logger.log("[INFO] GEMINI_PARSE: API Key not provided. Skipping Gemini call.");
    return null;
  }
  if ((!emailSubject || emailSubject.trim() === "") && (!emailBody || emailBody.trim() === "")) {
    Logger.log("[WARN] GEMINI_PARSE: Both email subject and body are empty. Skipping Gemini call.");
    return null;
  }

  // Using gemini-1.5-flash-latest as it's good for this kind of task and generally available
  const API_ENDPOINT = "https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash-latest:generateContent?key=" + apiKey;
  // Fallback option if Flash gives issues:
  // const API_ENDPOINT = "https://generativelanguage.googleapis.com/v1beta/models/gemini-1.0-pro:generateContent?key=" + apiKey;
  
  Logger.log(`[DEBUG] GEMINI_PARSE: Using API Endpoint: ${API_ENDPOINT.split('key=')[0] + "key=..."}`);

  const bodySnippet = emailBody ? emailBody.substring(0, 12000) : ""; // Max 12k chars for body snippet

  // ---- EXPANDED PROMPT ----
  const prompt = `You are a highly specialized AI assistant expert in parsing job application-related emails for a tracking system. Your sole purpose is to analyze the provided email Subject and Body, and extract three key pieces of information: "company_name", "job_title", and "status". You MUST return this information ONLY as a single, valid JSON object, with no surrounding text, explanations, apologies, or markdown.

CRITICAL INSTRUCTIONS - READ AND FOLLOW CAREFULLY:

**PRIORITY 1: Determine Relevance - IS THIS A JOB APPLICATION UPDATE FOR THE RECIPIENT?**
- Your FIRST task is to assess if the email DIRECTLY relates to a job application previously submitted by the recipient, or an update to such an application.
- **IF THE EMAIL IS NOT APPLICATION-RELATED:** This includes general newsletters, marketing or promotional emails, sales pitches, webinar invitations, event announcements, account security alerts, password resets, bills/invoices, platform notifications not tied to a specific submitted application (e.g., "new jobs you might like"), or spam.
    - In such cases, IMMEDIATELY set ALL three fields ("company_name", "job_title", "status") to the exact string "${MANUAL_REVIEW_NEEDED}".
    - Do NOT attempt to extract any information from these irrelevant emails.
    - Your output for these MUST be: {"company_name": "${MANUAL_REVIEW_NEEDED}","job_title": "${MANUAL_REVIEW_NEEDED}","status": "${MANUAL_REVIEW_NEEDED}"}

**PRIORITY 2: If Application-Related, Proceed with Extraction:**

1.  "company_name":
    *   **Goal**: Extract the full, official name of the HIRING COMPANY to which the user applied.
    *   **ATS Handling**: Emails often originate from Applicant Tracking Systems (ATS) like Greenhouse (notifications@greenhouse.io), Lever (no-reply@hire.lever.co), Workday, Taleo, iCIMS, Ashby, SmartRecruiters, etc. The sender domain may be the ATS. You MUST identify the actual hiring company mentioned WITHIN the email subject or body. Look for phrases like "Your application to [Hiring Company]", "Careers at [Hiring Company]", "Update from [Hiring Company]", or the company name near the job title.
    *   **Do NOT extract**: The name of the ATS (e.g., "Greenhouse", "Lever"), the name of the job board (e.g., "LinkedIn", "Indeed", "Wellfound" - unless the job board IS the direct hiring company), or generic terms.
    *   **Ambiguity**: If the hiring company name is genuinely unclear from an application context, or only an ATS name is present without the actual company, use "${MANUAL_REVIEW_NEEDED}".
    *   **Accuracy**: Prefer full legal names if available (e.g., "Acme Corporation" over "Acme").

2.  "job_title":
    *   **Goal**: Extract the SPECIFIC job title THE USER APPLIED FOR, as mentioned in THIS email. The title is often explicitly stated after phrases like "your application for...", "application to the position of...", "the ... role", or directly alongside the company name in application submission/viewed confirmations.
    *   **LinkedIn Emails ("Application Sent To..." / "Application Viewed By...")**: These emails (often from sender "LinkedIn") frequently state the company name AND the job title the user applied for directly in the main body or a prominent header within the email content. Scrutinize these carefully for both. Example: "Your application for **Senior Product Manager** was sent to **Innovate Corp**." or "A recruiter from **Innovate Corp** viewed your application for **Senior Product Manager**." Extract "Senior Product Manager".
    *   **ATS Confirmation Emails (e.g., from Greenhouse, Lever)**: These emails confirming receipt of an application (e.g., "We've received your application to [Company]") often DO NOT restate the specific job title within the body of *that specific confirmation email*. If the job title IS NOT restated, you MUST use "${MANUAL_REVIEW_NEEDED}" for the job_title. Do not assume it from the subject line unless the subject clearly states "Your application for [Job Title] at [Company]".
    *   **General Updates/Rejections**: Some updates or rejections may or may not restate the title. If the title of the specific application is not clearly present in THIS email, use "${MANUAL_REVIEW_NEEDED}".
    *   **Strict Rule**: Do NOT infer a job title from company career pages, other listed jobs, or generic phrases like "various roles" unless that phrase directly follows "your application for". Only extract what is stated for THIS specific application event in THIS email. If in doubt, or if only a very generic descriptor like "a role" is used without specifics, prefer "${MANUAL_REVIEW_NEEDED}".

3.  "status":
    *   **Goal**: Determine the current status of the application based on the content of THIS email.
    *   **Strictly Adhere to List**: You MUST choose a status ONLY from the following exact list. Do not invent new statuses or use variations:
        *   "${DEFAULT_STATUS}" (Maps to: Application submitted, application sent, successfully applied, application received - first confirmation)
        *   "${REJECTED_STATUS}" (Maps to: Not moving forward, unfortunately, decided not to proceed, position filled by other candidates, regret to inform)
        *   "${OFFER_STATUS}" (Maps to: Offer of employment, pleased to offer, job offer)
        *   "${INTERVIEW_STATUS}" (Maps to: Invitation to interview, schedule an interview, interview request, like to speak with you)
        *   "${ASSESSMENT_STATUS}" (Maps to: Online assessment, coding challenge, technical test, skills test, take-home assignment)
        *   "${APPLICATION_VIEWED_STATUS}" (Maps to: Application was viewed by recruiter/company, your profile was viewed for the role)
        *   "Update/Other" (Maps to: General updates like "still reviewing applications," "we're delayed," "thanks for your patience," status is mentioned but unclear which of the above it fits best.)
    *   **Exclusion**: "${ACCEPTED_STATUS}" is typically set manually by the user after they accept an offer; do not use it.
    *   **Last Resort**: If the email is clearly job-application-related for the recipient, but the status is absolutely ambiguous and doesn't fit "Update/Other" (very rare), then as a final fallback, use "${MANUAL_REVIEW_NEEDED}" for the status.

**Output Requirements**:
*   **ONLY JSON**: Your entire response must be a single, valid JSON object.
*   **NO Extra Text**: No explanations, greetings, apologies, summaries, or markdown formatting (like \`\`\`json\`\`\`).
*   **Structure**: {"company_name": "...", "job_title": "...", "status": "..."}
*   **Placeholder Usage**: Adhere strictly to using "${MANUAL_REVIEW_NEEDED}" when information is absent or criteria are not met, as instructed for each field.

--- EXAMPLES START ---
Example 1 (LinkedIn "Application Sent To Company - Title Clearly Stated"):
Subject: Francis, your application was sent to MycoWorks
Body: LinkedIn. Your application was sent to MycoWorks. MycoWorks - Emeryville, CA (On-Site). Data Architect/Analyst. Applied on May 16, 2025.
Output:
{"company_name": "MycoWorks","job_title": "Data Architect/Analyst","status": "${DEFAULT_STATUS}"}

Example 2 (Indeed "Application Submitted", title present):
Subject: Indeed Application: Senior Software Engineer
Body: indeed. Application submitted. Senior Software Engineer. Innovatech Solutions - Anytown, USA. The following items were sent to Innovatech Solutions.
Output:
{"company_name": "Innovatech Solutions","job_title": "Senior Software Engineer","status": "${DEFAULT_STATUS}"}

Example 3 (Rejection from ATS, title might be in subject, but not confirmed in this email body):
Subject: Update on your application for Product Manager at MegaEnterprises
Body: From: no-reply@greenhouse.io. Dear Applicant, Thank you for your interest in MegaEnterprises. After careful consideration, we have decided to move forward with other candidates for this position.
Output:
{"company_name": "MegaEnterprises","job_title": "Product Manager","status": "${REJECTED_STATUS}"} 
(Self-correction: Title "Product Manager" taken from subject if directly linked to "your application". If subject was generic like "Application Update", job_title would be ${MANUAL_REVIEW_NEEDED})

Example 4 (Interview Invitation via ATS, title present):
Subject: Invitation to Interview: Data Analyst at Beta Innovations (via Lever)
Body: We were impressed with your application for the Data Analyst role and would like to invite you to an interview...
Output:
{"company_name": "Beta Innovations","job_title": "Data Analyst","status": "${INTERVIEW_STATUS}"}

Example 5 (ATS Email - Application Received, NO specific title in THIS email body):
Subject: Thank you for applying to Handshake!
Body: no-reply@greenhouse.io. Hi Francis, Thank you for your interest in Handshake! We have received your application and will be reviewing your background shortly... Handshake Recruiting.
Output:
{"company_name": "Handshake","job_title": "${MANUAL_REVIEW_NEEDED}","status": "${DEFAULT_STATUS}"}

Example 6 (Unrelated Marketing):
Subject: Join our webinar on Future Tech!
Body: Hi User, Don't miss out on our exclusive webinar...
Output:
{"company_name": "${MANUAL_REVIEW_NEEDED}","job_title": "${MANUAL_REVIEW_NEEDED}","status": "${MANUAL_REVIEW_NEEDED}"}

Example 7 (LinkedIn "Application Viewed By..." - Title Clearly Stated):
Subject: Your application was viewed by Gotham Technology Group
Body: LinkedIn. Great job getting noticed by the hiring team at Gotham Technology Group. Gotham Technology Group - New York, United States. Business Analyst/Product Manager. Applied on May 14.
Output:
{"company_name": "Gotham Technology Group","job_title": "Business Analyst/Product Manager","status": "${APPLICATION_VIEWED_STATUS}"}

Example 8 (Wellfound "Application Submitted" - Often has title):
Subject: Application to LILT successfully submitted
Body: wellfound. Your application to LILT for the position of Lead Product Manager has been submitted! View your application. LILT.
Output:
{"company_name": "LILT","job_title": "Lead Product Manager","status": "${DEFAULT_STATUS}"}

Example 9 (Email indicating general interest/no specific role or company clear):
Subject: An interesting opportunity
Body: Hi Francis, Your profile on LinkedIn matches an opening we have. Would you be open to a quick chat? Regards, Recruiter.
Output:
{"company_name": "${MANUAL_REVIEW_NEEDED}","job_title": "${MANUAL_REVIEW_NEEDED}","status": "Update/Other"}
--- EXAMPLES END ---

--- START OF EMAIL TO PROCESS ---
Subject: ${emailSubject}
Body:
${bodySnippet}
--- END OF EMAIL TO PROCESS ---
Output JSON:
`; // End of prompt template literal
  // ---- END OF EXPANDED PROMPT ----

  const payload = {
    "contents": [{"parts": [{"text": prompt}]}],
    "generationConfig": {
      "temperature": 0.2, 
      "maxOutputTokens": 512, 
      "topP": 0.95, 
      "topK": 40
    },
    "safetySettings": [ 
      { "category": "HARM_CATEGORY_HARASSMENT", "threshold": "BLOCK_MEDIUM_AND_ABOVE" },
      { "category": "HARM_CATEGORY_HATE_SPEECH", "threshold": "BLOCK_MEDIUM_AND_ABOVE" },
      { "category": "HARM_CATEGORY_SEXUALLY_EXPLICIT", "threshold": "BLOCK_MEDIUM_AND_ABOVE" },
      { "category": "HARM_CATEGORY_DANGEROUS_CONTENT", "threshold": "BLOCK_MEDIUM_AND_ABOVE" }
    ]
  };
  const options = {'method':'post', 'contentType':'application/json', 'payload':JSON.stringify(payload), 'muteHttpExceptions':true};

  if(DEBUG_MODE)Logger.log(`[DEBUG] GEMINI_PARSE: Calling API for subj: "${emailSubject.substring(0,100)}". Prompt len (approx): ${prompt.length}`);
  let response; let attempt = 0; const maxAttempts = 2;

  while(attempt < maxAttempts){
    attempt++;
    try {
      response = UrlFetchApp.fetch(API_ENDPOINT, options);
      const responseCode = response.getResponseCode(); const responseBody = response.getContentText();
      if(DEBUG_MODE) Logger.log(`[DEBUG] GEMINI_PARSE (Attempt ${attempt}): RC: ${responseCode}. Body(start): ${responseBody.substring(0,200)}`);

      if (responseCode === 200) {
        const jsonResponse = JSON.parse(responseBody);
        if (jsonResponse.candidates && jsonResponse.candidates[0]?.content?.parts?.[0]?.text) {
          let extractedJsonString = jsonResponse.candidates[0].content.parts[0].text.trim();
          // Clean potential markdown code block formatting from the response
          if (extractedJsonString.startsWith("```json")) extractedJsonString = extractedJsonString.substring(7).trim();
          if (extractedJsonString.startsWith("```")) extractedJsonString = extractedJsonString.substring(3).trim();
          if (extractedJsonString.endsWith("```")) extractedJsonString = extractedJsonString.substring(0, extractedJsonString.length - 3).trim();
          
          if(DEBUG_MODE)Logger.log(`[DEBUG] GEMINI_PARSE: Cleaned JSON from API: ${extractedJsonString}`);
          try {
            const extractedData = JSON.parse(extractedJsonString);
            // Basic validation that all expected keys are present
            if (typeof extractedData.company_name !== 'undefined' && 
                typeof extractedData.job_title !== 'undefined' && 
                typeof extractedData.status !== 'undefined') {
              Logger.log(`[INFO] GEMINI_PARSE: Success. C:"${extractedData.company_name}", T:"${extractedData.job_title}", S:"${extractedData.status}"`);
              return {
                  company: extractedData.company_name || MANUAL_REVIEW_NEEDED, 
                  title: extractedData.job_title || MANUAL_REVIEW_NEEDED, 
                  status: extractedData.status || MANUAL_REVIEW_NEEDED
              };
            } else {
              Logger.log(`[WARN] GEMINI_PARSE: JSON from Gemini missing one or more expected fields. Output: ${extractedJsonString}`);
              return {company:MANUAL_REVIEW_NEEDED, title:MANUAL_REVIEW_NEEDED, status:MANUAL_REVIEW_NEEDED}; // Fallback
            }
          } catch (e) {
            Logger.log(`[ERROR] GEMINI_PARSE: Error parsing JSON string from Gemini: ${e.toString()}\nOffending String: >>>${extractedJsonString}<<<`);
            return {company:MANUAL_REVIEW_NEEDED, title:MANUAL_REVIEW_NEEDED, status:MANUAL_REVIEW_NEEDED}; // Fallback
          }
        } else if (jsonResponse.promptFeedback?.blockReason) {
          Logger.log(`[ERROR] GEMINI_PARSE: Prompt blocked by API. Reason: ${jsonResponse.promptFeedback.blockReason}. Details: ${JSON.stringify(jsonResponse.promptFeedback.safetyRatings)}`);
          return {company:MANUAL_REVIEW_NEEDED, title:MANUAL_REVIEW_NEEDED, status:`Blocked: ${jsonResponse.promptFeedback.blockReason}`}; // Include block reason in status for debugging
        } else {
          Logger.log(`[ERROR] GEMINI_PARSE: API response structure unexpected (no candidates/text part). Full Body (first 500 chars): ${responseBody.substring(0,500)}`);
          return null; 
        }
      } else if (responseCode === 429) { // Rate limit
        Logger.log(`[WARN] GEMINI_PARSE: Rate limit (HTTP 429). Attempt ${attempt}/${maxAttempts}. Waiting...`);
        if (attempt < maxAttempts) { Utilities.sleep(5000 + Math.floor(Math.random() * 5000)); continue; }
        else { Logger.log(`[ERROR] GEMINI_PARSE: Max retry attempts reached for rate limit.`); return null; }
      } else { // Other API errors (400, 404 model not found, 500, etc.)
        Logger.log(`[ERROR] GEMINI_PARSE: API call returned HTTP error. Code: ${responseCode}. Body (first 500 chars): ${responseBody.substring(0,500)}`);
        // Specific check for 404 model error, in case it switches during operation
        if (responseCode === 404 && responseBody.includes("is not found for API version")) {
            Logger.log(`[FATAL] GEMINI_MODEL_ERROR: The model specified (${API_ENDPOINT.split('/models/')[1].split(':')[0]}) may no longer be valid or available. Check model name and API version.`)
        }
        return null; // Indicates failure to parse
      }
    } catch (e) { // Catch network errors or other exceptions during UrlFetchApp.fetch
      Logger.log(`[ERROR] GEMINI_PARSE: Exception during API call (Attempt ${attempt}): ${e.toString()}\nStack: ${e.stack}`);
      if (attempt < maxAttempts) { Utilities.sleep(3000); continue; } // Basic retry wait
      return null;
    }
  }
  Logger.log(`[ERROR] GEMINI_PARSE: Failed to get a valid response from Gemini API after ${maxAttempts} attempts.`);
  return null; // Fallback if all attempts fail
}

// --- Main Email Processing Function ---
function processJobApplicationEmails() {
  const SCRIPT_START_TIME = new Date();
  Logger.log(`\n==== STARTING PROCESS JOB EMAILS (${SCRIPT_START_TIME.toLocaleString()}) ====`);

  const scriptProperties = PropertiesService.getScriptProperties(); // Use ScriptProperties for API Key
  const geminiApiKey = scriptProperties.getProperty(GEMINI_API_KEY_PROPERTY);
  let useGemini = false;
  Logger.log(`[DEBUG] API_KEY_CHECK: Attempting key: "${GEMINI_API_KEY_PROPERTY}" using getScriptProperties()`);
  if (geminiApiKey) {
    Logger.log(`[DEBUG] API_KEY_CHECK: Retrieved: "${geminiApiKey.substring(0,10)}..." (len: ${geminiApiKey.length})`);
    if (geminiApiKey.trim() !== "" && geminiApiKey.startsWith("AIza") && geminiApiKey.length > 30) {
      useGemini = true; Logger.log("[INFO] PROCESS_EMAIL: Gemini API Key VALID. Will use Gemini.");
    } else { Logger.log(`[WARN] PROCESS_EMAIL: Gemini API Key found but seems INVALID. Using regex.`); }
  } else { Logger.log("[WARN] PROCESS_EMAIL: Gemini API Key NOT FOUND. Using regex."); }
  if (!useGemini) Logger.log("[INFO] PROCESS_EMAIL: Final decision: Using regex-based parser.");

  const { spreadsheet: ss, sheet: dataSheet } = getOrCreateSpreadsheetAndSheet(); 
  if (!ss || !dataSheet) { Logger.log(`[FATAL ERROR] PROCESS_EMAIL: Sheet/Tab fail. Aborting.`); return; }
  Logger.log(`[INFO] PROCESS_EMAIL: Sheet OK: "${ss.getName()}" / "${dataSheet.getName()}"`);

  let procLbl, processedLblObj, manualLblObj;
  try { 
    procLbl = GmailApp.getUserLabelByName(GMAIL_LABEL_TO_PROCESS);
    processedLblObj = GmailApp.getUserLabelByName(GMAIL_LABEL_PROCESSED);
    manualLblObj = GmailApp.getUserLabelByName(GMAIL_LABEL_MANUAL_REVIEW);
    if (!procLbl || !processedLblObj || !manualLblObj) throw new Error("Core Gmail labels missing.");
    if (DEBUG_MODE) Logger.log(`[DEBUG] PROCESS_EMAIL: Core Gmail labels verified.`);
  } catch(e) {
    Logger.log(`[FATAL ERROR] PROCESS_EMAIL: Labels missing! Error: ${e.message}`); return;
  }

  const lastR = dataSheet.getLastRow(); 
  const existingDataCache = {}; 
  const processedEmailIds = new Set();
  if (lastR >= 2) {
    Logger.log(`[INFO] PRELOAD: Loading data from "${dataSheet.getName()}" (Rows 2 to ${lastR})...`);
    try {
      const colsToPreload = [COMPANY_COL, JOB_TITLE_COL, EMAIL_ID_COL, STATUS_COL, PEAK_STATUS_COL];
      const minCol = Math.min(...colsToPreload); const maxCol = Math.max(...colsToPreload);
      const numColsToRead = maxCol - minCol + 1;
      if (numColsToRead < 1 || minCol < 1) throw new Error("Invalid preload column calculation.");

      const preloadRange = dataSheet.getRange(2, minCol, lastR - 1, numColsToRead);
      const preloadValues = preloadRange.getValues();
      const coIdx = COMPANY_COL-minCol, tiIdx = JOB_TITLE_COL-minCol, idIdx = EMAIL_ID_COL-minCol, stIdx = STATUS_COL-minCol, pkIdx = PEAK_STATUS_COL-minCol;

      for (let i = 0; i < preloadValues.length; i++) {
        const rN = i + 2, rD = preloadValues[i];
        const eId = rD[idIdx]?.toString().trim()||"", oCo = rD[coIdx]?.toString().trim()||"", oTi = rD[tiIdx]?.toString().trim()||"", cS  = rD[stIdx]?.toString().trim()||"", cPkS = rD[pkIdx]?.toString().trim()||"";
        if(eId) processedEmailIds.add(eId);
        const cL = oCo.toLowerCase();
        if(cL && cL !== MANUAL_REVIEW_NEEDED.toLowerCase() && cL !== 'n/a'){
          if(!existingDataCache[cL]) existingDataCache[cL]=[];
          existingDataCache[cL].push({row:rN,emailId:eId,company:oCo,title:oTi,status:cS, peakStatus: cPkS});
        }
      }
      Logger.log(`[INFO] PRELOAD: Complete. Cached ${Object.keys(existingDataCache).length} co. ${processedEmailIds.size} IDs.`);
    } catch (e) { Logger.log(`[FATAL ERROR] Preload: ${e.toString()}\nStack:${e.stack}\nAbort.`); return; }
  } else { Logger.log(`[INFO] PRELOAD: Sheet empty or header only.`); }

  const THREAD_PROCESSING_LIMIT = 20; // Process up to 15 threads per run
  let threadsToProcess = [];
  try { threadsToProcess = procLbl.getThreads(0, THREAD_PROCESSING_LIMIT); } 
  catch (e) { Logger.log(`[ERROR] GATHER_THREADS: Failed for "${procLbl.getName()}": ${e}`); return; }

  const messagesToSort = []; let skippedCount = 0; let fetchErrorCount = 0;
  if (DEBUG_MODE) Logger.log(`[DEBUG] GATHER_THREADS: Found ${threadsToProcess.length} threads.`);
  for (const thread of threadsToProcess) {
    const tId = thread.getId();
    try {
      const mIT = thread.getMessages();
      for (const msg of mIT) {
        const mId = msg.getId();
        if (!processedEmailIds.has(mId)) { messagesToSort.push({ message: msg, date: msg.getDate(), threadId: tId }); } 
        else { skippedCount++; }
      }
    } catch (e) { Logger.log(`[ERROR] GATHER_MESSAGES: Thread ${tId}: ${e}`); fetchErrorCount++; }
  }
  Logger.log(`[INFO] GATHER_MESSAGES: New: ${messagesToSort.length}. Skipped: ${skippedCount}. Fetch errors: ${fetchErrorCount}.`);

  if (messagesToSort.length === 0) {
    Logger.log("[INFO] PROCESS_LOOP: No new messages.");
    try { updateDashboardMetrics(); } catch (e_dash) { Logger.log(`[ERROR] Dashboard update (no new msgs): ${e_dash.message}`); }
    Logger.log(`==== SCRIPT FINISHED (${new Date().toLocaleString()}) - No new messages ====`);
    return;
  }

  messagesToSort.sort((a, b) => new Date(a.date).getTime() - new Date(b.date).getTime());
  Logger.log(`[INFO] PROCESS_LOOP: Sorted ${messagesToSort.length} new messages.`);
  
  let threadProcessingOutcomes = {}; 
  let pTRC = 0, sUC = 0, nEC = 0, pEC = 0;

  for (let i = 0; i < messagesToSort.length; i++) {
    const elapsedTime = (new Date().getTime() - SCRIPT_START_TIME.getTime()) / 1000;
    if (elapsedTime > 320) { Logger.log(`[WARN] Time limit nearing (${elapsedTime}s). Stopping loop.`); break; } // Slightly increased margin

    const entry = messagesToSort[i];
    const { message, date: emailDateObj, threadId } = entry;
    const emailDate = new Date(emailDateObj); 
    const msgId = message.getId();
    const pSTM = new Date();
    if(DEBUG_MODE)Logger.log(`\n--- Processing Msg ${i+1}/${messagesToSort.length} (ID: ${msgId}, Thread: ${threadId}) ---`);

    let companyName=MANUAL_REVIEW_NEEDED, jobTitle=MANUAL_REVIEW_NEEDED, applicationStatus=null; 
    let plainBodyText=null, requiresManualReview=false, sheetWriteOpSuccess=false;

    try {
      const emailSubject=message.getSubject()||"", senderEmail=message.getFrom()||"", emailPermaLink=`https://mail.google.com/mail/u/0/#inbox/${msgId}`, currentTimestamp=new Date();
      let detectedPlatform=DEFAULT_PLATFORM;
      try{const eAM=senderEmail.match(/<([^>]+)>/);if(eAM&&eAM[1]){const sD=eAM[1].split('@')[1]?.toLowerCase();if(sD){for(const k in PLATFORM_DOMAIN_KEYWORDS){if(sD.includes(k)){detectedPlatform=PLATFORM_DOMAIN_KEYWORDS[k];break;}}}}if(DEBUG_MODE)Logger.log(`[DEBUG] Platform: ${detectedPlatform}`);}catch(ePlat){Logger.log(`WARN: Plat Detect Err: ${ePlat}`);}
      try{plainBodyText=message.getPlainBody();}catch(eBody){Logger.log(`WARN: Get Body Fail Msg ${msgId}: ${eBody}`);plainBodyText="";}

      if (useGemini && plainBodyText && plainBodyText.trim() !== "") {
        const gRes = parseEmailWithGemini(emailSubject, plainBodyText, geminiApiKey);
        if (gRes) { 
            companyName=gRes.company||MANUAL_REVIEW_NEEDED; jobTitle=gRes.title||MANUAL_REVIEW_NEEDED; applicationStatus=gRes.status;
            Logger.log(`[INFO] Gemini: C:"${companyName}", T:"${jobTitle}", S:"${applicationStatus}"`);
            if(!applicationStatus||applicationStatus===MANUAL_REVIEW_NEEDED||applicationStatus==="Update/Other"){const kS=parseBodyForStatus(plainBodyText); if(kS&&kS!==DEFAULT_STATUS){applicationStatus=kS;} else if(!applicationStatus&&kS===DEFAULT_STATUS){applicationStatus=DEFAULT_STATUS;}}
        } else { 
            Logger.log(`[WARN] Gemini fail Msg ${msgId}. Fallback regex.`);
            const rEx=extractCompanyAndTitle(message,detectedPlatform,emailSubject,plainBodyText);companyName=rEx.company;jobTitle=rEx.title;applicationStatus=parseBodyForStatus(plainBodyText);
        }
      } else { 
          const rEx=extractCompanyAndTitle(message,detectedPlatform,emailSubject,plainBodyText);companyName=rEx.company;jobTitle=rEx.title;applicationStatus=parseBodyForStatus(plainBodyText);
          if(DEBUG_MODE) Logger.log(`[DEBUG] Regex Parse: C:"${companyName}", T:"${jobTitle}", S (body scan):"${applicationStatus}"`);
      }
      
      requiresManualReview = (companyName === MANUAL_REVIEW_NEEDED || jobTitle === MANUAL_REVIEW_NEEDED);
      const finalStatusForSheet = applicationStatus || DEFAULT_STATUS;
      const companyCacheKey = (companyName !== MANUAL_REVIEW_NEEDED) ? companyName.toLowerCase() : `_manual_review_placeholder_${msgId}`;
      let existingRowInfo = null; let targetSheetRow = -1;

      if (companyName !== MANUAL_REVIEW_NEEDED && existingDataCache[companyCacheKey]) {
          const pM = existingDataCache[companyCacheKey];
          if (jobTitle !== MANUAL_REVIEW_NEEDED) existingRowInfo = pM.find(e => e.title && e.title.toLowerCase() === jobTitle.toLowerCase());
          if (!existingRowInfo && pM.length > 0) existingRowInfo = pM.reduce((l,c)=>(c.row > l.row ? c : l), pM[0]);
          if (existingRowInfo) targetSheetRow = existingRowInfo.row;
      }

      if (targetSheetRow !== -1 && existingRowInfo) { // UPDATE EXISTING ROW
        const rangeToUpdate = dataSheet.getRange(targetSheetRow, 1, 1, TOTAL_COLUMNS_IN_SHEET);
        const currentSheetValuesRow = rangeToUpdate.getValues()[0];
        let newSheetValues = [...currentSheetValuesRow];

        newSheetValues[PROCESSED_TIMESTAMP_COL - 1] = currentTimestamp;
        // EMAIL_DATE_COL: Update only if this email is newer than what's in sheet for this field specifically
        const existingEmailDate = currentSheetValuesRow[EMAIL_DATE_COL-1];
        if (!(existingEmailDate instanceof Date) || emailDate.getTime() > existingEmailDate.getTime()) {
            newSheetValues[EMAIL_DATE_COL - 1] = emailDate;
        }
        const existingLastUpdate = newSheetValues[LAST_UPDATE_DATE_COL-1];
        if(!(existingLastUpdate instanceof Date) || emailDate.getTime() > existingLastUpdate.getTime()){
            newSheetValues[LAST_UPDATE_DATE_COL-1] = emailDate;
        }
        newSheetValues[EMAIL_SUBJECT_COL-1]=emailSubject; newSheetValues[EMAIL_LINK_COL-1]=emailPermaLink; newSheetValues[EMAIL_ID_COL-1]=msgId; newSheetValues[PLATFORM_COL-1]=detectedPlatform;
        if(companyName!==MANUAL_REVIEW_NEEDED && (newSheetValues[COMPANY_COL-1]===MANUAL_REVIEW_NEEDED || companyName.toLowerCase()!==newSheetValues[COMPANY_COL-1]?.toLowerCase())) newSheetValues[COMPANY_COL-1]=companyName;
        if(jobTitle!==MANUAL_REVIEW_NEEDED && (newSheetValues[JOB_TITLE_COL-1]===MANUAL_REVIEW_NEEDED || jobTitle.toLowerCase()!==newSheetValues[JOB_TITLE_COL-1]?.toLowerCase())) newSheetValues[JOB_TITLE_COL-1]=jobTitle;
        
        const statusInSheetBeforeUpdate = currentSheetValuesRow[STATUS_COL-1]?.toString().trim() || DEFAULT_STATUS;
        let statusForThisUpdate = finalStatusForSheet; 
        // Refined Status Update Logic
        if (statusInSheetBeforeUpdate !== ACCEPTED_STATUS || statusForThisUpdate === ACCEPTED_STATUS) {
            const curRank = STATUS_HIERARCHY[statusInSheetBeforeUpdate] ?? 0;
            const newRank = STATUS_HIERARCHY[statusForThisUpdate] ?? 0;
            if (newRank >= curRank || statusForThisUpdate === REJECTED_STATUS || statusForThisUpdate === OFFER_STATUS ) {
                 newSheetValues[STATUS_COL - 1] = statusForThisUpdate;
            } else { Logger.log(`[DEBUG] Status not updated: "${statusForThisUpdate}" (rank ${newRank}) is not >= current "${statusInSheetBeforeUpdate}" (rank ${curRank}) and not final.`); }
        } else { Logger.log(`[DEBUG] Status is "${ACCEPTED_STATUS}", not changing to "${statusForThisUpdate}".`); }
        const statusAfterUpdate = newSheetValues[STATUS_COL - 1];

        // --- PEAK STATUS LOGIC for UPDATE ---
        let currentPeakFromSheet = existingRowInfo.peakStatus || currentSheetValuesRow[PEAK_STATUS_COL - 1]?.toString().trim();
        if (!currentPeakFromSheet || currentPeakFromSheet === MANUAL_REVIEW_NEEDED || currentPeakFromSheet === "") currentPeakFromSheet = DEFAULT_STATUS; 
        
        const currentPeakRank = STATUS_HIERARCHY[currentPeakFromSheet] ?? -2;
        const newStatusRankForPeak = STATUS_HIERARCHY[statusAfterUpdate] ?? -2; // Use the just-updated status for peak eval
        const excludedFromPeak = new Set([REJECTED_STATUS, ACCEPTED_STATUS, MANUAL_REVIEW_NEEDED, "Update/Other"]);

        let updatedPeakStatus = currentPeakFromSheet; 
        if (newStatusRankForPeak > currentPeakRank && !excludedFromPeak.has(statusAfterUpdate)) {
            updatedPeakStatus = statusAfterUpdate;
        } else if (currentPeakFromSheet === DEFAULT_STATUS && !excludedFromPeak.has(statusAfterUpdate) && STATUS_HIERARCHY[statusAfterUpdate] > STATUS_HIERARCHY[DEFAULT_STATUS]) {
            updatedPeakStatus = statusAfterUpdate; 
        }
        newSheetValues[PEAK_STATUS_COL - 1] = updatedPeakStatus;
        if(DEBUG_MODE) Logger.log(`[DEBUG] Peak Status Update: Row ${targetSheetRow}. Current Peak: "${currentPeakFromSheet}", New Current Status: "${statusAfterUpdate}", New Peak Set: "${updatedPeakStatus}"`);
        
        rangeToUpdate.setValues([newSheetValues]);
        Logger.log(`[INFO] SHEET WRITE: Updated Row ${targetSheetRow}. Status: "${statusAfterUpdate}", Peak: "${updatedPeakStatus}"`);
        sUC++; sheetWriteOpSuccess = true;
        const cacheKey = (newSheetValues[COMPANY_COL - 1] !== MANUAL_REVIEW_NEEDED) ? newSheetValues[COMPANY_COL - 1].toLowerCase() : companyCacheKey;
        if(existingDataCache[cacheKey]){existingDataCache[cacheKey]=existingDataCache[cacheKey].map(e=>e.row===targetSheetRow?{...e, status:statusAfterUpdate, peakStatus:updatedPeakStatus}:e);}

      } else { // APPEND NEW ROW
        const nRC = new Array(TOTAL_COLUMNS_IN_SHEET).fill("");
        nRC[PROCESSED_TIMESTAMP_COL-1]=currentTimestamp; nRC[EMAIL_DATE_COL-1]=emailDate; nRC[PLATFORM_COL-1]=detectedPlatform; nRC[COMPANY_COL-1]=companyName; nRC[JOB_TITLE_COL-1]=jobTitle; nRC[STATUS_COL-1]=finalStatusForSheet; nRC[LAST_UPDATE_DATE_COL-1]=emailDate; nRC[EMAIL_SUBJECT_COL-1]=emailSubject; nRC[EMAIL_LINK_COL-1]=emailPermaLink; nRC[EMAIL_ID_COL-1]=msgId;

        const excludedFromPeakInit = new Set([REJECTED_STATUS, ACCEPTED_STATUS, MANUAL_REVIEW_NEEDED, "Update/Other"]);
        if (!excludedFromPeakInit.has(finalStatusForSheet)) { nRC[PEAK_STATUS_COL - 1] = finalStatusForSheet; } 
        else { nRC[PEAK_STATUS_COL - 1] = DEFAULT_STATUS; }
        if(DEBUG_MODE) Logger.log(`[DEBUG] Peak Status (New Row): Initial Peak set to "${nRC[PEAK_STATUS_COL - 1]}" for initial status "${finalStatusForSheet}"`);
        
        dataSheet.appendRow(nRC);
        const nSRN = dataSheet.getLastRow();
        Logger.log(`[INFO] SHEET WRITE: Appended Row ${nSRN}. Status: "${finalStatusForSheet}", Peak: "${nRC[PEAK_STATUS_COL - 1]}"`);
        nEC++; sheetWriteOpSuccess = true;
        const newEntryCacheKey = (nRC[COMPANY_COL-1] !== MANUAL_REVIEW_NEEDED) ? nRC[COMPANY_COL-1].toLowerCase() : companyCacheKey; // Use actual company or placeholder
        if(!existingDataCache[newEntryCacheKey]) existingDataCache[newEntryCacheKey]=[];
        existingDataCache[newEntryCacheKey].push({row:nSRN,emailId:msgId,company:nRC[COMPANY_COL-1],title:nRC[JOB_TITLE_COL-1],status:nRC[STATUS_COL-1], peakStatus:nRC[PEAK_STATUS_COL-1]});
      }

      // --- Corrected threadProcessingOutcomes logic ---
      if (sheetWriteOpSuccess) {
        pTRC++;
        processedEmailIds.add(msgId);
        let messageOutcome = (requiresManualReview || companyName === MANUAL_REVIEW_NEEDED || jobTitle === MANUAL_REVIEW_NEEDED) ? 'manual' : 'done';
        
        // If thread is already marked 'manual' by a previous message in THIS run, it stays 'manual'.
        // Otherwise, set or update the outcome.
        if (threadProcessingOutcomes[threadId] !== 'manual') {
            threadProcessingOutcomes[threadId] = messageOutcome;
        }
        // If current message dictates 'manual', ensure thread outcome reflects that (it might be first message or override a 'done')
        if (messageOutcome === 'manual') {
             threadProcessingOutcomes[threadId] = 'manual';
        }
        // If threadProcessingOutcomes[threadId] was undefined, it's now set.
        if (DEBUG_MODE) Logger.log(`[DEBUG] Thread ${threadId} outcome for labeling set to: ${threadProcessingOutcomes[threadId]} (current message was: ${messageOutcome})`);
      } else {
        pEC++;
        threadProcessingOutcomes[threadId] = 'manual'; 
        Logger.log(`[ERROR] SHEET WRITE FAILED Msg ${msgId}. Thread ${threadId} auto-marked manual.`);
      }

    } catch (e) {
      Logger.log(`[FATAL ERROR] Processing Msg ${msgId} (Thread ${threadId}): ${e.message}\nStack: ${e.stack}`);
      threadProcessingOutcomes[threadId] = 'manual'; // Ensure thread is marked for manual review on error
      pEC++;
    }
    if(DEBUG_MODE){const pTTM=(new Date().getTime()-pSTM.getTime())/1000;Logger.log(`--- End Msg ${i+1}/${messagesToSort.length} --- Time: ${pTTM}s ---`);} 
    Utilities.sleep(200 + Math.floor(Math.random() * 150)); // Slightly reduced sleep
  }

  Logger.log(`\n[INFO] PROCESS_LOOP: Finished loop. Parsed: ${pTRC}, Sheet Updates: ${sUC}, New Entries: ${nEC}, Processing Errors: ${pEC}.`);
  if(DEBUG_MODE && Object.keys(threadProcessingOutcomes).length > 0) Logger.log(`[DEBUG] Final Thread Outcomes for Labeling: ${JSON.stringify(threadProcessingOutcomes)}`);
  else if (DEBUG_MODE) Logger.log(`[DEBUG] No thread outcomes recorded for labeling (threadProcessingOutcomes empty).`);

  applyFinalLabels(threadProcessingOutcomes, procLbl, processedLblObj, manualLblObj);
  
  try {
    Logger.log("[INFO] PROCESS_EMAIL: Attempting final dashboard metrics update...");
    updateDashboardMetrics();
  } catch (e_dash_final) {
    Logger.log(`[ERROR] PROCESS_EMAIL: Failed final dashboard update call: ${e_dash_final.message}`);
  }

  const SCRIPT_END_TIME = new Date();
  Logger.log(`\n==== SCRIPT FINISHED (${SCRIPT_END_TIME.toLocaleString()}) === Total Time: ${(SCRIPT_END_TIME.getTime() - SCRIPT_START_TIME.getTime())/1000}s ====`);
}


// --- Helper: Apply Labels After Processing ---
function applyFinalLabels(threadOutcomes, processingLabel, processedLabelObj, manualReviewLabelObj) { /* ... Same as v2.3 ... */ const tTU=Object.keys(threadOutcomes);if(tTU.length===0){Logger.log("[INFO] LABEL_MGMT: No thread outcomes.");return;}Logger.log(`[INFO] LABEL_MGMT: Applying labels for ${tTU.length} threads.`);let sLC=0;let lE=0;if(!processingLabel||!processedLabelObj||!manualReviewLabelObj){Logger.log(`[ERROR] LABEL_MGMT: Invalid label objects. Aborting.`);return;}const tPLN=processingLabel.getName();for(const tId of tTU){const o=threadOutcomes[tId];const tLTA=(o==='manual')?manualReviewLabelObj:processedLabelObj;const tLTAN=tLTA.getName();try{const th=GmailApp.getThreadById(tId);if(!th){Logger.log(`[WARN] LABEL_MGMT: Thread ${tId} not found. Skip.`);lE++;continue;}const cTL=th.getLabels().map(l=>l.getName());let lACTT=false;if(cTL.includes(tPLN)){try{th.removeLabel(processingLabel);if(DEBUG_MODE)Logger.log(`[DEBUG] LABEL_MGMT: Removed "${tPLN}" from ${tId}`);lACTT=true;}catch(e){Logger.log(`[WARN] LABEL_MGMT: Fail remove "${tPLN}" from ${tId}: ${e}`);}} if(!cTL.includes(tLTAN)){try{th.addLabel(tLTA);Logger.log(`[INFO] LABEL_MGMT: Added "${tLTAN}" to ${tId}`);lACTT=true;}catch(e){Logger.log(`[ERROR] LABEL_MGMT: Fail add "${tLTAN}" to ${tId}: ${e}`);lE++;continue;}}else if(DEBUG_MODE)Logger.log(`[DEBUG] LABEL_MGMT: Thread ${tId} already has "${tLTAN}".`); if(lACTT){sLC++;Utilities.sleep(150+Math.floor(Math.random()*100));}}catch(e){Logger.log(`[ERROR] LABEL_MGMT: General error for ${tId}: ${e}`);lE++;}} Logger.log(`[INFO] LABEL_MGMT: Finished. Success changes/verified: ${sLC}. Errors: ${lE}.`);}

// --- Helper: Apply Labels After Processing ---
function applyFinalLabels(threadOutcomes, processingLabel, processedLabelObj, manualReviewLabelObj) {
  const threadIdsToUpdate = Object.keys(threadOutcomes);
  if (threadIdsToUpdate.length === 0) { Logger.log("[INFO] LABEL_MGMT: No thread outcomes to process."); return; }
  Logger.log(`[INFO] LABEL_MGMT: Applying labels for ${threadIdsToUpdate.length} threads.`);
  let successfulLabelChanges = 0; let labelErrors = 0;
  if (!processingLabel || !processedLabelObj || !manualReviewLabelObj || typeof processingLabel.getName !== 'function' || typeof processedLabelObj.getName !== 'function' || typeof manualReviewLabelObj.getName !== 'function' ) { Logger.log(`[ERROR] LABEL_MGMT: Invalid label objects. Aborting.`); return; }
  const toProcessLabelName = processingLabel.getName();
  for (const threadId of threadIdsToUpdate) {
    const outcome = threadOutcomes[threadId]; const targetLabelToAdd = (outcome === 'manual') ? manualReviewLabelObj : processedLabelObj; const targetLabelNameToAdd = targetLabelToAdd.getName();
    try {
      const thread = GmailApp.getThreadById(threadId); if (!thread) { Logger.log(`[WARN] LABEL_MGMT: Thread ${threadId} not found. Skip.`); labelErrors++; continue; }
      const currentThreadLabels = thread.getLabels().map(l => l.getName()); let labelsActuallyChangedThisThread = false;
      if (currentThreadLabels.includes(toProcessLabelName)) { try { thread.removeLabel(processingLabel); if (DEBUG_MODE) Logger.log(`[DEBUG] LABEL_MGMT: Removed "${toProcessLabelName}" from ${threadId}`); labelsActuallyChangedThisThread = true; } catch (e) { Logger.log(`[WARN] LABEL_MGMT: Fail remove "${toProcessLabelName}" from ${threadId}: ${e}`); } }
      if (!currentThreadLabels.includes(targetLabelNameToAdd)) { try { thread.addLabel(targetLabelToAdd); Logger.log(`[INFO] LABEL_MGMT: Added "${targetLabelNameToAdd}" to ${threadId}`); labelsActuallyChangedThisThread = true; } catch (e) { Logger.log(`[ERROR] LABEL_MGMT: Fail add "${targetLabelNameToAdd}" to ${threadId}: ${e}`); labelErrors++; continue; } }
      else if (DEBUG_MODE) Logger.log(`[DEBUG] LABEL_MGMT: Thread ${threadId} already has "${targetLabelNameToAdd}".`);
      if (labelsActuallyChangedThisThread) { successfulLabelChanges++; Utilities.sleep(200 + Math.floor(Math.random() * 100)); } // Slightly longer pause for label changes
    } catch (e) { Logger.log(`[ERROR] LABEL_MGMT: General error for ${threadId}: ${e}`); labelErrors++; }
  }
  Logger.log(`[INFO] LABEL_MGMT: Finished. Success changes/verified: ${successfulLabelChanges}. Errors: ${labelErrors}.`);
}

// --- Auto-Reject Stale Applications Function ---
function markStaleApplicationsAsRejected() {
  const SSTA = new Date(); Logger.log(`\n==== AUTO-REJECT STALE START (${SSTA.toLocaleString()}) ====`);
  const { spreadsheet: ss, sheet: dataSheet } = getOrCreateSpreadsheetAndSheet();
  if (!ss || !dataSheet) { Logger.log(`[FATAL] AUTO-REJECT_STALE: No Sheet/Tab. Abort.`); return; }
  Logger.log(`[INFO] AUTO-REJECT_STALE: Using "${dataSheet.getName()}" in "${ss.getName()}".`);
  const dR = dataSheet.getDataRange(); const sV = dR.getValues();
  if (sV.length <= 1) { Logger.log("[INFO] AUTO-REJECT_STALE: No data rows."); return; }
  const cD = new Date(); const sThD = new Date(); sThD.setDate(cD.getDate() - (WEEKS_THRESHOLD * 7));
  Logger.log(`[INFO] AUTO-REJECT_STALE: Stale if Last Update < ${sThD.toLocaleDateString()}`);
  let uAC = 0;
  for (let i = 1; i < sV.length; i++) {
    const cR = sV[i]; const cSt = cR[STATUS_COL - 1]?.toString().trim(); const lUDV = cR[LAST_UPDATE_DATE_COL - 1]; let lUD;
    if (lUDV instanceof Date) { lUD = lUDV; } else if (lUDV && typeof lUDV === 'string' && !isNaN(new Date(lUDV).getTime())) { lUD = new Date(lUDV); }
    else { if (DEBUG_MODE && lUDV) Logger.log(`[DEBUG] AUTO-REJECT_STALE: Row ${i + 1} Skip: Invalid Last Update Date: "${lUDV}"`); continue; }
    if (FINAL_STATUSES_FOR_STALE_CHECK.has(cSt) || !cSt || cSt === MANUAL_REVIEW_NEEDED) { if (DEBUG_MODE && cSt) Logger.log(`[DEBUG] AUTO-REJECT_STALE: Row ${i + 1} Skip: Status "${cSt}" is final/manual/empty.`); continue; }
    if (lUD < sThD) {
      const currentPeakStatus = cR[PEAK_STATUS_COL - 1]?.toString().trim() || 'Not Set'; // Get peak status for logging
      Logger.log(`[INFO] AUTO-REJECT_STALE: Row ${i + 1} - Stale. LastUpd:${lUD.toLocaleDateString()}, OldStat:"${cSt}" -> New:"${REJECTED_STATUS}". Peak was: "${currentPeakStatus}"`);
      sV[i][STATUS_COL - 1] = REJECTED_STATUS; sV[i][LAST_UPDATE_DATE_COL - 1] = cD; sV[i][PROCESSED_TIMESTAMP_COL - 1] = cD; uAC++;
    }
  }
  if (uAC > 0) { Logger.log(`[INFO] AUTO-REJECT_STALE: Found ${uAC} stale apps. Writing...`); try { dR.setValues(sV); Logger.log(`[INFO] AUTO-REJECT_STALE: Updated ${uAC} stale apps.`); } catch (e) { Logger.log(`[ERROR] AUTO-REJECT_STALE: Write fail: ${e}\n${e.stack}`); } }
  else { Logger.log("[INFO] AUTO-REJECT_STALE: No stale apps found."); }
  const SETA = new Date(); Logger.log(`==== AUTO-REJECT STALE END (${SETA.toLocaleString()}) ==== Time: ${(SETA.getTime() - SSTA.getTime()) / 1000}s ====`);
}

// --- Part 5 of code blocks --- 

// --- Helper: Get or Create Dashboard Sheet and move to front ---
function getOrCreateDashboardSheet(spreadsheet) {
  let dashboardSheet = spreadsheet.getSheetByName(DASHBOARD_TAB_NAME);
  if (!dashboardSheet) {
    dashboardSheet = spreadsheet.insertSheet(DASHBOARD_TAB_NAME, 0); // Insert at the very first position
    Logger.log(`[INFO] SETUP_DASH: Created new dashboard sheet "${DASHBOARD_TAB_NAME}" at the first position.`);

    // Optional: Clean up default "Sheet1" if it exists and isn't the data or dashboard sheet
    // This part is similar to your existing logic for the data sheet.
    if (spreadsheet.getSheets().length > 2) { // Ensure we don't accidentally delete data/dashboard if only they exist
        const dataSheetName = SHEET_TAB_NAME; // from your global constants
        const defaultSheet = spreadsheet.getSheetByName('Sheet1');
        if (defaultSheet &&
            defaultSheet.getSheetId() !== spreadsheet.getSheetByName(dataSheetName).getSheetId() &&
            defaultSheet.getSheetId() !== dashboardSheet.getSheetId()) {
            try {
                spreadsheet.deleteSheet(defaultSheet);
                Logger.log(`[INFO] SETUP_DASH: Removed default 'Sheet1' after dashboard creation.`);
            } catch (eDeleteDefault) {
                Logger.log(`[WARN] SETUP_DASH: Failed to remove default 'Sheet1' after dashboard creation: ${eDeleteDefault.message}`);
            }
        }
    }
  } else {
    // If it exists, ensure it's the active sheet and then move it to the first position.
    spreadsheet.setActiveSheet(dashboardSheet);
    spreadsheet.moveActiveSheet(0); // Index 0 makes it the first sheet
    Logger.log(`[INFO] SETUP_DASH: Found existing dashboard sheet "${DASHBOARD_TAB_NAME}" and ensured it is the first tab.`);
  }
  return dashboardSheet;
}

// --- Helper: Format Dashboard Sheet (Initial Layout & Styling, plus Helper Sheet Formula Setup) ---
function formatDashboardSheet(dashboardSheet) {
  if (!dashboardSheet || typeof dashboardSheet.getName !== 'function') { 
    Logger.log(`[ERROR] FORMAT_DASH: Invalid dashboardSheet object provided. Cannot format.`); 
    return; 
  }
  Logger.log(`[INFO] FORMAT_DASH: Starting initial formatting for dashboard sheet "${dashboardSheet.getName()}".`);

  // Clear existing dashboard content and formatting
  dashboardSheet.clear(); 
  dashboardSheet.clearFormats(); 
  dashboardSheet.clearNotes();
  const conditionalFormatRules = dashboardSheet.getConditionalFormatRules();
  conditionalFormatRules.forEach(rule => {
    try { dashboardSheet.removeConditionalFormatRule(rule.getId()); }
    catch (e) { Logger.log(`[WARN] FORMAT_DASH: Could not remove a conditional format rule: ${e.message}`);}
  });
  
  try { 
    dashboardSheet.setHiddenGridlines(true); 
    Logger.log(`[INFO] FORMAT_DASH: Gridlines hidden for "${dashboardSheet.getName()}".`);
  } catch (e) { 
    Logger.log(`[ERROR] FORMAT_DASH: Error hiding gridlines: ${e.toString()}`); 
  }

  // --- Color Constants ---
  const TEAL_ACCENT_BG = "#26A69A", HEADER_TEXT_COLOR = "#FFFFFF", 
        LIGHT_GREY_BG = "#F5F5F5", DARK_GREY_TEXT = "#424242", 
        CARD_BORDER_COLOR = "#BDBDBD", VALUE_TEXT_COLOR = TEAL_ACCENT_BG, 
        METRIC_FONT_SIZE = 15, METRIC_FONT_WEIGHT = "bold", LABEL_FONT_WEIGHT = "bold";
  const SECONDARY_CARD_BG = "#FFFDE7", SECONDARY_VALUE_COLOR = "#FF8F00"; // Yellow for Manual Review
  const ORANGE_CARD_BG = "#FFF3E0", ORANGE_VALUE_COLOR = "#EF6C00";    // Orange for Direct Reject Rate
  
  const spacerColAWidth = 20; // Define consistent spacer width

  // --- Main Dashboard Title ---
  dashboardSheet.getRange("A1:M1").merge() 
                .setValue("Job Application Dashboard").setBackground(TEAL_ACCENT_BG)
                .setFontColor(HEADER_TEXT_COLOR).setFontSize(18).setFontWeight("bold")
                .setHorizontalAlignment("center").setVerticalAlignment("middle");
  dashboardSheet.setRowHeight(1, 45); 
  dashboardSheet.setRowHeight(2, 10); // Spacer

  // --- "Key Metrics Overview:" Sub-Title ---
  dashboardSheet.getRange("B3").setValue("Key Metrics Overview:")
                .setFontSize(14).setFontWeight(LABEL_FONT_WEIGHT).setFontColor(DARK_GREY_TEXT);
  dashboardSheet.setRowHeight(3, 30); 
  dashboardSheet.setRowHeight(4, 10); // Spacer

  // --- DEFINE Column Letters for Formulas ---
  const appSheetNameForFormula = `'${SHEET_TAB_NAME}'`; 
  const companyColLetter = columnToLetter(COMPANY_COL);     
  const jobTitleColLetter = columnToLetter(JOB_TITLE_COL);  
  const statusColLetter = columnToLetter(STATUS_COL);       
  const peakStatusColLetter = columnToLetter(PEAK_STATUS_COL); 
  const emailDateColLetter = columnToLetter(EMAIL_DATE_COL);
  const platformColLetter = columnToLetter(PLATFORM_COL); 

  // --- Scorecard Setup on Dashboard Sheet ---
  // Row 1 of Scorecards (Sheet Row 5)
  dashboardSheet.getRange("B5").setValue("Total Apps").setFontWeight(LABEL_FONT_WEIGHT).setFontColor(DARK_GREY_TEXT).setVerticalAlignment("middle");
  dashboardSheet.getRange("C5").setFormula(`=IFERROR(COUNTA(${appSheetNameForFormula}!${companyColLetter}2:${companyColLetter}), 0)`).setFontSize(METRIC_FONT_SIZE).setFontWeight(METRIC_FONT_WEIGHT).setFontColor(VALUE_TEXT_COLOR).setHorizontalAlignment("center").setVerticalAlignment("middle").setNumberFormat("0");
  dashboardSheet.getRange("E5").setValue("Peak Interviews").setFontWeight(LABEL_FONT_WEIGHT).setFontColor(DARK_GREY_TEXT).setVerticalAlignment("middle");
  dashboardSheet.getRange("F5").setFormula(`=IFERROR(COUNTIF(${appSheetNameForFormula}!${peakStatusColLetter}2:${peakStatusColLetter},"${INTERVIEW_STATUS}"), 0)`).setFontSize(METRIC_FONT_SIZE).setFontWeight(METRIC_FONT_WEIGHT).setFontColor(VALUE_TEXT_COLOR).setHorizontalAlignment("center").setVerticalAlignment("middle").setNumberFormat("0");
  dashboardSheet.getRange("H5").setValue("Interview Rate").setFontWeight(LABEL_FONT_WEIGHT).setFontColor(DARK_GREY_TEXT).setVerticalAlignment("middle");
  dashboardSheet.getRange("I5").setFormula(`=IFERROR(F5/C5, 0)`).setFontSize(METRIC_FONT_SIZE).setFontWeight(METRIC_FONT_WEIGHT).setFontColor(VALUE_TEXT_COLOR).setHorizontalAlignment("center").setVerticalAlignment("middle").setNumberFormat("0.00%");
  dashboardSheet.getRange("K5").setValue("Offer Rate").setFontWeight(LABEL_FONT_WEIGHT).setFontColor(DARK_GREY_TEXT).setVerticalAlignment("middle");
  dashboardSheet.getRange("L5").setFormula(`=IFERROR(F7/C5, 0)`).setFontSize(METRIC_FONT_SIZE).setFontWeight(METRIC_FONT_WEIGHT).setFontColor(VALUE_TEXT_COLOR).setHorizontalAlignment("center").setVerticalAlignment("middle").setNumberFormat("0.00%");
  dashboardSheet.setRowHeight(5, 40); 
  dashboardSheet.setRowHeight(6, 10); 

  // Row 2 of Scorecards (Sheet Row 7)
  dashboardSheet.getRange("B7").setValue("Active Apps").setFontWeight(LABEL_FONT_WEIGHT).setFontColor(DARK_GREY_TEXT).setVerticalAlignment("middle");
  let activeAppsFormula = `=IFERROR(COUNTIFS(${appSheetNameForFormula}!${statusColLetter}2:${statusColLetter}, "<>"&"", ${appSheetNameForFormula}!${statusColLetter}2:${statusColLetter}, "<>${REJECTED_STATUS}", ${appSheetNameForFormula}!${statusColLetter}2:${statusColLetter}, "<>${ACCEPTED_STATUS}"), 0)`;
  dashboardSheet.getRange("C7").setFormula(activeAppsFormula).setFontSize(METRIC_FONT_SIZE).setFontWeight(METRIC_FONT_WEIGHT).setFontColor(VALUE_TEXT_COLOR).setHorizontalAlignment("center").setVerticalAlignment("middle").setNumberFormat("0");
  dashboardSheet.getRange("E7").setValue("Peak Offers").setFontWeight(LABEL_FONT_WEIGHT).setFontColor(DARK_GREY_TEXT).setVerticalAlignment("middle");
  dashboardSheet.getRange("F7").setFormula(`=IFERROR(COUNTIF(${appSheetNameForFormula}!${peakStatusColLetter}2:${peakStatusColLetter},"${OFFER_STATUS}"), 0)`).setFontSize(METRIC_FONT_SIZE).setFontWeight(METRIC_FONT_WEIGHT).setFontColor(VALUE_TEXT_COLOR).setHorizontalAlignment("center").setVerticalAlignment("middle").setNumberFormat("0"); 
  dashboardSheet.getRange("H7").setValue("Current Interviews").setFontWeight(LABEL_FONT_WEIGHT).setFontColor(DARK_GREY_TEXT).setVerticalAlignment("middle");
  dashboardSheet.getRange("I7").setFormula(`=IFERROR(COUNTIF(${appSheetNameForFormula}!${statusColLetter}2:${statusColLetter},"${INTERVIEW_STATUS}"), 0)`).setFontSize(METRIC_FONT_SIZE).setFontWeight(METRIC_FONT_WEIGHT).setFontColor(VALUE_TEXT_COLOR).setHorizontalAlignment("center").setVerticalAlignment("middle").setNumberFormat("0");
  dashboardSheet.getRange("K7").setValue("Current Assessments").setFontWeight(LABEL_FONT_WEIGHT).setFontColor(DARK_GREY_TEXT).setVerticalAlignment("middle");
  dashboardSheet.getRange("L7").setFormula(`=IFERROR(COUNTIF(${appSheetNameForFormula}!${statusColLetter}2:${statusColLetter},"${ASSESSMENT_STATUS}"), 0)`).setFontSize(METRIC_FONT_SIZE).setFontWeight(METRIC_FONT_WEIGHT).setFontColor(VALUE_TEXT_COLOR).setHorizontalAlignment("center").setVerticalAlignment("middle").setNumberFormat("0");
  dashboardSheet.setRowHeight(7, 40); 
  dashboardSheet.setRowHeight(8, 10); 

  // Row 3 of Scorecards (Sheet Row 9)
  dashboardSheet.getRange("B9").setValue("Total Rejections").setFontWeight(LABEL_FONT_WEIGHT).setFontColor(DARK_GREY_TEXT).setVerticalAlignment("middle");
  dashboardSheet.getRange("C9").setFormula(`=IFERROR(COUNTIF(${appSheetNameForFormula}!${statusColLetter}2:${statusColLetter},"${REJECTED_STATUS}"), 0)`).setFontSize(METRIC_FONT_SIZE).setFontWeight(METRIC_FONT_WEIGHT).setFontColor(VALUE_TEXT_COLOR).setHorizontalAlignment("center").setVerticalAlignment("middle").setNumberFormat("0");
  dashboardSheet.getRange("E9").setValue("Apps Viewed (Peak)").setFontWeight(LABEL_FONT_WEIGHT).setFontColor(DARK_GREY_TEXT).setVerticalAlignment("middle");
  dashboardSheet.getRange("F9").setFormula(`=IFERROR(COUNTIF(${appSheetNameForFormula}!${peakStatusColLetter}2:${peakStatusColLetter},"${APPLICATION_VIEWED_STATUS}"),0)`).setFontSize(METRIC_FONT_SIZE).setFontWeight(METRIC_FONT_WEIGHT).setFontColor(VALUE_TEXT_COLOR).setHorizontalAlignment("center").setVerticalAlignment("middle").setNumberFormat("0");
  dashboardSheet.getRange("H9").setValue("Manual Review").setFontWeight(LABEL_FONT_WEIGHT).setFontColor(DARK_GREY_TEXT).setVerticalAlignment("middle");
  const compColManual = `${appSheetNameForFormula}!${companyColLetter}2:${companyColLetter}="${MANUAL_REVIEW_NEEDED}"`;
  const titleColManual = `${appSheetNameForFormula}!${jobTitleColLetter}2:${jobTitleColLetter}="${MANUAL_REVIEW_NEEDED}"`;
  const statusColManualForReview = `${appSheetNameForFormula}!${statusColLetter}2:${statusColLetter}="${MANUAL_REVIEW_NEEDED}"`;
  const finalManualReviewFormula = `=IFERROR(SUM(ARRAYFORMULA(SIGN((${compColManual})+(${titleColManual})+(${statusColManualForReview})))),0)`;
  dashboardSheet.getRange("I9").setFormula(finalManualReviewFormula).setFontSize(METRIC_FONT_SIZE).setFontWeight(METRIC_FONT_WEIGHT).setFontColor(SECONDARY_VALUE_COLOR).setHorizontalAlignment("center").setVerticalAlignment("middle").setNumberFormat("0");
  dashboardSheet.getRange("K9").setValue("Direct Reject Rate").setFontWeight(LABEL_FONT_WEIGHT).setFontColor(DARK_GREY_TEXT).setVerticalAlignment("middle");
  const directRejectFormula = `=IFERROR(COUNTIFS(${appSheetNameForFormula}!${peakStatusColLetter}2:${peakStatusColLetter},"${DEFAULT_STATUS}",${appSheetNameForFormula}!${statusColLetter}2:${statusColLetter},"${REJECTED_STATUS}")/C5, 0)`;
  dashboardSheet.getRange("L9").setFormula(directRejectFormula).setFontSize(METRIC_FONT_SIZE).setFontWeight(METRIC_FONT_WEIGHT).setFontColor(ORANGE_VALUE_COLOR).setHorizontalAlignment("center").setVerticalAlignment("middle").setNumberFormat("0.00%");
  dashboardSheet.setRowHeight(9, 40); 
  dashboardSheet.setRowHeight(10, 15); 

  // --- Apply Card Styling (Corrected for Borders & Colors) ---
  const scorecardRangesToStyle = [
      "B5:C5", "E5:F5", "H5:I5", "K5:L5", 
      "B7:C7", "E7:F7", "H7:I7", "K7:L7", 
      "B9:C9", "E9:F9", "H9:I9", "K9:L9"  
  ];
  scorecardRangesToStyle.forEach(rangeString => {
    const range = dashboardSheet.getRange(rangeString);
    range.setBackground(LIGHT_GREY_BG) 
         .setBorder(true, true, true, true, true, true, CARD_BORDER_COLOR, SpreadsheetApp.BorderStyle.SOLID_THIN);
  });
  dashboardSheet.getRange("H9:I9").setBackground(SECONDARY_CARD_BG); 
  dashboardSheet.getRange("K9:L9").setBackground(ORANGE_CARD_BG);    
  const primaryValueCells = ["C5", "F5", "I5", "L5", "C7", "F7", "I7", "L7", "C9", "F9"];
  primaryValueCells.forEach(cellAddress => {
    dashboardSheet.getRange(cellAddress).setFontColor(VALUE_TEXT_COLOR);
  });
  dashboardSheet.getRange("I9").setFontColor(SECONDARY_VALUE_COLOR); // Value for Manual Review
  dashboardSheet.getRange("L9").setFontColor(ORANGE_VALUE_COLOR);    // Value for Direct Reject
  const labelCellAddresses = ["B5", "E5", "H5", "K5", "B7", "E7", "H7", "K7", "B9", "E9", "H9", "K9"];
  labelCellAddresses.forEach(cellAddress => { // Ensure label colors are correct
      dashboardSheet.getRange(cellAddress).setFontColor(DARK_GREY_TEXT);
  });
  // --- End of Apply Card Styling (Corrected) ---

  // --- Chart Section Titles ---
  const chartSectionTitleRow1 = 11;
  dashboardSheet.getRange("B" + chartSectionTitleRow1).setValue("Platform & Weekly Trends").setFontSize(12).setFontWeight(LABEL_FONT_WEIGHT).setFontColor(DARK_GREY_TEXT);
  dashboardSheet.setRowHeight(chartSectionTitleRow1, 25);
  dashboardSheet.setRowHeight(chartSectionTitleRow1 + 1, 5); 

  const chartSectionTitleRow2 = 28; 
  dashboardSheet.getRange("B" + chartSectionTitleRow2).setValue("Application Funnel Analysis").setFontSize(12).setFontWeight(LABEL_FONT_WEIGHT).setFontColor(DARK_GREY_TEXT);
  dashboardSheet.setRowHeight(chartSectionTitleRow2, 25);
  dashboardSheet.setRowHeight(chartSectionTitleRow2 + 1, 5); 

  // --- Column Widths ---
  const labelWidth = 150; const valueWidth = 75; const spacerS = 15; 
  dashboardSheet.setColumnWidth(1, spacerColAWidth); 
  dashboardSheet.setColumnWidth(2, labelWidth); dashboardSheet.setColumnWidth(3, valueWidth); 
  dashboardSheet.setColumnWidth(4, spacerS); 
  dashboardSheet.setColumnWidth(5, labelWidth); dashboardSheet.setColumnWidth(6, valueWidth); 
  dashboardSheet.setColumnWidth(7, spacerS); 
  dashboardSheet.setColumnWidth(8, labelWidth); dashboardSheet.setColumnWidth(9, valueWidth); 
  dashboardSheet.setColumnWidth(10, spacerS); 
  dashboardSheet.setColumnWidth(11, labelWidth); dashboardSheet.setColumnWidth(12, valueWidth); 
  dashboardSheet.setColumnWidth(13, spacerColAWidth);

  // ---- START: Setup Formulas in Helper Sheet ----
  const ss = dashboardSheet.getParent();
  let helperSheet = ss.getSheetByName(HELPER_SHEET_NAME);
  if (!helperSheet) {
    helperSheet = getOrCreateHelperSheet(ss); 
    if (!helperSheet) {
        Logger.log(`[ERROR] FORMAT_DASH_HELPER: Helper sheet "${HELPER_SHEET_NAME}" missing and could not be created. Cannot set formulas.`);
        return; 
    }
  }
  if (helperSheet) { 
    Logger.log(`[INFO] FORMAT_DASH_HELPER: Setting up formulas in helper sheet "${helperSheet.getName()}".`);
    helperSheet.getRange("A1:B").clearContent(); 
    helperSheet.getRange("D1:E").clearContent(); 
    helperSheet.getRange("G1:H").clearContent();
    helperSheet.getRange("J1:K").clearContent(); 
    Logger.log(`[DEBUG] FORMAT_DASH_HELPER: Cleared A:B, D:E, G:H, J:K in helper sheet.`);
    
    helperSheet.getRange("A1").setValue("Platform");
    helperSheet.getRange("B1").setValue("Count");
    const platformQueryFormula = `=IFERROR(QUERY(${appSheetNameForFormula}!${platformColLetter}2:${platformColLetter}, "SELECT ${platformColLetter}, COUNT(${platformColLetter}) WHERE ${platformColLetter} IS NOT NULL AND ${platformColLetter} <> '' GROUP BY ${platformColLetter} ORDER BY COUNT(${platformColLetter}) DESC LABEL ${platformColLetter} '', COUNT(${platformColLetter}) ''", 0), {"No Platforms",0})`;
    helperSheet.getRange("A2").setFormula(platformQueryFormula);
    Logger.log(`[INFO] HELPER_FORMULA: Platform formula set in A2.`);

    helperSheet.getRange("J1").setValue("RAW_VALID_DATES_FOR_WEEKLY");
    const rawDatesFormula = `=IFERROR(FILTER(${appSheetNameForFormula}!${emailDateColLetter}2:${emailDateColLetter}, ISNUMBER(${appSheetNameForFormula}!${emailDateColLetter}2:${emailDateColLetter})), "")`;
    helperSheet.getRange("J2").setFormula(rawDatesFormula);
    helperSheet.getRange("J2:J").setNumberFormat("yyyy-mm-dd hh:mm:ss");
    Logger.log(`[INFO] HELPER_FORMULA: Raw Valid Dates formula set in J2.`);

    helperSheet.getRange("K1").setValue("CALCULATED_WEEK_STARTS");
    // CHOOSE ONE weekStartCalcFormula (Monday or Sunday start)
    const weekStartCalcFormula = `=ARRAYFORMULA(IF(ISBLANK(J2:J), "", DATE(YEAR(J2:J), MONTH(J2:J), DAY(J2:J) - WEEKDAY(J2:J, 2) + 1)))`; // Monday Start
    // const weekStartCalcFormula = `=ARRAYFORMULA(IF(ISBLANK(J2:J), "", DATE(YEAR(J2:J), MONTH(J2:J), DAY(J2:J) - WEEKDAY(J2:J, 1) + 1)))`; // Sunday Start
    helperSheet.getRange("K2").setFormula(weekStartCalcFormula);
    helperSheet.getRange("K2:K").setNumberFormat("yyyy-mm-dd");
    Logger.log(`[INFO] HELPER_FORMULA: Calculated Week Starts formula set in K2.`);

    helperSheet.getRange("D1").setValue("Week Starting");
    const uniqueWeeksFormula = `=IFERROR(SORT(UNIQUE(FILTER(K2:K, K2:K<>""))), {"No Data"})`;
    helperSheet.getRange("D2").setFormula(uniqueWeeksFormula);
    helperSheet.getRange("D2:D").setNumberFormat("yyyy-mm-dd");
    Logger.log(`[INFO] HELPER_FORMULA: Unique Weeks formula set in D2.`);

    helperSheet.getRange("E1").setValue("Applications");
    const weeklyCountsFormula = `=ARRAYFORMULA(IF(D2:D="", "", COUNTIF(K2:K, D2:D)))`;
    helperSheet.getRange("E2").setFormula(weeklyCountsFormula);
    helperSheet.getRange("E2:E").setNumberFormat("0");
    Logger.log(`[INFO] HELPER_FORMULA: Weekly Counts formula set in E2.`);
    
    helperSheet.getRange("G1").setValue("Stage"); helperSheet.getRange("H1").setValue("Count");
    const funnelStagesValues = [DEFAULT_STATUS, APPLICATION_VIEWED_STATUS, ASSESSMENT_STATUS, INTERVIEW_STATUS, OFFER_STATUS];
    helperSheet.getRange(2, 7, funnelStagesValues.length, 1).setValues(funnelStagesValues.map(stage => [stage]));
    helperSheet.getRange("H2").setFormula(`=IFERROR(COUNTA(${appSheetNameForFormula}!${companyColLetter}2:${companyColLetter}),0)`); 
    for (let i = 1; i < funnelStagesValues.length; i++) { 
      helperSheet.getRange(i + 2, 8).setFormula(`=IFERROR(COUNTIF(${appSheetNameForFormula}!${peakStatusColLetter}2:${peakStatusColLetter}, G${i + 2}),0)`);
    }
    Logger.log(`[INFO] HELPER_FORMULA: All helper formulas set.`);
  }
  // ---- END: Setup Formulas in Helper Sheet ----

  // --- Hide Unused Columns and Rows on Dashboard Sheet---
  const lastUsedDataColumn = 13; 
  const maxCols = dashboardSheet.getMaxColumns();
  if (maxCols > 0) { dashboardSheet.showColumns(1, maxCols); }
  if (maxCols > lastUsedDataColumn) {
      dashboardSheet.hideColumns(lastUsedDataColumn + 1, maxCols - lastUsedDataColumn);
  }

  const lastUsedDataRow = 45; 
  const maxRows = dashboardSheet.getMaxRows();
  if (maxRows > 1) { dashboardSheet.showRows(1, maxRows); }
  if (maxRows > lastUsedDataRow) {
      dashboardSheet.hideRows(lastUsedDataRow + 1, maxRows - lastUsedDataRow);
  }
  Logger.log(`[INFO] FORMAT_DASH: Formatting concluded for Dashboard. Visible cols to ${columnToLetter(lastUsedDataColumn)}, rows to ${lastUsedDataRow}.`);
}

// --- Dashboard: Update Metrics and Chart Data (Helper Sheet is Formula-Driven) ---
function updateDashboardMetrics() {
  const SCRIPT_START_TIME_DASH = new Date();
  Logger.log(`\n==== STARTING DASHBOARD METRICS UPDATE (Helper Sheet is Formula-Driven) (${SCRIPT_START_TIME_DASH.toLocaleString()}) ====`);

  const { spreadsheet: ss } = getOrCreateSpreadsheetAndSheet();
  if (!ss) { 
    Logger.log(`[ERROR] UPDATE_DASH: Could not get spreadsheet. Aborting.`);
    return; 
  }

  const dashboardSheet = ss.getSheetByName(DASHBOARD_TAB_NAME); 
  const helperSheet = ss.getSheetByName(HELPER_SHEET_NAME); 

  if (!dashboardSheet) {
      Logger.log(`[WARN] UPDATE_DASH: Dashboard sheet missing. Cannot create/update charts.`);
      // If no dashboard sheet, can't do chart updates
      const SCRIPT_END_TIME_DASH_NO_DASH = new Date();
      Logger.log(`\n==== DASHBOARD METRICS UPDATE FINISHED (No Dashboard Sheet) (${SCRIPT_END_TIME_DASH_NO_DASH.toLocaleString()}) ====`);
      return;
  }
  if (!helperSheet) {
    Logger.log(`[ERROR] UPDATE_DASH: Helper sheet missing. Cannot verify chart data sources or create/update charts.`);
    // If no helper sheet, chart data sources are invalid
    const SCRIPT_END_TIME_DASH_NO_HELP = new Date();
    Logger.log(`\n==== DASHBOARD METRICS UPDATE FINISHED (No Helper Sheet) (${SCRIPT_END_TIME_DASH_NO_HELP.toLocaleString()}) ====`);
    return; 
  }

  Logger.log(`[INFO] UPDATE_DASH: All scorecard metrics AND chart helper data are formula-based.`);
  
  // This function's primary role now is to ensure the chart OBJECTS exist on the dashboard
  // and are correctly configured to point to the formula-driven helper sheet data.
  // The actual data aggregation happens via formulas in the helper sheet itself.

  // --- Call Chart Update Functions ---
  // These functions will check if charts exist. If not, they create them using data
  // from helperSheet (which is now formula-driven and should be up-to-date).
  // If charts exist, they mainly ensure ranges are still correct (though this often
  // isn't strictly needed if ranges are static like 'HelperSheet!A1:B10').
  if (dashboardSheet && helperSheet) { // Redundant check, but safe
     Logger.log(`[INFO] UPDATE_DASH: Ensuring chart objects are present and configured...`);
     try {
        updatePlatformDistributionChart(dashboardSheet, helperSheet);
        updateApplicationsOverTimeChart(dashboardSheet, helperSheet);
        updateApplicationFunnelChart(dashboardSheet, helperSheet);
        Logger.log(`[INFO] UPDATE_DASH: Chart object presence and configuration check complete.`);
     } catch (e) {
        Logger.log(`[ERROR] UPDATE_DASH: Error during chart update/creation calls: ${e.toString()} \nStack: ${e.stack}`);
     }
  } else {
      // This case should ideally be caught by earlier checks for dashboardSheet and helperSheet
      Logger.log(`[WARN] UPDATE_DASH: Skipping chart object updates as dashboardSheet or helperSheet is unexpectedly missing at this stage.`);
  }

  const SCRIPT_END_TIME_DASH = new Date();
  Logger.log(`\n==== DASHBOARD METRICS UPDATE FINISHED (${SCRIPT_END_TIME_DASH.toLocaleString()}) === Total Time: ${(SCRIPT_END_TIME_DASH.getTime() - SCRIPT_START_TIME_DASH.getTime())/1000}s ====`);
}

// --- Helper: Get or Create Helper Sheet ---
function getOrCreateHelperSheet(spreadsheet) {
  let helperSheet = spreadsheet.getSheetByName(HELPER_SHEET_NAME);
  if (!helperSheet) {
    helperSheet = spreadsheet.insertSheet(HELPER_SHEET_NAME);
    Logger.log(`[INFO] SETUP_HELPER: Created new helper sheet "${HELPER_SHEET_NAME}".`);
    try {
      helperSheet.hideSheet(); // Hide it by default
      Logger.log(`[INFO] SETUP_HELPER: Sheet "${HELPER_SHEET_NAME}" has been hidden.`);
    } catch (e) {
      Logger.log(`[WARN] SETUP_HELPER: Could not hide helper sheet "${HELPER_SHEET_NAME}": ${e}`);
    }
  } else {
    Logger.log(`[INFO] SETUP_HELPER: Found existing helper sheet "${HELPER_SHEET_NAME}". Ensuring it's hidden if not already.`);
    if (!helperSheet.isSheetHidden()) {
        try {
            helperSheet.hideSheet();
            Logger.log(`[INFO] SETUP_HELPER: Sheet "${HELPER_SHEET_NAME}" was visible and has now been hidden.`);
        } catch (e) {
            Logger.log(`[WARN] SETUP_HELPER: Could not hide existing helper sheet "${HELPER_SHEET_NAME}": ${e}`);
        }
    }
  }
  return helperSheet;
}

// --- Dashboard Chart: Update Platform Distribution Pie Chart ---
function updatePlatformDistributionChart(dashboardSheet, helperSheet) {
  Logger.log(`[INFO] CHART_PLATFORM: Attempting to create/update Platform Distribution chart.`);
  const CHART_TITLE = "Platform Distribution";
  let existingChart = null;

  const charts = dashboardSheet.getCharts();
  for (let i = 0; i < charts.length; i++) {
    if (charts[i].getOptions().get('title') === CHART_TITLE && charts[i].getContainerInfo().getAnchorColumn() === 2) {
      existingChart = charts[i];
      break;
    }
  }
  if(existingChart) Logger.log(`[DEBUG] CHART_PLATFORM: Found existing chart.`);
  else Logger.log(`[DEBUG] CHART_PLATFORM: No existing chart with title '${CHART_TITLE}'. Will create new.`);

  // This check is CRITICAL. Number of rows with actual data in column A of helper sheet.
  const lastPlatformRow = helperSheet.getRange("A1:A").getValues().filter(String).length; 
  let dataRange;

  // We need at least a header AND one row of data (lastPlatformRow >= 2) for a chart
  if (helperSheet.getRange("A1").getValue() === "Platform" && lastPlatformRow >= 2) {
      dataRange = helperSheet.getRange(`A1:B${lastPlatformRow}`);
      Logger.log(`[INFO] CHART_PLATFORM: Data range for chart set to ${HELPER_SHEET_NAME}!A1:B${lastPlatformRow}`);
  } else {
      Logger.log(`[WARN] CHART_PLATFORM: Not enough data or invalid header for platform chart (found ${lastPlatformRow} rows with content in A). Chart will not be created/updated.`);
      if (existingChart) { 
          try { dashboardSheet.removeChart(existingChart); Logger.log(`[INFO] CHART_PLATFORM: Removed existing chart due to insufficient data.`);}
          catch (e) { Logger.log(`[ERROR] CHART_PLATFORM: Could not remove chart: ${e}`); }
      }
      return; // EXIT HERE if no valid data
  }

  const chartSectionTitleRow1 = 11; 
  const anchorRow = chartSectionTitleRow1 + 2; // Should be 13
  const anchorCol = 2;  // Column B
  const chartWidth = 460; 
  const chartHeight = 280; 

    const optionsToSet = {
    title: CHART_TITLE,
    legend: { position: Charts.Position.RIGHT },
    pieHole: 0.4,
    width: chartWidth,
    height: chartHeight,
    sliceVisibilityThreshold: 0 // Add this line
  };

  try { // Wrap chart operations in try-catch
    if (existingChart) {
      Logger.log(`[DEBUG] CHART_PLATFORM: Modifying existing chart.`);
      let chartBuilder = existingChart.modify();
      chartBuilder = chartBuilder.clearRanges().addRange(dataRange).setChartType(Charts.ChartType.PIE);
      for (const key in optionsToSet) { if (optionsToSet.hasOwnProperty(key)) chartBuilder = chartBuilder.setOption(key, optionsToSet[key]); }
      chartBuilder = chartBuilder.setPosition(anchorRow, anchorCol, 0, 0); 
      dashboardSheet.updateChart(chartBuilder.build());
      Logger.log(`[INFO] CHART_PLATFORM: Updated existing chart "${CHART_TITLE}".`);
    } else { 
      Logger.log(`[DEBUG] CHART_PLATFORM: Creating new chart.`);
      let newChartBuilder = dashboardSheet.newChart().setChartType(Charts.ChartType.PIE);
      for (const key in optionsToSet) { if (optionsToSet.hasOwnProperty(key)) newChartBuilder = newChartBuilder.setOption(key, optionsToSet[key]); }
      newChartBuilder = newChartBuilder.addRange(dataRange).setPosition(anchorRow, anchorCol, 0, 0);
      dashboardSheet.insertChart(newChartBuilder.build());
      Logger.log(`[INFO] CHART_PLATFORM: Created new chart "${CHART_TITLE}".`);
    }
  } catch (e) {
    Logger.log(`[ERROR] CHART_PLATFORM: Failed during chart build/insert/update: ${e.message} ${e.stack}`);
  }
}

// --- Helper: Get Week Start Date (Monday) ---
function getWeekStartDate(inputDate) {
  const date = new Date(inputDate.getTime()); // Clone the date to avoid modifying the original
  const day = date.getDay(); // Sunday - 0, Monday - 1, ..., Saturday - 6
  const diff = date.getDate() - day + (day === 0 ? -6 : 1); // Adjust when day is Sunday
  return new Date(date.setDate(diff));
}

// --- Dashboard Chart: Update Applications Over Time (Weekly) Line Chart ---
function updateApplicationsOverTimeChart(dashboardSheet, helperSheet) {
  Logger.log(`[INFO] CHART_APPS_TIME: Attempting to create/update Applications Over Time chart.`);
  const CHART_TITLE = "Applications Over Time (Weekly)";
  let existingChart = null;

  const charts = dashboardSheet.getCharts();
  for (let i = 0; i < charts.length; i++) {
    // Check by title and a known anchor column (e.g., Col H where it's expected to start)
    if (charts[i].getOptions().get('title') === CHART_TITLE && charts[i].getContainerInfo().getAnchorColumn() === 8) { 
      existingChart = charts[i];
      break;
    }
  }
  if(existingChart) Logger.log(`[DEBUG] CHART_APPS_TIME: Found existing chart.`);
  else Logger.log(`[DEBUG] CHART_APPS_TIME: No existing chart with title '${CHART_TITLE}'. Will create new.`);

  // Determine the actual last row of data in the helper sheet for this chart (D:E)
  // Counts non-empty cells in column D (Week Starting)
  const lastWeeklyDataRowInHelper = helperSheet.getRange("D1:D").getValues().filter(String).length; 
  let dataRange;

  // Check if D1 has the correct header AND if there is at least one data row (header is row 1, so need >=2 for data)
  if (helperSheet.getRange("D1").getValue() === "Week Starting" && lastWeeklyDataRowInHelper >= 2) {
      dataRange = helperSheet.getRange(`D1:E${lastWeeklyDataRowInHelper}`); // e.g., D1:E5 if header + 4 data weeks
      Logger.log(`[INFO] CHART_APPS_TIME: Data range for chart set to ${helperSheet.getName()}!D1:E${lastWeeklyDataRowInHelper}`);
  } else {
      Logger.log(`[WARN] CHART_APPS_TIME: Not enough data or invalid header for weekly chart (Helper Col D has ${lastWeeklyDataRowInHelper} content rows including header). Chart will not be created/updated.`);
      if (existingChart) { 
        try { dashboardSheet.removeChart(existingChart); Logger.log(`[INFO] CHART_APPS_TIME: Removed existing chart due to insufficient/invalid data in helper sheet.`); }
        catch (e) { Logger.log(`[ERROR] CHART_APPS_TIME: Could not remove chart: ${e}`); }
      }
      return; // EXIT HERE if no valid data to plot
  }

  // Anchor Row values should align with your formatDashboardSheet
  const chartSectionTitleRow1 = 11; // As set in your latest formatDashboardSheet
  const anchorRow = chartSectionTitleRow1 + 2; // Should be 13
  const anchorCol = 8;  // Column H

  const chartWidth = 460; // As determined previously
  const chartHeight = 280; 

  const optionsToSet = {
    title: CHART_TITLE,
    hAxis: { title: 'Week Starting', textStyle: { fontSize: 10 }, format: 'M/d' }, // Short date format for axis
    vAxis: { title: 'Number of Applications', textStyle: { fontSize: 10 }, viewWindow: { min: 0 } },
    legend: { position: 'none' }, 
    colors: ['#26A69A'], 
    width: chartWidth,
    height: chartHeight,
    // pointSize: 5, // Optionally add points to the line chart
    // curveType: 'function' // For a smoothed line, if desired
  };

  try { 
    if (existingChart) { 
      Logger.log(`[DEBUG] CHART_APPS_TIME: Modifying existing chart.`);
      let chartBuilder = existingChart.modify();
      chartBuilder = chartBuilder.clearRanges().addRange(dataRange).setChartType(Charts.ChartType.LINE);
      for (const key in optionsToSet) { if (optionsToSet.hasOwnProperty(key)) chartBuilder = chartBuilder.setOption(key, optionsToSet[key]); }
      chartBuilder = chartBuilder.setPosition(anchorRow, anchorCol, 0, 0);
      dashboardSheet.updateChart(chartBuilder.build());
      Logger.log(`[INFO] CHART_APPS_TIME: Updated existing chart "${CHART_TITLE}".`);
    } else { 
      Logger.log(`[DEBUG] CHART_APPS_TIME: Creating new chart.`);
      let newChartBuilder = dashboardSheet.newChart().setChartType(Charts.ChartType.LINE);
      for (const key in optionsToSet) { if (optionsToSet.hasOwnProperty(key)) newChartBuilder = newChartBuilder.setOption(key, optionsToSet[key]); }
      newChartBuilder = newChartBuilder.addRange(dataRange).setPosition(anchorRow, anchorCol, 0, 0);
      dashboardSheet.insertChart(newChartBuilder.build());
      Logger.log(`[INFO] CHART_APPS_TIME: Created new chart "${CHART_TITLE}".`);
    }
  } catch (e) {
      Logger.log(`[ERROR] CHART_APPS_TIME: Failed during chart build/insert/update: ${e.message} \nStack: ${e.stack}`);
  }
}

// --- Dashboard Chart: Update Application Funnel (Peak Stages) Column Chart ---
function updateApplicationFunnelChart(dashboardSheet, helperSheet) {
  Logger.log(`[INFO] CHART_FUNNEL: Attempting to create/update Application Funnel chart (using individual setOption).`);
  const CHART_TITLE = "Application Funnel (Peak Stages)";
  let existingChart = null;

  const charts = dashboardSheet.getCharts();
  for (let i = 0; i < charts.length; i++) {
    if (charts[i].getOptions().get('title') === CHART_TITLE && charts[i].getContainerInfo().getAnchorColumn() === 2) {
      existingChart = charts[i];
      break;
    }
  }
  if(existingChart) Logger.log(`[DEBUG] CHART_FUNNEL: Found existing chart.`);
  else Logger.log(`[DEBUG] CHART_FUNNEL: No existing chart found with title '${CHART_TITLE}'. Will create new.`);

  const lastFunnelDataRow = helperSheet.getRange("G:G").getValues().filter(String).length;
  let dataRange;
  if (helperSheet.getRange("G1").getValue() === "Stage" && lastFunnelDataRow >= 2) {
      dataRange = helperSheet.getRange(`G1:H${lastFunnelDataRow}`);
      Logger.log(`[INFO] CHART_FUNNEL: Data range for chart set to ${HELPER_SHEET_NAME}!G1:H${lastFunnelDataRow}`);
  } else {
      Logger.log(`[WARN] CHART_FUNNEL: No data/invalid header for funnel chart. Chart will not be created/updated.`);
      if (existingChart) { dashboardSheet.removeChart(existingChart); Logger.log(`[INFO] CHART_FUNNEL: Removed chart - no data.`);}
      return;
  }

  const chartSectionTitleRow2 = 28; 
  const anchorRow = chartSectionTitleRow2 + 2; // Should be 30
  const anchorCol = 2;  // Column B
  const chartWidth = 460; 
  const chartHeight = 280; 

  const optionsToSet = {
    title: CHART_TITLE,
    hAxis: { title: 'Application Stage', textStyle: { fontSize: 10 }, slantedText: true, slantedTextAngle: 30 },
    vAxis: { title: 'Number of Applications', textStyle: { fontSize: 10 }, viewWindow: { min: 0 } },
    legend: { position: 'none' }, // Or Charts.Position.NONE
    colors: ['#26A69A'], 
    bar: { groupWidth: '60%' }, 
    width: chartWidth,
    height: chartHeight,
  };
  Logger.log(`[DEBUG] CHART_FUNNEL: Options object for chart: ${JSON.stringify(optionsToSet)}`);

  if (existingChart) { 
    Logger.log(`[DEBUG] CHART_FUNNEL: Modifying existing chart.`);
    let chartBuilder = existingChart.modify();
    chartBuilder = chartBuilder.clearRanges();
    chartBuilder = chartBuilder.addRange(dataRange); // For modify, addRange can come before or after some setOptions
    chartBuilder = chartBuilder.setChartType(Charts.ChartType.COLUMN);
    
    // Loop to set options individually for modify
    for (const key in optionsToSet) {
      if (optionsToSet.hasOwnProperty(key)) {
        chartBuilder = chartBuilder.setOption(key, optionsToSet[key]);
      }
    }
    
    chartBuilder = chartBuilder.setPosition(anchorRow, anchorCol, 0, 0);
    const updatedChart = chartBuilder.build();
    dashboardSheet.updateChart(updatedChart);
    Logger.log(`[INFO] CHART_FUNNEL: Updated existing chart "${CHART_TITLE}".`);
  } else { // Creating a new chart
    Logger.log(`[DEBUG] CHART_FUNNEL: Creating new chart using individual setOption calls.`);
    let newChartBuilder = dashboardSheet.newChart();
    newChartBuilder = newChartBuilder.setChartType(Charts.ChartType.COLUMN);
    
    // Apply options individually
    for (const key in optionsToSet) {
      if (optionsToSet.hasOwnProperty(key)) {
        Logger.log(`[DEBUG] CHART_FUNNEL (New): Setting option ${key} = ${JSON.stringify(optionsToSet[key])}`);
        newChartBuilder = newChartBuilder.setOption(key, optionsToSet[key]);
      }
    }
    
    Logger.log(`[DEBUG] CHART_FUNNEL (New): After all individual setOption, type: ${typeof newChartBuilder}, has addRange: ${typeof newChartBuilder.addRange}`);
    newChartBuilder = newChartBuilder.addRange(dataRange);
    Logger.log(`[DEBUG] CHART_FUNNEL (New): After addRange, type: ${typeof newChartBuilder}, has setPosition: ${typeof newChartBuilder.setPosition}`);
    newChartBuilder = newChartBuilder.setPosition(anchorRow, anchorCol, 0, 0);
    const newChart = newChartBuilder.build();
    dashboardSheet.insertChart(newChart);
    Logger.log(`[INFO] CHART_FUNNEL: Created new chart "${CHART_TITLE}" using individual setOptions.`);
  }
}
