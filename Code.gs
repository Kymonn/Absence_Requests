// --- Configuration ---
/**
 * @fileoverview Absence Request Processing Script
 *
 * Handles staff absence requests submitted via a Google Form, manages approvals,
 * sends notifications, and logs absences to a Google Sheet and Calendar.
 *
 * MANDATORY: Replace placeholder values below with your actual school details.
 * IMPORTANT: Ensure the Header Names match EXACTLY (case-sensitive) the
 *            first row of your respective Google Sheets.
 */

// --- General Configuration ---
const HEADTEACHER_EMAIL        = "kmontiel.staff@sidneystringeracademy.org.uk"; // REPLACE: Headteacher's actual email address.
const ADMIN_EMAIL_FOR_ERRORS   = "kmontiel.staff@sidneystringeracademy.org.uk"; // Optional: Email for critical error notifications.
const SENDER_EMAIL             = ""; // IMPORTANT: If set, this will be the "from" address on all emails. It MUST be a configured alias for the user running the script.
const SENDER_NAME              = "Absence Request System";     // The name that appears as the sender (e.g., "School Office").
const TARGET_CALENDAR_ID       = "c_860cc463977d96c5225672c0d56a6f7585aaa64ddd3c1170a1f71e80c929425f@group.calendar.google.com"; // Specific calendar for absence events.
const ADMIN_FORM_ID            = "1XRwmIU4n4muXoq10-GLAb8M4RZLa8pKD2V0em96-sLM"; // For use in populating email dropdown
const HEADER_ADMIN_NOTIFY_HT   = "Notify Headteacher?";        // The exact question text for the HT notification option on the admin form.

// --- Sheet Configuration ---
const REQUEST_LOG_SHEET_NAME   = "Staff Absence Requests Log"; // Name of the master response sheet.
const ARCHIVE_SHEET_NAME       = "Request Log ARCHIVE";        // Name of the sheet to archive old requests to.
const HISTORY_SHEET_NAME       = "Absence History";            // Name of the sheet storing EXTERNALLY POPULATED approved absences.
const STAFF_RESPONSE_SHEET     = "_RawStaffResponses";         // Name of the sheet where staff responses land.
const ADMIN_RESPONSE_SHEET     = "_RawAdminResponses";         // Name of the sheet where admin responses land.

// --- Staff Directory Lookup Configuration ---
const USE_DIRECTORY_LOOKUP       = true;                                         // Set true to use Directory Sheet, false if Line Manager email is on the form.
const STAFF_DIRECTORY_SHEET_ID   = "1IY-hSAavgJM28o_pWE9GQaoLgm-_S8GYFtlkcytaZ8Y"; // REPLACE: ID of the Staff Directory Google Sheet.
const STAFF_DIRECTORY_SHEET_NAME = "Directory";                                  // Optional: Specify sheet name if not the first sheet.

// --- Header Names in Staff Directory Sheet (Only used if USE_DIRECTORY_LOOKUP = true) ---
const HEADER_DIR_STAFF_EMAIL       = "Staff Email";        // Staff email column in the directory.
const HEADER_DIR_LM_EMAIL          = "Line Manager Email"; // Line Manager email column in the directory.
const HEADER_DIR_STATUS            = "Status";             // Staff status (e.g., Active, Left).
const HEADER_DIR_LEAVING_DATE      = "Leaving Date";       // Staff leaving date.
const HEADER_DIR_STAFF_MEMBER_NAME = "Staff Member Name";  // Full name of the staff member.
const HEADER_DIR_TEACHING_SUPPORT  = "Teaching/Support";   // Category: Teaching or Support staff.
const HEADER_DIR_DEPARTMENT        = "Department";         // Staff member's department.
// Headers for working days/hours in the Staff Directory
const HEADER_DIR_MONDAY            = "Monday";
const HEADER_DIR_TUESDAY           = "Tuesday";
const HEADER_DIR_WEDNESDAY         = "Wednesday";
const HEADER_DIR_THURSDAY          = "Thursday";
const HEADER_DIR_FRIDAY            = "Friday";
const HEADER_DIR_SATURDAY          = "Saturday";
const HEADER_DIR_SUNDAY            = "Sunday";

// --- Header Names in "Staff Absence Requests Log" Sheet (Form Responses) ---
const HEADER_LOG_TIMESTAMP              = "Timestamp";
const HEADER_MASTER_LOG_EMAIL           = "Email Address";                  // The single, consolidated email column in the master log.
const HEADER_STAFF_FORM_EMAIL           = "Email Address";                  // The auto-collected email header from the staff form's raw response sheet.
const HEADER_ADMIN_FORM_EMAIL           = "Staff Member's Email Address";   // The manually entered email header from the admin form's raw response sheet.
const HEADER_LOG_STAFF_NAME             = "Staff Name";                     // Staff member's name (can be auto-populated from directory).
const HEADER_LOG_LM_EMAIL_FORM          = "Line Manager email";             // LM Email from the form (ONLY used if USE_DIRECTORY_LOOKUP = false).
const HEADER_LOG_ABSENCE_START_DATE     = "Absence Start Date";
const HEADER_LOG_ABSENCE_START_TIME     = "Absence Start Time";
const HEADER_LOG_ABSENCE_END_DATE       = "Absence End Date";
const HEADER_LOG_ABSENCE_END_TIME       = "Absence End Time";
const HEADER_LOG_ABSENCE_TYPE           = "Absence Type";
const HEADER_LOG_REASON                 = "Further Details/Cover Implications & Arrangements";
const HEADER_LOG_EVIDENCE_FILE          = "File upload for evidence (If needed)"; // Column for the evidence file URL/ID.
const HEADER_LOG_SUBMISSION_SOURCE      = "Submission Source";  // Tracks how the request was submitted (Form/Admin)
const HEADER_LOG_TEACHING_SUPPORT       = "Teaching/Support"; // Optional: Staff category in log (populated from Directory)
const HEADER_LOG_DEPARTMENT             = "Department";       // Optional: Staff department in log (populated from Directory)
const HEADER_LOG_DURATION_DAYS          = "Duration (Days)";  // Optional: Calculated duration in days (from schedule)

// --- Headers for Script-Added Tracking Columns in "Staff Absence Requests Log" Sheet ---
const HEADER_LOG_APPROVAL_STATUS        = "Approval Status";      // Tracks approval (Pending, Approved, Rejected, Error).
const HEADER_LOG_PAY_STATUS             = "Pay Status";           // Tracks pay status if approved.
const HEADER_LOG_LM_NOTIFIED            = "Line Manager Notified";    // Timestamp of LM notification.
const HEADER_LOG_HT_NOTIFIED            = "Headteacher Notified";     // Timestamp of HT notification.
const HEADER_LOG_APPROVAL_DATE          = "Approval Decision Date"; // Timestamp of approval/rejection.
const HEADER_LOG_APPROVER_EMAIL         = "Approver Email";       // Email of the person who approved/rejected.
const HEADER_LOG_DURATION_HOURS         = "Duration (Hours)";     // Calculated absence duration in hours.
const HEADER_LOG_APPROVER_COMMENT       = "Approver Comment";     // Approver's optional comment.
const HEADER_LOG_START_DATETIME         = "Start Date/Time";      // Programmatically combined start date/time.
const HEADER_LOG_END_DATETIME           = "End Date/Time";        // Programmatically combined end date/time.

// --- Header Names in "Absence History" Sheet (Used for READING historical data) ---
const HEADER_HIST_STAFF_EMAIL           = "EmailAddress";     // Email of the staff member in the history record.
const HEADER_HIST_STAFF_NAME            = "Full Name";        // Staff Name in the history record.
const HEADER_HIST_START_DATE            = "StartDate";        // Start Date of the historical absence.
const HEADER_HIST_END_DATE              = "EndDate";          // End Date of the historical absence.
const HEADER_HIST_ABSENCE_TYPE          = "Type";             // Type of absence in the history record.
const HEADER_HIST_DURATION_DAYS         = "DfE Duration";     // Duration in Days from the history record.
const HEADER_HIST_DATE_APPROVED         = "Date Approved";    // Date Approved (Optional for reading from history).

// --- Archiving Configuration ---
const ARCHIVE_OLDER_THAN_DAYS  = 1; // Archive requests older than this many days.

// --- Deployment Note ---
// This script relies on the Web App being deployed with:
//   Execute as: Me (the script owner, typically the Headteacher or a designated admin)
//   Who has access: Anyone [within your domain] (Or restrict as needed, but "Execute as: Me" is key for approver validation)
// Failure to deploy as "Me" (by the correct account) will break the approval link authorization.
// ---

// --- GLOBAL VARIABLES ---
const SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();
const SCRIPT_CACHE = CacheService.getScriptCache(); // Cache for column indices to improve performance.

// =========================================================================
// MAIN ENTRY POINTS & PRIMARY HANDLERS
// =========================================================================

/**
 * Routes a new submission from a raw response sheet to the master log sheet.
 * This function intelligently maps columns from the source sheet to the master
 * log, consolidating different email columns into a single one.
 *
 * @param {GoogleAppsScript.Events.SheetsOnFormSubmit} e The event object from the raw sheet.
 * @param {object} sourceOptions Options defining the submission source.
 */
function _routeFormResponseToMasterLog(e, sourceOptions) {
  if (!e || !e.values) {
    Logger.log(`Router function called with invalid event object for source: ${sourceOptions.submissionSource}. Exiting.`);
    return;
  }

  try {
    const sourceSheet = e.range.getSheet();
    const masterLogSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(REQUEST_LOG_SHEET_NAME);
    if (!masterLogSheet) throw new Error(`Master log sheet "${REQUEST_LOG_SHEET_NAME}" not found.`);

    // Get the column header-to-index mapping for both the source and master sheets.
    const sourceIndices = getColumnIndices(sourceSheet, sourceSheet.getName());
    const masterIndices = getColumnIndices(masterLogSheet, REQUEST_LOG_SHEET_NAME);
    if (!sourceIndices || !masterIndices) throw new Error("Could not get column indices for routing.");

    // Create a new, empty array that matches the width of the master log.
    const newMasterRow = Array(masterLogSheet.getLastColumn()).fill('');

    // --- Map values from source to master based on matching header names ---
    for (const header in masterIndices) {
      // Skip the consolidated email header for now; we'll handle it specially.
      if (header === HEADER_MASTER_LOG_EMAIL.toLowerCase()) continue;

      const masterColIndex = masterIndices[header] - 1; // 0-based for array
      const sourceColIndex = sourceIndices[header];     // 1-based from helper

      // If the header exists in the source sheet, copy its value to the master row.
      if (sourceColIndex) {
        newMasterRow[masterColIndex] = e.values[sourceColIndex - 1];
      }
    }

    // --- Special Handling for the Consolidated Email Column ---
    const masterEmailColIndex = masterIndices[HEADER_MASTER_LOG_EMAIL.toLowerCase()] - 1;
    let emailValue = '';

    if (sourceOptions.submissionSource === "Staff Form") {
      const sourceEmailIndex = sourceIndices[HEADER_STAFF_FORM_EMAIL.toLowerCase()] - 1;
      emailValue = e.values[sourceEmailIndex];
    } else if (sourceOptions.submissionSource === "Admin Form") {
      const sourceEmailIndex = sourceIndices[HEADER_ADMIN_FORM_EMAIL.toLowerCase()] - 1;
      emailValue = e.values[sourceEmailIndex];
    }

    // Place the correct email value into the master row.
    newMasterRow[masterEmailColIndex] = emailValue;

    // Append the newly constructed and correctly ordered row to the master log.
    masterLogSheet.appendRow(newMasterRow);
    const newRowIndex = masterLogSheet.getLastRow();

    Logger.log(`Routed and mapped submission from '${sourceSheet.getName()}' to master log row ${newRowIndex}.`);

    // Call the main handler, passing it the NEW row data from the MASTER sheet.
    _onFormSubmitHandler(newMasterRow, newRowIndex, masterLogSheet, sourceOptions);

  } catch (error) {
    Logger.log(`FATAL ERROR in _routeFormResponseToMasterLog: ${error.message}\nStack:\n${error.stack}`);
    notifyAdmin(`Absence Script CRITICAL Failure in routing function: ${error.message}`);
  }
}

/**
 * The SINGLE entry point for ALL Google Form submissions to this spreadsheet.
 * It inspects the event to determine which form was submitted and routes
 * the data to the master log with the correct options.
 *
 * @param {GoogleAppsScript.Events.SheetsOnFormSubmit} e The event object.
 */
function handleFormSubmissions(e) {
  // Get the name of the sheet where the new response was added.
  const sourceSheetName = e.range.getSheet().getName();
  Logger.log(`Submission received on sheet: "${sourceSheetName}".`);

  // Define the names of your raw response sheets.
  // IMPORTANT: These must exactly match the names of your tabs.
  const STAFF_RESPONSE_SHEET = "_RawStaffResponses";
  const ADMIN_RESPONSE_SHEET = "_RawAdminResponses";

  // Based on the sheet name, call the router with the correct options.
  if (sourceSheetName === STAFF_RESPONSE_SHEET) {
    _routeFormResponseToMasterLog(e, {
      submissionSource: "Staff Form",
      requiresApproval: true
    });
  } else if (sourceSheetName === ADMIN_RESPONSE_SHEET) {
    // If the Admin form is set to collect emails, the submitter's email will be available
    // in the event object's namedValues, under the "Email Address" header.
    let adminEmail = "";
    if (e.namedValues && e.namedValues[HEADER_STAFF_FORM_EMAIL] && e.namedValues[HEADER_STAFF_FORM_EMAIL][0]) {
        adminEmail = e.namedValues[HEADER_STAFF_FORM_EMAIL][0];
    }

    // Determine if HT approval is required based on the form response.
    let requiresApproval = false; // Default to the bypass behavior.
    if (e.namedValues && e.namedValues[HEADER_ADMIN_NOTIFY_HT] && e.namedValues[HEADER_ADMIN_NOTIFY_HT][0]) {
        const notifyHtResponse = e.namedValues[HEADER_ADMIN_NOTIFY_HT][0].toString().toUpperCase();
        if (notifyHtResponse === 'TRUE') {
            requiresApproval = true;
        }
    }

    _routeFormResponseToMasterLog(e, {
      submissionSource: "Admin Form",
      requiresApproval: requiresApproval,
      adminSubmitterEmail: adminEmail // Pass the captured email along
    });
  } else {
    // If the submission came from any other sheet, log it and ignore it.
    Logger.log(`Submission from an untracked sheet ("${sourceSheetName}"). Ignoring.`);
    return;
  }
}

/**
 * Shared handler for all form submissions AFTER they have been routed to the master log.
 * It validates the request and delegates to the core processing logic.
 * @param {Array} rowData The array of values for the new row.
 * @param {number} rowIndex The row number of the new submission in the master log.
 * @param {Sheet} requestSheet The sheet object for the master log.
 * @param {object} sourceOptions Options defining the submission source.
 */
function _onFormSubmitHandler(rowData, rowIndex, requestSheet, sourceOptions) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) {
    Logger.log("Could not acquire lock for form submission. Exiting.");
    return;
  }

  let logIndices = null;

  try {
    logIndices = getColumnIndices(requestSheet, REQUEST_LOG_SHEET_NAME);
    if (!logIndices) throw new Error(`Could not get column indices for "${REQUEST_LOG_SHEET_NAME}".`);

    const essentialHeaders = [HEADER_LOG_APPROVAL_STATUS, HEADER_MASTER_LOG_EMAIL, HEADER_LOG_ABSENCE_START_DATE, HEADER_LOG_ABSENCE_START_TIME];
    if (!validateRequiredHeaders(logIndices, essentialHeaders, REQUEST_LOG_SHEET_NAME)) {
      throw new Error(`Missing essential headers in "${REQUEST_LOG_SHEET_NAME}" for form processing.`);
    }

    // The status check is still valid, as the core processor might be re-run.
    const statusCol = logIndices[HEADER_LOG_APPROVAL_STATUS.toLowerCase()];
    const currentStatus = requestSheet.getRange(rowIndex, statusCol).getValue();
    if (currentStatus && currentStatus.toString().trim() !== '') {
      Logger.log(`Row ${rowIndex} already has status '${currentStatus}'. Skipping.`);
      lock.releaseLock();
      return;
    }

    Logger.log(`Processing submission from '${sourceOptions.submissionSource}' for master log row index: ${rowIndex}.`);

    const submitterEmail = rowData[logIndices[HEADER_MASTER_LOG_EMAIL.toLowerCase()] - 1];
    if (!submitterEmail || !isValidEmail(submitterEmail.toString())) {
      Logger.log(`ERROR: Missing or invalid staff email in row ${rowIndex}. Email: '${submitterEmail}'. Skipping.`);
      requestSheet.getRange(rowIndex, logIndices[HEADER_LOG_APPROVAL_STATUS.toLowerCase()]).setValue('Error - Invalid Staff Email');
      lock.releaseLock();
      return;
    }

    const webAppUrl = ScriptApp.getService().getUrl();
    if (sourceOptions.requiresApproval && !webAppUrl) {
      Logger.log("ERROR: Script must be deployed as Web App for approval links to work.");
      notifyAdmin("Absence Script ERROR: Not deployed as Web App. Approval links cannot be generated.");
      requestSheet.getRange(rowIndex, logIndices[HEADER_LOG_APPROVAL_STATUS.toLowerCase()]).setValue('Error - Deploy Script as Web App');
      lock.releaseLock();
      return;
    }

    // --- Data Gathering & Processing ---
    let totalRequestsOnThisDate = { approved: 0, pending: 0 };
    const startDateRaw = rowData[logIndices[HEADER_LOG_ABSENCE_START_DATE.toLowerCase()] - 1];
    const startTimeRaw = rowData[logIndices[HEADER_LOG_ABSENCE_START_TIME.toLowerCase()] - 1];
    const formStartDateTime = combineDateAndTime(startDateRaw, startTimeRaw);

    if (formStartDateTime instanceof Date && !isNaN(formStartDateTime)) {
      totalRequestsOnThisDate = countTotalRequestsForStartDate(requestSheet, logIndices, formStartDateTime);
    }

    const absenceDetailsOnDate = getAbsenceDetailsForDate(requestSheet, logIndices, formStartDateTime, rowIndex);

    const processingOptions = {
      notifyLM: true,
      notifyHT: sourceOptions.requiresApproval,
      submissionSource: sourceOptions.submissionSource,
      webAppUrl: webAppUrl,
      absenceDetailsOnDate: absenceDetailsOnDate,
      adminSubmitterEmail: sourceOptions.adminSubmitterEmail // Pass email through
    };

    _handleAbsenceProcessing(rowData, rowIndex, requestSheet, logIndices, processingOptions);

  } catch (error) {
    Logger.log(`FATAL ERROR in _onFormSubmitHandler: ${error.message}\nStack:\n${error.stack}`);
    notifyAdmin(`Absence Script CRITICAL Failure in form processing: ${error.message}`);
    try {
      if (requestSheet && rowIndex && logIndices && logIndices[HEADER_LOG_APPROVAL_STATUS.toLowerCase()]) {
        if (typeof rowIndex === 'number' && rowIndex > 1) {
          requestSheet.getRange(rowIndex, logIndices[HEADER_LOG_APPROVAL_STATUS.toLowerCase()]).setValue('Fatal Error - Check Logs');
        }
      }
    } catch (finalError) {
      Logger.log(`Failed to update status on fatal error: ${finalError.message}`);
    }
  } finally {
    lock.releaseLock();
  }
}

/**
 * Handles GET requests from email approval/rejection links.
 */
function doGet(e) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) {
    Logger.log("doGet could not acquire lock. Server may be busy.");
    return createHtmlResponse("Server Busy", "The server is busy processing another request. Please try again in a moment.", "#ff9800");
  }

  let message = "An error occurred processing your request. Please contact the administrator.";
  let title = "Absence Request Processing Error";
  let bgColor = "#f44336";
  let requestSheet;

  try {
    // --- Parameter and Security Validation ---
    const params = e.parameter;
    const { action, row: rowParam, approver: approverParam, comment } = params;
    const isCommentProvided = typeof comment !== 'undefined';
    const commentParam = isCommentProvided ? (comment || "").trim() : "";

    Logger.log(`doGet received: action=${action}, row=${rowParam}, approver=${approverParam}, comment_provided=${isCommentProvided}, comment='${commentParam}'`);

    if (!action || !rowParam || isNaN(parseInt(rowParam, 10)) || parseInt(rowParam, 10) < 2 || !approverParam) {
      throw new Error("Invalid or missing parameters in the request link. Please ensure the link was copied correctly.");
    }
    const rowIndex = parseInt(rowParam, 10);
    const decodedApprover = decodeURIComponent(approverParam);

    const activeUser = Session.getActiveUser()?.getEmail();
    const effectiveUser = Session.getEffectiveUser()?.getEmail();

    if (!effectiveUser || effectiveUser.toLowerCase() !== HEADTEACHER_EMAIL.toLowerCase()) {
      Logger.log(`SECURITY ALERT: doGet executed by effective user ${effectiveUser}, but script is configured for HEADTEACHER_EMAIL ${HEADTEACHER_EMAIL}. Check deployment settings (must be 'Execute as: Me' by the configured Headteacher).`);
      throw new Error(`Configuration Error: The script is not running as the designated Headteacher (${HEADTEACHER_EMAIL}). Please contact the administrator to correct the script deployment settings.`);
    }
    if (!activeUser || activeUser.toLowerCase() !== decodedApprover.toLowerCase()) {
      Logger.log(`Access Denied: Active user (${activeUser}) does not match the approver in the link (${decodedApprover}). Effective user: ${effectiveUser}.`);
      throw new Error(`Access Denied. This link was intended for ${decodedApprover}. Please ensure you are logged into Google as ${decodedApprover}.`);
    }

    // If the comment form hasn't been submitted yet, show it.
    if (!isCommentProvided) {
      Logger.log(`Comment not provided for action ${action}, row ${rowIndex}. Displaying comment prompt page.`);
      const webAppUrl = ScriptApp.getService().getUrl();
      return createCommentPromptPage(action, rowIndex, decodedApprover, webAppUrl);
    }

    // --- Sheet and Data Validation ---
    requestSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(REQUEST_LOG_SHEET_NAME);
    if (!requestSheet) throw new Error(`Sheet "${REQUEST_LOG_SHEET_NAME}" not found.`);

    const logIndices = getColumnIndices(requestSheet, REQUEST_LOG_SHEET_NAME);
    const requiredDoGetHeaders = [HEADER_LOG_APPROVAL_STATUS, HEADER_LOG_PAY_STATUS, HEADER_LOG_APPROVAL_DATE, HEADER_LOG_APPROVER_EMAIL, HEADER_MASTER_LOG_EMAIL, HEADER_LOG_ABSENCE_TYPE, HEADER_LOG_ABSENCE_START_DATE, HEADER_LOG_ABSENCE_START_TIME, HEADER_LOG_ABSENCE_END_DATE, HEADER_LOG_ABSENCE_END_TIME, HEADER_LOG_APPROVER_COMMENT, HEADER_LOG_STAFF_NAME, HEADER_LOG_DURATION_DAYS, HEADER_LOG_DURATION_HOURS];
    if (!validateRequiredHeaders(logIndices, requiredDoGetHeaders, REQUEST_LOG_SHEET_NAME)) {
      throw new Error(`Missing required headers in "${REQUEST_LOG_SHEET_NAME}" for doGet processing. Cannot proceed.`);
    }

    if (rowIndex > requestSheet.getMaxRows()) {
      throw new Error(`Invalid row number (${rowIndex}). Row does not exist in the sheet "${REQUEST_LOG_SHEET_NAME}". Max rows: ${requestSheet.getMaxRows()}`);
    }

    const statusCol = logIndices[HEADER_LOG_APPROVAL_STATUS.toLowerCase()];
    const statusRange = requestSheet.getRange(rowIndex, statusCol);
    const currentStatus = statusRange.getValue();

    if (currentStatus === 'Approved' || currentStatus === 'Rejected') {
      Logger.log(`Action '${action}' on row ${rowIndex} ignored. Status already '${currentStatus}'.`);
      title = "Absence Request Already Processed";
      message = `Request for row ${rowIndex} has already been processed with status: ${currentStatus}.\nNo further action has been taken.`;
      bgColor = (currentStatus === 'Approved') ? '#4CAF50' : '#ff9800'; // Orange for already rejected
      return createHtmlResponse(title, message, bgColor);
    }

    // --- Process Action ---
    let newApprovalStatus = '', newPayStatus = '', decisionText = '';
    switch (action) {
      case 'approve_paid':
        newApprovalStatus = 'Approved'; newPayStatus = 'Paid'; decisionText = 'Approved with Pay';
        title = "Absence Request Approved (Paid)"; bgColor = "#4CAF50";
        break;
      case 'approve_unpaid':
        newApprovalStatus = 'Approved'; newPayStatus = 'Unpaid'; decisionText = 'Approved without Pay';
        title = "Absence Request Approved (Unpaid)"; bgColor = "#4CAF50";
        break;
      case 'reject':
        newApprovalStatus = 'Rejected'; newPayStatus = ''; decisionText = 'Rejected';
        title = "Absence Request Rejected"; bgColor = "#ff9800";
        break;
      case 'other_specify':
        newApprovalStatus = 'Other'; newPayStatus = 'Other'; decisionText = 'Other (See Comment)';
        title = "Absence Request - Other"; bgColor = "#607d8b";
        break;
      default:
        throw new Error(`Invalid action parameter received: '${action}'.`);
    }
    message = `Request for row ${rowIndex} has been marked as ${decisionText}.`;

    const approvalDate = new Date();
    const timestampFormat = "yyyy-MM-dd HH:mm:ss";
    requestSheet.getRange(rowIndex, statusCol).setValue(newApprovalStatus);
    requestSheet.getRange(rowIndex, logIndices[HEADER_LOG_PAY_STATUS.toLowerCase()]).setValue(newPayStatus);
    requestSheet.getRange(rowIndex, logIndices[HEADER_LOG_APPROVAL_DATE.toLowerCase()]).setValue(approvalDate).setNumberFormat(timestampFormat);
    requestSheet.getRange(rowIndex, logIndices[HEADER_LOG_APPROVER_EMAIL.toLowerCase()]).setValue(activeUser);
    requestSheet.getRange(rowIndex, logIndices[HEADER_LOG_APPROVER_COMMENT.toLowerCase()]).setValue(commentParam);
    Logger.log(`Updated row ${rowIndex}: Status='${newApprovalStatus}', Pay='${newPayStatus}', Approver='${activeUser}', Comment='${commentParam}'`);

    // --- Post-Processing (Calendar & Notifications) ---
    if (newApprovalStatus === 'Approved' || newApprovalStatus === 'Other') {
      try {
        const staffNameValue = requestSheet.getRange(rowIndex, logIndices[HEADER_LOG_STAFF_NAME.toLowerCase()]).getValue();
        const eventStartDateRaw = requestSheet.getRange(rowIndex, logIndices[HEADER_LOG_ABSENCE_START_DATE.toLowerCase()]).getValue();
        const eventStartTimeRaw = requestSheet.getRange(rowIndex, logIndices[HEADER_LOG_ABSENCE_START_TIME.toLowerCase()]).getValue();
        const eventEndDateRaw = requestSheet.getRange(rowIndex, logIndices[HEADER_LOG_ABSENCE_END_DATE.toLowerCase()]).getValue();
        const eventEndTimeRaw = requestSheet.getRange(rowIndex, logIndices[HEADER_LOG_ABSENCE_END_TIME.toLowerCase()]).getValue();
        const eventAbsenceType = requestSheet.getRange(rowIndex, logIndices[HEADER_LOG_ABSENCE_TYPE.toLowerCase()]).getValue();

        const eventStartDateTime = combineDateAndTime(eventStartDateRaw, eventStartTimeRaw);
        const eventEndDateTime = combineDateAndTime(eventEndDateRaw, eventEndTimeRaw);

        if (staffNameValue && eventStartDateTime instanceof Date && !isNaN(eventStartDateTime) && eventEndDateTime instanceof Date && !isNaN(eventEndDateTime)) {
          createCalendarEvent(staffNameValue.toString(), eventStartDateTime, eventEndDateTime, eventAbsenceType.toString(), activeUser, commentParam);
        } else {
          Logger.log(`Skipping calendar event creation for row ${rowIndex} due to missing data. Staff: '${staffNameValue}', Start: ${eventStartDateTime}, End: ${eventEndDateTime}`);
          notifyAdmin(`Absence Script Calendar Warning: Could not create event for row ${rowIndex} due to missing staff name or invalid dates.`);
        }
      } catch (calError) {
        Logger.log(`Error during calendar event creation call for row ${rowIndex}: ${calError.message}\nStack: ${calError.stack}`);
        notifyAdmin(`Absence Script ERROR: Failed to initiate calendar event creation for row ${rowIndex}. Error: ${calError.message}`);
      }
    }

    try {
      const requesterEmail = requestSheet.getRange(rowIndex, logIndices[HEADER_MASTER_LOG_EMAIL.toLowerCase()]).getValue();
      const reqStartDateRaw = requestSheet.getRange(rowIndex, logIndices[HEADER_LOG_ABSENCE_START_DATE.toLowerCase()]).getValue();
      const reqStartTimeRaw = requestSheet.getRange(rowIndex, logIndices[HEADER_LOG_ABSENCE_START_TIME.toLowerCase()]).getValue();
      const reqEndDateRaw = requestSheet.getRange(rowIndex, logIndices[HEADER_LOG_ABSENCE_END_DATE.toLowerCase()]).getValue();
      const reqEndTimeRaw = requestSheet.getRange(rowIndex, logIndices[HEADER_LOG_ABSENCE_END_TIME.toLowerCase()]).getValue();
      const reqAbsenceType = requestSheet.getRange(rowIndex, logIndices[HEADER_LOG_ABSENCE_TYPE.toLowerCase()]).getValue();

      const reqStartDateTime = combineDateAndTime(reqStartDateRaw, reqStartTimeRaw);
      const reqEndDateTime = combineDateAndTime(reqEndDateRaw, reqEndTimeRaw);

      const tz = Session.getScriptTimeZone();
      const formattedReqStart = (reqStartDateTime instanceof Date && !isNaN(reqStartDateTime)) ? Utilities.formatDate(reqStartDateTime, tz, "dd/MM/yyyy HH:mm") : "[Invalid Start Date/Time]";
      const formattedReqEnd = (reqEndDateTime instanceof Date && !isNaN(reqEndDateTime)) ? Utilities.formatDate(reqEndDateTime, tz, "dd/MM/yyyy HH:mm") : "[Invalid End Date/Time]";
      const formattedReqStartDateOnly = (reqStartDateTime instanceof Date && !isNaN(reqStartDateTime)) ? Utilities.formatDate(reqStartDateTime, tz, "dd/MM/yyyy") : "[Invalid Start Date]";

      const durationDays = requestSheet.getRange(rowIndex, logIndices[HEADER_LOG_DURATION_DAYS.toLowerCase()]).getValue();
      const durationHours = requestSheet.getRange(rowIndex, logIndices[HEADER_LOG_DURATION_HOURS.toLowerCase()]).getValue();
      let durationForEmail = 'N/A';
      if (durationDays && typeof durationDays === 'number' && durationDays > 0) {
        durationForEmail = `${durationDays} day(s)`;
      } else if (durationHours && typeof durationHours === 'number' && durationHours > 0) {
        durationForEmail = `${durationHours.toFixed(2)} hours`;
      }

      if (isValidEmail(requesterEmail.toString())) {
        const requesterSubject = `Update on your Absence Request (${formattedReqStartDateOnly})`;
        const htmlRequesterBody = _createDecisionEmailHtml(decisionText, formattedReqStart, formattedReqEnd, reqAbsenceType, commentParam, durationForEmail);

        const mailOptions = {
          to: requesterEmail.toString(),
          subject: requesterSubject,
          htmlBody: htmlRequesterBody
        };
        if (SENDER_EMAIL && isValidEmail(SENDER_EMAIL)) {
            mailOptions.from = SENDER_EMAIL;
            mailOptions.name = SENDER_NAME;
        }
        MailApp.sendEmail(mailOptions);
        Logger.log(`Sent HTML decision notification ('${decisionText}') to requester ${requesterEmail} for row ${rowIndex}. Comment included: ${!!commentParam}`);
        message += `\n\nThe staff member (${requesterEmail}) has been notified.`;
      } else {
        Logger.log(`WARNING: Could not notify requester for row ${rowIndex}. Invalid email found: '${requesterEmail}'`);
        message += `\n\nWARNING: Could not notify staff member (invalid email: '${requesterEmail}'). Please inform them manually.`;
        bgColor = "#ff9800"; // Orange warning
        notifyAdmin(`Absence Script WARNING: Failed to send decision email for row ${rowIndex} to invalid address: '${requesterEmail}'`);
      }
    } catch (notifyError) {
      Logger.log(`ERROR sending notification to requester for row ${rowIndex}: ${notifyError.message}\nStack: ${notifyError.stack}`);
      message += `\n\nWARNING: An error occurred sending the notification email to the staff member. Error: ${notifyError.message}`;
      bgColor = "#ff9800"; // Orange warning
      notifyAdmin(`Absence Script ERROR: Failed to send decision email for row ${rowIndex}. Error: ${notifyError.message}`);
    }

  } catch (error) {
    Logger.log(`ERROR in doGet: ${error.message}\nParameters: ${e ? JSON.stringify(e.parameter) : 'N/A'}\nStack: ${error.stack}`);
    message = `An error occurred: ${error.message}. Please contact the administrator. Check script logs for details.`;
    title = "Absence Request Processing Error";
    bgColor = "#f44336"; // Red error
    notifyAdmin(`Absence Script CRITICAL Failure in doGet: ${error.message}. Parameters: ${e ? JSON.stringify(e.parameter) : 'N/A'}`);
  } finally {
    lock.releaseLock();
  }

  return createHtmlResponse(title, message, bgColor);
}

// =========================================================================
// CORE PROCESSING LOGIC (Called by Form Submit)
// =========================================================================

/**
 * CORE LOGIC for processing an absence row.
 * Handles data validation, calculations, notifications, and sheet updates.
 */
function _handleAbsenceProcessing(rowData, rowIndex, requestSheet, logIndices, options) {
  try {
    // --- Header Validation ---
    const requiredCoreHeaders = [HEADER_LOG_TIMESTAMP, HEADER_MASTER_LOG_EMAIL, HEADER_LOG_STAFF_NAME, HEADER_LOG_ABSENCE_START_DATE, HEADER_LOG_ABSENCE_START_TIME, HEADER_LOG_ABSENCE_END_DATE, HEADER_LOG_ABSENCE_END_TIME, HEADER_LOG_ABSENCE_TYPE, HEADER_LOG_REASON, HEADER_LOG_APPROVAL_STATUS, HEADER_LOG_PAY_STATUS, HEADER_LOG_LM_NOTIFIED, HEADER_LOG_HT_NOTIFIED, HEADER_LOG_APPROVAL_DATE, HEADER_LOG_APPROVER_EMAIL, HEADER_LOG_DURATION_HOURS, HEADER_LOG_APPROVER_COMMENT, HEADER_LOG_EVIDENCE_FILE, HEADER_LOG_SUBMISSION_SOURCE];
    if (!USE_DIRECTORY_LOOKUP && !logIndices[HEADER_LOG_LM_EMAIL_FORM.toLowerCase()]) {
      requiredCoreHeaders.push(HEADER_LOG_LM_EMAIL_FORM);
    }
    if (!validateRequiredHeaders(logIndices, requiredCoreHeaders, REQUEST_LOG_SHEET_NAME)) {
      throw new Error(`Missing one or more required headers in "${REQUEST_LOG_SHEET_NAME}" for core processing of row ${rowIndex}.`);
    }

    // --- Data Extraction and Enrichment ---
    if (logIndices[HEADER_LOG_SUBMISSION_SOURCE.toLowerCase()]) {
      requestSheet.getRange(rowIndex, logIndices[HEADER_LOG_SUBMISSION_SOURCE.toLowerCase()]).setValue(options.submissionSource || 'Unknown');
    }

    const submitterEmail = rowData[logIndices[HEADER_MASTER_LOG_EMAIL.toLowerCase()] - 1];
    
    // Determine the correct recipient for invalid request notifications.
    // If an admin submitted it, notify the admin. Otherwise, notify the staff member.
    let emailForInvalidNotification = submitterEmail;
    if (options.submissionSource === 'Admin Form' && options.adminSubmitterEmail && isValidEmail(options.adminSubmitterEmail)) {
      emailForInvalidNotification = options.adminSubmitterEmail;
      Logger.log(`Admin submission detected. Invalid request notifications for row ${rowIndex} will be sent to admin: ${emailForInvalidNotification}`);
    }

    const staffDirectoryInfo = USE_DIRECTORY_LOOKUP ? getStaffDirectoryInfoByEmail(submitterEmail) : null;
    let staffDisplayName = rowData[logIndices[HEADER_LOG_STAFF_NAME.toLowerCase()] - 1];

    if ((!staffDisplayName || String(staffDisplayName).trim() === "") && staffDirectoryInfo && staffDirectoryInfo.staffMemberName) {
      staffDisplayName = staffDirectoryInfo.staffMemberName.trim();
      requestSheet.getRange(rowIndex, logIndices[HEADER_LOG_STAFF_NAME.toLowerCase()]).setValue(staffDisplayName);
    } else if (!staffDisplayName || String(staffDisplayName).trim() === "") {
      staffDisplayName = submitterEmail;
      requestSheet.getRange(rowIndex, logIndices[HEADER_LOG_STAFF_NAME.toLowerCase()]).setValue(staffDisplayName);
    }

    if (staffDirectoryInfo) {
      if (logIndices[HEADER_LOG_TEACHING_SUPPORT.toLowerCase()] && staffDirectoryInfo.teachingOrSupport) {
        requestSheet.getRange(rowIndex, logIndices[HEADER_LOG_TEACHING_SUPPORT.toLowerCase()]).setValue(staffDirectoryInfo.teachingOrSupport);
      }
      if (logIndices[HEADER_LOG_DEPARTMENT.toLowerCase()] && staffDirectoryInfo.department) {
        requestSheet.getRange(rowIndex, logIndices[HEADER_LOG_DEPARTMENT.toLowerCase()]).setValue(staffDirectoryInfo.department);
      }
    }

    const startDateRaw = rowData[logIndices[HEADER_LOG_ABSENCE_START_DATE.toLowerCase()] - 1];
    const startTimeRaw = rowData[logIndices[HEADER_LOG_ABSENCE_START_TIME.toLowerCase()] - 1];
    const endDateRaw = rowData[logIndices[HEADER_LOG_ABSENCE_END_DATE.toLowerCase()] - 1];
    const endTimeRaw = rowData[logIndices[HEADER_LOG_ABSENCE_END_TIME.toLowerCase()] - 1];
    const startDateTime = combineDateAndTime(startDateRaw, startTimeRaw);
    const endDateTime = combineDateAndTime(endDateRaw, endTimeRaw);

    // --- Validation Guard Clauses ---
    if (!(startDateTime instanceof Date && !isNaN(startDateTime))) {
      requestSheet.getRange(rowIndex, logIndices[HEADER_LOG_APPROVAL_STATUS.toLowerCase()]).setValue('Error - Invalid Start Date');
      Logger.log(`Row ${rowIndex}: Invalid start date/time. Raw: '${startDateRaw}', '${startTimeRaw}'.`);
      return;
    }
    if (!(endDateTime instanceof Date && !isNaN(endDateTime))) {
      requestSheet.getRange(rowIndex, logIndices[HEADER_LOG_APPROVAL_STATUS.toLowerCase()]).setValue('Error - Invalid End Date');
      Logger.log(`Row ${rowIndex}: Invalid end date/time. Raw: '${endDateRaw}', '${endTimeRaw}'.`);
      return;
    }

    // --- Programmatically set combined Start and End Date/Time columns ---
    try {
        const tz = Session.getScriptTimeZone();
        const dateTimeFormat = "dd/MM/yyyy HH:mm";

        if (logIndices[HEADER_LOG_START_DATETIME.toLowerCase()]) {
            const formattedStartDateTime = Utilities.formatDate(startDateTime, tz, dateTimeFormat);
            requestSheet.getRange(rowIndex, logIndices[HEADER_LOG_START_DATETIME.toLowerCase()]).setValue(formattedStartDateTime);
        }
        if (logIndices[HEADER_LOG_END_DATETIME.toLowerCase()]) {
            const formattedEndDateTime = Utilities.formatDate(endDateTime, tz, dateTimeFormat);
            requestSheet.getRange(rowIndex, logIndices[HEADER_LOG_END_DATETIME.toLowerCase()]).setValue(formattedEndDateTime);
        }
    } catch(e) {
        Logger.log(`Row ${rowIndex}: Could not set combined date/time fields. Error: ${e.message}`);
        // Non-fatal error, so we continue processing.
    }


    // Request is more than a month old
    const now = new Date();
    const oneMonthAgo = new Date();
    oneMonthAgo.setMonth(now.getMonth() - 1);
    oneMonthAgo.setHours(0, 0, 0, 0);
    if (startDateTime < oneMonthAgo) {
      requestSheet.getRange(rowIndex, logIndices[HEADER_LOG_APPROVAL_STATUS.toLowerCase()]).setValue('Invalid - Too Far In Past');
      const tz = Session.getScriptTimeZone();
      const errorMessage = `Absence start date (${Utilities.formatDate(startDateTime, tz, "dd/MM/yyyy HH:mm")}) is more than one month in the past.`;
      if (logIndices[HEADER_LOG_APPROVER_COMMENT.toLowerCase()]) requestSheet.getRange(rowIndex, logIndices[HEADER_LOG_APPROVER_COMMENT.toLowerCase()]).setValue(errorMessage);
      if (isValidEmail(emailForInvalidNotification)) {
        const template = HtmlService.createTemplateFromFile('invalid-request-email.html');
        template.reason = 'Start Date Too Far In Past: The start date of your request is more than one month in the past.';
        template.instruction = 'Please contact administration if this is an urgent historical correction.';
        const htmlBody = template.evaluate().getContent();
        const mailOptions = {
            to: emailForInvalidNotification,
            subject: 'Absence Request Invalid: Start Date Too Far In Past',
            htmlBody: htmlBody
        };
        if (SENDER_EMAIL && isValidEmail(SENDER_EMAIL)) {
            mailOptions.from = SENDER_EMAIL;
            mailOptions.name = SENDER_NAME;
        }
        MailApp.sendEmail(mailOptions);
      }
      Logger.log(`Row ${rowIndex}: ${errorMessage}`);
      return;
    }

    // Request is over a year in the future
    const oneYearFromNow = new Date();
    oneYearFromNow.setFullYear(now.getFullYear() + 1);
    oneYearFromNow.setHours(23, 59, 59, 999);
    if (startDateTime > oneYearFromNow) {
      requestSheet.getRange(rowIndex, logIndices[HEADER_LOG_APPROVAL_STATUS.toLowerCase()]).setValue('Invalid - Too Far In Future');
      const tz = Session.getScriptTimeZone();
      const errorMessage = `Absence start date (${Utilities.formatDate(startDateTime, tz, "dd/MM/yyyy HH:mm")}) is more than one year in the future.`;
      if (logIndices[HEADER_LOG_APPROVER_COMMENT.toLowerCase()]) requestSheet.getRange(rowIndex, logIndices[HEADER_LOG_APPROVER_COMMENT.toLowerCase()]).setValue(errorMessage);
      if (isValidEmail(emailForInvalidNotification)) {
        const template = HtmlService.createTemplateFromFile('invalid-request-email.html');
        template.reason = 'Start Date Too Far In Future: The start date of your request is more than one year in the future.';
        template.instruction = 'Please submit a new request with a valid start date.';
        const htmlBody = template.evaluate().getContent();
        const mailOptions = {
            to: emailForInvalidNotification,
            subject: 'Absence Request Invalid: Start Date Too Far In Future',
            htmlBody: htmlBody
        };
        if (SENDER_EMAIL && isValidEmail(SENDER_EMAIL)) {
            mailOptions.from = SENDER_EMAIL;
            mailOptions.name = SENDER_NAME;
        }
        MailApp.sendEmail(mailOptions);
      }
      Logger.log(`Row ${rowIndex}: ${errorMessage}`);
      return;
    }

    // Request starts on a non-working day
    // A day is considered a working day if there is at least one valid time range.
    if (USE_DIRECTORY_LOOKUP && staffDirectoryInfo && staffDirectoryInfo.days) {
      const dayOfWeek = Utilities.formatDate(startDateTime, Session.getScriptTimeZone(), 'EEEE').toLowerCase();
      const workScheduleForDay = staffDirectoryInfo.days[dayOfWeek];

      let isWorkingDay = false;
      if (workScheduleForDay && typeof workScheduleForDay === 'string' && workScheduleForDay.trim() !== '') {
        const shiftStrings = workScheduleForDay.split(',').map(s => s.trim());
        for (const shiftStr of shiftStrings) {
            try {
                const parsed = parseDayTimeString(startDateTime, shiftStr);
                // A valid shift has a start time before an end time.
                if (parsed.schedStart < parsed.schedEnd) {
                    isWorkingDay = true;
                    break; // Found at least one valid shift, so it's a working day.
                }
            } catch (e) {
                // Invalid format for this part of the string, just ignore and check the next.
            }
        }
      }

      if (!isWorkingDay) {
        requestSheet.getRange(rowIndex, logIndices[HEADER_LOG_APPROVAL_STATUS.toLowerCase()]).setValue('Invalid - Non-working Day');
        const tz = Session.getScriptTimeZone();
        const errorMessage = `Absence start date (${Utilities.formatDate(startDateTime, tz, "dd/MM/yyyy HH:mm")}) falls on a non-working day (${dayOfWeek}).`;
        if (logIndices[HEADER_LOG_APPROVER_COMMENT.toLowerCase()]) requestSheet.getRange(rowIndex, logIndices[HEADER_LOG_APPROVER_COMMENT.toLowerCase()]).setValue(errorMessage);
        if (isValidEmail(emailForInvalidNotification)) {
            const template = HtmlService.createTemplateFromFile('invalid-request-email.html');
            template.reason = `Non-working Day: Your request falls on a day you are not scheduled to work (${dayOfWeek}).`;
            template.instruction = 'Please check your submission or contact administration if your work schedule is incorrect.';
            const htmlBody = template.evaluate().getContent();
            const mailOptions = {
                to: emailForInvalidNotification,
                subject: 'Absence Request Invalid: Non-working Day',
                htmlBody: htmlBody
            };
            if (SENDER_EMAIL && isValidEmail(SENDER_EMAIL)) {
                mailOptions.from = SENDER_EMAIL;
                mailOptions.name = SENDER_NAME;
            }
            MailApp.sendEmail(mailOptions);
        }
        Logger.log(`Row ${rowIndex}: ${errorMessage}`);
        return;
      }
    }

    // Request end date is before start date
    if (endDateTime < startDateTime) {
      requestSheet.getRange(rowIndex, logIndices[HEADER_LOG_APPROVAL_STATUS.toLowerCase()]).setValue('Invalid - End Before Start');
      if (logIndices[HEADER_LOG_APPROVER_COMMENT.toLowerCase()]) requestSheet.getRange(rowIndex, logIndices[HEADER_LOG_APPROVER_COMMENT.toLowerCase()]).setValue('End date/time is before start date/time.');
      if (isValidEmail(emailForInvalidNotification)) {
        const template = HtmlService.createTemplateFromFile('invalid-request-email.html');
        template.reason = 'End date/time is before start date/time.';
        template.instruction = 'Please check your entry and resubmit.';
        const htmlBody = template.evaluate().getContent();
        const mailOptions = {
            to: emailForInvalidNotification,
            subject: 'Absence Request Invalid: End date/time is before start date/time',
            htmlBody: htmlBody
        };
        if (SENDER_EMAIL && isValidEmail(SENDER_EMAIL)) {
            mailOptions.from = SENDER_EMAIL;
            mailOptions.name = SENDER_NAME;
        }
        MailApp.sendEmail(mailOptions);
      }
      Logger.log(`Row ${rowIndex}: End date/time (${endDateTime}) is before start date/time (${startDateTime}).`);
      return;
    }

    // Request start date and time equals end date and time (Zero Length)
    if (startDateTime.getTime() === endDateTime.getTime()) {
      requestSheet.getRange(rowIndex, logIndices[HEADER_LOG_APPROVAL_STATUS.toLowerCase()]).setValue('Invalid - Zero Length');
      const errorMessage = `Absence has zero duration (start and end times are identical).`;
      if (logIndices[HEADER_LOG_APPROVER_COMMENT.toLowerCase()]) requestSheet.getRange(rowIndex, logIndices[HEADER_LOG_APPROVER_COMMENT.toLowerCase()]).setValue(errorMessage);
      if (isValidEmail(emailForInvalidNotification)) {
        const template = HtmlService.createTemplateFromFile('invalid-request-email.html');
        template.reason = 'Zero-Length Absence: The start and end times are the same.';
        template.instruction = 'Please submit a new request with a valid duration.';
        const htmlBody = template.evaluate().getContent();
        const mailOptions = {
            to: emailForInvalidNotification,
            subject: 'Absence Request Invalid: Zero-Length Absence',
            htmlBody: htmlBody
        };
        if (SENDER_EMAIL && isValidEmail(SENDER_EMAIL)) {
            mailOptions.from = SENDER_EMAIL;
            mailOptions.name = SENDER_NAME;
        }
        MailApp.sendEmail(mailOptions);
      }
      Logger.log(`Row ${rowIndex}: ${errorMessage}`);
      return;
    }


    // Request overlaps with previous pending or approved requests
    const allExistingRows = requestSheet.getRange(2, 1, requestSheet.getLastRow() - 1, requestSheet.getLastColumn()).getValues();
    const overlappingSummaries = [];
    for (let i = 0; i < allExistingRows.length; i++) {
      const sheetRowIdx = i + 2;
      if (sheetRowIdx === rowIndex) continue;
      const otherRow = allExistingRows[i];
      const otherRowEmail = otherRow[logIndices[HEADER_MASTER_LOG_EMAIL.toLowerCase()] - 1];
      const otherRowStatus = otherRow[logIndices[HEADER_LOG_APPROVAL_STATUS.toLowerCase()] - 1];
      if (!otherRowEmail || otherRowEmail.toString().trim().toLowerCase() !== submitterEmail.toString().trim().toLowerCase()) continue;
      if (otherRowStatus !== 'Pending Approval' && otherRowStatus !== 'Approved') continue;
      const otherRowStartDateTime = combineDateAndTime(otherRow[logIndices[HEADER_LOG_ABSENCE_START_DATE.toLowerCase()] - 1], otherRow[logIndices[HEADER_LOG_ABSENCE_START_TIME.toLowerCase()] - 1]);
      const otherRowEndDateTime = combineDateAndTime(otherRow[logIndices[HEADER_LOG_ABSENCE_END_DATE.toLowerCase()] - 1], otherRow[logIndices[HEADER_LOG_ABSENCE_END_TIME.toLowerCase()] - 1]);
      if (!(otherRowStartDateTime instanceof Date) || isNaN(otherRowStartDateTime) || !(otherRowEndDateTime instanceof Date) || isNaN(otherRowEndDateTime)) continue;
      if (doDateRangesOverlap(startDateTime, endDateTime, otherRowStartDateTime, otherRowEndDateTime)) {
        const tz = Session.getScriptTimeZone();
        overlappingSummaries.push(`Type: ${otherRow[logIndices[HEADER_LOG_ABSENCE_TYPE.toLowerCase()]-1] || 'N/A'} | Status: ${otherRowStatus} | ${Utilities.formatDate(otherRowStartDateTime, tz, "dd/MM/yyyy HH:mm")} - ${Utilities.formatDate(otherRowEndDateTime, tz, "dd/MM/yyyy HH:mm")}`);
      }
    }
    if (overlappingSummaries.length > 0) {
      requestSheet.getRange(rowIndex, logIndices[HEADER_LOG_APPROVAL_STATUS.toLowerCase()]).setValue('Invalid - Overlap');
      const summaryText = 'Overlapping absence request(s) detected: ' + overlappingSummaries.join(' ; ');
      if (logIndices[HEADER_LOG_APPROVER_COMMENT.toLowerCase()]) requestSheet.getRange(rowIndex, logIndices[HEADER_LOG_APPROVER_COMMENT.toLowerCase()]).setValue(summaryText);
      if (isValidEmail(emailForInvalidNotification)) {
        const template = HtmlService.createTemplateFromFile('invalid-request-email.html');
        template.reason = `Overlapping Absence Detected: Your request overlaps with existing requests:\n${overlappingSummaries.join('\n')}`;
        template.instruction = 'Please contact admin to clarify.';
        const htmlBody = template.evaluate().getContent();
        const mailOptions = {
            to: emailForInvalidNotification,
            subject: 'Absence Request Invalid: Overlapping Absence Detected',
            htmlBody: htmlBody.replace(/\n/g, '<br>') // For the list of overlaps
        };
        if (SENDER_EMAIL && isValidEmail(SENDER_EMAIL)) {
            mailOptions.from = SENDER_EMAIL;
            mailOptions.name = SENDER_NAME;
        }
        MailApp.sendEmail(mailOptions);
      }
      Logger.log(`Row ${rowIndex}: Overlapping request(s) found. Details: ${summaryText}`);
      return;
    }

    // --- Calculations ---
    let durationHours = 0;
    try {
      durationHours = calculateAbsenceHoursFromSchedule(submitterEmail, startDateTime, endDateTime);
      requestSheet.getRange(rowIndex, logIndices[HEADER_LOG_DURATION_HOURS.toLowerCase()]).setValue(parseFloat(durationHours.toFixed(2)));
    } catch (calcError) {
      Logger.log(`Error calculating duration (hours) for row ${rowIndex}: ${calcError.message}`);
      requestSheet.getRange(rowIndex, logIndices[HEADER_LOG_DURATION_HOURS.toLowerCase()]).setValue("Calculation Error");
      durationHours = "Error"; // Signal that the calculation failed.
    }

    let roundedDurationDays = null;
    if (logIndices[HEADER_LOG_DURATION_DAYS.toLowerCase()]) {
      try {
        const durationDays = calculateAbsenceDays(submitterEmail, startDateTime, endDateTime);
        roundedDurationDays = Math.round(durationDays * 100) / 100;
        requestSheet.getRange(rowIndex, logIndices[HEADER_LOG_DURATION_DAYS.toLowerCase()]).setValue(roundedDurationDays);
      } catch (err) {
        requestSheet.getRange(rowIndex, logIndices[HEADER_LOG_DURATION_DAYS.toLowerCase()]).setValue("Error");
        roundedDurationDays = "Error";
        Logger.log(`Error calculating duration (days) for row ${rowIndex}: ${err.message}`);
      }
    }

    // --- Notification Preparation ---
    let lineManagerEmail = null;
    if (USE_DIRECTORY_LOOKUP) {
      lineManagerEmail = staffDirectoryInfo ? staffDirectoryInfo.lineManagerEmail : null;
    } else if (logIndices[HEADER_LOG_LM_EMAIL_FORM.toLowerCase()]) {
      lineManagerEmail = rowData[logIndices[HEADER_LOG_LM_EMAIL_FORM.toLowerCase()] - 1];
    }
    if (lineManagerEmail && typeof lineManagerEmail === 'string' && !isValidEmail(lineManagerEmail.trim())) {
      lineManagerEmail = null;
    }
    if (USE_DIRECTORY_LOOKUP && !lineManagerEmail && options.notifyLM) {
      notifyAdmin(`Absence Alert: Could not find Line Manager for ${submitterEmail} (Row ${rowIndex}) in Staff Directory for notification.`);
    }

    const historySummary = calculateAbsenceHistoryCategorized(HISTORY_SHEET_NAME, submitterEmail, startDateTime);
    const timestampFormat = "yyyy-MM-dd HH:mm:ss";
    const tz = Session.getScriptTimeZone();
    const formattedStartDate = Utilities.formatDate(startDateTime, tz, "dd/MM/yyyy");
    const formattedStartTime = Utilities.formatDate(startDateTime, tz, "HH:mm");
    const formattedEndDate = Utilities.formatDate(endDateTime, tz, "dd/MM/yyyy");
    const formattedEndTime = Utilities.formatDate(endDateTime, tz, "HH:mm");
    const durationForEmail = (roundedDurationDays !== null && roundedDurationDays !== "Error") 
      ? `${roundedDurationDays} day(s)` 
      : (typeof durationHours === 'number' ? `${durationHours.toFixed(2)} hours` : 'N/A');
    const absenceType = rowData[logIndices[HEADER_LOG_ABSENCE_TYPE.toLowerCase()] - 1];
    const reason = rowData[logIndices[HEADER_LOG_REASON.toLowerCase()] - 1];
    const requestDetails = `\n  Staff Member: ${staffDisplayName} (${submitterEmail})\n  Type of Absence: ${absenceType || 'N/A'}\n\n  Start: ${formattedStartDate} at ${formattedStartTime}\n  End: ${formattedEndDate} at ${formattedEndTime}\n  \n  Calculated Duration: ${durationForEmail}\n  Reason Provided: ${reason || '(None provided)'}\n  Row Number: ${rowIndex}`;

    // --- Send Notifications ---

    // If the submission does not require Headteacher approval (e.g., from an Admin Form),
    // log it directly and notify the Line Manager for information only.
    if (!options.notifyHT) {
      const adminSubmitter = options.adminSubmitterEmail || Session.getActiveUser().getEmail() || 'Admin Action';
      requestSheet.getRange(rowIndex, logIndices[HEADER_LOG_APPROVAL_STATUS.toLowerCase()]).setValue('Approved - Admin Logged');
      
      // For admin-logged requests, the pay status is determined here.
      let finalPayStatus = rowData[logIndices[HEADER_LOG_PAY_STATUS.toLowerCase()] - 1];
      if (!finalPayStatus || String(finalPayStatus).trim() === '') {
        finalPayStatus = ''; // Default if blank
      }
      requestSheet.getRange(rowIndex, logIndices[HEADER_LOG_PAY_STATUS.toLowerCase()]).setValue(finalPayStatus);
      
      requestSheet.getRange(rowIndex, logIndices[HEADER_LOG_APPROVAL_DATE.toLowerCase()]).setValue(now).setNumberFormat(timestampFormat);
      requestSheet.getRange(rowIndex, logIndices[HEADER_LOG_APPROVER_EMAIL.toLowerCase()]).setValue(adminSubmitter);
      
      // Create the calendar event immediately
      createCalendarEvent(staffDisplayName, startDateTime, endDateTime, absenceType, adminSubmitter, 'Logged directly by admin.');
      
      // Notify Line Manager for info
      if (options.notifyLM && lineManagerEmail) {
        try {
          const template = HtmlService.createTemplateFromFile('lm-info-email.html');
          template.infoMessage = 'An absence has been logged by administration for your staff member:';
          template.actionMessage = 'This has been automatically approved and logged. No action is needed.';
          template.staffDisplayName = staffDisplayName;
          template.submitterEmail = submitterEmail;
          template.absenceType = absenceType;
          template.formattedStart = `${formattedStartDate} at ${formattedStartTime}`;
          template.formattedEnd = `${formattedEndDate} at ${formattedEndTime}`;
          template.durationForEmail = durationForEmail;
          template.reason = reason || '(None provided)';
          template.rowIndex = rowIndex;
          const htmlBody = template.evaluate().getContent();

          const mailOptions = {
            to: lineManagerEmail,
            subject: `Absence Logged (Info Only): ${staffDisplayName}`,
            htmlBody: htmlBody
          };
          if (SENDER_EMAIL && isValidEmail(SENDER_EMAIL)) {
              mailOptions.from = SENDER_EMAIL;
              mailOptions.name = SENDER_NAME;
          }
          MailApp.sendEmail(mailOptions);

          if (logIndices[HEADER_LOG_LM_NOTIFIED.toLowerCase()]) requestSheet.getRange(rowIndex, logIndices[HEADER_LOG_LM_NOTIFIED.toLowerCase()]).setValue(now).setNumberFormat(timestampFormat);
        } catch (errLm) { Logger.log(`Row ${rowIndex}: ERROR sending admin-logged LM email: ${errLm.message}`); }
      }

      // Notify Staff Member of Admin-Logged Submission
      if (isValidEmail(submitterEmail) && options.submissionSource === 'Admin Form') {
        try {
          const template = HtmlService.createTemplateFromFile('submission-confirmation-email.html');
          template.infoMessage = 'This email confirms that an absence has been logged on your behalf by school administration:';
          template.actionMessage = 'This request has been automatically approved and recorded in the system. No further action is required from you.';
          template.staffName = staffDisplayName;
          template.submitterEmail = submitterEmail;
          template.absenceType = absenceType;
          template.formattedStart = `${formattedStartDate} at ${formattedStartTime}`;
          template.formattedEnd = `${formattedEndDate} at ${formattedEndTime}`;
          template.reason = reason || '(None provided)';
          template.payStatus = finalPayStatus;
          template.durationForEmail = durationForEmail;
          const htmlBody = template.evaluate().getContent();

          const mailOptions = {
            to: submitterEmail,
            subject: `Absence Logged on Your Behalf (${formattedStartDate})`,
            htmlBody: htmlBody
          };
          if (SENDER_EMAIL && isValidEmail(SENDER_EMAIL)) {
              mailOptions.from = SENDER_EMAIL;
              mailOptions.name = SENDER_NAME;
          }
          MailApp.sendEmail(mailOptions);
          Logger.log(`Row ${rowIndex}: Sent admin-logged submission confirmation to staff member ${submitterEmail}.`);
        } catch (errStaff) {
          Logger.log(`Row ${rowIndex}: ERROR sending admin-logged staff confirmation email: ${errStaff.message}`);
          notifyAdmin(`Absence Script ERROR: Failed to send admin-logged confirmation to ${submitterEmail} for row ${rowIndex}. Error: ${errStaff.message}`);
        }
      }
      
      Logger.log(`Row ${rowIndex}: Admin-logged absence processed directly. Calendar event created.`);
      return; // End processing for this request
    }

    // Standard request notification process
    let htNotificationSent = false;
    if (options.notifyHT && isValidEmail(HEADTEACHER_EMAIL) && options.webAppUrl) {
      const emailOptionsForHT = { to: HEADTEACHER_EMAIL, subject: `Absence Request ACTION REQUIRED: ${staffDisplayName}`, attachments: [] };
      
      const detailsMap = {
        "Staff Member": `${staffDisplayName} (${submitterEmail})`,
        "Type of Absence": absenceType || 'N/A',
        "Start": `${formattedStartDate} at ${formattedStartTime}`,
        "End": `${formattedEndDate} at ${formattedEndTime}`,
        "Calculated Duration": durationForEmail,
        "Reason Provided": reason || '(None provided)',
        "Log Row Number": rowIndex
      };

      let requestDetailsTableHtml = Object.entries(detailsMap).map(([key, value]) => 
        `<tr><th>${key}</th><td>${value}</td></tr>`
      ).join('');

      const template = HtmlService.createTemplateFromFile('headteacher-approval-email.html');
      template.requestDetailsTableHtml = requestDetailsTableHtml;
      template.historySummary = historySummary;
      template.formattedStartDate = formattedStartDate;
      template.approvePaidUrl = `${options.webAppUrl}?action=approve_paid&row=${rowIndex}&approver=${encodeURIComponent(HEADTEACHER_EMAIL)}`;
      template.approveUnpaidUrl = `${options.webAppUrl}?action=approve_unpaid&row=${rowIndex}&approver=${encodeURIComponent(HEADTEACHER_EMAIL)}`;
      template.rejectUrl = `${options.webAppUrl}?action=reject&row=${rowIndex}&approver=${encodeURIComponent(HEADTEACHER_EMAIL)}`;
      template.otherUrl = `${options.webAppUrl}?action=other_specify&row=${rowIndex}&approver=${encodeURIComponent(HEADTEACHER_EMAIL)}`;
      template.approvedListHtml = _generateRequestListHtml(options.absenceDetailsOnDate.approved);
      template.pendingListHtml = _generateRequestListHtml(options.absenceDetailsOnDate.pending);
      template.calendarLink = `https://calendar.google.com/calendar/u/0/embed?src=${encodeURIComponent(TARGET_CALENDAR_ID)}&ctz=Europe/London`;
      template.evidenceAttached = false;
      template.evidenceError = '';

      const evidenceFileUrlOrId = logIndices[HEADER_LOG_EVIDENCE_FILE.toLowerCase()] ? (rowData[logIndices[HEADER_LOG_EVIDENCE_FILE.toLowerCase()] - 1] || '') : '';
      if (evidenceFileUrlOrId) {
        try {
          const fileId = extractFileIdFromDriveUrl(evidenceFileUrlOrId);
          if (fileId) {
            emailOptionsForHT.attachments.push(DriveApp.getFileById(fileId).getBlob());
            template.evidenceAttached = true;
          } else {
            template.evidenceError = `Evidence link provided but could not be attached: ${evidenceFileUrlOrId}`;
          }
        } catch (fileError) {
          template.evidenceError = `Failed to attach evidence: ${evidenceFileUrlOrId}. Error: ${fileError.message}`;
          notifyAdmin(`Absence Script WARNING: Failed to attach evidence for row ${rowIndex}. Error: ${fileError.message}`);
        }
      }
      
      emailOptionsForHT.htmlBody = template.evaluate().getContent();

      try {
        if (SENDER_EMAIL && isValidEmail(SENDER_EMAIL)) {
            emailOptionsForHT.from = SENDER_EMAIL;
            emailOptionsForHT.name = SENDER_NAME;
        }
        MailApp.sendEmail(emailOptionsForHT);
        if (logIndices[HEADER_LOG_HT_NOTIFIED.toLowerCase()]) requestSheet.getRange(rowIndex, logIndices[HEADER_LOG_HT_NOTIFIED.toLowerCase()]).setValue(new Date()).setNumberFormat(timestampFormat);
        htNotificationSent = true;
      } catch (errHt) {
        notifyAdmin(`Absence Script ERROR: Failed to send approval email to HT for row ${rowIndex}. Error: ${errHt.message}`);
      }
    }

    // After successfully notifying HT, send confirmation to staff member
    if (htNotificationSent && (options.submissionSource === 'Staff Form' || options.submissionSource === 'Admin Form')) {
      try {
        const template = HtmlService.createTemplateFromFile('submission-confirmation-email.html');
        template.staffName = staffDisplayName;

        let infoMessage = '';
        let subject = '';

        if (options.submissionSource === 'Staff Form') {
          infoMessage = 'This is a confirmation that your absence request has been successfully submitted and is now pending approval.';
          subject = `Absence Request Submitted Successfully (${formattedStartDate})`;
        } else { // Admin Form
          infoMessage = 'This email confirms an absence request was submitted on your behalf by administration and is now pending Headteacher approval.';
          subject = `Absence Request Submitted on Your Behalf (${formattedStartDate})`;
        }

        template.infoMessage = infoMessage;
        template.actionMessage = 'You will receive another email once a decision has been made by the Headteacher.';
        template.absenceType = absenceType;
        template.formattedStart = `${formattedStartDate} at ${formattedStartTime}`;
        template.formattedEnd = `${formattedEndDate} at ${formattedEndTime}`;
        template.reason = reason || '(None provided)';
        template.durationForEmail = durationForEmail;
        const htmlBody = template.evaluate().getContent();

        const mailOptions = {
            to: submitterEmail,
            subject: subject,
            htmlBody: htmlBody
        };
        if (SENDER_EMAIL && isValidEmail(SENDER_EMAIL)) {
            mailOptions.from = SENDER_EMAIL;
            mailOptions.name = SENDER_NAME;
        }
        MailApp.sendEmail(mailOptions);
        Logger.log(`Row ${rowIndex}: Sent pending submission confirmation email to ${submitterEmail} for source '${options.submissionSource}'.`);
      } catch (errConfirm) {
        Logger.log(`Row ${rowIndex}: ERROR sending pending submission confirmation email: ${errConfirm.message}`);
      }
    }

    requestSheet.getRange(rowIndex, logIndices[HEADER_LOG_APPROVAL_STATUS.toLowerCase()]).setValue(htNotificationSent ? 'Pending Approval' : 'Error - Check Logs');

    if (options.notifyLM && lineManagerEmail) {
      try {
        const template = HtmlService.createTemplateFromFile('lm-info-email.html');
        template.infoMessage = 'Please note the following absence request for your information:';
        template.actionMessage = 'This request requires Headteacher approval. No action is needed from you via this email.';
        template.staffDisplayName = staffDisplayName;
        template.submitterEmail = submitterEmail;
        template.absenceType = absenceType;
        template.formattedStart = `${formattedStartDate} at ${formattedStartTime}`;
        template.formattedEnd = `${formattedEndDate} at ${formattedEndTime}`;
        template.durationForEmail = durationForEmail;
        template.reason = reason || '(None provided)';
        template.rowIndex = rowIndex;
        const htmlBody = template.evaluate().getContent();
        
        const mailOptions = {
            to: lineManagerEmail,
            subject: `Absence Request Submitted (Info Only): ${staffDisplayName}`,
            htmlBody: htmlBody
        };
        if (SENDER_EMAIL && isValidEmail(SENDER_EMAIL)) {
            mailOptions.from = SENDER_EMAIL;
            mailOptions.name = SENDER_NAME;
        }
        MailApp.sendEmail(mailOptions);
        
        if (logIndices[HEADER_LOG_LM_NOTIFIED.toLowerCase()]) requestSheet.getRange(rowIndex, logIndices[HEADER_LOG_LM_NOTIFIED.toLowerCase()]).setValue(new Date()).setNumberFormat(timestampFormat);
      } catch (errLm) {
        Logger.log(`Row ${rowIndex}: ERROR sending LM email: ${errLm.message}`);
      }
    }
    Logger.log(`Row ${rowIndex}: Standard approval workflow. HT Notified: ${htNotificationSent}, LM Notified: ${!!(options.notifyLM && lineManagerEmail)}.`);

  } catch (coreError) {
    Logger.log(`CORE PROCESSING ERROR for Row ${rowIndex}: ${coreError.message}\nStack: ${coreError.stack}`);
    notifyAdmin(`Absence Script CRITICAL Failure in _handleAbsenceProcessing for row ${rowIndex}: ${coreError.message}`);
    try {
      if (requestSheet && logIndices && logIndices[HEADER_LOG_APPROVAL_STATUS.toLowerCase()]) {
        requestSheet.getRange(rowIndex, logIndices[HEADER_LOG_APPROVAL_STATUS.toLowerCase()]).setValue('Fatal Error - Check Core Logs');
      }
    } catch (finalUpdateError) { /* Silent fail */ }
  }
}

// =========================================================================
// EMAIL & UI HELPERS
// =========================================================================

/**
 * Creates the HTML body for the decision notification email sent to the staff member.
 * @param {string} decisionText - The decision text (e.g., "Approved with Pay").
 * @param {string} formattedStart - The formatted start date/time.
 * @param {string} formattedEnd - The formatted end date/time.
 * @param {string} absenceType - The type of absence.
 * @param {string} comment - The approver's comment, if any.
 * @returns {string} The complete HTML for the email body.
 */
function _createDecisionEmailHtml(decisionText, formattedStart, formattedEnd, absenceType, comment, durationForEmail) {
  let decisionColor = "#000000"; // Default black
  if (decisionText === 'Approved without Pay') decisionColor = "#ff9800"; // Orange
  else if (decisionText.includes("Approved")) decisionColor = "#4CAF50"; // Green
  else if (decisionText.includes("Rejected")) decisionColor = "#f44336"; // Red

  const template = HtmlService.createTemplateFromFile('decision-email.html');
  template.decisionText = decisionText;
  template.formattedStart = formattedStart;
  template.formattedEnd = formattedEnd;
  template.absenceType = absenceType;
  template.comment = comment;
  template.decisionColor = decisionColor;
  template.durationForEmail = durationForEmail;
  
  return template.evaluate().getContent();
}

/**
 * Generates an HTML unordered list for a given array of absence request details.
 * @param {Array<object>} requests - Array of request info objects {name, type, startTime, endTime}.
 * @returns {string} An HTML string representing a <ul> list, or a "None" paragraph.
 * @private
 */
function _generateRequestListHtml(requests) {
    if (!requests || requests.length === 0) {
        return '<p style="margin:0;padding-left:20px;font-style:italic;">None</p>';
    }

    let listItems = requests.map(req => {
        const name = req.name || 'Unknown';
        const type = req.type || 'N/A';
        const times = (req.startTime && req.endTime) ? ` - ${req.startTime} to ${req.endTime}` : '';
        return `<li>${name} (${type})${times}</li>`;
    }).join('');

    return `<ul style="margin: 0; padding-left: 40px; list-style-type: disc;">${listItems}</ul>`;
}

function createHtmlResponse(title, message, bgColor) {
  const displayMessage = message.replace(/\n/g, '<br>');
  const template = HtmlService.createTemplateFromFile('html-response.html');
  template.title = title;
  template.displayMessage = displayMessage;
  template.bgColor = bgColor;
  template.SPREADSHEET_ID = SPREADSHEET_ID;

  const html = template.evaluate();
  html.setTitle(title).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
  return html;
}

function createCommentPromptPage(action, row, approverEmail, webAppUrl) {
  let actionText = "process";
  if (action.startsWith('approve_paid')) actionText = "approve with pay";
  else if (action.startsWith('approve_unpaid')) actionText = "approve without pay";
  else if (action === 'reject') actionText = "reject";
  else if (action === 'other_specify') actionText = "set status to 'Other' and specify details";
  const title = `Confirm Action & Add Comment (Row ${row})`;

  const template = HtmlService.createTemplateFromFile('comment-prompt-page.html');
  template.title = title;
  template.actionText = actionText;
  template.row = row;
  template.webAppUrl = webAppUrl;
  template.action = action;
  template.approverEmail = encodeURIComponent(approverEmail);

  const html = template.evaluate();
  html.setTitle(title).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
  return html;
}

// =========================================================================
// SHEET INTERACTION & DATA VALIDATION HELPERS
// =========================================================================
function getColumnIndices(sheet, sheetIdentifier) {
  if (!sheet || !sheetIdentifier) { Logger.log("Error: Invalid sheet or identifier for getColumnIndices."); return null; }
  const cacheKey = `columnIndices_${sheetIdentifier}_v2`;
  const cached = SCRIPT_CACHE.get(cacheKey);
  if (cached) { try { return JSON.parse(cached); } catch (e) { SCRIPT_CACHE.remove(cacheKey); } }
  try {
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const indices = {};
    headers.forEach((h, i) => { if (h && String(h).trim()) indices[String(h).trim().toLowerCase()] = i + 1; });
    if (Object.keys(indices).length === 0) { Logger.log(`WARN: No headers found in ${sheet.getName()}`); return null; }
    SCRIPT_CACHE.put(cacheKey, JSON.stringify(indices), 21600); // 6 hours
    return indices;
  } catch (e) { Logger.log(`Error reading headers from ${sheet.getName()}: ${e}`); notifyAdmin(`Absence Script: Error reading headers from ${sheet.getName()}: ${e.message}`); return null; }
}

function validateRequiredHeaders(indices, requiredHeaders, sheetName) {
  if (!indices) { Logger.log(`Header validation failed for "${sheetName}": Indices object is null.`); return false; }
  for (const header of requiredHeaders) {
    if (!indices[header.toLowerCase()]) {
      Logger.log(`Header validation failed for "${sheetName}": Required header "${header}" not found.`);
      notifyAdmin(`Absence Script CONFIG ERROR: Required header "${header}" not found in sheet "${sheetName}".`);
      return false;
    }
  }
  return true;
}

function isValidEmail(email) {
  if (!email || typeof email !== 'string') return false;
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email.trim());
}

/**
 * Counts absence requests for a specific date, categorized by approval status.
 * It tallies requests that are 'Approved', 'Approved - Admin Logged', or 'Other'
 * into an 'approved' count, and those 'Pending Approval' into a 'pending' count.
 *
 * @param {Sheet} sheet The master log sheet object.
 * @param {object} logIndices An object mapping header names to column indices for the log sheet.
 * @param {Date} targetRequestStartDate The specific start date to count requests for.
 * @returns {{approved: number, pending: number}} An object containing the counts of approved and pending requests.
 */
function countTotalRequestsForStartDate(sheet, logIndices, targetRequestStartDate) {
  const counts = {
    approved: 0,
    pending: 0
  };

  if (!sheet || !logIndices || !(targetRequestStartDate instanceof Date) || isNaN(targetRequestStartDate)) {
    Logger.log(`countTotalRequestsForStartDate: Invalid params. Sheet: ${sheet ? sheet.getName() : 'null'}, logIndices: ${logIndices ? 'present' : 'null'}, Target: ${targetRequestStartDate}`);
    return counts;
  }

  try {
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return counts;
    }

    const startDateColIdx = logIndices[HEADER_LOG_ABSENCE_START_DATE.toLowerCase()] - 1;
    const statusColIdx = logIndices[HEADER_LOG_APPROVAL_STATUS.toLowerCase()] - 1;

    if (isNaN(startDateColIdx) || isNaN(statusColIdx)) {
      notifyAdmin(`Absence Script Count ERROR: Could not find Start Date or Status column in "${sheet.getName()}".`);
      return counts;
    }

    const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

    const targetY = targetRequestStartDate.getFullYear();
    const targetM = targetRequestStartDate.getMonth();
    const targetD = targetRequestStartDate.getDate();

    for (const row of data) {
      const cellVal = row[startDateColIdx];

      if (cellVal) {
        const entryDate = combineDateAndTime(cellVal, null);

        if (entryDate && !isNaN(entryDate.getTime())) {
          // Compare year, month, and day directly to avoid timezone issues.
          if (targetY === entryDate.getFullYear() && targetM === entryDate.getMonth() && targetD === entryDate.getDate()) {
            const status = row[statusColIdx] ? String(row[statusColIdx]).trim() : '';
            if (status === 'Approved' || status === 'Approved - Admin Logged' || status === 'Other') {
              counts.approved++;
            } else if (status === 'Pending Approval') {
              counts.pending++;
            }
          }
        }
      }
    }
  } catch (e) {
    Logger.log(`Error in countTotalRequestsForStartDate (${sheet.getName()}): ${e}`);
    return { approved: 0, pending: 0 };
  }
  return counts;
}

/**
 * an 'approved' count, and those 'Pending Approval' into a 'pending' count.
 *
 * @param {Sheet} sheet The master log sheet object.
 * @param {object} logIndices An object mapping header names to column indices for the log sheet.
 * @param {Date} targetRequestStartDate The specific start date to count requests for.
 * @param {number} currentRowIndex The row index of the current request being processed, to exclude it from the lists.
 * @returns {{approved: Array<object>, pending: Array<object>}} An object containing arrays of detailed request info.
 */
function getAbsenceDetailsForDate(sheet, logIndices, targetRequestStartDate, currentRowIndex) {
  const details = {
    approved: [],
    pending: []
  };

  if (!sheet || !logIndices || !(targetRequestStartDate instanceof Date) || isNaN(targetRequestStartDate)) {
    return details;
  }

  try {
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      return details;
    }

    const requiredHeaders = [HEADER_LOG_STAFF_NAME, HEADER_LOG_ABSENCE_TYPE, HEADER_LOG_ABSENCE_START_TIME, HEADER_LOG_ABSENCE_END_TIME, HEADER_LOG_ABSENCE_START_DATE, HEADER_LOG_APPROVAL_STATUS];
    if (!validateRequiredHeaders(logIndices, requiredHeaders, sheet.getName())) {
      notifyAdmin(`Absence Script Count ERROR: Missing required headers for details gathering in "${sheet.getName()}".`);
      return details;
    }
    
    const nameColIdx = logIndices[HEADER_LOG_STAFF_NAME.toLowerCase()] - 1;
    const typeColIdx = logIndices[HEADER_LOG_ABSENCE_TYPE.toLowerCase()] - 1;
    const startTimeColIdx = logIndices[HEADER_LOG_ABSENCE_START_TIME.toLowerCase()] - 1;
    const endTimeColIdx = logIndices[HEADER_LOG_ABSENCE_END_TIME.toLowerCase()] - 1;
    const startDateColIdx = logIndices[HEADER_LOG_ABSENCE_START_DATE.toLowerCase()] - 1;
    const statusColIdx = logIndices[HEADER_LOG_APPROVAL_STATUS.toLowerCase()] - 1;
    
    const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

    const targetY = targetRequestStartDate.getFullYear();
    const targetM = targetRequestStartDate.getMonth();
    const targetD = targetRequestStartDate.getDate();

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const sheetRowIndex = i + 2; // i is 0-based, rows are 1-based, plus header.
      if (sheetRowIndex === currentRowIndex) continue; // Skip the request being processed

      const cellVal = row[startDateColIdx];

      if (cellVal) {
        const entryDate = combineDateAndTime(cellVal, null);

        if (entryDate && !isNaN(entryDate.getTime())) {
          if (targetY === entryDate.getFullYear() && targetM === entryDate.getMonth() && targetD === entryDate.getDate()) {
            const tz = Session.getScriptTimeZone();
            const startTimeObj = combineDateAndTime(entryDate, row[startTimeColIdx]);
            const endTimeObj = combineDateAndTime(entryDate, row[endTimeColIdx]);

            const requestInfo = {
              name: row[nameColIdx] || 'N/A',
              type: row[typeColIdx] || 'N/A',
              startTime: (startTimeObj && !isNaN(startTimeObj.getTime())) ? Utilities.formatDate(startTimeObj, tz, "HH:mm") : 'N/A',
              endTime: (endTimeObj && !isNaN(endTimeObj.getTime())) ? Utilities.formatDate(endTimeObj, tz, "HH:mm") : 'N/A',
              sortTime: (startTimeObj && !isNaN(startTimeObj.getTime())) ? startTimeObj : null
            };

            const status = row[statusColIdx] ? String(row[statusColIdx]).trim() : '';
            if (status === 'Approved' || status === 'Approved - Admin Logged' || status === 'Other') {
              details.approved.push(requestInfo);
            } else if (status === 'Pending Approval') {
              details.pending.push(requestInfo);
            }
          }
        }
      }
    }

    const sorter = (a, b) => {
        if (!a.sortTime) return 1;
        if (!b.sortTime) return -1;
        return a.sortTime.getTime() - b.sortTime.getTime();
    };

    details.approved.sort(sorter);
    details.pending.sort(sorter);

  } catch (e) {
    Logger.log(`Error in getAbsenceDetailsForDate (${sheet.getName()}): ${e}`);
    notifyAdmin(`Absence Script CRITICAL Failure in getAbsenceDetailsForDate: ${e.message}`);
    return { approved: [], pending: [] };
  }
  return details;
}

// =========================================================================
// STAFF DIRECTORY & USER INFORMATION HELPERS
// =========================================================================
function getStaffDirectoryInfoByEmail(staffEmail) {
  if (!staffEmail || !USE_DIRECTORY_LOOKUP || !STAFF_DIRECTORY_SHEET_ID) return null;
  staffEmail = String(staffEmail).trim().toLowerCase();
  let dirSheet;
  try {
    dirSheet = STAFF_DIRECTORY_SHEET_NAME ? SpreadsheetApp.openById(STAFF_DIRECTORY_SHEET_ID).getSheetByName(STAFF_DIRECTORY_SHEET_NAME) : SpreadsheetApp.openById(STAFF_DIRECTORY_SHEET_ID).getSheets()[0];
    if (!dirSheet) { Logger.log(`Dir sheet "${STAFF_DIRECTORY_SHEET_NAME || 'First'}" not found.`); return null; }
  } catch (e) { Logger.log(`Error accessing Dir Sheet: ${e}`); notifyAdmin(`Absence: Error accessing Staff Directory: ${e.message}`); return null; }
  try {
    const dirIndices = getColumnIndices(dirSheet, `dir_${STAFF_DIRECTORY_SHEET_ID}_${STAFF_DIRECTORY_SHEET_NAME || 'FirstSheet'}`);
    const reqDirHeaders = [HEADER_DIR_STAFF_EMAIL, HEADER_DIR_LM_EMAIL, HEADER_DIR_STATUS, HEADER_DIR_LEAVING_DATE, HEADER_DIR_STAFF_MEMBER_NAME, HEADER_DIR_TEACHING_SUPPORT, HEADER_DIR_DEPARTMENT, HEADER_DIR_MONDAY, HEADER_DIR_TUESDAY, HEADER_DIR_WEDNESDAY, HEADER_DIR_THURSDAY, HEADER_DIR_FRIDAY, HEADER_DIR_SATURDAY, HEADER_DIR_SUNDAY];
    if (!validateRequiredHeaders(dirIndices, reqDirHeaders, dirSheet.getName())) return null;
    const lastRow = dirSheet.getLastRow();
    if (lastRow < 2) return null;
    const data = dirSheet.getRange(2, 1, lastRow - 1, dirSheet.getLastColumn()).getValues();
    const emailCol_0 = dirIndices[HEADER_DIR_STAFF_EMAIL.toLowerCase()] - 1;
    for (let r = 0; r < data.length; r++) {
      const row = data[r];
      if (row[emailCol_0] && String(row[emailCol_0]).trim().toLowerCase() === staffEmail) {
        return {
          staffEmail: row[emailCol_0],
          lineManagerEmail: row[dirIndices[HEADER_DIR_LM_EMAIL.toLowerCase()] - 1],
          status: row[dirIndices[HEADER_DIR_STATUS.toLowerCase()] - 1],
          leavingDate: row[dirIndices[HEADER_DIR_LEAVING_DATE.toLowerCase()] - 1],
          staffMemberName: row[dirIndices[HEADER_DIR_STAFF_MEMBER_NAME.toLowerCase()] - 1],
          teachingOrSupport: row[dirIndices[HEADER_DIR_TEACHING_SUPPORT.toLowerCase()] - 1],
          department: row[dirIndices[HEADER_DIR_DEPARTMENT.toLowerCase()] - 1],
          days: {
            monday: row[dirIndices[HEADER_DIR_MONDAY.toLowerCase()] - 1],
            tuesday: row[dirIndices[HEADER_DIR_TUESDAY.toLowerCase()] - 1],
            wednesday: row[dirIndices[HEADER_DIR_WEDNESDAY.toLowerCase()] - 1],
            thursday: row[dirIndices[HEADER_DIR_THURSDAY.toLowerCase()] - 1],
            friday: row[dirIndices[HEADER_DIR_FRIDAY.toLowerCase()] - 1],
            saturday: row[dirIndices[HEADER_DIR_SATURDAY.toLowerCase()] - 1],
            sunday: row[dirIndices[HEADER_DIR_SUNDAY.toLowerCase()] - 1]
          }
        };
      }
    }
    return null;
  } catch (e) { Logger.log(`Error processing Dir Sheet: ${e}\n${e.stack}`); notifyAdmin(`Absence: Error processing Staff Directory for ${staffEmail}: ${e.message}`); return null; }
}

// =========================================================================
// DATE, TIME, & DURATION CALCULATION HELPERS
// =========================================================================

/**
 * Combines a date and time value into a single, valid JavaScript Date object.
 * Handles Date objects, DD/MM/YYYY strings, and other formats.
 *
 * @param {Date|string} dateValue The date, which can be a Date object or a string.
 * @param {Date|string|number} timeValue The time value.
 * @returns {Date} A valid Date object, or a Date object with a value of NaN if parsing fails.
 */
function combineDateAndTime(dateValue, timeValue) {
  if (!dateValue) return new Date(NaN);

  let baseDate;

  // 1. Best case: The value is already a valid Date object.
  if (dateValue instanceof Date && !isNaN(dateValue)) {
    baseDate = new Date(dateValue.getFullYear(), dateValue.getMonth(), dateValue.getDate());
  }
  // 2. Handle the 'DD/MM/YYYY' string format explicitly.
  else if (typeof dateValue === 'string' && /^\d{1,2}\/\d{1,2}\/\d{4}$/.test(dateValue.trim())) {
    const parts = dateValue.trim().split('/');
    const day = parseInt(parts[0], 10);
    const month = parseInt(parts[1], 10) - 1; // Month is 0-indexed in JavaScript
    const year = parseInt(parts[2], 10);
    baseDate = new Date(year, month, day);
  }
  // 3. Fallback for other string formats (like ISO YYYY-MM-DD).
  else {
    const parsedDate = new Date(dateValue);
    if (isNaN(parsedDate.getTime())) return new Date(NaN);
    baseDate = new Date(parsedDate.getFullYear(), parsedDate.getMonth(), parsedDate.getDate());
  }

  // If baseDate is still invalid after all attempts, exit.
  if (isNaN(baseDate.getTime())) return new Date(NaN);

  let h = 0, m = 0, s = 0;
  // Handle if time is already a valid Date object (e.g., from a sheet cell)
  if (timeValue instanceof Date && !isNaN(timeValue.getTime())) {
    h = timeValue.getHours(); m = timeValue.getMinutes(); s = timeValue.getSeconds();
  }
  // Handle if time is a string like "HH:mm" or "HH:mm:ss"
  else if (typeof timeValue === 'string' && timeValue.trim() !== '') {
    const parts = timeValue.split(':');
    if (parts.length >= 2) {
      h = parseInt(parts[0], 10); m = parseInt(parts[1], 10); if (parts.length === 3) s = parseInt(parts[2], 10);
      if (isNaN(h) || isNaN(m) || isNaN(s) || h < 0 || h > 23 || m < 0 || m > 59 || s < 0 || s > 59) return new Date(NaN);
    } else return new Date(NaN);
  }
  // Handle if time is a Google Sheets serial number (a fraction of a day)
  else if (typeof timeValue === 'number' && !isNaN(timeValue) && timeValue >= 0 && timeValue < 1) {
    const totalSec = timeValue * 86400;
    h = Math.floor(totalSec / 3600); m = Math.floor((totalSec % 3600) / 60); s = Math.floor(totalSec % 60);
  }
  // Handle if time is empty/null, defaulting to midnight
  else if (timeValue === null || timeValue === undefined || (typeof timeValue === 'string' && timeValue.trim() === '')) {
    h = 0; m = 0; s = 0;
  }
  // If none of the above, the format is invalid
  else return new Date(NaN);

  baseDate.setHours(h, m, s, 0);
  return baseDate;
}

function doDateRangesOverlap(startA, endA, startB, endB) {
  if (!(startA instanceof Date && !isNaN(startA)) || !(endA instanceof Date && !isNaN(endA)) || !(startB instanceof Date && !isNaN(startB)) || !(endB instanceof Date && !isNaN(endB))) return false;
  if (startA > endA || startB > endB) return false; // Invalid ranges
  return (startA < endB) && (endA > startB);
}

function calculateAbsenceDays(staffEmail, startDateTime, endDateTime) {
  if (!staffEmail || !(startDateTime instanceof Date) || isNaN(startDateTime) || !(endDateTime instanceof Date) || isNaN(endDateTime)) throw new Error('Invalid params for calculateAbsenceDays.');
  if (endDateTime <= startDateTime) return 0;
  const staffInfo = getStaffDirectoryInfoByEmail(staffEmail);
  if (!staffInfo || !staffInfo.days) throw new Error(`Staff '${staffEmail}' not found or schedule missing for calculateAbsenceDays.`);
  let totalDaysMissed = 0;
  let currentDateIter = new Date(startDateTime.getFullYear(), startDateTime.getMonth(), startDateTime.getDate());
  const finalBoundary = new Date(endDateTime.getFullYear(), endDateTime.getMonth(), endDateTime.getDate());
  finalBoundary.setHours(23, 59, 59, 999);

  // Iterate through each calendar day within the absence period.
  while (currentDateIter <= finalBoundary && currentDateIter < endDateTime) {
    const dayOfWeek = Utilities.formatDate(currentDateIter, Session.getScriptTimeZone(), 'EEEE').toLowerCase();
    const workStr = staffInfo.days[dayOfWeek];

    // Handle new split-shift format "HH:mm-HH:mm, HH:mm-HH:mm"
    if (workStr && typeof workStr === 'string' && workStr.trim() !== '') {
      const shiftStrings = workStr.split(',').map(s => s.trim());
      let totalScheduledHoursForDay = 0;
      let totalMissedHoursOnDay = 0;
      const shifts = [];

      // First, parse all valid shifts and calculate the total scheduled hours for the day.
      for (const shiftStr of shiftStrings) {
        // Basic validation for the "HH:mm-HH:mm" format.
        if (shiftStr.match(/^\d{1,2}:\d{2}-\d{1,2}:\d{2}$/)) {
          try {
            const parsed = parseDayTimeString(currentDateIter, shiftStr);
            if (parsed.schedEnd > parsed.schedStart) {
              shifts.push(parsed);
              totalScheduledHoursForDay += (parsed.schedEnd.getTime() - parsed.schedStart.getTime()) / 3600000;
            }
          } catch (e) {
            // Log or ignore invalid parts of the string. For robustness, we'll ignore.
            Logger.log(`Skipping invalid time segment "${shiftStr}" for ${staffEmail} on ${dayOfWeek}.`);
          }
        }
      }

      // If there are any valid shifts, calculate the portion of the workday missed.
      if (totalScheduledHoursForDay > 0) {
        for (const shift of shifts) {
          // Calculate the actual overlap between the absence period and this specific shift.
          const effectiveStart = new Date(Math.max(shift.schedStart.getTime(), startDateTime.getTime()));
          const effectiveEnd = new Date(Math.min(shift.schedEnd.getTime(), endDateTime.getTime()));

          // If there's a positive overlap, add the missed hours.
          if (effectiveEnd > effectiveStart) {
            totalMissedHoursOnDay += (effectiveEnd.getTime() - effectiveStart.getTime()) / 3600000;
          }
        }

        // Add the fraction of the day that was missed to the total.
        if (totalMissedHoursOnDay > 0) {
          totalDaysMissed += (totalMissedHoursOnDay / totalScheduledHoursForDay);
        }
      }
    }

    currentDateIter.setDate(currentDateIter.getDate() + 1);
  }
  return parseFloat(totalDaysMissed.toFixed(4));
}

/**
 * Calculates the total number of scheduled work hours missed during an absence period.
 * This function iterates through each day of the absence, parses the staff member's
 * daily schedule (including split shifts), and sums the duration of any overlap
 * between the scheduled work time and the absence period.
 *
 * @param {string} staffEmail The email of the staff member.
 * @param {Date} startDateTime The start date and time of the absence.
 * @param {Date} endDateTime The end date and time of the absence.
 * @returns {number} The total number of missed work hours, rounded to 4 decimal places.
 * @throws {Error} If staff member is not found or their schedule is missing.
 */
function calculateAbsenceHoursFromSchedule(staffEmail, startDateTime, endDateTime) {
  if (!staffEmail || !(startDateTime instanceof Date) || isNaN(startDateTime) || !(endDateTime instanceof Date) || isNaN(endDateTime)) {
    throw new Error('Invalid params for calculateAbsenceHoursFromSchedule.');
  }
  if (endDateTime <= startDateTime) return 0;

  const staffInfo = getStaffDirectoryInfoByEmail(staffEmail);
  if (!staffInfo || !staffInfo.days) {
    throw new Error(`Staff '${staffEmail}' not found or schedule missing for calculateAbsenceHoursFromSchedule.`);
  }

  let totalMissedHours = 0;
  let currentDateIter = new Date(startDateTime.getFullYear(), startDateTime.getMonth(), startDateTime.getDate());
  const finalBoundary = new Date(endDateTime.getFullYear(), endDateTime.getMonth(), endDateTime.getDate());
  finalBoundary.setHours(23, 59, 59, 999);

  // Iterate through each calendar day within the absence period.
  while (currentDateIter <= finalBoundary && currentDateIter < endDateTime) {
    const dayOfWeek = Utilities.formatDate(currentDateIter, Session.getScriptTimeZone(), 'EEEE').toLowerCase();
    const workStr = staffInfo.days[dayOfWeek];

    if (workStr && typeof workStr === 'string' && workStr.trim() !== '') {
      const shiftStrings = workStr.split(',').map(s => s.trim());
      
      for (const shiftStr of shiftStrings) {
        if (shiftStr.match(/^\d{1,2}:\d{2}-\d{1,2}:\d{2}$/)) {
          try {
            const parsed = parseDayTimeString(currentDateIter, shiftStr);
            if (parsed.schedEnd > parsed.schedStart) {
              // Calculate the actual overlap between the absence period and this specific shift.
              const effectiveStart = new Date(Math.max(parsed.schedStart.getTime(), startDateTime.getTime()));
              const effectiveEnd = new Date(Math.min(parsed.schedEnd.getTime(), endDateTime.getTime()));

              // If there's a positive overlap, add the missed hours.
              if (effectiveEnd > effectiveStart) {
                totalMissedHours += (effectiveEnd.getTime() - effectiveStart.getTime()) / 3600000;
              }
            }
          } catch (e) {
            Logger.log(`Skipping invalid time segment "${shiftStr}" for ${staffEmail} on ${dayOfWeek} during hour calculation.`);
          }
        }
      }
    }
    currentDateIter.setDate(currentDateIter.getDate() + 1);
  }
  return parseFloat(totalMissedHours.toFixed(4));
}

function parseDayTimeString(date, timeString) {
  const match = String(timeString).match(/^(\d{1,2}):(\d{2})-(\d{1,2}):(\d{2})$/);
  if (!match) throw new Error(`Invalid time string: "${timeString}".`);
  const sH = parseInt(match[1], 10), sM = parseInt(match[2], 10);
  const eH = parseInt(match[3], 10), eM = parseInt(match[4], 10);
  const schedStart = new Date(date); schedStart.setHours(sH, sM, 0, 0);
  const schedEnd = new Date(date); schedEnd.setHours(eH, eM, 0, 0);
  return { schedStart, schedEnd };
}

function calculateAbsenceHistoryCategorized(sheetName, staffEmail, currentRequestStartDate) {
  let totalDays = 0; const catTotals = {};
  let summary = `No approved absences found starting in 12 months prior to ${Utilities.formatDate(currentRequestStartDate, Session.getScriptTimeZone(), "dd/MM/yyyy")} within '${sheetName}'.\n`;
  if (!sheetName || !staffEmail || !(currentRequestStartDate instanceof Date) || isNaN(currentRequestStartDate)) return "Error: Invalid input for history.";
  staffEmail = String(staffEmail).trim().toLowerCase();
  let histSheet;
  try {
    histSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
    if (!histSheet || histSheet.getLastRow() < 2) return summary;
    const histIndices = getColumnIndices(histSheet, `hist_${sheetName}`);
    const reqHistHeaders = [HEADER_HIST_STAFF_EMAIL, HEADER_HIST_START_DATE, HEADER_HIST_DURATION_DAYS, HEADER_HIST_ABSENCE_TYPE];
    if (!validateRequiredHeaders(histIndices, reqHistHeaders, sheetName)) return `Error: Missing headers in ${sheetName}.`;
    const emailCol = histIndices[HEADER_HIST_STAFF_EMAIL.toLowerCase()] - 1, startCol = histIndices[HEADER_HIST_START_DATE.toLowerCase()] - 1, durCol = histIndices[HEADER_HIST_DURATION_DAYS.toLowerCase()] - 1, typeCol = histIndices[HEADER_HIST_ABSENCE_TYPE.toLowerCase()] - 1;
    const data = histSheet.getRange(2, 1, histSheet.getLastRow() - 1, histSheet.getLastColumn()).getValues();
    const periodEnd = new Date(currentRequestStartDate); periodEnd.setHours(0, 0, 0, 0);
    const periodStart = new Date(currentRequestStartDate); periodStart.setFullYear(periodStart.getFullYear() - 1); periodStart.setHours(0, 0, 0, 0);

    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      if (!row[emailCol] || String(row[emailCol]).trim().toLowerCase() !== staffEmail) continue;
      const recStart = new Date(row[startCol]);
      if (!(recStart instanceof Date) || isNaN(recStart)) continue;
      recStart.setHours(0, 0, 0, 0);
      if (recStart >= periodStart && recStart < periodEnd) {
        const dur = (typeof row[durCol] === 'number' && !isNaN(row[durCol]) && row[durCol] > 0) ? row[durCol] : 0;
        if (dur > 0) {
          totalDays += dur;
          const type = String(row[typeCol] || 'Uncategorized').trim();
          catTotals[type] = (catTotals[type] || 0) + dur;
        }
      }
    }
    if (totalDays > 0) {
      summary = `Total Days in Period: ${totalDays.toFixed(2)}\nBreakdown by Type:\n`;
      Object.keys(catTotals).sort().forEach(cat => {
        summary += ` - ${cat}: ${catTotals[cat].toFixed(2).replace(/\.00$/, '')} day(s)\n`;
      });
    }
  } catch (e) { summary = `Error processing history from '${sheetName}': ${e.message}`; notifyAdmin(`Absence: Error calculating history from ${sheetName}: ${e.message}`); }
  return summary.trim();
}

// =========================================================================
// NOTIFICATION & EXTERNAL SERVICE HELPERS
// =========================================================================
function notifyAdmin(message, subjectPrefix = "Absence System Notification") {
  try {
    if (ADMIN_EMAIL_FOR_ERRORS && isValidEmail(ADMIN_EMAIL_FOR_ERRORS)) {
      const mailOptions = {
        to: ADMIN_EMAIL_FOR_ERRORS,
        subject: `${subjectPrefix}: ${message.substring(0, 80)}`,
        body: message
      };
      if (SENDER_EMAIL && isValidEmail(SENDER_EMAIL)) {
        mailOptions.from = SENDER_EMAIL;
        mailOptions.name = SENDER_NAME;
      }
      MailApp.sendEmail(mailOptions);
    } else if (ADMIN_EMAIL_FOR_ERRORS) Logger.log(`Admin notify FAIL: invalid email '${ADMIN_EMAIL_FOR_ERRORS}'. Msg: ${message}`);
    else Logger.log(`Admin notify SKIP: no email configured. Msg: ${message}`);
  } catch (e) { Logger.log(`CRITICAL: Admin notify send FAIL: ${e}. Original msg: ${message}`); }
}

function createCalendarEvent(staffName, startTime, endTime, absenceType, approverEmail, approverComment = "") {
  try {
    if (!(startTime instanceof Date && !isNaN(startTime)) || !(endTime instanceof Date && !isNaN(endTime))) {
      Logger.log(`Calendar Error for "${staffName}": Invalid dates. Start: ${startTime}, End: ${endTime}`);
      notifyAdmin(`Absence Calendar Error: Invalid dates for ${staffName}.`); return;
    }
    if (endTime <= startTime) { Logger.log(`Calendar WARN for "${staffName}": End <= Start. Event not created.`); return; }

    const title = `Absence: ${staffName}`;
    let desc = `Staff: ${staffName}\nType: ${absenceType || 'N/A'}\nProcessed by: ${approverEmail || 'N/A'}`;
    if (approverComment && String(approverComment).trim()) desc += `\nComment: ${String(approverComment).trim()}`;

    let cal; const effUser = Session.getEffectiveUser().getEmail();
    if (TARGET_CALENDAR_ID && String(TARGET_CALENDAR_ID).trim()) {
      cal = CalendarApp.getCalendarById(String(TARGET_CALENDAR_ID).trim());
      if (!cal) { Logger.log(`Calendar Error: Cannot get cal ID "${TARGET_CALENDAR_ID}". User: ${effUser}. Fallback to default.`); cal = CalendarApp.getDefaultCalendar(); }
    } else { cal = CalendarApp.getDefaultCalendar(); Logger.log(`Using default calendar for ${effUser}.`); }
    if (!cal) { Logger.log(`Calendar Error: No calendar found for ${effUser}.`); notifyAdmin(`Absence Calendar Error: No calendar for ${effUser} to log ${staffName}.`); return; }

    // An event is "all-day" if it starts and ends at midnight on different days.
    const isAllDay = startTime.getHours() === 0 && startTime.getMinutes() === 0 && startTime.getSeconds() === 0 &&
      endTime.getHours() === 0 && endTime.getMinutes() === 0 && endTime.getSeconds() === 0 &&
      (endTime.getTime() > startTime.getTime()) && ((endTime.getTime() - startTime.getTime()) % 86400000 === 0);

    let event;
    if (isAllDay) {
      // For Google Calendar, the end date for all-day events is exclusive.
      event = cal.createAllDayEvent(title, startTime, endTime, { description: desc });
      Logger.log(`Created ALL-DAY event for ${staffName} from ${startTime.toLocaleDateString()} to ${new Date(endTime.getTime() - 86400000).toLocaleDateString()}. ID: ${event.getId()}`);
    } else {
      event = cal.createEvent(title, startTime, endTime, { description: desc });
      Logger.log(`Created event for ${staffName} from ${startTime.toLocaleString()} to ${endTime.toLocaleString()}. ID: ${event.getId()}`);
    }
  } catch (e) { Logger.log(`Calendar Create Error for ${staffName}: ${e}\n${e.stack}`); notifyAdmin(`Absence Calendar Error for ${staffName}: ${e.message}`); }
}

function extractFileIdFromDriveUrl(urlOrId) {
  if (!urlOrId || typeof urlOrId !== 'string') return null;
  urlOrId = String(urlOrId).trim();

  // First, check if the input is just a valid-looking ID itself.
  if (/^[a-zA-Z0-9_-]{25,50}$/.test(urlOrId)) {
    try { DriveApp.getFileById(urlOrId); return urlOrId; } catch (e) { /* Not a valid ID, continue to URL parsing */ }
  }

  // Test against various known Google Drive/Docs URL formats.
  const regexes = [
    /drive\.google\.com\/file\/d\/([a-zA-Z0-9_-]+)/,         // Standard file viewer URL
    /drive\.google\.com\/open\?id=([a-zA-Z0-9_-]+)/,         // Older "open" sharing URL
    /drive\.google\.com\/uc\?id=([a-zA-Z0-9_-]+)/,           // Download/user content URL
    /docs\.google\.com\/(?:document|spreadsheets|presentation|file|forms)\/d\/([a-zA-Z0-9_-]+)/ // G-Suite editor URLs
  ];
  for (const regex of regexes) {
    const match = urlOrId.match(regex);
    if (match && match[1]) return match[1];
  }
  return null;
}

// =========================================================================
// FORM MANAGEMENT HELPERS
// =========================================================================

/**
 * Updates the staff email dropdown on the Admin Form with a current list of active staff.
 * This function should be run on a daily time-driven trigger.
 */
function updateAdminFormStaffList() {
  try {
    if (!ADMIN_FORM_ID) {
      throw new Error("ADMIN_FORM_ID is not configured in the script.");
    }

    const activeEmails = _getActiveStaffEmailsFromDirectory();
    if (!activeEmails || activeEmails.length === 0) {
      Logger.log("No active staff emails found in the directory. The form will not be updated.");
      notifyAdmin("Absence System Warning: Could not find any active staff in the directory to update the Admin Form dropdown.");
      return;
    }

    const form = FormApp.openById(ADMIN_FORM_ID);
    const questionTitle = HEADER_ADMIN_FORM_EMAIL; // Uses the existing constant
    const questionItem = _findQuestionByTitle(form, questionTitle);

    if (!questionItem) {
      throw new Error(`Could not find a question with the exact title "${questionTitle}" in the Admin Form.`);
    }

    // Ensure the question is a dropdown (List) or multiple choice type
    const itemType = questionItem.getType();
    if (itemType !== FormApp.ItemType.LIST && itemType !== FormApp.ItemType.MULTIPLE_CHOICE) {
      throw new Error(`The question "${questionTitle}" is not a Dropdown or Multiple Choice. Its type is ${itemType}.`);
    }

    // Update the choices in the dropdown list
    questionItem.asListItem().setChoiceValues(activeEmails);

    Logger.log(`Successfully updated the Admin Form dropdown with ${activeEmails.length} staff emails.`);

  } catch (error) {
    Logger.log(`ERROR in updateAdminFormStaffList: ${error.message}\nStack: ${error.stack}`);
    notifyAdmin(`Absence Script CRITICAL Failure: Failed to update Admin Form dropdown. Error: ${error.message}`);
  }
}

/**
 * Fetches a sorted list of email addresses for all "Active" staff from the directory.
 * @returns {string[]|null} An array of email strings, or null if an error occurs.
 * @private
 */
function _getActiveStaffEmailsFromDirectory() {
  if (!USE_DIRECTORY_LOOKUP || !STAFF_DIRECTORY_SHEET_ID) return null;

  let dirSheet;
  try {
    dirSheet = STAFF_DIRECTORY_SHEET_NAME ? SpreadsheetApp.openById(STAFF_DIRECTORY_SHEET_ID).getSheetByName(STAFF_DIRECTORY_SHEET_NAME) : SpreadsheetApp.openById(STAFF_DIRECTORY_SHEET_ID).getSheets()[0];
    if (!dirSheet) {
      Logger.log(`Directory sheet "${STAFF_DIRECTORY_SHEET_NAME || 'First'}" not found.`);
      return null;
    }
  } catch (e) {
    Logger.log(`Error accessing Directory Sheet: ${e}`);
    return null;
  }

  const dirIndices = getColumnIndices(dirSheet, `dir_${STAFF_DIRECTORY_SHEET_ID}_${STAFF_DIRECTORY_SHEET_NAME || 'FirstSheet'}`);
  const reqDirHeaders = [HEADER_DIR_STAFF_EMAIL, HEADER_DIR_STATUS];
  if (!validateRequiredHeaders(dirIndices, reqDirHeaders, dirSheet.getName())) return null;

  const lastRow = dirSheet.getLastRow();
  if (lastRow < 2) return [];

  const data = dirSheet.getRange(2, 1, lastRow - 1, dirSheet.getLastColumn()).getValues();
  const emailColIdx = dirIndices[HEADER_DIR_STAFF_EMAIL.toLowerCase()] - 1;
  const statusColIdx = dirIndices[HEADER_DIR_STATUS.toLowerCase()] - 1;

  const emails = data
    .filter(row => {
      const status = row[statusColIdx] ? String(row[statusColIdx]).trim().toLowerCase() : '';
      const email = row[emailColIdx] ? String(row[emailColIdx]).trim() : '';
      // Include row only if status is 'active' and email is valid
      return status === 'active' && isValidEmail(email);
    })
    .map(row => String(row[emailColIdx]).trim());

  return emails.sort(); // Return a sorted list for a user-friendly dropdown
}

/**
 * Finds a question item in a Google Form by its exact title.
 * @param {Form} form The Google Form object.
 * @param {string} title The exact title of the question to find.
 * @returns {Item|null} The form item object, or null if not found.
 * @private
 */
function _findQuestionByTitle(form, title) {
  const items = form.getItems();
  for (const item of items) {
    if (item.getTitle() === title) {
      return item;
    }
  }
  return null;
}

// =========================================================================
// ARCHIVING
// =========================================================================

/**
 * Archives old, completed requests from the main log to an archive sheet.
 * This version includes UI prompts and is intended for manual execution.
 * Use this when running manually from the menu or script editor.
 */
function archiveOldRequests() {
  const ui = SpreadsheetApp.getUi();
  const result = ui.alert(
    'Confirm Archiving',
    `This will move all completed requests older than ${ARCHIVE_OLDER_THAN_DAYS} days from "${REQUEST_LOG_SHEET_NAME}" to "${ARCHIVE_SHEET_NAME}".\n\nThis action cannot be undone. Are you sure you want to proceed?`,
    ui.ButtonSet.YES_NO
  );

  if (result !== ui.Button.YES) {
    ui.alert('Archiving Cancelled', 'No changes have been made.', ui.ButtonSet.OK);
    return;
  }

  try {
    const result = _performArchiving();
    ui.alert('Archiving Complete', result.message, ui.ButtonSet.OK);
    Logger.log(result.message);
  } catch (error) {
    Logger.log(`ERROR in archiveOldRequests: ${error.message}\nStack: ${error.stack}`);
    notifyAdmin(`Absence Script CRITICAL Failure: Failed to archive old requests. Error: ${error.message}`);
    ui.alert('Archiving Failed', `An error occurred: ${error.message}. Please check the logs.`, ui.ButtonSet.OK);
  }
}

/**
 * Archives old requests automatically without user prompts.
 * This version is designed to run on time triggers.
 * It will archive requests automatically and send email notifications about the results.
 */
function archiveOldRequestsAutomatically() {
  try {
    Logger.log(`Starting automatic archiving of requests older than ${ARCHIVE_OLDER_THAN_DAYS} days...`);
    const result = _performArchiving();
    
    // Log the result
    Logger.log(result.message);
    
    // Optionally notify admin of successful archiving (only if requests were actually archived)
    if (result.archivedCount > 0) {
      notifyAdmin(`Automatic Archiving Complete: ${result.message}`, "Absence System - Automatic Archiving");
    }
    
  } catch (error) {
    Logger.log(`ERROR in archiveOldRequestsAutomatically: ${error.message}\nStack: ${error.stack}`);
    notifyAdmin(`Absence Script CRITICAL Failure: Automatic archiving failed. Error: ${error.message}`);
  }
}

/**
 * Core archiving logic that both manual and automatic functions can use.
 * This function contains no UI elements and returns results that can be handled differently.
 * @returns {object} Object containing success status, message, and archived count
 * @private
 */
function _performArchiving() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName(REQUEST_LOG_SHEET_NAME);
  if (!logSheet) {
    throw new Error(`Log sheet "${REQUEST_LOG_SHEET_NAME}" not found.`);
  }

  let archiveSheet = ss.getSheetByName(ARCHIVE_SHEET_NAME);
  if (!archiveSheet) {
    archiveSheet = ss.insertSheet(ARCHIVE_SHEET_NAME);
    const headers = logSheet.getRange(1, 1, 1, logSheet.getLastColumn()).getValues();
    archiveSheet.getRange(1, 1, 1, headers[0].length).setValues(headers);
    archiveSheet.setFrozenRows(1);
    Logger.log(`Created archive sheet: "${ARCHIVE_SHEET_NAME}"`);
  } else {
    // Ensure archive sheet has the same column structure as the log sheet
    const logHeaders = logSheet.getRange(1, 1, 1, logSheet.getLastColumn()).getValues()[0];
    const archiveHeaders = archiveSheet.getRange(1, 1, 1, archiveSheet.getLastColumn()).getValues()[0];
    
    // If column counts don't match, update the archive sheet headers
    if (logHeaders.length !== archiveHeaders.length) {
      archiveSheet.getRange(1, 1, 1, logHeaders.length).setValues([logHeaders]);
      Logger.log(`Updated archive sheet headers to match log sheet (${logHeaders.length} columns)`);
    }
  }

  const logIndices = getColumnIndices(logSheet, REQUEST_LOG_SHEET_NAME);
  if (!logIndices) {
    throw new Error(`Could not get column indices for "${REQUEST_LOG_SHEET_NAME}".`);
  }

  const requiredArchiveHeaders = [HEADER_LOG_ABSENCE_START_DATE, HEADER_LOG_APPROVAL_STATUS];
  if (!validateRequiredHeaders(logIndices, requiredArchiveHeaders, REQUEST_LOG_SHEET_NAME)) {
    throw new Error("Missing required headers for archiving.");
  }
  
  const absenceStartDateColIdx = logIndices[HEADER_LOG_ABSENCE_START_DATE.toLowerCase()] - 1;
  const statusColIdx = logIndices[HEADER_LOG_APPROVAL_STATUS.toLowerCase()] - 1;

  const lastRow = logSheet.getLastRow();
  if (lastRow < 2) {
    return {
      success: true,
      message: 'The log sheet has no data to archive.',
      archivedCount: 0
    };
  }

  // Get ALL data from the log sheet (all columns)
  const dataRange = logSheet.getRange(2, 1, lastRow - 1, logSheet.getLastColumn());
  const allData = dataRange.getValues();

  const archiveThreshold = new Date();
  archiveThreshold.setDate(archiveThreshold.getDate() - ARCHIVE_OLDER_THAN_DAYS);
  archiveThreshold.setHours(0, 0, 0, 0);

  const rowsToArchive = [];

  // First pass: identify which rows need to be archived
  allData.forEach((row, index) => {
    const status = row[statusColIdx] ? String(row[statusColIdx]).trim() : '';
    const absenceDateVal = row[absenceStartDateColIdx];
    const absenceDate = combineDateAndTime(absenceDateVal, null); // Use helper to parse date robustly

    const isFinalStatus = status && status !== 'Pending Approval';
    const isOldEnough = absenceDate && !isNaN(absenceDate) && absenceDate < archiveThreshold;

    if (isFinalStatus && isOldEnough) {
      rowsToArchive.push({
        data: row,
        originalRowIndex: index + 2 // Convert to 1-based row number (accounting for header)
      });
    }
  });

  if (rowsToArchive.length > 0) {
    // Step 1: Append archived rows to archive sheet
    const archiveStartRow = archiveSheet.getLastRow() + 1;
    const rowDataToArchive = rowsToArchive.map(item => item.data);
    const numCols = Math.max(rowDataToArchive[0].length, logSheet.getLastColumn());
    archiveSheet.getRange(archiveStartRow, 1, rowDataToArchive.length, numCols).setValues(rowDataToArchive);
    Logger.log(`Added ${rowDataToArchive.length} rows to archive sheet starting at row ${archiveStartRow}`);
    
    // Step 2: Extract row indices and sort them in descending order for deletion
    const rowIndicesToDelete = rowsToArchive.map(item => item.originalRowIndex);
    rowIndicesToDelete.sort((a, b) => b - a); // Sort descending (delete from bottom up)
    
    Logger.log(`Deleting ${rowIndicesToDelete.length} rows from log sheet: ${rowIndicesToDelete.join(', ')}`);
    
    // Step 3: Delete rows from log sheet (from bottom to top to avoid row shifting issues)
    let deletedCount = 0;
    rowIndicesToDelete.forEach(rowIndex => {
      try {
        logSheet.deleteRow(rowIndex);
        deletedCount++;
      } catch (deleteError) {
        Logger.log(`Warning: Could not delete row ${rowIndex}: ${deleteError.message}`);
        // Continue with other deletions even if one fails
      }
    });
    
    const message = `Successfully archived ${rowDataToArchive.length} request(s) and deleted ${deletedCount} rows from the log sheet.`;
    Logger.log(`Archived ${rowDataToArchive.length} rows with ${numCols} columns each and deleted ${deletedCount} corresponding rows from log sheet.`);
    
    return {
      success: true,
      message: message,
      archivedCount: rowDataToArchive.length
    };
  } else {
    return {
      success: true,
      message: 'No requests met the archiving criteria.',
      archivedCount: 0
    };
  }
}

/**
 * Sets up a time-driven trigger to run automatic archiving daily.
 * Run this function once to install the trigger.
 */
function setupAutomaticArchivingTrigger() {
  // Delete any existing archiving triggers first
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'archiveOldRequestsAutomatically') {
      ScriptApp.deleteTrigger(trigger);
      Logger.log('Deleted existing automatic archiving trigger');
    }
  });

  // Create new daily trigger at 2 AM
  ScriptApp.newTrigger('archiveOldRequestsAutomatically')
    .timeBased()
    .everyDays(1)
    .atHour(2) // 2 AM
    .create();
    
  Logger.log('Created automatic archiving trigger to run daily at 2 AM');
  
  // Optionally notify admin
  notifyAdmin('Automatic archiving trigger has been set up to run daily at 2 AM', 'Absence System - Trigger Setup');
}

/**
 * Removes the automatic archiving trigger.
 * Run this function if you want to disable automatic archiving.
 */
function removeAutomaticArchivingTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  let removed = 0;
  
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'archiveOldRequestsAutomatically') {
      ScriptApp.deleteTrigger(trigger);
      removed++;
    }
  });
  
  if (removed > 0) {
    Logger.log(`Removed ${removed} automatic archiving trigger(s)`);
    notifyAdmin(`Removed ${removed} automatic archiving trigger(s)`, 'Absence System - Trigger Removed');
  } else {
    Logger.log('No automatic archiving triggers found to remove');
  }
}

// =========================================================================
// SCRIPT MANAGEMENT & ADMIN UI
// =========================================================================

/**
 * Flushes all known column index caches used by the script.
 * This is essential to run after changing the column order on any of the
 * tracked sheets (Master Log, Raw Response sheets, Directory, or History).
 */
function flushColumnIndexCache() {
    const ui = SpreadsheetApp.getUi();
    const cache = CacheService.getScriptCache();
    if (!cache) {
        ui.alert("Cache Error", "Cannot access the script cache service.", ui.ButtonSet.OK);
        return;
    }

    // --- Define all cache keys used in the script ---
    const keysToFlush = [
        // Master Log Sheet
        `columnIndices_${REQUEST_LOG_SHEET_NAME}_v2`,
        // Raw Response Sheets (NEWLY ADDED)
        `columnIndices_${STAFF_RESPONSE_SHEET}_v2`,
        `columnIndices_${ADMIN_RESPONSE_SHEET}_v2`,
        // History Sheet
        `hist_${HISTORY_SHEET_NAME}`,
        // Legacy keys for good measure
        `columnIndices_${REQUEST_LOG_SHEET_NAME}`
    ];

    // Conditionally add the directory sheet key
    if (USE_DIRECTORY_LOOKUP && STAFF_DIRECTORY_SHEET_ID) {
        const dirSheetIdentifier = `dir_${STAFF_DIRECTORY_SHEET_ID}_${STAFF_DIRECTORY_SHEET_NAME || 'FirstSheet'}`;
        keysToFlush.push(dirSheetIdentifier);
    }

    // --- Attempt to remove each key ---
    let removed = 0;
    let failed = 0;
    Logger.log(`Attempting to flush the following keys: ${keysToFlush.join(', ')}`);

    keysToFlush.forEach(key => {
        try {
            // Check if the key exists before trying to remove it
            if (cache.get(key) !== null) {
                cache.remove(key);
                // Verify removal
                if (cache.get(key) === null) {
                    removed++;
                    Logger.log(`Successfully removed key: ${key}`);
                } else {
                    failed++;
                    Logger.log(`FAILED to confirm removal of key: ${key}`);
                }
            }
        } catch (e) {
            failed++;
            Logger.log(`Error removing cache key ${key}: ${e}`);
        }
    });

    let msg = `Cache flush complete.\n\n${removed} key(s) successfully removed.`;
    if (failed > 0) {
        msg += `\n${failed} key(s) failed to confirm removal (check script logs for details).`;
    }
    ui.alert('Cache Flush Results', msg, ui.ButtonSet.OK);
    Logger.log(msg);
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Admin Absence Tools')
    .addItem('Update Admin Form Staff List', 'updateAdminFormStaffList')
    .addItem('Flush Column Indices Cache', 'flushColumnIndexCache')
    .addSeparator()
    .addItem('Archive Old Requests (Manual)', 'archiveOldRequests')
    .addItem('Setup Automatic Archiving', 'setupAutomaticArchivingTrigger')
    .addItem('Remove Automatic Archiving', 'removeAutomaticArchivingTrigger')
    .addToUi();
}
