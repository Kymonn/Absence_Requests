/**
 * @fileoverview This is the loader script bound to the spreadsheet.
 * It creates the Admin UI and calls the main logic from the CoreLogic library.
 * 
 * This script should replace the current Code.js in your spreadsheet's Apps Script editor.
 * Before using this, you must:
 * 1. Deploy the standalone script as a library
 * 2. Add the library to this project with identifier "CoreLogic"
 * 3. Run setupAllTriggers() from the Admin menu to create required triggers
 */

function onOpen() {
  // This function runs automatically when the spreadsheet is opened.
  // It calls the onOpen function in our library to create the menu.
  CoreLogic.onOpen();
}

// --- Menu Function Stubs ---
// These functions are called by the menu items. They simply pass the call
// along to the corresponding function in the CoreLogic library.

function updateAbsenceTypeDropdownsForCurrentSite() {
  CoreLogic.updateAbsenceTypeDropdownsForCurrentSite();
}

function updateAdminFormStaffListForCurrentSite() {
  CoreLogic.updateAdminFormStaffListForCurrentSite();
}

function flushColumnIndexCache() {
  CoreLogic.flushColumnIndexCache();
}

function rerunProcessingOnRow() {
  CoreLogic.rerunProcessingOnRow();
}

function archiveOldRequests() {
  CoreLogic.archiveOldRequests();
}

// --- Trigger Management Function Stubs ---

function setupAllTriggers() {
  CoreLogic.setupAllTriggers();
}

function deleteAllTriggers() {
  CoreLogic.deleteAllTriggers();
}

// --- Form Submission Handler ---
// This function must remain in the bound script to handle form submissions

function handleFormSubmissions(e) {
  // Pass the form submission event to the library for processing
  CoreLogic.handleFormSubmissions(e);
}

// --- Edit Handler ---
// This function must remain in the bound script to handle sheet edits

function onBromcomEntryEdit(e) {
  // Pass the edit event to the library for processing
  CoreLogic.onBromcomEntryEdit(e);
}

// --- Time-driven Function Stubs ---
// These will be called by triggers set up in the standalone script

function archiveOldRequestsAutomatically() {
  CoreLogic.archiveOldRequestsAutomatically();
} 
