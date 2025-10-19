/* global SpreadsheetApp, fParseA1Notation, fShowMessage, fShowToast, fEndToast */
/* exported fToggleDesignerVisibility, fGetVisibilityState, fCheckAndSetVisibility */

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// End - n/a
// Start - Designer Visibility Toggles
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

/* function fToggleDesignerVisibility
   Purpose: Toggles the visibility of all designer sheets, rows, and columns.
   Assumptions: A sheet named "Hide>" may exist to serve as a state marker.
   Notes: This is the main orchestrator for the Show/Hide All feature. It now handles cases where "Hide>" is missing.
   @returns {void}
*/
function fToggleDesignerVisibility() {
  fShowToast('⏳ Toggling visibility...', 'Show/Hide All');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const allSheets = ss.getSheets();
  const currentState = fGetVisibilityState(ss, allSheets);

  if (currentState === 'unknown') {
    fEndToast();
    fShowMessage('ℹ️ No Action', 'No designer elements with "Hide:" notes were found to toggle.');
    return;
  }

  const shouldHide = currentState === 'shown';
  const hideMarkerSheet = ss.getSheetByName('Hide>');
  const hideMarkerIndex = hideMarkerSheet ? hideMarkerSheet.getIndex() : -1;

  SpreadsheetApp.flush(); // Apply pending changes before proceeding

  if (shouldHide) {
    fHideAllElements(allSheets, hideMarkerIndex);
    fEndToast();
    fShowMessage('✅ Success', 'All designer elements have been hidden.');
  } else {
    fUnhideAllElements(allSheets, hideMarkerIndex);
    fEndToast();
    fShowMessage('✅ Success', 'All designer elements have been shown.');
  }
} // End function fToggleDesignerVisibility

/* function fGetVisibilityState
   Purpose: Checks if the designer elements are currently 'hidden', 'shown', or 'unknown'.
   Assumptions: None.
   Notes: This is the definitive state checker for the visibility toggle system.
   @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [ss=SpreadsheetApp.getActiveSpreadsheet()] - The spreadsheet object.
   @param {GoogleAppsScript.Spreadsheet.Sheet[]} [allSheets=ss.getSheets()] - An array of all sheets.
   @returns {string} 'hidden', 'shown', or 'unknown'.
*/
function fGetVisibilityState(ss = SpreadsheetApp.getActiveSpreadsheet(), allSheets = ss.getSheets()) {
  const hideMarkerSheet = ss.getSheetByName('Hide>');

  // 1. Check the 'Hide>' marker sheet first.
  if (hideMarkerSheet) {
    return hideMarkerSheet.isSheetHidden() ? 'hidden' : 'shown';
  }

  // 2. If no marker sheet, check the A1 note of all sheets.
  for (const sheet of allSheets) {
    const note = sheet.getRange('A1').getNote();
    if (note.includes('Hide: ')) {
      const hideString = note.split('Hide: ')[1].split('\n')[0];
      const ranges = fParseA1Notation(hideString);
      if (ranges.rows.length > 0) {
        return sheet.isRowHiddenByUser(ranges.rows[0]) ? 'hidden' : 'shown';
      }
      if (ranges.cols.length > 0) {
        return sheet.isColumnHiddenByUser(ranges.cols[0]) ? 'hidden' : 'shown';
      }
    }
  }

  // 3. If no 'Hide:' notes are found anywhere, the state is unknown.
  return 'unknown';
} // End function fGetVisibilityState

/* function fCheckAndSetVisibility
   Purpose: Automatically shows or hides designer elements based on the desired state.
   Assumptions: None.
   Notes: This is called by onOpen triggers to enforce visibility rules.
   @param {boolean} shouldShow - If true, ensures elements are visible. If false, ensures they are hidden.
   @returns {void}
*/
function fCheckAndSetVisibility(shouldShow) {
  const currentState = fGetVisibilityState();

  if (shouldShow && currentState === 'hidden') {
    // Admin user, and elements are hidden, so show them.
    fToggleDesignerVisibility();
  } else if (!shouldShow && currentState === 'shown') {
    // Player user, and elements are visible, so hide them.
    fToggleDesignerVisibility();
  }
} // End function fCheckAndSetVisibility

/* function fHideAllElements
   Purpose: Hides all designated designer elements.
   Assumptions: Called by  fToggleDesignerVisibility.
   Notes: Iterates through sheets to hide them and their specified rows/cols.
   @param {GoogleAppsScript.Spreadsheet.Sheet[]} allSheets - An array of all sheets in the spreadsheet.
   @param {number} hideMarkerIndex - The 1-based index of the "Hide>" sheet, or -1 if not found.
   @returns {void}
*/
function fHideAllElements(allSheets, hideMarkerIndex) {
  allSheets.forEach((sheet, index) => {
    const currentIndex = index + 1;
    // Only hide the sheet if a marker exists and we are at or after it
    if (hideMarkerIndex !== -1 && currentIndex >= hideMarkerIndex) {
      sheet.hideSheet();
    }

    // Only process rows/cols for sheets BEFORE the marker (or all sheets if no marker)
    if (hideMarkerIndex === -1 || currentIndex < hideMarkerIndex) {
      const note = sheet.getRange('A1').getNote();
      if (note.includes('Hide: ')) {
        const hideString = note.split('Hide: ')[1].split('\n')[0];
        const ranges = fParseA1Notation(hideString);
        ranges.rows.forEach(row => sheet.hideRows(row));
        ranges.cols.forEach(col => sheet.hideColumns(col));
      }
    }
  });
} // End function fHideAllElements

/* function fUnhideAllElements
   Purpose: Unhides all designated designer elements.
   Assumptions: Called by function fToggleDesignerVisibility.
   Notes: Iterates through sheets to unhide them and their specified rows/cols.
   @param {GoogleAppsScript.Spreadsheet.Sheet[]} allSheets - An array of all sheets in the spreadsheet.
   @param {number} hideMarkerIndex - The 1-based index of the "Hide>" sheet, or -1 if not found.
   @returns {void}
*/
function fUnhideAllElements(allSheets, hideMarkerIndex) {
  allSheets.forEach((sheet, index) => {
    const currentIndex = index + 1;
    // Only unhide the sheet if a marker exists and we are at or after it
    if (hideMarkerIndex !== -1 && currentIndex >= hideMarkerIndex) {
      sheet.showSheet();
    }

    // Only process rows/cols for sheets BEFORE the marker (or all sheets if no marker)
    if (hideMarkerIndex === -1 || currentIndex < hideMarkerIndex) {
      const note = sheet.getRange('A1').getNote();
      if (note.includes('Hide: ')) {
        const hideString = note.split('Hide: ')[1].split('\n')[0];
        const ranges = fParseA1Notation(hideString);
        ranges.rows.forEach(row => sheet.unhideRow(sheet.getRange(row, 1)));
        ranges.cols.forEach(col => sheet.unhideColumn(sheet.getRange(1, col)));
      }
    }
  });

  // Scroll to the top-left of the active sheet
  SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().getRange('A1').activate();
} // End function fUnhideAllElements