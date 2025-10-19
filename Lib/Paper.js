/* global SpreadsheetApp, fShowMessage, fParseA1Notation */
/* exported fPrepGameForPaper */

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// End - n/a
// Start - Paper & Print Utilities
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

/* function fPrepGameForPaper
   Purpose: Creates a clean, print-friendly copy of the <Game> sheet by copying it to the developer's master "Paper" file.
   Assumptions: Run from a Character Sheet. This is a designer-only function.
   Notes: This function reads the master 'Paper' sheet ID from the Versions sheet, which only designers can access.
   @returns {void}
*/
function fPrepGameForPaper() {
  fShowToast('⏳ Preparing <Paper> sheet...', 'Print Prep');

  // --- Admin Check ---
  const adminEmails = [g.ADMIN_EMAIL, g.DEV_EMAIL].map(e => e.toLowerCase());
  const isAdmin = adminEmails.includes(Session.getActiveUser().getEmail().toLowerCase());
  if (!isAdmin) {
    fEndToast();
    fShowMessage('❌ Permission Denied', 'This function is available for designers only.');
    return;
  }
  // --- End Admin Check ---

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = ss.getSheetByName('Game');

  if (!sourceSheet) {
    fEndToast();
    fShowMessage('❌ Error', "A sheet named 'Game' must exist to run this function.");
    return;
  }

  // 1. Get the Developer's Master Paper Sheet ID
  let paperSheetId;
  try {
    fShowToast('Locating master paper sheet...', 'Print Prep');
    paperSheetId = fGetMasterSheetId(g.CURRENT_VERSION, 'Paper'); // <-- YOUR 'Paper' ssabbr
  } catch (e) {
    fEndToast();
    fShowMessage('❌ Error', `Could not find master 'Paper' sheet ID. Error: ${e.message}`);
    return;
  }

  if (!paperSheetId) {
    fEndToast();
    fShowMessage('❌ Error', "Could not find the master 'Paper' sheet ID in the <Versions> sheet.");
    return;
  }

  // 2. Open the external Master Paper spreadsheet
  let paperSS;
  try {
    paperSS = SpreadsheetApp.openById(paperSheetId);
  } catch (e) {
    fEndToast();
    fShowMessage('❌ Error', "Could not open the master 'Paper' file. It may have been deleted or permissions may have changed.");
    return;
  }

  // 3. Duplicate the <Game> sheet *into the external spreadsheet* FIRST.
  fShowToast('Copying <Game> data...', 'Print Prep');
  const newPaperSheet = sourceSheet.copyTo(paperSS);

  // 4. Now that there are two sheets, it is safe to delete the old one.
  const oldPaperSheet = paperSS.getSheetByName('Paper');
  if (oldPaperSheet) {
    paperSS.deleteSheet(oldPaperSheet);
  }

  // 5. Rename and position the new sheet.
  newPaperSheet.setName('Paper');
  newPaperSheet.setTabColor(null); // Remove any custom tab color
  paperSS.moveActiveSheet(1);


  // 6. Read the 'Hide:' notation from the A1 cell.
  const note = newPaperSheet.getRange('A1').getNote();
  if (!note.includes('Hide: ')) {
    fEndToast();
    fShowMessage('✅ Success', 'The master <Paper> sheet has been updated, but no "Hide:" notation was found to remove designer elements.');
    // Activate the new sheet for the user
    SpreadApp.openById(paperSheetId);
    return;
  }
  const hideString = note.split('Hide: ')[1].split('\n')[0];
  const rangesToHide = fParseA1Notation(hideString);

  // 7. Delete the specified columns and rows from the external sheet.
  fShowToast('Removing designer elements...', 'Print Prep');
  rangesToHide.cols.sort((a, b) => b - a).forEach(col => newPaperSheet.deleteColumn(col));
  rangesToHide.rows.sort((a, b) => b - a).forEach(row => newPaperSheet.deleteRow(row));

  fEndToast();
  fShowMessage('✅ Success', "The master 'Paper' sheet has been successfully updated with a clean copy of this character's <Game> sheet.");

  // Activate the new sheet for the user. This will open the Paper CS file.
  SpreadsheetApp.openById(paperSheetId);
} // End function fPrepGameForPaper