/* global SpreadsheetApp, g, fGetSheetData */
/* exported fShowMessage, fLoadSheetToArray, fNormalizeTags, fActivateSheetByName, fClearAndWriteData */

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// End - n/a
// Start - User Interface Utilities
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


/* function fActivateSheetByName
   Purpose: Activates a sheet by its name, making it visible to the user.
   Assumptions: The sheet with the given name exists in the active spreadsheet.
   Notes: A generic helper to guide the user's focus.
   @param {string} sheetName - The name of the sheet to activate.
   @returns {void}
*/
function fActivateSheetByName(sheetName) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (sheet) {
    sheet.activate();
  }
} // End function fActivateSheetByName



/* function fTrimSheet
   Purpose: Trims all empty rows and columns from the active sheet based on cell content.
   Assumptions: The user has triggered this from a menu item on an active sheet.
   Notes: This is a destructive action that removes rows/columns permanently. It ignores formatting and formulas that result in empty strings.
   @returns {void}
*/
function fTrimSheet() {
  fShowToast('⏳ Analyzing sheet...', 'Trim Sheet');
  const sheet = SpreadsheetApp.getActiveSheet();
  const sheetName = sheet.getName();

  let lastRow = sheet.getLastRow();
  let lastCol = sheet.getLastColumn();
  const maxRows = sheet.getMaxRows();
  const maxCols = sheet.getMaxColumns();

  // --- THIS IS THE FIX ---
  // If the sheet is completely empty, getLastRow/getLastColumn return 0.
  // In this case, we want to trim down to a 1x1 sheet, not a 0x0 sheet.
  if (lastRow === 0) {
    lastRow = 1;
  }
  if (lastCol === 0) {
    lastCol = 1;
  }
  // --- END FIX ---

  const rowsToDelete = maxRows - lastRow;
  const colsToDelete = maxCols - lastCol;

  if (rowsToDelete <= 0 && colsToDelete <= 0) {
    fEndToast();
    fShowMessage('ℹ️ No Action Needed', 'The active sheet has no empty rows or columns to trim.');
    return;
  }

  // Build a confirmation message
  let confirmMessage = `This will permanently delete empty rows and columns from the sheet "${sheetName}".\n\n`;
  if (rowsToDelete > 0) {
    confirmMessage += `➡️ Rows to delete: ${rowsToDelete} (from row ${lastRow + 1} to ${maxRows})\n`;
  }
  if (colsToDelete > 0) {
    const startColA1 = sheet.getRange(1, lastCol + 1).getA1Notation().replace('1', '');
    const endColA1 = sheet.getRange(1, maxCols).getA1Notation().replace('1', '');
    confirmMessage += `➡️ Columns to delete: ${colsToDelete} (from column ${startColA1} to ${endColA1})\n`;
  }
  confirmMessage += '\n⚠️ IMPORTANT: This action is based on cell CONTENT only and does not consider formatting. It cannot be undone.\n\nTo proceed, type TRIM below.';

  fShowToast('Waiting for your confirmation...', 'Trim Sheet');
  const confirmationText = fPromptWithInput('Confirm Trim', confirmMessage);

  if (confirmationText === null || confirmationText.toLowerCase().trim() !== 'trim') {
    fEndToast();
    fShowMessage('ℹ️ Canceled', 'Trim operation canceled.');
    return;
  }

  fShowToast('Trimming sheet...', 'Trim Sheet');
  // Delete columns first to avoid potential errors if both are at max
  if (colsToDelete > 0) {
    sheet.deleteColumns(lastCol + 1, colsToDelete);
  }
  if (rowsToDelete > 0) {
    sheet.deleteRows(lastRow + 1, rowsToDelete);
  }

  fEndToast();
  fShowMessage('✅ Success', `Successfully trimmed ${rowsToDelete} row(s) and ${colsToDelete} column(s) from "${sheetName}".`);
} // End function fTrimSheet

/* function fShowMessage
   Purpose: Displays a simple modal pop-up message to the user.
   Assumptions: None.
   Notes: This is our standard method for all user-facing modal alerts.
   @param {string} title - The title to display in the message box header.
   @param {string} message - The main body of the message.
   @returns {void}
*/
function fShowMessage(title, message) {
  SpreadsheetApp.getUi().alert(title, message, SpreadsheetApp.getUi().ButtonSet.OK);
} // End function fShowMessage

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// End - User Interface Utilities
// Start - Data Handling Utilities
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

/* function fClearAndWriteData
   Purpose: A generic helper to clear old data from a sheet and write new, sorted data, preserving the header and template row formatting.
   Assumptions: The destination sheet has a 'Header' row tag.
   Notes: This is a reusable utility for any 'Build...' process.
   @param {GoogleAppsScript.Spreadsheet.Sheet} destSheet - The destination sheet object.
   @param {Array<Array<string>>} dataToWrite - A 2D array of the new data to be written.
   @param {object} destColTags - The column tag map for the destination sheet.
   @returns {void}
*/
function fClearAndWriteData(destSheet, dataToWrite, destColTags) {
  const { rowTags } = fGetSheetData('Temp', destSheet.getName(), destSheet.getParent(), true);
  const headerRowIndex = rowTags.header;
  const firstDataRow = headerRowIndex + 2;
  const lastRow = destSheet.getLastRow();

  // 1. Clear old data
  if (lastRow >= firstDataRow) {
    destSheet.getRange(firstDataRow, 2, lastRow - firstDataRow + 1, destSheet.getMaxColumns() - 1).clearContent();
    if (lastRow > firstDataRow) {
      destSheet.deleteRows(firstDataRow + 1, lastRow - firstDataRow);
    }
  }

  // 2. Write new data
  const newRowCount = dataToWrite.length;
  if (newRowCount > 0) {
    if (newRowCount > 1) {
      destSheet.insertRowsAfter(firstDataRow, newRowCount - 1);
      const formatSourceRange = destSheet.getRange(firstDataRow, 1, 1, destSheet.getMaxColumns());
      const formatDestRange = destSheet.getRange(firstDataRow + 1, 1, newRowCount - 1, destSheet.getMaxColumns());
      formatSourceRange.copyTo(formatDestRange, { formatOnly: true });
    }

    // Map the sparse array to a full 2D array that matches the destination sheet's column order
    const outputData = dataToWrite.map(row => {
      const outputRow = [];
      for (const tag in destColTags) {
        const colIndex = destColTags[tag];
        outputRow[colIndex] = row[colIndex] || ''; // Use empty string for any undefined values
      }
      return outputRow.slice(1); // Remove the first column (row tags)
    });

    destSheet.getRange(firstDataRow, 2, newRowCount, outputData[0].length).setValues(outputData);
  }
} // End function fClearAndWriteData


/* function fColumnToNumber
   Purpose: Converts a column letter string (A, B, AA, AB) to its 1-based column number.
   Assumptions: None.
   Notes: A helper for fParseA1Notation.
   @param {string} colString - The column letter string.
   @returns {number} The 1-based column number.
*/
function fColumnToNumber(colString) {
  let num = 0;
  const upperColString = colString.toUpperCase();
  for (let i = 0; i < upperColString.length; i++) {
    num = num * 26 + (upperColString.charCodeAt(i) - 64);
  }
  return num;
} // End function fColumnToNumber

/* function fParseA1Notation
   Purpose: Parses a custom A1 notation string into an object of rows and columns.
   Assumptions: The input string format is "A,1,3-4,D-F,BI-BK".
   Notes: This is the core parser for the Show/Hide All feature.
   @param {string} notationString - The string to parse.
   @returns {{rows: number[], cols: number[]}} An object containing arrays of row and column numbers.
*/
function fParseA1Notation(notationString) {
  const output = { rows: [], cols: [] };
  if (!notationString) return output;

  const parts = notationString.split(',');

  parts.forEach(part => {
    part = part.trim();
    // Handle row ranges (e.g., "3-5")
    if (part.includes('-') && !isNaN(part.split('-')[0])) {
      const [start, end] = part.split('-').map(Number);
      for (let i = start; i <= end; i++) {
        output.rows.push(i);
      }
    }
    // Handle single rows
    else if (!isNaN(part)) {
      output.rows.push(Number(part));
    }
    // Handle column ranges (e.g., "D-F" or "AZ-BC")
    else if (part.includes('-')) {
      const [start, end] = part.split('-').map(fColumnToNumber);
      for (let i = start; i <= end; i++) {
        output.cols.push(i);
      }
    }
    // Handle single columns (e.g., "A" or "BI")
    else {
      output.cols.push(fColumnToNumber(part));
    }
  });

  // Remove duplicates and sort
  output.rows = [...new Set(output.rows)].sort((a, b) => a - b);
  output.cols = [...new Set(output.cols)].sort((a, b) => a - b);

  return output;
} // End function fParseA1Notation



/* function fLoadSheetToArray
   Purpose: Loads an entire sheet's data into the global g object for in-memory processing.
   Assumptions: The sheet with the specified sheetName exists in the provided spreadsheet object.
   Notes: Creates the necessary object structure within g if it doesn't exist.
   @param {string} spreadsheetName - The key to use for the spreadsheet in the g object (e.g., 'Ver').
   @param {string} sheetName - The exact, case-sensitive name of the sheet to load.
   @param {GoogleAppsScript.Spreadsheet.Spreadsheet} [ss=SpreadsheetApp.getActiveSpreadsheet()] - The spreadsheet object to load from.
   @returns {void}
*/
function fLoadSheetToArray(spreadsheetName, sheetName, ss = SpreadsheetApp.getActiveSpreadsheet()) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error(`Sheet "${sheetName}" could not be found.`);
  }

  // Ensure the object structure exists
  if (!g[spreadsheetName]) {
    g[spreadsheetName] = {};
  }
  if (!g[spreadsheetName][sheetName]) {
    g[spreadsheetName][sheetName] = {};
  }

  g[spreadsheetName][sheetName].arr = sheet.getDataRange().getValues();
} // End function fLoadSheetToArray

/* function fNormalizeTags
   Purpose: Converts a raw tag string into an array of standardized tags.
   Assumptions: None.
   Notes: Processes tags to be case-insensitive, space-insensitive, and comma-separated.
   @param {string} tagString - The raw string from a tag cell (e.g., "Character Name, ID").
   @returns {string[]} An array of normalized tags (e.g., ['charactername', 'id']).
*/
function fNormalizeTags(tagString) {
  if (!tagString || typeof tagString !== 'string') {
    return [];
  }
  return tagString
    .toLowerCase()
    .replace(/\s+/g, '')
    .split(',')
    .filter(tag => tag); // Filter out any empty strings that result from ",," or trailing commas
} // End function fNormalizeTags