/* global g, fGetSheetId, SpreadsheetApp, fBuildTagMaps, fShowMessage, fShowToast, fActivateSheetByName, fGetSheetData, fEndToast, fGetVerifiedLocalFile, fGetCodexSpreadsheet, fDeleteTableRow, fGetMasterSheetId, fClearAndWriteData */
/* exported fBuildPowers */

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// End - n/a
// Start - Power List Generation
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


/* function fClearAllFilterCheckboxes
   Purpose: Unchecks all 'isactive' checkboxes on a given filter sheet.
   Assumptions: The sheet has a column tagged 'isactive'.
   Notes: A reusable helper for user convenience.
   @param {string} sheetName - The name of the sheet to clear (e.g., 'Filter Powers').
   @returns {void}
*/
function fClearAllFilterCheckboxes(sheetName) {
  fActivateSheetByName(sheetName);
  fShowToast(`⏳ Clearing all selections in <${sheetName}>...`, 'Clear Selections');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    fEndToast();
    fShowMessage('❌ Error', `Could not find the <${sheetName}> sheet.`);
    return;
  }

  const { rowTags, colTags } = fGetSheetData('CS', sheetName, ss, true);
  const headerRow = rowTags.header;
  const isActiveCol = colTags.isactive;

  if (headerRow === undefined || isActiveCol === undefined) {
    fEndToast();
    fShowMessage('❌ Error', `The <${sheetName}> sheet is not tagged correctly.`);
    return;
  }

  const firstDataRow = headerRow + 2;
  const lastRow = sheet.getLastRow();
  const numRows = lastRow - firstDataRow + 1;

  if (numRows > 0) {
    sheet.getRange(firstDataRow, isActiveCol + 1, numRows, 1).uncheck();
  }

  fEndToast();
  fShowMessage('✅ Success', 'All selections have been cleared.\n\n⚠️ You must now select at least one table and run the "Filter..." command or click the green "Refresh" button to update your character\'s dropdowns.');
} // End function fClearAllFilterCheckboxes

/* function fGetAllPowerTablesList
   Purpose: A helper function to get a definitive, aggregated list of all available power tables from DB and Custom sources, including SubType.
   Assumptions: None.
   Notes: This is the central source of truth for what power tables currently exist. Includes SubType for sorting.
   @returns {{allPowerTables: Array<{tableName: string, source: string, subType: string}>}} An object containing the aggregated list.
*/
function fGetAllPowerTablesList() {
  const dbPowerTables = [];
  const customPowerTables = [];

  // 1a. Get standard tables from the PLAYER'S LOCAL DB copy.
  const dbFile = fGetVerifiedLocalFile(g.CURRENT_VERSION, 'DB');
  if (dbFile) {
    const sourceSS = SpreadsheetApp.open(dbFile);
    const { arr, rowTags, colTags } = fGetSheetData('DB', 'Powers', sourceSS);
    const headerRow = rowTags.header;
    if (headerRow !== undefined) {
      const tableNameCol = colTags.tablename;
      const subTypeCol = colTags.subtype; // <-- Get SubType column index
      const uniqueTables = {}; // Use an object to track unique tables and their subtypes

      arr.slice(headerRow + 1).forEach(row => {
        const tableName = row[tableNameCol];
        const subType = row[subTypeCol];
        if (tableName && !uniqueTables[tableName]) { // Only add if not already seen
          uniqueTables[tableName] = subType || 'Unknown'; // Store subtype, default if blank
        }
      });

      // Convert the uniqueTables object into the desired array format
      for (const tableName in uniqueTables) {
        dbPowerTables.push({ tableName: tableName, source: 'DB', subType: uniqueTables[tableName] });
      }
    }
  }

  // 1b. Get custom tables from all registered sources in the Codex.
  const codexSS = fGetCodexSpreadsheet();
  const { arr: sourcesArr, rowTags: sourcesRowTags, colTags: sourcesColTags } = fGetSheetData('Codex', 'Custom Abilities', codexSS, true);
  const sourcesHeader = sourcesRowTags.header;
  if (sourcesHeader !== undefined) {
    for (let r = sourcesHeader + 1; r < sourcesArr.length; r++) {
      const sourceRow = sourcesArr[r];
      if (sourceRow && sourceRow[sourcesColTags.sheetid]) {
        const sourceId = sourceRow[sourcesColTags.sheetid];
        const sourceName = sourceRow[sourcesColTags.custabilitiesname];
        try {
          const customSS = SpreadsheetApp.openById(sourceId);
          const { arr, rowTags, colTags } = fGetSheetData(`Cust_${sourceId}`, 'VerifiedPowers', customSS);
          const headerRow = rowTags.header;
          if (headerRow !== undefined) {
            const tableNameCol = colTags.tablename;
            const subTypeCol = colTags.subtype; // <-- Get SubType column index
            const uniqueTables = {};

            arr.slice(headerRow + 1).forEach(row => {
              const tableName = row[tableNameCol];
              const subType = row[subTypeCol];
              if (tableName && !uniqueTables[tableName]) {
                uniqueTables[tableName] = subType || 'Unknown';
              }
            });

            for (const tableName in uniqueTables) {
              customPowerTables.push({ tableName: `Cust - ${tableName}`, source: sourceName, subType: uniqueTables[tableName] });
            }
          }
        } catch (e) {
          // Fail silently during the health check
          console.error(`Could not access custom source "${sourceName}" with ID ${sourceId}. Error: ${e}`);
        }
      }
    }
  }

  // --- NEW SORTING LOGIC ---
  // Define the desired order for SubTypes
  const subTypeOrder = ['Class', 'Race', 'Combat Style', 'Luck', 'Unknown']; // Add others as needed

  // Sort DB tables first
  dbPowerTables.sort((a, b) => {
    const subTypeIndexA = subTypeOrder.indexOf(a.subType);
    const subTypeIndexB = subTypeOrder.indexOf(b.subType);

    // Sort by SubType order
    if (subTypeIndexA !== subTypeIndexB) {
      return (subTypeIndexA === -1 ? Infinity : subTypeIndexA) - (subTypeIndexB === -1 ? Infinity : subTypeIndexB);
    }
    // If SubTypes are the same, sort by TableName
    return a.tableName.localeCompare(b.tableName);
  });

  // Sort Custom tables alphabetically by TableName
  customPowerTables.sort((a, b) => a.tableName.localeCompare(b.tableName));
  // --- END NEW SORTING LOGIC ---

  // Combine DB tables first, then Custom tables
  return { allPowerTables: [...dbPowerTables, ...customPowerTables] };
} // End function fGetAllPowerTablesList

/* function fUpdatePowerTablesList
   Purpose: Updates the <Filter Powers> sheet with a unique list of all TableNames from the PLAYER'S LOCAL DB and all registered custom sources, sorted by SubType then TableName.
   Assumptions: The user is running this from a Character Sheet. The <Filter Powers> sheet has a 'subtype' column tag.
   Notes: Aggregates from multiple sources and sorts them into logical groups. Can be run silently.
   @param {boolean} [isSilent=false] - If true, suppresses the final success message.
   @returns {void}
*/
function fUpdatePowerTablesList(isSilent = false) {
  fActivateSheetByName('Filter Powers');
  fShowToast('⏳ Syncing power tables...', isSilent ? '⚙️ Onboarding' : 'Sync Power Tables');

  const destSS = SpreadsheetApp.getActiveSpreadsheet();
  const destSheet = destSS.getSheetByName('Filter Powers');
  if (!destSheet) {
    if (!isSilent) fEndToast();
    fShowMessage('❌ Error', 'Could not find the <Filter Powers> sheet in this spreadsheet.');
    return;
  }

  // --- Preserve checked state ---
  const { arr: oldArr, rowTags: oldRowTags, colTags: oldColTags } = fGetSheetData('CS', 'Filter Powers', destSS, true);
  const oldHeaderRow = oldRowTags.header;
  const previouslyChecked = new Set();
  // Check if necessary tags exist before trying to access them
  if (oldHeaderRow !== undefined && oldColTags.isactive !== undefined && oldColTags.tablename !== undefined) {
    for (let r = oldHeaderRow + 1; r < oldArr.length; r++) {
      if (oldArr[r] && oldArr[r][oldColTags.isactive] === true && oldArr[r][oldColTags.tablename]) { // Check tablename exists
        previouslyChecked.add(oldArr[r][oldColTags.tablename]);
      }
    }
  }
  // --- END Preserve checked state ---

  // --- Get sorted data including SubType ---
  const { allPowerTables } = fGetAllPowerTablesList();
  // --- END Get sorted data ---

  const { rowTags: destRowTags, colTags: destColTags } = fGetSheetData('CS', 'Filter Powers', destSS, true);
  const destHeaderRow = destRowTags.header; // 0-based index
  if (destHeaderRow === undefined) {
    if (!isSilent) fEndToast();
    fShowMessage('❌ Error', 'Could not find a "Header" tag in the <Filter Powers> sheet.');
    return;
  }
  if (destColTags.subtype === undefined) {
    if (!isSilent) fEndToast();
    fShowMessage('❌ Error', 'The <Filter Powers> sheet is missing the required "SubType" column tag.');
    return;
  }

  const firstDataRow = destHeaderRow + 2; // 1-based row number for the first *data* row (template row)
  const lastRow = destSheet.getLastRow(); // 1-based row number of the last row with content

  // --- REVISED CLEARING LOGIC ---
  // 1. Delete extra rows: If there's more than one data row currently, delete rows AFTER the first data row.
  if (lastRow > firstDataRow) {
    destSheet.deleteRows(firstDataRow + 1, lastRow - firstDataRow);
  }
  // 2. Clear content of the first data row (template row), EXCLUDING the row tag in column A.
  // Check if firstDataRow actually has content before clearing
  if (lastRow >= firstDataRow && destSheet.getMaxColumns() > 1) {
    destSheet.getRange(firstDataRow, 2, 1, destSheet.getMaxColumns() - 1).clearContent();
  }
  // --- END REVISED CLEARING LOGIC ---


  const newRowCount = allPowerTables.length;
  if (newRowCount > 0) {
    // Add new rows if needed (if more than one item needs to be written)
    if (newRowCount > 1) {
      destSheet.insertRowsAfter(firstDataRow, newRowCount - 1);
      // Copy formatting from the (now potentially empty) template row
      const formatSourceRange = destSheet.getRange(firstDataRow, 1, 1, destSheet.getMaxColumns());
      const formatDestRange = destSheet.getRange(firstDataRow + 1, 1, newRowCount - 1, destSheet.getMaxColumns());
      formatSourceRange.copyTo(formatDestRange, { formatOnly: true });
    }

    // --- SIMPLIFIED DATA PREPARATION ---
    // Create a 2D array directly mapping the sorted data to the sheet columns (excluding column A)
    const outputData = allPowerTables.map(item => {
      const rowArray = [];
      // Create an empty array representing the row structure based on max columns
      for (let i = 1; i < destSheet.getMaxColumns(); i++) { // Start from 1 to skip tag column A
        rowArray.push('');
      }
      // Place data according to colTags indices (adjusting because we sliced column A)
      if (destColTags.tablename !== undefined) rowArray[destColTags.tablename - 1] = item.tableName;
      if (destColTags.subtype !== undefined) rowArray[destColTags.subtype - 1] = item.subType;
      if (destColTags.source !== undefined) rowArray[destColTags.source - 1] = item.source;
      // 'isactive' (checkbox) is handled after writing
      return rowArray;
    });
    // --- END SIMPLIFIED DATA PREPARATION ---


    // Write the data starting from column B (index 1) of the first data row
    const writeRange = destSheet.getRange(firstDataRow, 2, newRowCount, outputData[0].length);
    writeRange.setValues(outputData);

    // --- Re-apply checked state and ensure checkboxes ---
    const newIsActiveCol = destColTags.isactive + 1; // 1-based column for getRange
    const newTableNameCol = destColTags.tablename + 1; // 1-based column for getRange

    // Iterate through the newly written rows
    for (let i = 0; i < newRowCount; i++) {
      const currentRow = firstDataRow + i;
      const tableName = destSheet.getRange(currentRow, newTableNameCol).getValue();
      const checkboxRange = destSheet.getRange(currentRow, newIsActiveCol);
      // Ensure the checkboxRange is valid before trying to insert/check
      if (checkboxRange) {
        if (previouslyChecked.has(tableName)) {
          checkboxRange.check();
        } else {
          checkboxRange.insertCheckboxes(); // Ensure even unchecked rows get a box
        }
      }
    }
    // --- END Re-apply checked state ---
  }

  if (isSilent) {
    fShowToast('✅ Power tables synced.', '⚙️ Onboarding');
  } else {
    fEndToast();
    fShowMessage('✅ Success', `The <Filter Powers> sheet has been updated with ${newRowCount} power tables, sorted by SubType.\n\nYour previous selections have been preserved.`);
  }
} // End function fUpdatePowerTablesList

/* function fGetPowerSourceData
   Purpose: A helper to fetch, process, and aggregate all power data from the master Tables file.
   Assumptions: The 'Tables' file ID is valid and the source sheets exist.
   Notes: This is a helper for the fBuildPowers refactor. Filters out header/template rows.
   @param {object} destColTags - The column tag map from the destination <Powers> sheet.
   @returns {Array<Array<string>>} A 2D array of the aggregated and processed power data.
*/
function fGetPowerSourceData(destColTags) {
  const tablesId = fGetMasterSheetId(g.CURRENT_VERSION, 'Tables');
  if (!tablesId) {
    throw new Error('Could not find the ID for the "Tables" spreadsheet in the master <Versions> sheet.');
  }

  const sourceSS = SpreadsheetApp.openById(tablesId);
  const sourceSheetNames = ['Class', 'Race', 'CombatStyles', 'Luck'];
  const allPowersData = [];
  g.Tbls = {}; // Ensure a fresh cache namespace

  sourceSheetNames.forEach(sourceSheetName => {
    fShowToast(`⏳ Processing <${sourceSheetName}>...`, 'Build Powers');
    const sourceSheet = sourceSS.getSheetByName(sourceSheetName);
    if (!sourceSheet) {
      fShowToast(`⚠️ Could not find sheet: ${sourceSheetName}. Skipping.`, 'Build Powers', 10);
      return;
    }

    const { arr: sourceArr, rowTags: sourceRowTags, colTags: sourceColTags } = fGetSheetData('Tbls', sourceSheetName, sourceSS);
    const sourceHeaderIndex = sourceRowTags.header;
    if (sourceHeaderIndex === undefined) {
      fShowToast(`⚠️ No "Header" tag in <${sourceSheetName}>. Skipping.`, 'Build Powers', 10);
      return;
    }

    // --- Define required column tags ---
    const abilityNameCol = sourceColTags.abilityname;
    const subTypeCol = sourceColTags.subtype;
    const tableNameCol = sourceColTags.tablename;
    const usageCol = sourceColTags.usage;
    const actionCol = sourceColTags.action;
    const effectCol = sourceColTags.effect;
    const typeCol = sourceColTags.type;
    const sourceCol = sourceColTags.source;

    // Check if essential columns are missing
    if (abilityNameCol === undefined || subTypeCol === undefined || tableNameCol === undefined || usageCol === undefined || actionCol === undefined || effectCol === undefined) {
      fShowToast(`⚠️ Missing required column tags (abilityname, subtype, etc.) in <${sourceSheetName}>. Skipping.`, 'Build Powers', 10);
      return;
    }
    // --- End required column tags check ---


    for (let r = sourceHeaderIndex + 1; r < sourceArr.length; r++) {
      const row = sourceArr[r];
      const abilityName = row[abilityNameCol];
      const subType = row[subTypeCol]; // Get SubType for filtering

      // --- THIS IS THE FIX ---
      // Skip row if AbilityName is 'Power' OR SubType is 'SubType' (indicating a header/template row)
      if (!abilityName || abilityName === 'Power' || subType === 'SubType') {
        continue;
      }
      // --- END FIX ---

      const tableName = row[tableNameCol];
      const usage = row[usageCol];
      const action = row[actionCol];
      const effect = row[effectCol];
      const dropDownValue = `${tableName} - ${abilityName}⚡ (${usage}, ${action}) ➡ ${effect}`;

      const newRow = [];
      newRow[destColTags.dropdown] = dropDownValue;
      newRow[destColTags.type] = row[typeCol]; // Use typeCol index
      newRow[destColTags.subtype] = subType; // Already have subType
      newRow[destColTags.tablename] = tableName; // Already have tableName
      newRow[destColTags.source] = row[sourceCol]; // Use sourceCol index
      newRow[destColTags.usage] = usage; // Already have usage
      newRow[destColTags.action] = action; // Already have action
      newRow[destColTags.abilityname] = abilityName; // Already have abilityName
      newRow[destColTags.effect] = effect; // Already have effect

      allPowersData.push(newRow);
    }
  });

  // Sort the combined array
  fShowToast('⏳ Sorting all powers...', 'Build Powers');
  allPowersData.sort((a, b) => a[destColTags.dropdown].localeCompare(b[destColTags.dropdown]));

  return allPowersData;
} // End function fGetPowerSourceData



/* function fBuildPowers
   Purpose: The master function to rebuild the <Powers> sheet in the DB file from the master Tables file.
   Assumptions: The user is running this from the DB spreadsheet.
   Notes: This is a destructive and regenerative process that now reads from multiple source sheets.
   @returns {void}
*/
function fBuildPowers() {
  fShowToast('⏳ Initializing power build...', 'Build Powers');
  const destSheetName = 'Powers';
  fActivateSheetByName(destSheetName);

  try {
    const destSS = SpreadsheetApp.getActiveSpreadsheet();
    const destSheet = destSS.getSheetByName(destSheetName);
    if (!destSheet) {
      throw new Error(`Could not find the <${destSheetName}> sheet in the current spreadsheet.`);
    }

    g.DB = {}; // Ensure a fresh cache namespace
    const { colTags: destColTags } = fGetSheetData('DB', destSheetName, destSS, true);

    const allPowersData = fGetPowerSourceData(destColTags);

    fShowToast(`⏳ Writing ${allPowersData.length} new powers...`, 'Build Powers');
    fClearAndWriteData(destSheet, allPowersData, destColTags);

    fEndToast();
    fShowMessage('✅ Success', `The <${destSheetName}> sheet has been successfully rebuilt with ${allPowersData.length} powers from all sources.`);
  } catch (e) {
    console.error(`❌ CRITICAL ERROR in fBuildPowers: ${e.message}\n${e.stack}`);
    fEndToast();
    fShowMessage('❌ Error', `A critical error occurred. Please check the execution logs. Error: ${e.message}`);
  }
} // End function fBuildPowers


/* function fPerformPowerHealthCheck
   Purpose: A helper to find and remove any stale ("orphaned") power tables from the <Filter Powers> sheet.
   Assumptions: None.
   Notes: This is part of the fFilterPowers workflow.
   @returns {void}
*/
function fPerformPowerHealthCheck() {
  fShowToast('⚕️ Verifying power sources...', 'Filter Powers');
  const csSS = SpreadsheetApp.getActiveSpreadsheet();
  const { allPowerTables } = fGetAllPowerTablesList(); // Get a fresh list of ALL valid tables
  const validTableNames = new Set(allPowerTables.map(t => t.tableName));

  const filterSheet = csSS.getSheetByName('Filter Powers');
  const { arr: choicesArr, rowTags: choicesRowTags, colTags: choicesColTags } = fGetSheetData('CS', 'Filter Powers', csSS, true);
  const choicesHeaderRow = choicesRowTags.header;

  const orphanRows = [];
  for (let r = choicesHeaderRow + 1; r < choicesArr.length; r++) {
    const tableName = choicesArr[r][choicesColTags.tablename];
    if (tableName && !validTableNames.has(tableName)) {
      orphanRows.push({ row: r + 1, name: tableName });
    }
  }

  if (orphanRows.length > 0) {
    fShowToast('🧹 Cleaning up stale entries...', 'Filter Powers');
    orphanRows.sort((a, b) => b.row - a.row).forEach(orphan => {
      fDeleteTableRow(filterSheet, orphan.row);
    });
    const orphanNames = orphanRows.map(o => `- ${o.name}`).join('\n');
    fShowMessage('ℹ️ List Cleaned', `The following power tables could no longer be found and have been removed from your list:\n\n${orphanNames}`);
  }
} // End function fPerformPowerHealthCheck


/* function fGetSelectedPowerTables
   Purpose: A helper to read the <Filter Powers> sheet and return an array of user-selected tables.
   Assumptions: None.
   Notes: This is part of the fFilterPowers workflow.
   @returns {Array<object>|null} An array of selected table objects, or null if none are selected or the sheet is invalid.
*/
function fGetSelectedPowerTables() {
  const csSS = SpreadsheetApp.getActiveSpreadsheet();
  const { arr, rowTags, colTags } = fGetSheetData('CS', 'Filter Powers', csSS, true); // Force refresh after health check
  const headerRow = rowTags.header;

  const tableNameCol = colTags.tablename;
  const hasContent = arr.slice(headerRow + 1).some(row => row[tableNameCol]);
  if (!hasContent) {
    fEndToast();
    fUpdatePowerTablesList(); // No tables listed, so run the sync process for the user.
    return null;
  }

  if (headerRow === undefined) {
    fEndToast();
    fShowMessage('❌ Error', 'Could not find a "Header" tag in the <Filter Powers> sheet.');
    return null;
  }

  const selectedTables = arr
    .slice(headerRow + 1)
    .filter(row => row[colTags.isactive] === true)
    .map(row => ({ tableName: row[tableNameCol], source: row[colTags.source] }));

  if (selectedTables.length === 0) {
    fEndToast();
    fShowMessage('ℹ️ No Filters Selected', 'Please check one or more boxes on the <Filter Powers> sheet before filtering.');
    return null;
  }

  return selectedTables;
} // End function fGetSelectedPowerTables


/* function fFetchAllPowerData
   Purpose: A helper to fetch and aggregate all power data from the DB and selected custom sources.
   Assumptions: None.
   Notes: This is part of the fFilterPowers workflow.
   @param {Array<object>} selectedTables - The array of table objects returned by fGetSelectedPowerTables.
   @returns {{allPowersData: Array<Array<string>>, dbHeader: Array<string>}|null} An object containing the aggregated data and header, or null on error.
*/
function fFetchAllPowerData(selectedTables) {
  fShowToast('Fetching all selected powers...', 'Filter Powers');
  let allPowersData = [];
  let dbHeader = [];
  const codexSS = fGetCodexSpreadsheet();

  const dbFile = fGetVerifiedLocalFile(g.CURRENT_VERSION, 'DB');
  if (!dbFile) {
    fEndToast();
    fShowMessage('❌ Error', 'Could not find or restore your local "DB" file to get power data from. Please run initial setup.');
    return null;
  }
  const dbSS = SpreadsheetApp.open(dbFile);
  const { arr: allDbPowers, rowTags: dbRowTags, colTags: dbColTags } = fGetSheetData('DB', 'Powers', dbSS);
  
  // --- THIS IS THE FIX ---
  // The header for the cache file MUST be based on the colTag row (row 0), not the human-readable "Header" row.
  dbHeader = allDbPowers[0];
  // --- END FIX ---

  // Fetch from the local DB if selected
  const selectedDbTables = selectedTables.filter(t => t.source === 'DB').map(t => t.tableName);
  if (selectedDbTables.length > 0) {
    const dbPowers = allDbPowers
      .slice(dbRowTags.header + 1)
      .filter(row => selectedDbTables.includes(row[dbColTags.tablename]));
    allPowersData = allPowersData.concat(dbPowers);
  }

  // Fetch from Custom Sources
  const selectedCustomTables = selectedTables.filter(t => t.source !== 'DB');
  if (selectedCustomTables.length > 0) {
    const { arr: sourcesArr, colTags: sourcesColTags } = fGetSheetData('Codex', 'Custom Abilities', codexSS, true);
    for (const customTable of selectedCustomTables) {
      const sourceInfo = sourcesArr.find(row => row[sourcesColTags.custabilitiesname] === customTable.source);
      if (sourceInfo) {
        const sourceId = sourceInfo[sourcesColTags.sheetid];
        fShowToast(`Fetching from "${customTable.source}"...`, 'Filter Powers');
        try {
          const customSS = SpreadsheetApp.openById(sourceId);
          const { arr: customSheetPowers, rowTags: custRowTags, colTags: custColTags } = fGetSheetData(`Cust_${sourceId}`, 'VerifiedPowers', customSS);
          if (dbHeader.length === 0) dbHeader = customSheetPowers[0]; // Also use row 0 for custom headers

          const cleanTableName = customTable.tableName.replace('Cust - ', '');
          const filteredCustomPowers = customSheetPowers
            .slice(custRowTags.header + 1)
            .filter(row => row[custColTags.tablename] === cleanTableName);

          const mappedCustomPowers = filteredCustomPowers.map(row => {
            const newRow = [];
            newRow[dbColTags.dropdown] = row[custColTags.dropdown];
            newRow[dbColTags.type] = row[custColTags.type];
            newRow[dbColTags.subtype] = row[custColTags.subtype];
            newRow[dbColTags.tablename] = row[custColTags.tablename];
            newRow[dbColTags.source] = row[custColTags.source];
            newRow[dbColTags.usage] = row[custColTags.usage];
            newRow[dbColTags.action] = row[custColTags.action];
            newRow[dbColTags.abilityname] = row[custColTags.abilityname];
            newRow[dbColTags.effect] = row[custColTags.effect];
            return newRow;
          });

          allPowersData = allPowersData.concat(mappedCustomPowers);
        } catch (e) {
          console.error(`Could not access custom source "${customTable.source}". Error: ${e}`);
          fShowMessage('⚠️ Warning', `Could not access the custom source "${customTable.source}". Skipping.`);
        }
      }
    }
  }
  return { allPowersData, dbHeader };
} // End function fFetchAllPowerData


/* function fCachePowerData
   Purpose: A helper to write aggregated power data to the <PowerDataCache> sheet.
   Assumptions: None.
   Notes: This is part of the fFilterPowers workflow.
   @param {Array<Array<string>>} allPowersData - The aggregated power data.
   @param {Array<string>} dbHeader - The header row for the data.
   @returns {void}
*/
function fCachePowerData(allPowersData, dbHeader) {
  const csSS = SpreadsheetApp.getActiveSpreadsheet();
  const cacheSheet = csSS.getSheetByName('PowerDataCache');
  if (!cacheSheet) {
    fEndToast();
    fShowMessage('❌ Error', 'Could not find the <PowerDataCache> sheet.');
    return;
  }
  cacheSheet.clear();
  if (allPowersData.length > 0) {
    const dataToCache = [dbHeader, ...allPowersData];
    cacheSheet.getRange(1, 1, dataToCache.length, dataToCache[0].length).setValues(dataToCache);
  }
  fShowToast('⚡ Power data cached locally.', 'Filter Powers');
} // End function fCachePowerData


/* function fApplyPowerDropdowns
   Purpose: A helper to build and apply the final data validation dropdowns to the <Game> sheet.
   Assumptions: None.
   Notes: This is part of the fFilterPowers workflow.
   @param {Array<Array<string>>} allPowersData - The aggregated power data.
   @returns {number} The number of powers added to the dropdowns.
*/
function fApplyPowerDropdowns(allPowersData) {
  const csSS = SpreadsheetApp.getActiveSpreadsheet();
  const { colTags: dbColTags } = fGetSheetData('DB', 'Powers', SpreadsheetApp.open(fGetVerifiedLocalFile(g.CURRENT_VERSION, 'DB')));
  const filteredPowerList = allPowersData.map(row => row[dbColTags.dropdown]);
  const gameSheet = csSS.getSheetByName('Game');
  if (!gameSheet) {
    fEndToast();
    fShowMessage('❌ Error', 'Could not find the <Game> sheet.');
    return 0;
  }

  const { rowTags: gameRowTags, colTags: gameColTags } = fGetSheetData('CS', 'Game', csSS);
  const startRow = gameRowTags.powertablestart + 1;
  const endRow = gameRowTags.powertableend + 1;
  const numRows = endRow - startRow + 1;
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(filteredPowerList.length > 0 ? filteredPowerList : [' '], true).setAllowInvalid(false).build();

  if (gameColTags.powerdropdown1 !== undefined) {
    const colIndex = gameColTags.powerdropdown1 + 1;
    gameSheet.getRange(startRow, colIndex, numRows, 1).setDataValidation(rule);
  }
  if (gameColTags.powerdropdown2 !== undefined) {
    const colIndex = gameColTags.powerdropdown2 + 1;
    gameSheet.getRange(startRow, colIndex, numRows, 1).setDataValidation(rule);
  }
  
  return filteredPowerList.length;
} // End function fApplyPowerDropdowns


/* function fFilterPowers
   Purpose: Builds custom power selection dropdowns on the Character Sheet based on the player's choices in <Filter Powers>, aggregating from DB and Custom sources.
   Assumptions: The user is running this from a Character Sheet.
   Notes: This is the primary player-facing function for customizing their power list. It now also populates a local cache sheet.
   @param {boolean} [isSilent=false] - If true, suppresses the final success message.
   @returns {void}
*/
function fFilterPowers(isSilent = false) {
  fActivateSheetByName('Filter Powers');
  
  fPerformPowerHealthCheck();

  const selectedTables = fGetSelectedPowerTables();
  if (!selectedTables) return; // Exit if no tables are selected or an error occurred.

  const powerData = fFetchAllPowerData(selectedTables);
  if (!powerData) return; // Exit if there was an error fetching data.

  const { allPowersData, dbHeader } = powerData;
  fCachePowerData(allPowersData, dbHeader);

  const finalCount = fApplyPowerDropdowns(allPowersData);

  if (isSilent) {
    fShowToast('✅ Power dropdowns updated.', '⚙️ Onboarding');
  } else {
    fEndToast();
    fShowMessage('✅ Success!', `Your power selection dropdowns have been updated with ${finalCount} powers.`);
  }
} // End function fFilterPowers