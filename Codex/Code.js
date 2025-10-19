// 💪MS4
/* global FlexLib, PropertiesService, SpreadsheetApp, Session */

const SCRIPT_INITIALIZED_KEY = 'CODEX_INITIALIZED';

/* function onOpen
   Purpose: Simple trigger that runs automatically when the spreadsheet is opened.
   Assumptions: None.
   Notes: Builds menus based on authorization status and user identity (player vs. designer). Handles admins opening templates correctly.
   @returns {void}
*/
function onOpen() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const isInitialized = scriptProperties.getProperty(SCRIPT_INITIALIZED_KEY);
  const g = FlexLib.getGlobals();
  const adminEmails = [g.ADMIN_EMAIL, g.DEV_EMAIL].map(e => e.toLowerCase());
  const currentUserEmail = Session.getActiveUser().getEmail().toLowerCase(); // Get email once
  const isAdmin = adminEmails.includes(currentUserEmail);

  // --- REVISED LOGIC v3 ---
  if (isAdmin) {
    // If the user is an admin, ALWAYS show the full admin menus immediately.
    // This prevents admins from accidentally triggering player setup on templates.
    FlexLib.fCreateCodexMenu();
    FlexLib.fCreateDesignerMenu('Codex');
    // Admin visibility state is not auto-changed
  } else if (!isInitialized) {
    // If it's a regular user AND the sheet is NOT initialized, show only the activation menu.
    SpreadsheetApp.getUi()
      .createMenu(g.VersionName)
      .addItem(`▶️ Activate ${g.VersionName} Menus`, 'fActivateCodex')
      .addToUi();
  } else {
    // If it's a regular user AND the sheet IS initialized, show the player menu.
    FlexLib.fCreateCodexMenu();
    FlexLib.fCheckAndSetVisibility(false); // Ensure elements are HIDDEN for players
  }
  // --- END REVISED LOGIC v3 ---
} // End function onOpen

/* function fActivateCodex
   Purpose: Runs the first-time authorization and setup, then prompts the user to refresh.
   Assumptions: Triggered by a user clicking the 'Activate' menu item.
   Notes: This function's execution triggers the Google Auth prompt and the one-time setup.
   @returns {void}
*/
function fActivateCodex() {
  // First, call the library to run the initial setup.
  FlexLib.run('InitialSetup');

  // Once setup is complete, set the property so the full menus appear next time.
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty(SCRIPT_INITIALIZED_KEY, 'true');

  // Display a consistent success message.
  const title = 'IMPORTANT - Please Refresh Browser Tab';
  const message = '✅ Success! The script has been authorized and setup is complete.\n\nPlease refresh this browser tab now to load the full custom menus.';
  SpreadsheetApp.getUi().alert(title, message, SpreadsheetApp.getUi().ButtonSet.OK);
} // End function fActivateCodex

/* function fMenuTrimSheet
   Purpose: Local trigger for the "Trim Empty Rows/Cols" menu item.
   Assumptions: None.
   Notes: Acts as a pass-through to the central dispatcher in FlexLib.
   @returns {void}
*/
function fMenuTrimSheet() {
  FlexLib.run('TrimSheet');
} // End function fMenuTrimSheet

/* function fMenuTagVerification
   Purpose: The local trigger function called by the "Tag Verification" menu item.
   Assumptions: None.
   Notes: This function acts as a simple pass-through to the central dispatcher in FlexLib.
   @returns {void}
*/
function fMenuTagVerification() {
  FlexLib.run('TagVerification');
} // End function fMenuTagVerification

/* function fMenuToggleVisibility
   Purpose: Local trigger for the "Show/Hide All" menu item.
   Assumptions: None.
   Notes: Acts as a pass-through to the central dispatcher in FlexLib.
   @returns {void}
*/
function fMenuToggleVisibility() {
  FlexLib.run('ToggleVisibility');
} // End function fMenuToggleVisibility


/* function fMenuTest
   Purpose: Local trigger for the "Test" menu item.
   Assumptions: None.
   Notes: Acts as a pass-through to the central dispatcher in FlexLib.
   @returns {void}
*/
function fMenuTest() {
  FlexLib.run('Test');
} // End function fMenuTest


/* function fMenuCreateLatestCharacter
   Purpose: Local trigger for the "Create New Character > Latest Version" menu item.
   Assumptions: None.
   Notes: Acts as a pass-through to the central dispatcher in FlexLib.
   @returns {void}
*/
function fMenuCreateLatestCharacter() {
  FlexLib.run('CreateLatestCharacter', 'Characters');
} // End function fMenuCreateLatestCharacter

/* function fMenuCreateLegacyCharacter
   Purpose: Local trigger for the "Create New Character > Older Legacy Version" menu item.
   Assumptions: None.
   Notes: Acts as a pass-through to the central dispatcher in FlexLib.
   @returns {void}
*/
function fMenuCreateLegacyCharacter() {
  FlexLib.run('CreateLegacyCharacter', 'Characters');
} // End function fMenuCreateLegacyCharacter

/* function fMenuRenameCharacter
   Purpose: Local trigger for the "Rename Character" menu item.
   Assumptions: None.
   Notes: Acts as a pass-through to the central dispatcher in FlexLib.
   @returns {void}
*/
function fMenuRenameCharacter() {
  FlexLib.run('RenameCharacter', 'Characters');
} // End function fMenuRenameCharacter

/* function fMenuCreateCustomList
   Purpose: Local trigger for the "Create New Custom Ability List..." menu item.
   Assumptions: None.
   Notes: Acts as a pass-through to the central dispatcher in FlexLib.
   @returns {void}
*/
function fMenuCreateCustomList() {
  FlexLib.run('CreateCustomList', 'Custom Abilities');
} // End function fMenuCreateCustomList

/* function fMenuAddNewCustomSource
   Purpose: Local trigger for the "Add New Source..." menu item.
   Assumptions: None.
   Notes: Acts as a pass-through to the central dispatcher in FlexLib.
   @returns {void}
*/
function fMenuAddNewCustomSource() {
  FlexLib.run('AddNewCustomSource', 'Custom Abilities');
} // End function fMenuAddNewCustomSource


/* function fMenuRenameCustomList
   Purpose: Local trigger for the "Rename Custom List..." menu item.
   Assumptions: None.
   Notes: Acts as a pass-through to the central dispatcher in FlexLib.
   @returns {void}
*/
function fMenuRenameCustomList() {
  FlexLib.run('RenameCustomList', 'Custom Abilities');
} // End function fMenuRenameCustomList

/* function fMenuDeleteCustomList
   Purpose: Local trigger for the "Delete Custom List(s)..." menu item.
   Assumptions: None.
   Notes: Acts as a pass-through to the central dispatcher in FlexLib.
   @returns {void}
*/
function fMenuDeleteCustomList() {
  FlexLib.run('DeleteCustomList', 'Custom Abilities');
} // End function fMenuDeleteCustomList

/* function fMenuShareCustomLists
   Purpose: Local trigger for the "Share Custom List(s)..." menu item.
   Assumptions: None.
   Notes: Acts as a pass-through to the central dispatcher in FlexLib.
   @returns {void}
*/
function fMenuShareCustomLists() {
  FlexLib.run('ShareCustomLists', 'Custom Abilities');
} // End function fMenuShareCustomLists


/* function fMenuDeleteCharacter
   Purpose: Local trigger for the "Delete Character(s)" menu item.
   Assumptions: None.
   Notes: Acts as a pass-through to the central dispatcher in FlexLib.
   @returns {void}
*/
function fMenuDeleteCharacter() {
  FlexLib.run('DeleteCharacter', 'Characters');
} // End function fMenuDeleteCharacter

/* function fMenuCreateCharacter
   Purpose: Local trigger for the "Create New Character" menu item.
   Assumptions: None.
   Notes: Acts as a pass-through to the central dispatcher in FlexLib.
   @returns {void}
*/
function fMenuCreateCharacter() {
  FlexLib.run('CreateCharacter');
} // End function fMenuCreateCharacter
