// 💪MS4
/* global FlexLib, PropertiesService, SpreadsheetApp, Session */

const SCRIPT_INITIALIZED_KEY = 'CUST_INITIALIZED';

/* function onOpen
   Purpose: Simple trigger that runs automatically when the spreadsheet is opened.
   Assumptions: None.
   Notes: Builds the full menu if authorized, otherwise provides an activation option. Also applies data validation rules.
   @returns {void}
*/
function onOpen() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const isInitialized = scriptProperties.getProperty(SCRIPT_INITIALIZED_KEY);
  const g = FlexLib.getGlobals();
  const adminEmails = [g.ADMIN_EMAIL, g.DEV_EMAIL].map(e => e.toLowerCase());
  const isAdmin = adminEmails.includes(Session.getActiveUser().getEmail().toLowerCase());

  // --- REVISED LOGIC ---
  if (!isInitialized) {
    // If not initialized, ALWAYS show the activation menu.
    SpreadsheetApp.getUi()
      .createMenu(g.VersionName)
      .addItem(`▶️ Activate ${g.VersionName} Menus`, 'fActivateMenus')
      .addToUi();
  } else {
    // If initialized, create menus based on user role and apply validations.
    FlexLib.fApplyPowerValidations();
    FlexLib.fApplyMagicItemValidations();
    FlexLib.fApplySkillSetValidations();

    FlexLib.fCreateCustMenu(); // Standard player menu

    if (isAdmin) {
      FlexLib.fCreateDesignerMenu('Cust'); // Add designer menu for admins
      // Admin visibility state is no longer auto-changed
    } else {
      FlexLib.fCheckAndSetVisibility(false); // Ensure elements are HIDDEN for players
    }
  }
  // --- END REVISED LOGIC ---
} // End function onOpen

/* function fActivateMenus
   Purpose: Runs the first-time authorization and menu setup.
   Assumptions: Triggered by a user clicking the 'Activate' menu item.
   Notes: This function's execution by a user triggers the Google Auth prompt if needed.
   @returns {void}
*/
function fActivateMenus() {
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty(SCRIPT_INITIALIZED_KEY, 'true');

  const title = 'IMPORTANT - Please Refresh Browser Tab';
  const message = '✅ Success! The script has been authorized.\n\nPlease refresh this browser tab now to load the full custom menus.';
  SpreadsheetApp.getUi().alert(title, message, SpreadsheetApp.getUi().ButtonSet.OK);
} // End function fActivateMenus


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

/* function fMenuDeleteSelectedPowers
   Purpose: Local trigger for the "Delete Selected Powers" menu item.
   Assumptions: None.
   Notes: Acts as a pass-through to the central dispatcher in FlexLib.
   @returns {void}
*/
function fMenuDeleteSelectedPowers() {
  FlexLib.run('DeleteSelectedPowers', 'Powers');
} // End function fMenuDeleteSelectedPowers

/* function fMenuVerifyAndPublish
   Purpose: Local trigger for the "Verify & Publish Powers" menu item.
   Assumptions: None.
   Notes: Acts as a pass-through to the central dispatcher in FlexLib.
   @returns {void}
*/
function fMenuVerifyAndPublish() {
  FlexLib.run('VerifyAndPublish', 'Powers');
} // End function fMenuVerifyAndPublish

/* function fMenuVerifyAndPublishMagicItems
   Purpose: Local trigger for the "Verify & Publish Items" menu item.
   Assumptions: None.
   Notes: Acts as a pass-through to the central dispatcher in FlexLib.
   @returns {void}
*/
function fMenuVerifyAndPublishMagicItems() {
  FlexLib.run('VerifyAndPublishMagicItems', 'Magic Items');
} // End function fMenuVerifyAndPublishMagicItems

/* function fMenuDeleteSelectedMagicItems
   Purpose: Local trigger for the "Delete Selected Items" menu item.
   Assumptions: None.
   Notes: Acts as a pass-through to the central dispatcher in FlexLib.
   @returns {void}
*/
function fMenuDeleteSelectedMagicItems() {
  FlexLib.run('DeleteSelectedMagicItems', 'Magic Items');
} // End function fMenuDeleteSelectedMagicItems


/* function fMenuVerifyAndPublishSkillSets
   Purpose: Local trigger for the "Verify & Publish Skill Sets" menu item.
   Assumptions: None.
   Notes: Acts as a pass-through to the central dispatcher in FlexLib.
   @returns {void}
*/
function fMenuVerifyAndPublishSkillSets() {
  FlexLib.run('VerifyAndPublishSkillSets', 'SkillSets');
} // End function fMenuVerifyAndPublishSkillSets

/* function fMenuDeleteSelectedSkillSets
   Purpose: Local trigger for the "Delete Selected Skill Sets" menu item.
   Assumptions: None.
   Notes: Acts as a pass-through to the central dispatcher in FlexLib.
   @returns {void}
*/
function fMenuDeleteSelectedSkillSets() {
  FlexLib.run('DeleteSelectedSkillSets', 'SkillSets');
} // End function fMenuDeleteSelectedSkillSets

