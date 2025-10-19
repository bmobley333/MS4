// 💪MS4
/* global FlexLib, PropertiesService, SpreadsheetApp, Session */

const SCRIPT_INITIALIZED_KEY = 'CUST_INITIALIZED';

/* function onOpen
   Purpose: Simple trigger that runs automatically when the spreadsheet is opened.
   Assumptions: None.
   Notes: Builds menus based on authorization status and user identity (player vs. designer). Also applies data validation rules.
   @returns {void}
*/
function onOpen() {
  // --- DEBUGGING ---
  console.log('Cust onOpen triggered.');
  const currentUserEmail = Session.getActiveUser().getEmail().toLowerCase();
  console.log(`Current User: ${currentUserEmail}`);
  // --- END DEBUGGING ---

  const scriptProperties = PropertiesService.getScriptProperties();
  const isInitialized = scriptProperties.getProperty(SCRIPT_INITIALIZED_KEY);
  const g = FlexLib.getGlobals();
  // Define primary vs secondary admin
  const primaryAdminEmail = g.ADMIN_EMAIL.toLowerCase();
  const secondaryAdminEmail = g.DEV_EMAIL.toLowerCase();
  const isAdmin = currentUserEmail === primaryAdminEmail || currentUserEmail === secondaryAdminEmail;
  const isPrimaryAdmin = currentUserEmail === primaryAdminEmail;

  // --- DEBUGGING ---
  console.log(`Is Initialized: ${isInitialized}`);
  console.log(`Is Admin: ${isAdmin}`);
  console.log(`Is Primary Admin: ${isPrimaryAdmin}`);
  // --- END DEBUGGING ---

  // --- REVISED LOGIC v3 ---
  if (isPrimaryAdmin) {
    // If the user is the PRIMARY admin, ALWAYS show the full admin menus immediately.
    console.log('User is Primary Admin, creating full menus.'); // DEBUG
    FlexLib.fApplyPowerValidations();
    FlexLib.fApplyMagicItemValidations();
    FlexLib.fApplySkillSetValidations();
    FlexLib.fCreateCustMenu();
    FlexLib.fCreateDesignerMenu('Cust');
    // Admin visibility state is not auto-changed
  } else if (!isInitialized) {
    // If the sheet is NOT initialized, ALWAYS show only the activation menu.
    // This now applies to regular users AND the secondary admin (thebmobley@gmail.com).
    console.log('Sheet NOT initialized, creating activation menu.'); // DEBUG
    SpreadsheetApp.getUi()
      .createMenu(g.VersionName)
      .addItem(`▶️ Activate ${g.VersionName} Menus`, 'fActivateMenus')
      .addToUi();
  } else if (isAdmin) {
    // If the sheet IS initialized and the user is the SECONDARY admin.
    console.log('User is Secondary Admin, sheet IS initialized, creating full menus.'); // DEBUG
    FlexLib.fApplyPowerValidations();
    FlexLib.fApplyMagicItemValidations();
    FlexLib.fApplySkillSetValidations();
    FlexLib.fCreateCustMenu();
    FlexLib.fCreateDesignerMenu('Cust');
     // Admin visibility state is not auto-changed (secondary admin can see hidden things)
  } else {
    // If the sheet IS initialized and it's a regular user.
    console.log('User is Regular Player, sheet IS initialized, creating player menu.'); // DEBUG
    FlexLib.fApplyPowerValidations();
    FlexLib.fApplyMagicItemValidations();
    FlexLib.fApplySkillSetValidations();
    FlexLib.fCreateCustMenu();
    FlexLib.fCheckAndSetVisibility(false); // Ensure elements are HIDDEN for players
  }
  // --- END REVISED LOGIC v3 ---
  console.log('Cust onOpen finished.'); // DEBUG
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

/* function buttonVerifyAndPublishPowers
   Purpose: Local trigger for a button, mimics the "Verify & Publish Powers" menu item.
   Assumptions: None.
   Notes: Assign this function name to a button on the <Powers> sheet.
   @returns {void}
*/
function buttonVerifyAndPublishPowers() {
  FlexLib.run('VerifyAndPublish', 'Powers');
} // End function buttonVerifyAndPublishPowers

/* function buttonVerifyAndPublishMagicItems
   Purpose: Local trigger for a button, mimics the "Verify & Publish Items" menu item.
   Assumptions: None.
   Notes: Assign this function name to a button on the <Magic Items> sheet.
   @returns {void}
*/
function buttonVerifyAndPublishMagicItems() {
  FlexLib.run('VerifyAndPublishMagicItems', 'Magic Items');
} // End function buttonVerifyAndPublishMagicItems

/* function buttonVerifyAndPublishSkillSets
   Purpose: Local trigger for a button, mimics the "Verify & Publish Skill Sets" menu item.
   Assumptions: None.
   Notes: Assign this function name to a button on the <SkillSets> sheet.
   @returns {void}
*/
function buttonVerifyAndPublishSkillSets() {
  FlexLib.run('VerifyAndPublishSkillSets', 'SkillSets');
} // End function buttonVerifyAndPublishSkillSets

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

