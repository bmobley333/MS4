// 💪MS4
/* global FlexLib, SpreadsheetApp, PropertiesService */

/* function onOpen
   Purpose: Simple trigger that runs automatically when the spreadsheet is opened.
   Assumptions: None.
   Notes: Its sole job is to call the library to build the custom menu.
   @returns {void}
*/
function onOpen() {
  const g = FlexLib.getGlobals();
  const adminEmails = [g.ADMIN_EMAIL, g.DEV_EMAIL].map(e => e.toLowerCase());
  const isAdmin = adminEmails.includes(Session.getActiveUser().getEmail().toLowerCase());

  if (isAdmin) {
    FlexLib.fCreateDesignerMenu('Tables');
    // Admin visibility state is no longer auto-changed
  } else {
    FlexLib.fCheckAndSetVisibility(false); // Ensure elements are HIDDEN for players
  }
} // End function onOpen


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

/* function fMenuClearProperties
   Purpose: Local trigger for the "Clear PropertiesService" menu item.
   Assumptions: None.
   Notes: Acts as a pass-through to the central dispatcher in FlexLib.
   @returns {void}
*/
function fMenuClearProperties() {
  FlexLib.run('ClearProperties');
} // End function fMenuClearProperties

/* function fMenuTest
   Purpose: Local trigger for the "Test" menu item.
   Assumptions: None.
   Notes: Acts as a pass-through to the central dispatcher in FlexLib.
   @returns {void}
*/
function fMenuTest() {
  FlexLib.run('Test');
} // End function fMenuTest

/* function fMenuVerifySkillSetLists
   Purpose: Local trigger for the "Verify" menu item under "Skill Sets".
   Assumptions: None.
   Notes: Acts as a pass-through to the central dispatcher in FlexLib.
   @returns {void}
*/
function fMenuVerifySkillSetLists() {
  FlexLib.run('VerifySkillSetLists', 'SkillSets');
} // End function fMenuVerifySkillSetLists

/* function fMenuVerifyIndividualSkills
   Purpose: Local trigger for the "Verify Skill Types" menu item.
   Assumptions: None.
   Notes: Acts as a pass-through to the central dispatcher in FlexLib.
   @returns {void}
*/
function fMenuVerifyIndividualSkills() {
  FlexLib.run('VerifyIndividualSkills');
} // End function fMenuVerifyIndividualSkills