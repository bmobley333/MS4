// 💪MS4
/* global g */
/* exported g, getGlobals */

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// End - n/a
// Start - Global Constants & State
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

const g = {
  // Developer Info
  ADMIN_EMAIL: 'metascapegame@gmail.com',
  DEV_EMAIL: 'TheBMobley@gmail.com',
  CURRENT_VERSION: '4',
  VersionName: '💪MScape-4',
  // The master source of truth for all game versions
  MASTER_VER_ID: '1wjX3P2GzDm733205yosvmSro7r0gbdOR2cEcvU7C3gE',
  MASTER_VER_INFO: {
    version: 'N/A',
    ssabbr: 'Ver',
    ssid: '1wjX3P2GzDm733205yosvmSro7r0gbdOR2cEcvU7C3gE',
    ssfullname: 'Versions',
  },

  // --- Session Caches ---
  codexSS: null, // Caches the Spreadsheet object for the Player's Codex.
  sheetIDs: {}, // Caches the mapping of Version -> Abbr -> full data object.

  // Object structures for sheet data (arrays and tag maps)
  Ver: {},
  Codex: {},
  CS: {},
  DB: {},
  Tbls: {},
}; // End const g

/* function getGlobals
   Purpose: A simple getter to make the global constants object (g) accessible to scripts that use FlexLib.
   Assumptions: None.
   Notes: Libraries do not expose global variables directly, so a getter function is required.
   @returns {object} The global g object.
*/
function getGlobals() {
  return g;
} // End function getGlobals