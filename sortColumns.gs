/**
 * This Google Sheets script keeps data in the specified column sorted any time
 * the data changes.
 *
 * After much research, there wasn't an easy way to automatically keep a column
 * sorted in Google Sheets, and creating a second sheet to act as a "view" to
 * my primary one in order to achieve that was not an option. Instead, I
 * created a script that watches for when a cell is edited and triggers
 * an auto sort.
 *
 * To Install:
 *   1. Open your Google Sheet.
 *   2. Navigate to Tools > Script editor…
 *   3. Copy and paste this script in the editor.
 *   4. Change the three constants at the start of the code below to reflect
 *      your preferences.
 *      - Note: My goal is to move these settings to a GUI and have this script
 *              be installable as an add-on.
 *   5. Give the script a name (e.g. "Keep Data Sorted") and hit save.
 *
 * To Use:
 *   Simply edit your Google Sheet like normal. Any time you edit data in your
 *   sort column (specified in `SORT_COLUMN_INDEX`), the script will re-sort
 *   your rows.
 *
 *   If you are having trouble getting it to work, try the following in order:
 *     1. Reload your spreadsheet.
 *     2. Open the script editor (Tools > Script editor…), click the "Select
 *        function" dropdown, choose `onInstall`, and hit Debug (the bug icon
 *        that precedes the dropdown).
 *     3. If that doesn't work, reach out via GitHub (link below) and ask for
 *        help. You may also find that others have run into the same issue
 *        and have already posted a solution.
 *
 * @author Mike Branski (@mikebranski)
 * @link https://gist.github.com/mikebranski/285b60aa5ec3da8638e5
 *
 * @OnlyCurrentDoc Limits the script to only accessing the current spreadsheet.
 */

// The numeric index of the column you wish to keep auto-sorted. A = 1, B = 2,
// and so on.
var SORT_COLUMN_INDEX = 2;
// Whether to sort the data in ascending or descending order.
var ASCENDING = false;
// If you have header rows in your sheet, specify how many to exclude them from
// the sort.
var NUMBER_OF_HEADER_ROWS = 1;

// No need to edit anything below this line for general use.
// Make an improvement? Ping me on GitHub and let me know!

// Keep track of the active sheet.
var activeSheet;

/**
 * Automatically sorts on the pre-defined column.
 *
 * @param {Sheet} sheet The sheet to sort.
 */
function autoSort(sheet) {
  // Get the entire set of data for this sheet.
  var range = sheet.getDataRange();

  // Then, if there are any header rows, offset our range to remove them from
  // it; otherwise, they will end up being sorted as well.
  if (NUMBER_OF_HEADER_ROWS > 0) {
    // Setting the second parameter of offset() to 0 to prevents it from
    // shifting any columns. Note that row headers wouldn't make much
    // sense here, but this is where you would modify it if you
    // wanted support for those as well.
    range = range.offset(NUMBER_OF_HEADER_ROWS, 0);
  }

  // Perform the actual sort.
  range.sort( {
    column: SORT_COLUMN_INDEX,
    ascending: ASCENDING
  } );
}

/**
 * Triggers when a sheet is edited, and calls the auto sort function if the
 * edited cell is in the column we're looking to sort.
 *
 * @param {Object} event The triggering event.
 */
function onEdit(event) {
  var editedCell;

  // Update the active sheet in case it changed.
  activeSheet = SpreadsheetApp.getActiveSheet();
  // Get the cell that was just modified.
  editedCell = activeSheet.getActiveCell();

  // Only trigger a re-sort if the user edited data in the column they're
  // sorting by; otherwise, we perform unnecessary additional sorts if
  // the targeted sort column's data didn't change.
  if (editedCell.getColumn() == SORT_COLUMN_INDEX) {
    autoSort(activeSheet);
  }
}

/**
 * Runs when the sheet is opened.
 *
 * @param {Object} event The triggering event.
 */
function onOpen(event) {
  activeSheet = SpreadsheetApp.getActiveSheet();
  autoSort(activeSheet);
}

/**
 * Runs when the add-on is installed; calls onOpen() to ensure any initializion
 * work is done immediately.
 *
 * @param {Object} event The triggering event.
 */
function onInstall(event) {
  onOpen(event);
}