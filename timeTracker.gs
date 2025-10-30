// The numeric index of the column you wish to keep auto-sorted. A = 1, B = 2,
// and so on.
var SORT_COLUMN_INDEX = 2;
// Whether to sort the data in ascending or descending order.
var ASCENDING = false;
// If you have header rows in your sheet, specify how many to exclude them from
// the sort.
var NUMBER_OF_HEADER_ROWS = 1;

var USER = 'Espen';
var PUNCH_OUT_ROW = NUMBER_OF_HEADER_ROWS + 1;
var PUNCH_OUT_CELL = 'C2';
var HOURS_CELL = 'D2';
var WORKDAY_CELL = 'F2';
var COMPANY_CELL = 'G2';
var YEAR_CELL = 'H2';

function getActiveSheet() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function autoSort(sheet, columnIndex) {
  // * @author Mike Branski (@mikebranski)
  // * @link https://gist.github.com/mikebranski/285b60aa5ec3da8638e5
  // Get the entire set of data for this sheet.
  var range = sheet.getDataRange();
  var name = sheet.getName();

  // Then, if there are any header rows, offset our range to remove them from
  // it; otherwise, they will end up being sorted as well.
  if (NUMBER_OF_HEADER_ROWS > 0) {
    // Setting the second parameter of offset() to 0 to prevents it from
    // shifting any columns. Note that row headers wouldn't make much
    // sense here, but this is where you would modify it if you
    // wanted support for those as well.
    range = range.offset(NUMBER_OF_HEADER_ROWS, 0);
  }
  console.log("testlog");
  console.log(name);
  // Perform the actual sort.
  if (name === 'Timetracker Horus') {
    range.sort( {
      column: columnIndex,
      ascending: ASCENDING
    });
  }
  console.log('Autosort done')
}

function setValue(cellName, value) {
  return getActiveSheet().getRange(cellName).setValue(value);
}

function getValue(cellName) {
  return getActiveSheet().getRange(cellName).getValue();
}

function getNextRow() {
  return getActiveSheet().getLastRow() + 1;
}

function formatDate(date) {
  return Utilities.formatDate(date, 'Europe/Oslo', 'yyyy/MM/dd HH:mm')
}

function validatePunchIn() {
  var firstPunchIn = getActiveSheet().getRange('A2').isBlank();
  if (firstPunchIn) {
    return
  }
  
  var alreadyPunchedIn = getActiveSheet().getRange(PUNCH_OUT_CELL).isBlank();
  if (alreadyPunchedIn) {
    throw ('You must punch out before you can punch in: Cell ' + PUNCH_OUT_CELL + ' is empty.');
  }
}

function checkIfCanPunchOut() {
  var alreadyPunchedOut = !getActiveSheet().getRange(PUNCH_OUT_CELL).isBlank();
  if (alreadyPunchedOut) {
    throw ('You must punch in before you can punch out: Cell ' + PUNCH_OUT_CELL + ' is not empty.');
  }
}

function calculateHours() {
  var punchedIn = new Date(getValue('B2'));
  var punchedOut = new Date(getValue(PUNCH_OUT_CELL));
  return (punchedOut - punchedIn ) / (1000 * 3600);
}

function punchIn () {
  console.log('starting punch in')
  validatePunchIn();
  var row = getNextRow();
  setValue('A' + row, USER);
  setValue('B' + row, formatDate(new Date()));
  autoSort(getActiveSheet(), 2);
}

function punchOut() {
  checkIfCanPunchOut();
  var now = new Date();
  setValue(PUNCH_OUT_CELL, formatDate(now));
  var hours = calculateHours();
  setValue(HOURS_CELL, hours.toFixed(2));
  setValue(WORKDAY_CELL, hours > 4)
  setValue(COMPANY_CELL, 'Horus')
  setValue(YEAR_CELL, now.getFullYear())
}
