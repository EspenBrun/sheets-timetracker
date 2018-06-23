var PUNCH_OUT_ROW = NUMBER_OF_HEADER_ROWS + 1;
var PUNCH_OUT_CELL = 'C2';
var HOURS_CELL = 'D2';

function getActiveSheet() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function autoSort(sheet, columnIndex) {
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
    column: columnIndex,
    ascending: ASCENDING
  } );
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

function now() {
  var now = new Date();
  return Utilities.formatDate(now, 'Europe/Oslo', 'yyyy/MM/dd HH:mm')
}

function validatePunchIn() {
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
  var hours = (punchedOut - punchedIn ) / (1000 * 3600);
  setValue(HOURS_CELL, hours.toFixed(2));
}

function punchIn () {
  validatePunchIn();
  var row = getNextRow();
  setValue('A' + row, 'Espen');
  setValue('B' + row, now());
  autoSort(getActiveSheet(), 2);
  
}

function punchOut() {
  checkIfCanPunchOut();
  setValue(PUNCH_OUT_CELL, now());
  calculateHours();
  autoSort(getActiveSheet(), 4);
}

