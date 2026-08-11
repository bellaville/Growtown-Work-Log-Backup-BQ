/**
 * Copies data rows from a source spreadsheet to a destination spreadsheet.
 *
 * The first row is treated as the header and is not copied.
 *
 * @param {string} sheetName The name of the source and destination sheet.
 */
function copyTableRows(sheetName) {

  const DESTINATION_SPREADSHEET_ID = SHEET_IDS[sheetName];

  if (!sheetName || typeof sheetName !== 'string') {
    throw new Error('A valid sheetName parameter is required.');
  }

  const sourceSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  const destinationSpreadsheet =
    SpreadsheetApp.openById(DESTINATION_SPREADSHEET_ID);

  const sourceSheet = sourceSpreadsheet.getSheetByName(SOURCE_SHEET_NAMES[sheetName]);
  const destinationSheet = destinationSpreadsheet.getSheetByName(sheetName + "_Archive");

  if (!sourceSheet) {
    throw new Error(`Source sheet not found: ${sheetName}`);
  }

  if (!destinationSheet) {
    throw new Error(`Destination sheet not found: ${sheetName}`);
  }

  const sourceLastRow = sourceSheet.getLastRow();
  const sourceLastColumn = sourceSheet.getLastColumn();

  // The sheet contains only a header row or is empty.
  if (sourceLastRow <= 1 || sourceLastColumn === 0) {
    return {
      sheetName,
      rowsCopied: 0,
      message: 'No data rows found.'
    };
  }

  // Validate that both sheets have the same number of columns.
  const destinationLastColumn = destinationSheet.getLastColumn();

  if (sourceLastColumn !== destinationLastColumn) {
    throw new Error(
      `Column mismatch. Source has ${sourceLastColumn} columns, ` +
      `but destination has ${destinationLastColumn} columns.`
    );
  }

  // Read rows 2 through the last source row.
  const sourceValues = sourceSheet
    .getRange(
      2,                         // Start below the header
      1,                         // First column
      sourceLastRow - 1,         // Number of data rows
      sourceLastColumn           // Number of columns
    )
    .getValues();

  console.log(sourceValues);

  // Remove completely blank rows.
  const rowsToCopy = sourceValues.filter(row =>
    row.some(value => value !== '' && value !== null)
  );

  if (rowsToCopy.length === 0) {
    return {
      sheetName,
      rowsCopied: 0,
      message: 'Only blank rows were found.'
    };
  }

  // Append after the existing destination data.
  const destinationStartRow = destinationSheet.getLastRow() + 1;

  destinationSheet
    .getRange(
      destinationStartRow,
      1,
      rowsToCopy.length,
      sourceLastColumn
    )
    .setValues(rowsToCopy);

  return {
    sheetName,
    rowsCopied: rowsToCopy.length,
    destinationStartRow
  };
}

function backupWorkOrders() {
  copyTableRows("Work_Orders");
}

function backupInputs() {
  copyTableRows("Inputs");
}

function backupOutputs() {
  copyTableRows("Outputs");
}
