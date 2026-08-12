function markWorkOrdersApproved() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ===== CONFIG =====
  const sourceSheetName = 'Work Orders for BQ';
  const workOrderSheetName = 'Work Orders';

  const sourceIdColumn = 1;    // Column in source sheet with work order IDs (A=1, B=2, etc.)
  const workOrderIdColumn = 1; // Column in work order sheet containing the work order ID
  const approvalColumn = 20;    // Column in work order sheet where "Approved" should be written
  const startRow = 2;          // Data starts on row 2 if row 1 is headers
  // ==================

  const sourceSheet = ss.getSheetByName(sourceSheetName);
  const workOrderSheet = ss.getSheetByName(workOrderSheetName);

  if (!sourceSheet || !workOrderSheet) {
    throw new Error('One or both sheet names are incorrect.');
  }

  const sourceLastRow = sourceSheet.getLastRow();
  const workOrderLastRow = workOrderSheet.getLastRow();

  if (sourceLastRow < startRow || workOrderLastRow < startRow) {
    return;
  }

  // Get IDs from source sheet
  const sourceIds = sourceSheet
    .getRange(startRow, sourceIdColumn, sourceLastRow - startRow + 1, 1)
    .getValues()
    .flat()
    .filter(id => id !== '' && id !== null);

  if (sourceIds.length === 0) {
    return;
  }

  const sourceIdSet = new Set(sourceIds.map(String));

  // Get all work order table rows
  const workOrderData = workOrderSheet
    .getRange(startRow, 1, workOrderLastRow - startRow + 1, workOrderSheet.getLastColumn())
    .getValues();

  // Get current values from approval column
  const approvalValues = workOrderSheet
    .getRange(startRow, approvalColumn, workOrderLastRow - startRow + 1, 1)
    .getValues();

  // Match IDs and set Approved
  for (let i = 0; i < workOrderData.length; i++) {
    const workOrderId = workOrderData[i][workOrderIdColumn - 1];
    if (sourceIdSet.has(String(workOrderId))) {
      approvalValues[i][0] = 'Approved';
      console.log(workOrderId + " marked as Approved (backed up)") // log for traceability
    }
  }

  // Write updated approval column back in one batch
  workOrderSheet
    .getRange(startRow, approvalColumn, approvalValues.length, 1)
    .setValues(approvalValues);

}

function markBackedUp(sourceSheetName, SheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const sourceIdColumn = 1;
  const IdColumn = 1;
  const BackedUpColumn = 11;
  const startRow = 2;
  // ==================

  const sourceSheet = ss.getSheetByName(sourceSheetName);
  const editSheet = ss.getSheetByName(SheetName);

  if (!sourceSheet || !editSheet) {
    throw new Error('One or both sheet names are incorrect.');
  }

  const sourceLastRow = sourceSheet.getLastRow();
  const LastRow = editSheet.getLastRow();

  if (sourceLastRow < startRow || LastRow < startRow) {
    return;
  }

  // Get IDs from source sheet
  const sourceIds = sourceSheet
    .getRange(startRow, sourceIdColumn, sourceLastRow - startRow + 1, 1)
    .getValues()
    .flat()
    .filter(id => id !== '' && id !== null);

  if (sourceIds.length === 0) {
    return;
  }

  const sourceIdSet = new Set(sourceIds.map(String));


  const IpOpData = editSheet
    .getRange(startRow, 1, LastRow - startRow + 1, editSheet.getLastColumn())
    .getValues();

  // Get current values from backed up column
  const backedUpValues = editSheet
    .getRange(startRow, BackedUpColumn, LastRow - startRow + 1, 1)
    .getValues();

  // Match IDs and set TRUE
  for (let i = 0; i < IpOpData.length; i++) {
    const id = IpOpData[i][IdColumn - 1];
    if (sourceIdSet.has(String(id))) {
      backedUpValues[i][0] = true;
      console.log(id + " marked as true (backed up)")
    }
  }

  // Write updated backed up column back in one batch
  editSheet
    .getRange(startRow, BackedUpColumn, backedUpValues.length, 1)
    .setValues(backedUpValues);

    
}


function markInputsBackedUp() {
  markBackedUp("Inputs for BQ", "Inputs");
}

function markOutputsBackedUp() {
  markBackedUp("Outputs for BQ", "Outputs");
}

