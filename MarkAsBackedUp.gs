function markBackedUp(sourceSpreadSheet, startRow, sourceSheetName, sourceIdCol, editSheetName, editIdCol, backedUpCol, backedUpValue) {
  
  const ss = SPREADSHEETS[sourceSpreadSheet];
  const sourceSheet = ss.getSheetByName(sourceSheetName);
  const editSheet = ss.getSheetByName(editSheetName);

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
    .getRange(startRow, sourceIdCol, sourceLastRow - startRow + 1, 1)
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
    .getRange(startRow, backedUpCol, LastRow - startRow + 1, 1)
    .getValues();

  // Match IDs and set TRUE
  for (let i = 0; i < IpOpData.length; i++) {
    const id = IpOpData[i][editIdCol - 1];
    if (sourceIdSet.has(String(id))) {
      backedUpValues[i][0] = backedUpValue;
      console.log(id + " marked as backed up)")
    }
  }

  // Write updated backed up column back in one batch
  editSheet
    .getRange(startRow, backedUpCol, backedUpValues.length, 1)
    .setValues(backedUpValues);

    
}


function markWorkOrdersBackedUp(){
  markBackedUp("Work_Log", 2, "Work Orders for BQ", 1, "Work Orders", 1, 20, "Approved");
}

function markInputsBackedUp() {
  markBackedUp("Work_Log", 2, "Inputs for BQ", 1, "Inputs", 1, 11, true);
}

function markOutputsBackedUp() {
  markBackedUp("Work_Log", 2, "Outputs for BQ", 1, "Outputs", 1, 11, true);
}

