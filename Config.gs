props = PropertiesService.getScriptProperties();

const ARCHIVE_SHEET_IDS = {
  "Work_Orders" : props.getProperty('WORK_ORDERS'),
  "Inputs" : props.getProperty('INPUTS'),
  "Outputs" : props.getProperty('OUTPUTS'),
  "LotChanges" : props.getProperty('LOT_CHANGES')
};

const SOURCE_SHEET_NAMES = {
  "Work_Orders" : "Work Orders for BQ",
  "Inputs" : "Inputs for BQ",
  "Outputs" : "Outputs for BQ",
  "LotChanges" : "NewLotRequests"
};

const SPREADSHEETS = {
  "Work_Log" : SpreadsheetApp.getActiveSpreadsheet(),
}

