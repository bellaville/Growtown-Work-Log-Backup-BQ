props = PropertiesService.getScriptProperties();

const ARCHIVE_SHEET_IDS = {
  "Work_Orders" : props.getProperty('WORK_ORDERS'),
  "Inputs" : props.getProperty('INPUTS'),
  "Outputs" : props.getProperty('OUTPUTS'),
  "Order_Forms" : props.getProperty('ORDER_FORMS')
};

const SOURCE_SHEET_NAMES = {
  "Work_Orders" : "Work Orders for BQ",
  "Inputs" : "Inputs for BQ",
  "Outputs" : "Outputs for BQ",
  "Order_Forms" : "Order Forms for BQ"
};

const SPREADSHEETS = {
  "Work_Log" : SpreadsheetApp.getActiveSpreadsheet(),
}

