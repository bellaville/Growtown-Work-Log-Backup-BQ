props = PropertiesService.getScriptProperties();

const SHEET_IDS = {
  "Work_Orders" : props.getProperty('WORK_ORDERS'),
  "Inputs" : props.getProperty('INPUTS'),
  "Outputs" : props.getProperty('OUTPUTS')
};

const SOURCE_SHEET_NAMES = {
  "Work_Orders" : "Work Orders for BQ",
  "Inputs" : "Inputs for BQ",
  "Outputs" : "Outputs for BQ"
};

