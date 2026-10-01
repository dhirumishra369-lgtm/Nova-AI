function createMasterMonthlyLedger() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Check or create sheets
  var sheetRM = ss.getSheetByName("Daily_RM_Issue");
  if (!sheetRM) {
    sheetRM = ss.insertSheet("Daily_RM_Issue");
    sheetRM.appendRow(["Date", "Item Code", "Item Description", "Qty Issued", "Unit", "Total Cost (₹)", "Photo Link"]);
  }
  
  var sheetProd = ss.getSheetByName("Daily_Production_Dispatch");
  if (!sheetProd) {
    sheetProd = ss.insertSheet("Daily_Production_Dispatch");
    sheetProd.appendRow(["Date", "Item Code", "Item Description", "Produced Qty", "Dispatched Qty", "Total Cost (₹)", "Total MRP (₹)", "Photo Link"]);
  }
  
  var sheetMaster = ss.getSheetByName("Master_Dashboard");
  if (!sheetMaster) {
    sheetMaster = ss.insertSheet("Master_Dashboard");
    sheetMaster.appendRow(["Item Description", "Total 30-Days Issued Qty", "Total 30-Days Produced Qty", "Total 30-Days Dispatched Qty", "Total Cost (₹)", "Total MRP (₹)"]);
  } else {
    sheetMaster.clear();
    sheetMaster.appendRow(["Item Description", "Total 30-Days Issued Qty", "Total 30-Days Produced Qty", "Total 30-Days Dispatched Qty", "Total Cost (₹)", "Total MRP (₹)"]);
  }
  
  // Formula to aggregate production and dispatch item-wise
  sheetMaster.getRange("A2").setFormula('=QUERY({Daily_Production_Dispatch!C2:G}, "SELECT Col2, SUM(Col3), SUM(Col4), SUM(Col5), SUM(Col6) WHERE Col2 IS NOT NULL GROUP BY Col2 LABEL SUM(Col3)\'\', SUM(Col4)\'\', SUM(Col5)\'\', SUM(Col6)\'\'", 0)');
  
  SpreadsheetApp.getUi().alert("Professional Master Ledger & Dashboard initialized successfully!");
}

// Function to log daily photo/image attachment links inside cells
function attachDailySlipPhoto(sheetName, rowIndex, columnIndex, imageUrl) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetName(sheetName);
  var cell = sheet.getRange(rowIndex, columnIndex);
  cell.setFormula('=HYPERLINK("' + imageUrl + '", "View Photo")');
}
