function onEditA() {
  createClocking("Schedule_A", "Clocking_A");
  backgrounder("Schedule_A", "Clocking_A");
}

function onEditB() {
  createClocking("Schedule_B", "Clocking_B");
  backgrounder("Schedule_B", "Clocking_B");
}

function DeleteNewSheets() {
  var newSheetName = /^Sheet[\d]+$/
  var ssdoc = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = ssdoc.getSheets();
  
  // delete all unauthorised sheets
  for (var i = 0; i < sheets.length; i++) {
    if (newSheetName.test(sheets[i].getName())) {
      ssdoc.deleteSheet(sheets[i])
    }
  }
}

