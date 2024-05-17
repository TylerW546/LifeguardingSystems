function fillAllSchedules() {
  var informationSheet = SpreadsheetApp.getActive().getSheetByName("Availability Information");
  var informationData = informationSheet.getDataRange().getValues();
  
  var availabilitySheet = SpreadsheetApp.getActive().getSheetByName("Coded Availability Schedule");

  var endRow = informationSheet
      .getRange(4, 1)
      .getNextDataCell(SpreadsheetApp.Direction.DOWN)
      .getRow();


  var scheduleRanges = informationData[1][5].replaceAll(" ","").split(",");

  var column = 1;
  for (var schedule = 0; schedule<scheduleRanges.length; schedule++) {
    var range = scheduleRanges[schedule].split("-");
    var start = new Date(range[0]+"/"+informationData[1][3]);
    var end = new Date(range[1]+"/"+informationData[1][3]); 

    fillSchedule(informationData, availabilitySheet, endRow, column, column + (end.getTime()-start.getTime())/(24 * 60 * 60 * 1000)+1, start);

    column += (end.getTime()-start.getTime())/(24 * 60 * 60 * 1000)+2;
  }
}

function assessWorkers() {
  var informationSheet = SpreadsheetApp.getActive().getSheetByName("Availability Information");
  var informationData = informationSheet.getDataRange().getValues();
  
  var availabilitySheet = SpreadsheetApp.getActive().getSheetByName("Coded Availability Schedule");
  var availabilityData = availabilitySheet.getDataRange().getValues();

  var endRow = informationSheet
      .getRange(4, 1)
      .getNextDataCell(SpreadsheetApp.Direction.DOWN)
      .getRow();


  var scheduleRanges = informationData[1][5].replaceAll(" ","").split(",");

  var column = 0;
  for (var schedule = 0; schedule<scheduleRanges.length; schedule++) {
    var range = scheduleRanges[schedule].split("-");
    var start = new Date(range[0]+"/"+informationData[1][3]);
    var end = new Date(range[1]+"/"+informationData[1][3]);

    thisColumn = [];
    for (var row = 1; row<=endRow-3; row++) {
      thisColumn.push([assessWorker(informationData, availabilityData, row, column, column + (end.getTime()-start.getTime())/(24 * 60 * 60 * 1000)+1, start)]);
    }
    availabilitySheet.getRange(2,column+1,endRow-3,1).setBackgrounds(thisColumn);

    dayRow = [];
    for (var col = column+1; col<=column + (end.getTime()-start.getTime())/(24 * 60 * 60 * 1000)+1; col++) {
      dayRow.push(assessDay(availabilityData, column, col, endRow));
    }
    availabilitySheet.getRange(1,column+2,1,(end.getTime()-start.getTime())/(24 * 60 * 60 * 1000)+1).setBackgrounds([dayRow]);

    column += (end.getTime()-start.getTime())/(24 * 60 * 60 * 1000)+2;
  }
}

function assessDay(availabilityData, nameCol, col, endRow) {
  var onHead = true;
  var headCount = 0;
  var otherCount = 0;

  for (var row = 2; row<=endRow-3; row++) {
    if (availabilityData[row][nameCol].includes(",")) {
      if (!(availabilityData[row][col]==="-" || availabilityData[row][col].toLowerCase()==="off")) {  
        if (onHead) {
          headCount++;
        } else {
          otherCount++;
        }
      }
    } else {
      onHead = false;
    }
  }
  if (headCount + otherCount < 24) {
    return "#ff6060";
  }
  if (headCount < 2) {
    return "#ffff00";
  }
  return "#96ff96";
}

function assessWorker(informationData, availabilityData, row, startColumn, endColumn) {
  if (availabilityData[row][startColumn].includes(",")) {
    var count = 0;
    for (var col = startColumn+1; col<=endColumn; col++) {
      if (!(availabilityData[row][col]==="-" || availabilityData[row][col].toLowerCase()==="off")) {
        count++;
      }
    }
    if (count > 10) {
      return "#ff6060";
    } else {
      return "#96ff96";
    }
  } else {
    return "#ffffff";
  }
}