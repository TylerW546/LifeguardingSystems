function createClocking(scheduleName, clockingName) {
  var clockingSheet = SpreadsheetApp.getActive().getSheetByName(clockingName);
  if (!clockingSheet.getRange(1,1,1,1).isChecked()) {
    return;
  }
  clockingSheet.getRange(1,1,1,1).uncheck();
  

  var protections = clockingSheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
  for (var i = 0; i < protections.length; i++) {
    protections[i].remove();
  }

  var protection = clockingSheet.getRange(1,1,500,2).protect();
  protection.removeEditors(protection.getEditors());
  var protection = clockingSheet.getRange(1,1,1,26).protect();
  protection.removeEditors(protection.getEditors());

  var scheduleSheet = SpreadsheetApp.getActive().getSheetByName(scheduleName);
  var contactSheet = SpreadsheetApp.getActive().getSheetByName('Contacts');

  var nameLines = copyNameCells(scheduleSheet, clockingSheet);
  clockingSheet.getRange(1,2,nameLines+1,1).setBackground("white");
  var days = copyDateCells(scheduleSheet, clockingSheet, scheduleName);
  fillEmails(clockingSheet, contactSheet, nameLines, days);
  protect(clockingSheet, contactSheet, nameLines, days);
}

function fillEmails(clockingSheet, contactSheet, nameLines, days) {
  var nameRange = clockingSheet.getRange(2,1,nameLines,1);
  var contactRange = contactSheet.getDataRange();
  var nameData = nameRange.getValues();
  var data = contactRange.getValues();

  var secondCol = [];
  for (var nameIndex = 0; nameIndex<nameLines; nameIndex++) {
    if (nameData[nameIndex][0].includes(",")) {
      secondCol.push(["=INDIRECT(ADDRESS(ROW(),"+(days+3)+"))+SUM({0,ARRAYFORMULA( LAMBDA(times, LAMBDA(fixedTime, IF(REGEXMATCH(fixedTime,\"-\"), LAMBDA(subtraction, LAMBDA(fixedSub, HOUR(fixedSub) + MINUTE(fixedSub)/60 + SECOND(fixedSub)/3600 )(IF(subtraction<0,subtraction+TIME(12,0,0),subtraction)))(TIMEVALUE(RIGHT(fixedTime,LEN(fixedTime)-SEARCH(\"-\", fixedTime)))-TIMEVALUE(LEFT(fixedTime,SEARCH(\"-\", fixedTime)-1))),\"\"))(times))(SPLIT(SUBSTITUTE(JOIN(\",\",INDIRECT(ADDRESS(ROW(),3)):INDIRECT(ADDRESS(ROW(),"+(days+2)+"))),\" \",\"\"),\",\")))})"]);
    } else {
      secondCol.push([""]);
    }
  }
  clockingSheet.getRange(2,2,nameLines,1).setValues(secondCol);
}

function protect(clockingSheet, contactSheet, nameLines, days) {
  var nameRange = clockingSheet.getRange(2,1,nameLines,1);
  var contactRange = contactSheet.getDataRange();
  var nameData = nameRange.getValues();
  var data = contactRange.getValues();

  for (var nameIndex = 0; nameIndex<nameLines; nameIndex++) {
    var entryRange = clockingSheet.getRange(nameIndex+2,3,1,days+2);
    
    var found = false;
    var row;
    for (row = 0; row < data.length; row++) {
      if (data[row][0].replace(/\s+/g,"").toLowerCase() === nameData[nameIndex][0].replace(/\s+/g,"").toLowerCase()) {
        found = true;
        break;
      }
    }
    var protection = entryRange.protect().setDescription("Protection for " + nameData[nameIndex][0]);
    protection.removeEditors(protection.getEditors());
    var color = "";
    if (nameData[nameIndex][0].includes(",")) {
      // Email not found
      color = "#CCCCCC";
    } else {
      // Not name
      color = "black"
    }
    if (found) {
      try {
        // Protected
        protection.addEditor(data[row][1]);
        color = "#96ff96";
      } catch {
        // Email failed, unprotected
        protection.remove();
        color = "#ff9696";
      }
    }
    clockingSheet.getRange(nameIndex+2,2,1,1).setBackground(color);
  }
}

function copyNameCells(scheduleSheet, clockingSheet) {
    var lines = scheduleSheet
      .getRange(1, 1)
      .getNextDataCell(SpreadsheetApp.Direction.DOWN)
      .getRow()-1;
    
    var fromRange = scheduleSheet.getRange(2,1,lines,1);
    var toRange = clockingSheet.getRange(2,1,lines,1);
    
    fromRange.copyTo(toRange, SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
    fromRange.copyTo(toRange, SpreadsheetApp.CopyPasteType.PASTE_COLUMN_WIDTHS, false);

    return lines;
}

function copyDateCells(scheduleSheet, clockingSheet, scheduleName) {
  var days = scheduleSheet
      .getRange(1, 1)
      .getNextDataCell(SpreadsheetApp.Direction.NEXT)
      .getColumn()-1;
    
    var fromRange = scheduleSheet.getRange(1,2,1,days);
    var toRange = clockingSheet.getRange(1,3,1,days);
  
    var fontColors = fromRange.getFontColors();
    var fonts = fromRange.getFontFamilies();
    var fontWeights = fromRange.getFontWeights();
    var fontStyles = fromRange.getFontStyles();
    
    toRange.setFontColors(fontColors);
    toRange.setFontFamilies(fonts);
    toRange.setFontWeights(fontWeights);
    toRange.setFontStyles(fontStyles);

    var toRange = clockingSheet.getRange(1,3+days,1,1);
    toRange.setValue("Extra Time");
    var toRange = clockingSheet.getRange(1,4+days,1,1);
    toRange.setValue("Extra Time Description");
    var toRange = clockingSheet.getRange(1,3,1,1);
    toRange.setValue("=ARRAYFORMULA(LAMBDA(cell,INDEX(SPLIT('"+scheduleName+"'!$A$1,\"-\"),0,1)+COLUMN(cell)-COLUMN())(INDIRECT(ADDRESS(ROW(),COLUMN(),,,\""+(scheduleName)+"\")):INDIRECT(ADDRESS(ROW(),COLUMN()+INDEX(SPLIT('"+scheduleName+"'!$A$1,\"-\"),0,2)-INDEX(SPLIT('"+scheduleName+"'!$A$1,\"-\"),0,1),,,\""+(scheduleName)+"\"))))");

    return days;
}
