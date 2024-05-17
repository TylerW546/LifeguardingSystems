function recreationTrigger() {
  var informationSheet = SpreadsheetApp.getActive().getSheetByName("Availability Information");
  var informationData = informationSheet.getDataRange().getValues();
  
  var availabilitySheet = SpreadsheetApp.getActive().getSheetByName("Coded Availability Schedule");
  var availabilityData = availabilitySheet.getDataRange().getValues();

  var endRow = informationSheet
      .getRange(4, 1)
      .getNextDataCell(SpreadsheetApp.Direction.DOWN)
      .getRow();

  if (informationSheet.getRange(2,14,1,1).isChecked()) {
    availabilitySheet.getDataRange().clearContent();
    recreateFullSchedule(informationSheet, informationData, availabilitySheet, endRow);
    informationSheet.getRange(2,14,1,1).setValue(false);
    assessWorkers();
  } else {
    if (!(SpreadsheetApp.getActiveSheet().getName() === "Coded Availability Schedule")) {
      fillAllSchedules();
    }
  }
}

function fixDate(input, informationData) {
  input = input.replaceAll(" ", "");
  if (input.toLowerCase() === "start") {
    return informationData[1][0]+"/"+informationData[1][3];
  } else if (input.toLowerCase() === "end") {
    return informationData[1][1]+"/"+informationData[1][3];
  } else {
    return input+"/"+informationData[1][3];
  }
}

function fillSchedule(informationData, availabilitySheet, endRow, nameCol, endCol, firstDate, overrideWithRestriction=false) {
  outputRange = availabilitySheet.getRange(2, nameCol, endRow-3, endCol-nameCol+1);
  outData = outputRange.getValues();
  availabilitySheet.setColumnWidths(nameCol,1,150);
  availabilitySheet.setColumnWidths(nameCol+1,endCol-nameCol,100);

  for (var row=0; row<=endRow-4; row++) {
    if (informationData[3+row][0].includes(",")) {
      for (var col=0; col<endCol-nameCol; col++) {
        var todayTime = firstDate.getTime() + col*24 * 60 * 60 * 1000;

        var resolved = false;

        //Check if day on
        var daysOn = informationData[3+row][13].split(",");
        for (var i=0; i<daysOn.length; i++) {
          if (todayTime === new Date(fixDate(daysOn[i], informationData)).getTime()) {
            resolved = true;
          }
        }

        //Check if in working range
        if (!resolved) {
          var startDate;
          if (!informationData[3+row][3].replaceAll(" ", "")) {
            startDate = new Date (fixDate("start", informationData)).getTime();
          } else {
            startDate = new Date (fixDate(informationData[3+row][3],informationData)).getTime();
          }
          
          var endDate;
          if (!informationData[3+row][4].replaceAll(" ", "")) {
            endDate = fixDate("end", informationData);
          } else {
            endDate = new Date (fixDate(informationData[3+row][4],informationData)).getTime();
          }

          if (todayTime < startDate || todayTime >= endDate) {
            outData[row][col+1]="-";
            resolved = true;
          }

        }

        //Check if day off
        
        if (!resolved) {
          var daysOff = informationData[3+row][5].split(",");
          for (var i=0; i<daysOff.length; i++) {
            if (daysOff[i].includes("-")) {
              var start = new Date(fixDate(daysOff[i].split("-")[0], informationData)).getTime();
              var end = new Date(fixDate(daysOff[i].split("-")[1], informationData)).getTime();
              if (todayTime >= start && todayTime <= end) {
                outData[row][col+1]="-";
                resolved = true;
              }
            } else {
              var day = new Date(fixDate(daysOff[i], informationData)).getTime();
              if (day === todayTime) {
                outData[row][col+1]="-";
                resolved = true;
              }
            }
          }

        }

        //Check if weekday off
        if (!resolved) {

          var weekday = new Date(todayTime).getDay();

          var daysOff = informationData[3+row][6+weekday].split(",");
          for (var i=0; i<daysOff.length; i++) {
            if (daysOff[i].includes("-")) {
              var start = new Date(fixDate(daysOff[i].split("-")[0], informationData));
              var end = new Date(fixDate(daysOff[i].split("-")[1], informationData));
              if (todayTime >= start && todayTime <= end) {
                outData[row][col+1]="-";
                resolved = true;
              }
            } else {
              var day = new Date(fixDate(daysOff[i], informationData));
              if (day === todayTime) {
                outData[row][col+1]="-";
                resolved = true;
              }
            }
          }

        }


        // Restrict Beach
        if ((overrideWithRestriction || outData[row][col+1]==="") && !(informationData[3+row][14]==="") && !(outData[row][col+1]==="-")) {
          var options = informationData[3+row][14].split(",");
          for (var i=0; i<options.length; i++) {
            var option = options[i];
            var dateRange = option.split(":")[1].slice(1,-1);
            var beaches = option.split(":")[0].slice(1,-1);

            if (dateRange.includes("-")) {
              var start = new Date(fixDate(dateRange.split("-")[0], informationData));
              var end = new Date(fixDate(dateRange.split("-")[1], informationData));
              if (todayTime >= start && todayTime <= end) {
                outData[row][col+1]=beaches;
                resolved = true;
              }
            } else {
              var day = new Date(fixDate(dateRange, informationData));
              if (day === todayTime) {
                outData[row][col+1]=beaches;
                resolved = true;
              }
            }
          }
        }

      }
    } else {
      if (row>1) { 
        for (var col=0; col<endCol-nameCol; col++) {
          outData[row][col+1] = "B";
        }
      } else {
        for (var col=0; col<endCol-nameCol; col++) {
          outData[row][col+1] = "";
        }
      }
    }

  }
  outputRange.setValues(outData);
}

function recreateFullSchedule(informationSheet, informationData, availabilitySheet, endRow) {
  var scheduleRanges = informationData[1][5].replaceAll(" ","").split(",");

  var column = 1;
  for (var schedule = 0; schedule<scheduleRanges.length; schedule++) {
    fillNames(informationSheet, availabilitySheet, endRow, column);
    addExtraTitles(availabilitySheet, endRow, column);
    
    var range = scheduleRanges[schedule].split("-");
    var start = new Date(range[0]+"/"+informationData[1][3]);
    var end = new Date(range[1]+"/"+informationData[1][3]);

    var rangeString = range[0]+"/"+informationData[1][3]+"-"+range[1]+"/"+informationData[1][3];
    
    
    var days = [];
    for (var day = start.getTime(); day<=end.getTime(); day += 24 * 60 * 60 * 1000) {
      days.push(new Date(day));
    }
    
    
    availabilitySheet.getRange(1,column,1,1).setValues([[rangeString]]);    
    
    
    toRange = availabilitySheet.getRange(1,column+1,1,1);    
    toRange.setValues([["=ARRAYFORMULA(LAMBDA(cell,INDEX(SPLIT(INDIRECT(ADDRESS(ROW(),COLUMN()-1)),\"-\"),0,1)+COLUMN(cell)-COLUMN())(INDIRECT(ADDRESS(ROW(),COLUMN())):INDIRECT(ADDRESS(ROW(),COLUMN()+INDEX(SPLIT(INDIRECT(ADDRESS(ROW(),COLUMN()-1)),\"-\"),0,2)-INDEX(SPLIT(INDIRECT(ADDRESS(ROW(),COLUMN()-1)),\"-\"),0,1)))))"]]);

    fillSchedule(informationData, availabilitySheet, endRow, column, column + (end.getTime()-start.getTime())/(24 * 60 * 60 * 1000)+1, start, true);
    addExtraInfo(availabilitySheet, endRow, column+1, column + (end.getTime()-start.getTime())/(24 * 60 * 60 * 1000)+1);

    column += (end.getTime()-start.getTime())/(24 * 60 * 60 * 1000)+2;
  }
}


function fillNames(informationSheet, availabilitySheet, endRow, toColumn) {
  var fromRange = informationSheet.getRange(4,1,endRow-3,1);
  var toRange = availabilitySheet.getRange(2,toColumn,endRow-3,1);
  
  fromRange.copyTo(toRange, SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  fromRange.copyTo(toRange, SpreadsheetApp.CopyPasteType.PASTE_COLUMN_WIDTHS, false);
}

function addExtraTitles(availabilitySheet, endRow, toColumn) {
  var toRange = availabilitySheet.getRange(endRow,toColumn,6,1);
  toRange.setValues([["People Scheduled"], ["Head Guards Scheduled"], ["People Available"], [""], ["Daily Score"], ["Score Per Guard"]]);
}

function addExtraInfo(availabilitySheet, endRow, startColumn, endColumn) {
  var scheduled = "=LAMBDA(lines,LEN(JOIN(\"\",ARRAYFORMULA(IF(REGEXMATCH(INDIRECT(ADDRESS(3,COLUMN())):INDIRECT(ADDRESS(lines,COLUMN())),\"[-|OFF]\"),\"\",IF(REGEXMATCH(INDIRECT(ADDRESS(3,1)):INDIRECT(ADDRESS(lines,1)),\",\"),\"n\",\"\"))))))(MATCH(\"@\",ARRAYFORMULA($A$2:$A&\"@\"),0))";
  var available = "=LAMBDA(lines,LEN(JOIN(\"\",ARRAYFORMULA(IF(REGEXMATCH(INDIRECT(ADDRESS(3,COLUMN())):INDIRECT(ADDRESS(lines,COLUMN())),\"[-]\"),\"\",IF(REGEXMATCH(INDIRECT(ADDRESS(3,1)):INDIRECT(ADDRESS(lines,1)),\",\"),\"n\",\"\"))))))(MATCH(\"@\",ARRAYFORMULA($A$2:$A&\"@\"),0))";
  var headGuardsScheduled = "=LAMBDA(lines,LEN(JOIN(\"\",ARRAYFORMULA(IF(REGEXMATCH(INDIRECT(ADDRESS(3,COLUMN())):INDIRECT(ADDRESS(lines,COLUMN())),\"[-|OFF]\"),\"\",IF(REGEXMATCH(INDIRECT(ADDRESS(3,1)):INDIRECT(ADDRESS(lines,1)),\",\"),\"n\",\"\"))))))(MATCH(\"B\",ARRAYFORMULA($B$2:$B20),0))";
  var score = "=LAMBDA(lines, SUM(ARRAYFORMULA(IF(REGEXMATCH(INDIRECT(ADDRESS(3,COLUMN())):INDIRECT(ADDRESS(lines,COLUMN())),\"[-|OFF]\"),\"\",IF(REGEXMATCH(INDIRECT(ADDRESS(3,1)):INDIRECT(ADDRESS(lines,1)),\",\"),INDIRECT(ADDRESS(5,18,,,\"Availability Information\")):INDIRECT(ADDRESS(5+lines,18,,,\"Availability Information\")),\"\")))))(MATCH(\"@\",ARRAYFORMULA($A$2:$A&\"@\"),0))"
  var scorePerScheduledGuard = "=ROUND(LAMBDA(lines, SUM(ARRAYFORMULA(IF(REGEXMATCH(INDIRECT(ADDRESS(3,COLUMN())):INDIRECT(ADDRESS(lines,COLUMN())),\"[-|OFF]\"),\"\",IF(REGEXMATCH(INDIRECT(ADDRESS(3,1)):INDIRECT(ADDRESS(lines,1)),\",\"),INDIRECT(ADDRESS(5,18,,,\"Availability Information\")):INDIRECT(ADDRESS(5+lines,18,,,\"Availability Information\")),\"\")))))(MATCH(\"@\",ARRAYFORMULA($A$2:$A&\"@\"),0))/LAMBDA(lines,LEN(JOIN(\"\",ARRAYFORMULA(IF(REGEXMATCH(INDIRECT(ADDRESS(3,COLUMN())):INDIRECT(ADDRESS(lines,COLUMN())),\"[-|OFF]\"),\"\",IF(REGEXMATCH(INDIRECT(ADDRESS(3,1)):INDIRECT(ADDRESS(lines,1)),\",\"),\"n\",\"\"))))))(MATCH(\"@\",ARRAYFORMULA($A$2:$A&\"@\"),0)),2)"
  
  var range = availabilitySheet.getRange(endRow,startColumn,1,endColumn-startColumn+1);
  range.setValue(scheduled);
  range = availabilitySheet.getRange(endRow+1,startColumn,1,endColumn-startColumn+1);
  range.setValue(headGuardsScheduled);
  range = availabilitySheet.getRange(endRow+2,startColumn,1,endColumn-startColumn+1);
  range.setValue(available);
  range = availabilitySheet.getRange(endRow+4,startColumn,1,endColumn-startColumn+1);
  range.setValue(score);
  range = availabilitySheet.getRange(endRow+5,startColumn,1,endColumn-startColumn+1);
  range.setValue(scorePerScheduledGuard);
}
