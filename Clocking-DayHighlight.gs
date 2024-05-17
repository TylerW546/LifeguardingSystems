function backgrounder(scheduleName, clockingName) {
  var scheduleSheet = SpreadsheetApp.getActive().getSheetByName(scheduleName);
  var scheduleData = scheduleSheet.getDataRange().getValues();
  var clockingSheet = SpreadsheetApp.getActive().getSheetByName(clockingName);
  var clockingData = clockingSheet.getDataRange().getValues();
  

  var lines = scheduleSheet
      .getRange(1, 1)
      .getNextDataCell(SpreadsheetApp.Direction.DOWN)
      .getRow()-1;
  var days = scheduleSheet
      .getRange(1, 1)
      .getNextDataCell(SpreadsheetApp.Direction.NEXT)
      .getColumn()-1;

  var formatArray = [];
  for (var line = 0; line < lines; line++) {
    formatArray.push([]);
    //if (clockingData[line+1][0].includes(",")) {
    //  formatArray.push(["white"]);
    //} else {
    //  formatArray.push(["black"]);
    //}
    for (var day = 0; day < days; day++) {
      if (clockingData[line+1][0].includes(",")) {
        if (!(scheduleData[line+1][day+1].includes("G") || scheduleData[line+1][day+1].includes("W") || scheduleData[line+1][day+1].includes("SFP") || scheduleData[line+1][day+1].includes("PC") || scheduleData[line+1][day+1].includes("N"))) {
          if (scheduleData[1][day+1].includes("Training")) {
            formatArray[line].push("#DDDDDD");
          } else {
            formatArray[line].push("#999999");
          }
           
        } else {
          if (scheduleData[line+1][day+1].includes("*")) {
            formatArray[line].push("#aa0088");
          } else if (scheduleData[line+1][day+1].includes("/")) {
                formatArray[line].push("#DDDDDD");
          } else {
            if (clockingData[line+1][day+2].includes("-") && clockingData[line+1][day+2].split("-")[0] != "" && clockingData[line+1][day+2].split("-")[1] != 0 && clockingData[line+1][day+2].replace(new RegExp('[0|1|2|3|4|5|6|7|8|9|:|,|-]', 'gi'), "").replaceAll(" ","") == "" && clockingData[line+1][day+2].replace(new RegExp('[0|1|2|3|4|5|6|7|8|9|,|-]', 'gi'), "").replaceAll(" ","").length%2==0 && clockingData[line+1][day+2].includes(":")) {
              if (scheduleData[1][day+1].includes("Rain")) {
                formatArray[line].push("#DDDDDD");
              } else {
                formatArray[line].push("#bbffbb");
              }
            } else if (clockingData[line+1][day+2] != "") {
              formatArray[line].push("#ff6060");
            } else {
              if (scheduleData[1][day+1].includes("Rain")) {
                formatArray[line].push("#DDDDDD");
              } else {
                formatArray[line].push("white");
              }
            }
          }
        }
      } else {
        formatArray[line].push("black");
      }
    }
  }
  toRange = clockingSheet.getRange(2,3,lines,days);
  toRange.setBackgrounds(formatArray);
}
