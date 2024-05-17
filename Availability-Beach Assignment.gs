function fillBeaches() {
  var availabilitySheet = SpreadsheetApp.getActive().getSheetByName("Coded Availability Schedule");
  var availabilityData = availabilitySheet.getDataRange().getValues();

  var beachSheet = SpreadsheetApp.getActive().getSheetByName("Beach Assignment");

  var infoSheet = SpreadsheetApp.getActive().getSheetByName("Availability Information");

  nameRows = infoSheet
      .getRange(4, 1)
      .getNextDataCell(SpreadsheetApp.Direction.DOWN)
      .getRow()-3;

  createOutline(availabilitySheet, beachSheet, nameRows);
  fillEmpties(beachSheet, infoSheet, nameRows);
  generateExtras(beachSheet, nameRows);
}

function createOutline(availabilitySheet, beachSheet, nameRows) {
  fromRange = availabilitySheet.getRange(1,1,nameRows+1,availabilitySheet.getDataRange().getNumColumns());
  
  toRange = beachSheet.getRange(1,1,fromRange.getNumRows()+1, fromRange.getNumColumns());

  fromRange.copyTo(toRange, SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  fromRange.copyTo(toRange, SpreadsheetApp.CopyPasteType.PASTE_COLUMN_WIDTHS, false);
  fromRange.copyTo(toRange, SpreadsheetApp.CopyPasteType.PASTE_CONDITIONAL_FORMATTING, false);
  toRange.setBackground("#ffffff");
}

function fillEmpties(beachSheet, infoSheet, nameRows) {
  var lastRow = nameRows;
  var lastColumn = beachSheet.getDataRange().getLastColumn()-1;
  
  var allPeople = beachSheet.getRange(2,1,lastRow,1).getValues().join("-+-").split("-+-");
  var peopleScores = {};
  
  for (var i=0; i< allPeople.length; i++) {
    if (allPeople[i].includes(",")) {
      peopleScores[allPeople[i]] = infoSheet.getDataRange().getValues()[3+i][17];
    }
  }
  
  var beachRange = beachSheet.getDataRange();
  var beachData = beachRange.getValues();
  var outData = beachData;

  for (var column=0; column<=lastColumn; column++) {
    if (beachData[0][column] instanceof Date) {
      var dailyPeopleScores = generateDailyShuffle(beachData[0][column], peopleScores);
      var peopleList = dailyPeopleScores.map(function(value,nn) { return value[0]; });
      var workingPeople = [];

      for (var i=0; i<peopleList.length; i++) {
        for (var row=1; row<=lastRow; row++) {
          if (beachData[row][0]===peopleList[i] && beachData[row][column]==="") {
            workingPeople.push(peopleList[i])
          }
        }
      }


      var gScorers = [];
      var wScorers = [];

      for (var row=1; row<=lastRow; row++) {
        if (beachData[row][0].includes(",")) {
          if (beachData[row][column]==="G") {
            gScorers.push(Math.floor(peopleScores[beachData[row][0]]/5)*5);
          } else if (beachData[row][column]==="W") {
            wScorers.push(Math.floor(peopleScores[beachData[row][0]]/5)*5);
          }
        }
      }

      for (var i=0; i<gScorers.length; i++) {
        if (wScorers.includes(gScorers[i])) {
          wScorers.splice(wScorers.indexOf(gScorers[i]), 1);
          gScorers.splice(i,1);
          i--;
        }
      }

      newWs = [];
      newGs = [];

      workingPeopleShuffled = workingPeople.map((x) => x);
      workingPeopleShuffled.sort((x) => (randomSeed(column*workingPeople.length+workingPeople.indexOf(x)) > .5) ? 1 : -1);
      for (var counter=0; counter<3 && (wScorers || gScorers); counter++) {
        for (var j=0; j<workingPeopleShuffled.length; j++) {
          if (gScorers.includes(Math.floor(peopleScores[workingPeopleShuffled[j]]/5)*5)) {
            newWs.push(workingPeopleShuffled[j]);
            gScorers.splice(gScorers.indexOf(peopleScores[workingPeopleShuffled[j]]), 1);
            workingPeople.splice(workingPeople.indexOf(workingPeopleShuffled[j]),1);
            j--;
          } else if (wScorers.includes(Math.floor(peopleScores[workingPeopleShuffled[j]]/5)*5)) {
            newGs.push(workingPeopleShuffled[j]);
            wScorers.splice(wScorers.indexOf(peopleScores[workingPeopleShuffled[j]]), 1);
            workingPeople.splice(workingPeople.indexOf(workingPeopleShuffled[j]),1);
            j--;
          }
        }
        for (var i=0; i<gScorers.length; i++) {
          gScorers[i] -= 3;
          wScorers[i] -= 3;
        }
      }

      for (var row=1; row<=lastRow; row++) {
        if (beachData[row][0].includes(",") && workingPeople.includes(beachData[row][0])) {
          if (workingPeople.indexOf(beachData[row][0])%2==Math.floor(2*randomSeed(beachData[0][column]))) {
            outData[row][column] = "G";
          } else {
            outData[row][column] = "W";
          }
        }
        else if (newGs.includes(beachData[row][0])) {
          outData[row][column] = "G";
        } else if (newWs.includes(beachData[row][0])){
          outData[row][column] = "W";
        }
      }
    }
  }

  beachRange.setValues(outData);
}

function generateExtras(beachSheet, nameRows) {
  var lastColumn = beachSheet.getDataRange().getLastColumn()-1;
  var beachData = beachSheet.getDataRange().getValues();
  
  var outArray = [];

  for (var i=0; i<12; i++) {
    outArray.push([]);
  }

  outRange = beachSheet.getRange(nameRows+4,1,12,lastColumn+1);
  
  for (var column=0; column<=lastColumn; column++) {
    for (var row=0; row<12; row++) {
      if (beachData[0][column] instanceof Date) {
        if (row<=6) {
          outArray[row][column]=[
            "=LAMBDA(lines, countif(INDIRECT(ADDRESS(3,COLUMN())):INDIRECT(ADDRESS(lines,COLUMN())),\"G\"))(MATCH(\"@\",ARRAYFORMULA($A$2:$A&\"@\"),0)-1)",
            "=LAMBDA(lines, countif(INDIRECT(ADDRESS(3,COLUMN())):INDIRECT(ADDRESS(lines,COLUMN())),\"W\"))(MATCH(\"@\",ARRAYFORMULA($A$2:$A&\"@\"),0)-1)",
            "=LAMBDA(lines, countif(INDIRECT(ADDRESS(3,COLUMN())):INDIRECT(ADDRESS(lines,COLUMN())),\"SFP\"))(MATCH(\"@\",ARRAYFORMULA($A$2:$A&\"@\"),0)-1)",
            "=LAMBDA(lines, countif(INDIRECT(ADDRESS(3,COLUMN())):INDIRECT(ADDRESS(lines,COLUMN())),\"N\"))(MATCH(\"@\",ARRAYFORMULA($A$2:$A&\"@\"),0)-1)",
            "=LAMBDA(lines, countif(INDIRECT(ADDRESS(3,COLUMN())):INDIRECT(ADDRESS(lines,COLUMN())),\"PC\"))(MATCH(\"@\",ARRAYFORMULA($A$2:$A&\"@\"),0)-1)",
            "",
            ""][row]
        } else {
          outArray[row][column] = "=LAMBDA(lines, TRIM(JOIN(\"\",ARRAYFORMULA(IF(REGEXMATCH(INDIRECT(ADDRESS(3,COLUMN())):INDIRECT(ADDRESS(lines,COLUMN())),INDIRECT(ADDRESS(ROW(),1))), IF(ISNUMBER(SEARCH(\"~*\",INDIRECT(ADDRESS(3,COLUMN())):INDIRECT(ADDRESS(lines,COLUMN())))),\"\",CHAR(10) & SUBSTITUTE(INDIRECT(ADDRESS(3,COLUMN())):INDIRECT(ADDRESS(lines,COLUMN())),INDIRECT(ADDRESS(ROW(),1)),PROPER(TRIM(INDEX(SPLIT(INDIRECT(ADDRESS(3,1)):INDIRECT(ADDRESS(lines,1)), \",\"),0,2))) & \" \" & TRIM(LEFT(TRIM(INDEX(SPLIT(INDIRECT(ADDRESS(3,1)):INDIRECT(ADDRESS(lines,1)), \",\"),0,1)),1)))),\"\")))))(MATCH(\"@\",ARRAYFORMULA($A$2:$A&\"@\"),0))";
        }
      } else {
        outArray[row][column]=["TOTAL GHB", "TOTAL WINGA", "TOTAL SFP", "TOTAL NILES", "TOTAL PC", "", "", "G", "W", "SFP", "N", "PC"][row];
      }
    }
  }

  outRange.setValues(outArray);
}

function generateDailyShuffle(day, peopleScores) {
  var dailyPeopleScores = {};
      
  var i = 0;
  for (const [key, value] of Object.entries(peopleScores)) {
    dailyPeopleScores[key] = value + 1.9*randomSeed(day*Object.keys(peopleScores).length+i);
    i++;
  }

  // Create items array
  var items = Object.keys(dailyPeopleScores).map(function(key) {
    return [key, dailyPeopleScores[key]];
  });

  // Sort the array based on the second element
  items.sort(function(first, second) {
    return second[1] - first[1];
  });

  return items;
}

function randomSeed(seed) {
    return (Math.round(2147483647*16807*(Math.round((seed*2.71828183%1)*2147483647*16807)%2147483647)/2147483647)%2147483647)/2147483647
}
