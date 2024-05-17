function availabilityUpdater() {
  colorCodeFoundNames();
  scoreCodes();
}

function test() {
  const extremity = ((24-12)/24*128+128).toString(16)
  console.log("#" + extremity + "6060");
}

function colorCodeFoundNames() {
  var informationSheet = SpreadsheetApp.getActive().getSheetByName("Availability Information");
  var informationData = informationSheet.getDataRange().getValues();

  var contactSheet = SpreadsheetApp.getActive().getSheetByName('Contacts');
  var contactData = contactSheet.getDataRange().getValues();


  colors = [];
  for (var row=3; row<informationData.length; row++) {
    if (informationData[row][0].includes(",")) {
      var found = false;
      var row;
      for (contactsRow = 0; contactsRow < contactData.length; contactsRow++) {
        if (contactData[contactsRow][0].replace(/\s+/g,"").toLowerCase() === informationData[row][0].replace(/\s+/g,"").toLowerCase()) {
          found = true;
          break;
        }
      }
      if (found) {
        colors.push(["#96ff96"]);
      } else {
        colors.push(["#ff9696"]);
      }
    }
    else {
      colors.push(["#ffffff"]);
    }
  }

  informationSheet.getRange(4,1,colors.length,1).setBackgrounds(colors);
}

function scoreCodes() {
  var informationSheet = SpreadsheetApp.getActive().getSheetByName("Availability Information");
  var informationData = informationSheet.getDataRange().getValues();

  values = [];
  for (var row=3; row<informationData.length; row++) {
    if (informationData[row][0].includes(",")) {
      values.push(["=IF(REGEXMATCH(INDIRECT(ADDRESS(ROW(),1)),\",\"),IF(INDIRECT(ADDRESS(ROW(),3)),9,LAMBDA(scoreBeforeSwim,LAMBDA(swimScore,IF(swimScore=4,scoreBeforeSwim,IF(scoreBeforeSwim=8,scoreBeforeSwim-(swimScore*swimScore-(swimScore-1)*(swimScore-1)),IF(scoreBeforeSwim=6,scoreBeforeSwim-(swimScore*2-FLOOR(swimScore/3)),scoreBeforeSwim))))(4-INDIRECT(ADDRESS(ROW(),2))))(LAMBDA(ranking,IF(10-2*ranking<6,0,12-2*(ranking)-IF(ranking>0,2,0)))(LEN(JOIN(\"\",ARRAYFORMULA(IF(REGEXMATCH(INDIRECT(ADDRESS(4,1)):INDIRECT(ADDRESS(ROW(),1)),\",\"),\"\",\"n\"))))-1))),\"B\")"]);
    } else {
      values.push([""]);
    }
  }

  informationSheet.getRange(4,18,values.length,1).setValues(values);
}