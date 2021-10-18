/*

-- This script necessitates the following spreadsheet schema:

Sheet for auto-generated summary of risks called "Summary"
Sheet with checklist of hazards called "Generic Hazard Checklist"
 - First col contains checkboxes for each hazard
 - Second col contains ID for the hazard
Sheets with ID numbers in the first col matching those ID numbers in the checklist

All must be done within very secific row and column ranges which need to be
  de-cyphered from this code.

*/

function updateSummarySheet() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  var allSheets = spreadsheet.getSheets();
  var summarySheet = spreadsheet.getSheetByName("Summary");
  var checklistSheet = spreadsheet.getSheetByName("Generic Hazard Checklist");

  var lastRow = checklistSheet.getDataRange().getLastRow();
  var checklistData = checklistSheet.getSheetValues(11, 1, lastRow - 10, 2);
  
  var sheetNumberRegex = /^[0-9]+$/;
  var sheetNumber = 1;
  
  var hazardData = {};
  var hazardSheet = allSheets[sheetNumber];
  var summary = [];
  
  for (var row = 0; row < checklistData.length; row++) {
    var isTicked = checklistData[row][0];
    var id = checklistData[row][1].toString();
    
    if (!id) {
      continue;
    }
    
    if (id.match(sheetNumberRegex)) {
      // Move onto the next sheet
      hazardSheet = allSheets[++sheetNumber];
      hazardData = parseHazardSheet(hazardSheet);
      continue;
    }
    
    if (isTicked) {
      summary = summary.concat(hazardData[id]);
    }
  }
  var currentLength = Math.max(1, summarySheet.getDataRange().getLastRow() - 5);
  summarySheet.getRange(6, 1, currentLength, 19).clearContent();
  summarySheet.getRange(6, 1, summary.length, 19).setValues(summary);
}

function parseHazardSheet(sheet) {
  // Hack for now...
  var lastRow = sheet.getDataRange().getLastRow();
  var data = sheet.getRange(20, 1, lastRow - 19, 19).getValues();
  var info = {};
  
  var id = "";
  var endRow = data.length;
  while (!data[endRow-1][6] && endRow > 1){
    endRow--; 
  }
  
  for (var row = endRow-1; row >= 0; row--) {
    id = data[row][0].toString();
    if (id) {
      info[id] = data.slice(row, endRow);
      endRow = row;
    }
  }
  
  return info;
}
