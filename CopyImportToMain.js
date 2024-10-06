// CopyImportToMain.js

function copyImportToMain() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var importSheet = ss.getSheetByName('Import');
    var mainSheet = ss.getSheetByName('Main');
    
    if (!importSheet) {
      SpreadsheetApp.getActiveSpreadsheet().toast('Import sheet not found.', 'Error', 3);
      return;
    }
    
    if (!mainSheet) {
      SpreadsheetApp.getActiveSpreadsheet().toast('Main sheet not found.', 'Error', 3);
      return;
    }
  
    // Get all data from Import sheet
    var importData = importSheet.getDataRange().getValues();
    
    // Check if the Main sheet is empty or contains headers only
    var lastRowInMain = mainSheet.getLastRow();
    var isEmpty = lastRowInMain === 0 || (lastRowInMain === 1 && mainSheet.getRange(1, 1).isBlank());
  
    // If Main sheet is empty or only has headers, add headers from Import sheet
    if (isEmpty) {
      var headers = importSheet.getRange(1, 1, 1, importSheet.getLastColumn()).getValues();
      mainSheet.getRange(1, 1, 1, headers[0].length).setValues(headers);
      mainSheet.getRange(2, 1, importData.length - 1, importData[0].length).setValues(importData.slice(1));
      insertAndFillColumns(mainSheet);
    } else {
      var newDataStartRow = lastRowInMain + 1;
      mainSheet.getRange(newDataStartRow, 1, importData.length, importData[0].length).setValues(importData);
      setExistingColumnsValues(mainSheet, newDataStartRow, importData.length);
    }
  
    mainSheet.setFrozenRows(1);
    SpreadsheetApp.getActiveSpreadsheet().toast('Data from Import copied to Main sheet.', 'Done', 3);
  }
  