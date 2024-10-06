// ClearAndImportSheet.js

function clearAndImportSheet() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    // Display a toast notification to indicate that clearing is starting
    SpreadsheetApp.getActiveSpreadsheet().toast('Clearing sheet...', 'Update', 3);
    
    // Clear the content of the entire sheet
    sheet.clear();
    
    // The file ID of the CSV file in Google Drive
    var fileId = '1RpW1GnEqm3KghxRh3mNvGK5KPtNRcpMR';
    
    // Import the CSV data from Google Drive
    importCsvData(fileId, sheet);
  
    // Display a toast notification to indicate that data import is complete
    SpreadsheetApp.getActiveSpreadsheet().toast('Data import completed.', 'Done', 3);
  }
  
  // Function to import CSV data
  function importCsvData(fileId, sheet) {
    var file = DriveApp.getFileById(fileId);
    var csvData = file.getBlob().getDataAsString();  // Get file content as a string
    var data = Utilities.parseCsv(csvData);  // Convert CSV string into an array
    
    // Place the CSV data into the sheet, starting from cell A1
    sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
  }
  