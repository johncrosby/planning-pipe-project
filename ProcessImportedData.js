// ProcessImportedData.js

function processImportedData() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    // Stage 1: Cleans up the imported sheet without adding new columns
    cleanUpImportedSheet(sheet);
    
    // Stage 2: Add necessary columns and process the data.
    addColumnsAndProcessData(sheet);
    
    // Notify that the entire process is complete
    SpreadsheetApp.getActiveSpreadsheet().toast('Data cleaned up and processed successfully.', 'Done', 3);
  }
  
  // Function to clean up the imported sheet (format and sanitize the data)
  function cleanUpImportedSheet(sheet) {
    formatColumns(sheet, ["Stage", "Category", "Region", "County"], "title");
    formatColumns(sheet, ["Heading"], "lower");
    concatenateAppRawAddr(sheet);
    movePersonalNamesToAppContact(sheet);
    cleanUpHeadingColumn(sheet); // New function to tidy up the Heading column
    SpreadsheetApp.getActiveSpreadsheet().toast('Data cleanup completed.', 'Done', 3);
  }
  
  // Function to add necessary columns and process the data
  function addColumnsAndProcessData(sheet) {
    // Process "Agt Contact" first as it is further to the right and won't affect "App Contact"
    processNamesInContactColumn(sheet, "Agt Contact");
    
    // Process "App Contact" after adding columns for "Agt Contact"
    processNamesInContactColumn(sheet, "App Contact");
  
    // Move "Site Raw Addr" to the first column
    moveSiteRawAddrToFirstColumn(sheet);
  
    // Notify that the columns have been added and the data has been processed
    SpreadsheetApp.getActiveSpreadsheet().toast('Columns added and contact names processed successfully.', 'Done', 3);
  }
  
  // Function to clean up the "Heading" column by removing everything before the number
function cleanUpHeadingColumn(sheet) {
  var headingColIndex = getColumnIndexByHeader(sheet, "Heading");
  if (headingColIndex === -1) {
      SpreadsheetApp.getActiveSpreadsheet().toast('Heading column not found.', 'Error', 3);
      return;
  }

  var lastRow = sheet.getLastRow();
  for (var row = 2; row <= lastRow; row++) { // Assuming headers are in the first row
      var headingValue = sheet.getRange(row, headingColIndex).getValue();
      if (headingValue) {
          // Match any number and everything that comes after it
          var cleanedValue = headingValue.replace(/^.*?(\d+)/, '$1');
          sheet.getRange(row, headingColIndex).setValue(cleanedValue);
      }
  }

  SpreadsheetApp.getActiveSpreadsheet().toast('Heading column cleaned up.', 'Done', 3);
}