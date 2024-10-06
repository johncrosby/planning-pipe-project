function copyImportToMain() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var importSheet = ss.getSheetByName('Import');
    var mainSheet = ss.getSheetByName('Main');
    
    if (!importSheet || !mainSheet) {
      SpreadsheetApp.getActiveSpreadsheet().toast('Import or Main sheet not found.', 'Error', 3);
      return;
    }
  
    // Get the data from the Import sheet
    var importData = importSheet.getDataRange().getValues();
    
    // Check if Main sheet has no data (first run)
    var lastRowInMain = mainSheet.getLastRow();
    var isFirstRun = lastRowInMain === 0 || (lastRowInMain === 1 && mainSheet.getRange(1, 1).isBlank());
    
    if (isFirstRun) {
      // First run: Copy headers and data, add new columns, and initialize values
      copyHeadersAndData(importSheet, mainSheet, importData);
      insertAndInitializeNewColumns(mainSheet, importData.length);
    } else {
      // Subsequent runs: Copy only data based on headers, initialize new columns for new rows
      var lastRowInImport = importData.length;
      var newDataStartRow = lastRowInMain + 1;
  
      // Dynamically map and copy data to corresponding columns, excluding Smart Rating and Letter
      copyDataByHeaders(importSheet, mainSheet, newDataStartRow, importData.slice(1), ["Smart Rating", "Letter"]);
  
      // Initialize new columns for new rows
      initializeExistingColumns(mainSheet, newDataStartRow, lastRowInImport - 1);
    }
    
    // Freeze the top row with headers
    mainSheet.setFrozenRows(1);
    
    SpreadsheetApp.getActiveSpreadsheet().toast('Data from Import copied to Main sheet.', 'Done', 3);
  }
  
  // Helper function to copy headers and data from Import to Main
  function copyHeadersAndData(importSheet, mainSheet, importData) {
    // Copy headers from Import sheet
    var headers = importSheet.getRange(1, 1, 1, importSheet.getLastColumn()).getValues();
    mainSheet.getRange(1, 1, 1, headers[0].length).setValues(headers);
    
    // Copy data below the header in Main sheet
    mainSheet.getRange(2, 1, importData.length - 1, importData[0].length).setValues(importData.slice(1));
  }
  
  function insertAndInitializeNewColumns(sheet, numRows) {
    var siteRawAddrIndex = getColumnIndexByHeader(sheet, "Site Raw Addr");
  
    if (siteRawAddrIndex !== -1) {
      // Insert "Smart Rating" column to the left of "Site Raw Addr"
      sheet.insertColumnBefore(siteRawAddrIndex);
      sheet.getRange(1, siteRawAddrIndex).setValue('Smart Rating');
      
      // Insert "Letter" column to the right of "Site Raw Addr"
      sheet.insertColumnAfter(siteRawAddrIndex + 1);
      sheet.getRange(1, siteRawAddrIndex + 2).setValue('Letter');
  
      // Insert "Creation" column to the right of "Letter"
      sheet.insertColumnAfter(siteRawAddrIndex + 2);
      sheet.getRange(1, siteRawAddrIndex + 3).setValue('Creation');
      
      // Initialize the newly added columns and apply dropdown + formatting
      initializeNewColumns(sheet, 2, numRows - 1, siteRawAddrIndex, siteRawAddrIndex + 2, siteRawAddrIndex + 3);
    } else {
      SpreadsheetApp.getActiveSpreadsheet().toast('Site Raw Addr column not found.', 'Error', 3);
    }
  }
  
  // Helper function to initialize newly added columns with dropdown and conditional formatting
  function initializeNewColumns(sheet, startRow, numRows, ratingColIndex, letterColIndex, creationColIndex) {
    // Initialize "Smart Rating" with 0
    sheet.getRange(startRow, ratingColIndex, numRows, 1).setValue(0);
    
    // Initialize "Letter" with "Pending" and add dropdown + conditional formatting
    createDropdownWithFormatting(sheet, letterColIndex, startRow, numRows);
    
    // Initialize "Creation" with the current date and time
    var currentDateTime = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'MMM d HH:mm:ss');
    sheet.getRange(startRow, creationColIndex, numRows, 1).setValue(currentDateTime);
  }
  
  
// Helper function to initialize existing columns for new rows (subsequent runs)
function initializeExistingColumns(sheet, startRow, numRows) {
    var siteRawAddrIndex = getColumnIndexByHeader(sheet, "Site Raw Addr");
  
    if (siteRawAddrIndex !== -1) {
      var ratingColIndex = siteRawAddrIndex - 1;
      var letterColIndex = siteRawAddrIndex + 1;
      var creationColIndex = siteRawAddrIndex + 2;
  
      // Only initialize if all indices are valid
      if (ratingColIndex > 0 && letterColIndex > 0 && creationColIndex > 0) {
        initializeNewColumns(sheet, startRow, numRows, ratingColIndex, letterColIndex, creationColIndex);
      } else {
        SpreadsheetApp.getActiveSpreadsheet().toast('Column indices not found for one or more columns.', 'Error', 3);
      }
    } else {
      SpreadsheetApp.getActiveSpreadsheet().toast('Site Raw Addr column not found in Main sheet.', 'Error', 3);
    }
  }
  
  
  
  // Helper function to dynamically map and copy data based on headers, excluding specific headers
  function copyDataByHeaders(importSheet, mainSheet, startRow, importData, excludeHeaders) {
    var mainHeaders = mainSheet.getRange(1, 1, 1, mainSheet.getLastColumn()).getValues()[0];
    var importHeaders = importSheet.getRange(1, 1, 1, importSheet.getLastColumn()).getValues()[0];
  
    // Create a mapping of import header to main header indices
    var headerMap = {};
    for (var i = 0; i < importHeaders.length; i++) {
      var headerName = importHeaders[i].toLowerCase();
      if (excludeHeaders.indexOf(headerName) === -1) { // Skip excluded headers
        var mainIndex = mainHeaders.indexOf(importHeaders[i]);
        if (mainIndex !== -1) {
          headerMap[i] = mainIndex + 1; // Map Import column index to Main column index (1-based)
        }
      }
    }
  
    // Copy data from Import to Main based on header mapping
    for (var i = 0; i < importData.length; i++) {
      var rowData = importData[i];
      for (var j in headerMap) {
        var mainCol = headerMap[j];
        mainSheet.getRange(startRow + i, mainCol).setValue(rowData[j]);
      }
    }
  }
  
  // Helper function to get the column index by header name
  function getColumnIndexByHeader(sheet, headerName) {
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    for (var col = 0; col < headers.length; col++) {
      if (headers[col].toLowerCase() === headerName.toLowerCase()) {
        return col + 1;
      }
    }
    return -1; // Return -1 if header not found
  }
  
  function createDropdownWithFormatting(sheet, letterColIndex, startRow, numRows) {
    const letterRange = sheet.getRange(startRow, letterColIndex, numRows, 1);
  
    // Set the default value to "Pending" for all rows in the range
    letterRange.setValue('Pending');
  
    // Create the dropdown data validation rule
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Pending', 'Sent', 'Ignore'], true)
      .setAllowInvalid(false)
      .build();
    letterRange.setDataValidation(rule);
  
    // Define the conditional formatting rules
    const rules = sheet.getConditionalFormatRules();
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Pending')
      .setBackground('#FFFF00') // Yellow
      .setRanges([letterRange])
      .build());
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Sent')
      .setBackground('#00FF00') // Green
      .setRanges([letterRange])
      .build());
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Ignore')
      .setBackground('#D3D3D3') // Grey
      .setRanges([letterRange])
      .build());
    sheet.setConditionalFormatRules(rules);
  }
  