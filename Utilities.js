// Utilities.js

function testUtilities() {
  SpreadsheetApp.getActiveSpreadsheet().toast('Utilities.js is loaded!');
}


// Helper: Format columns based on the mode (e.g., Title Case or Lower Case)
function formatColumns(sheet, columnNames, mode) {
    columnNames.forEach(function(columnName) {
      var columnIndex = getColumnIndexByHeader(sheet, columnName);
      
      if (columnIndex !== -1) {
        var lastRow = sheet.getLastRow();
        var range = sheet.getRange(2, columnIndex, lastRow - 1, 1); 
        var values = range.getValues();
  
        for (var i = 0; i < values.length; i++) {
          if (typeof values[i][0] === 'string') {  
            if (mode === "title") {
              values[i][0] = toTitleCase(values[i][0]);
            } else if (mode === "lower") {
              values[i][0] = values[i][0].toLowerCase();
            }
          }
        }
        range.setValues(values);
      } else {
        SpreadsheetApp.getActiveSpreadsheet().toast(columnName + ' header not found.', 'Error', 3);
      }
    });
  }
  
  // Helper: Convert text to Title Case
  function toTitleCase(str) {
    return str.toLowerCase().replace(/\b\w/g, function(char) {
      return char.toUpperCase();
    });
  }
  
  // Helper: Concatenate "App Addr" fields into "App Raw Addr"
  function concatenateAppRawAddr(sheet) {
    var appAddr1Index = getColumnIndexByHeader(sheet, "App Addr1");
    var appAddr2Index = getColumnIndexByHeader(sheet, "App Addr2");
    var appAddr3Index = getColumnIndexByHeader(sheet, "App Addr3");
    var appAddr4Index = getColumnIndexByHeader(sheet, "App Addr4");
    var appPcodeIndex = getColumnIndexByHeader(sheet, "App Pcode");
    var appRawAddrIndex = getColumnIndexByHeader(sheet, "App Raw Addr");
  
    if (appAddr1Index === -1 || appPcodeIndex === -1 || appRawAddrIndex === -1) {
      SpreadsheetApp.getActiveSpreadsheet().toast('Address or Raw Addr headers not found.', 'Error', 5);
      return;
    }
  
    var lastRow = sheet.getLastRow();
  
    for (var row = 2; row <= lastRow; row++) {
      var appAddr1 = sheet.getRange(row, appAddr1Index).getValue();
      var appAddr2 = sheet.getRange(row, appAddr2Index).getValue();
      var appAddr3 = sheet.getRange(row, appAddr3Index).getValue();
      var appAddr4 = sheet.getRange(row, appAddr4Index).getValue();
      var appPcode = sheet.getRange(row, appPcodeIndex).getValue();
      
      var appRawAddr = [appAddr1, appAddr2, appAddr3, appAddr4, appPcode]
        .filter(Boolean)
        .join(", ");
      
      sheet.getRange(row, appRawAddrIndex).setValue(appRawAddr);
    }
  }
  
  // Helper: Move personal names from "App Name" to "App Contact"
  function movePersonalNamesToAppContact(sheet) {
    var appNameCol = getColumnIndexByHeader(sheet, "App Name");
    var appContactCol = getColumnIndexByHeader(sheet, "App Contact");
    
    if (appNameCol === -1 || appContactCol === -1) {
      Logger.log('App Name or App Contact column not found.');
      return;
    }
    
    var lastRow = sheet.getLastRow();
    var namePattern = /^(Mr|Mrs|Ms|Miss|Sir|Madam|Mr\/s|Mr\s&\sMrs|Mr\/Mrs)\b/i;
    
    for (var row = 2; row <= lastRow; row++) {
      var appName = sheet.getRange(row, appNameCol).getValue();
      if (namePattern.test(appName)) {
        sheet.getRange(row, appContactCol).setValue(appName);
        sheet.getRange(row, appNameCol).setValue('');
      }
    }
  }
  

  function getColumnIndexByHeader(sheet, headerName) {
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    for (var col = 0; col < headers.length; col++) {
      if (headers[col].toLowerCase() === headerName.toLowerCase()) {
        return col + 1;
      }
    }
    return -1; // Return -1 if header not found
  }

  
  // Helper: Process contact names in "App Contact" or "Agt Contact" columns
  function processNamesInContactColumn(sheet, contactColumnName) {
    var contactColIndex = getColumnIndexByHeader(sheet, contactColumnName);
    
    if (contactColIndex === -1) {
      SpreadsheetApp.getActiveSpreadsheet().toast(contactColumnName + ' column not found.', 'Error', 3);
      return;
    }
    
    var lastRow = sheet.getLastRow();
  
    // Add 3 columns to the right of the contact column
    sheet.insertColumnsAfter(contactColIndex, 3);
    
    // Set headers for the new columns
    sheet.getRange(1, contactColIndex + 1).setValue(contactColumnName + ' Title');
    sheet.getRange(1, contactColIndex + 2).setValue(contactColumnName + ' Firstname');
    sheet.getRange(1, contactColIndex + 3).setValue(contactColumnName + ' Lastname');
    
    for (var row = 2; row <= lastRow; row++) {
      var contactValue = sheet.getRange(row, contactColIndex).getValue().trim();
      
      // Skip if the contact value is empty
      if (!contactValue) {
        continue;
      }
      
      // Try to split the contact value into title, firstname, and lastname
      var nameParts = splitName(contactValue);
      
      if (nameParts) {  // If the split was successful, fill the new columns
        sheet.getRange(row, contactColIndex + 1).setValue(nameParts.title || '');
        sheet.getRange(row, contactColIndex + 2).setValue(nameParts.firstname || '');
        sheet.getRange(row, contactColIndex + 3).setValue(nameParts.lastname || '');
      }
    }
  
    // Notify that the process is complete for the column
    SpreadsheetApp.getActiveSpreadsheet().toast(contactColumnName + ' processing completed.', 'Done', 3);
  }
  
  // Helper: Split contact name into title, firstname, and lastname
  function splitName(contactValue) {
    var namePattern = /^(Mr|Mrs|Ms|Miss|Dr|Sir|Lady|Prof|Mr\/s)\s+(\w+)\s+(\w+)$/i;  // A pattern to match simple cases
    var match = contactValue.match(namePattern);
    
    if (match) {
      return {
        title: match[1],      // Title, e.g., "Mr"
        firstname: match[2],  // Firstname, e.g., "John"
        lastname: match[3]    // Lastname, e.g., "Doe"
      };
    }
    
    // If the pattern doesn't match, return null to indicate an ambiguous name
    return null;
  }

  // Utilities.js

// Existing utility functions...

// Helper: Move "Site Raw Addr" to the first column
// function moveSiteRawAddrToFirstColumn(sheet) {
//   var siteRawAddrIndex = getColumnIndexByHeader(sheet, "Site Raw Addr");

//   if (siteRawAddrIndex === -1) {
//     SpreadsheetApp.getActiveSpreadsheet().toast('Site Raw Addr column not found.', 'Error', 3);
//     return;
//   }

//   // Cut the "Site Raw Addr" column
//   var lastRow = sheet.getLastRow();
//   var siteRawAddrRange = sheet.getRange(1, siteRawAddrIndex, lastRow, 1);  // Include the header
//   var siteRawAddrValues = siteRawAddrRange.getValues();
  
//   // Delete the original column
//   sheet.deleteColumn(siteRawAddrIndex);

//   // Insert a new column at the first position
//   sheet.insertColumnBefore(1);

//   // Set the values of "Site Raw Addr" in the first column
//   sheet.getRange(1, 1, lastRow, 1).setValues(siteRawAddrValues);

//   // Notify that the move is complete
//   SpreadsheetApp.getActiveSpreadsheet().toast('Site Raw Addr moved to first column.', 'Done', 3);
// }

// Helper: Move "Site Raw Addr" to the first column
function moveSiteRawAddrToFirstColumn(sheet) {
  var siteRawAddrIndex = getColumnIndexByHeader(sheet, "Site Raw Addr");

  if (siteRawAddrIndex === -1) {
    SpreadsheetApp.getActiveSpreadsheet().toast('Site Raw Addr column not found.', 'Error', 3);
    return;
  }

  // Cut the "Site Raw Addr" column
  var lastRow = sheet.getLastRow();
  var siteRawAddrRange = sheet.getRange(1, siteRawAddrIndex, lastRow, 1);  // Include the header
  var siteRawAddrValues = siteRawAddrRange.getValues();
  
  // Delete the original column
  sheet.deleteColumn(siteRawAddrIndex);

  // Insert a new column at the first position
  sheet.insertColumnBefore(1);

  // Set the values of "Site Raw Addr" in the first column
  sheet.getRange(1, 1, lastRow, 1).setValues(siteRawAddrValues);

  // Notify that the move is complete
  SpreadsheetApp.getActiveSpreadsheet().toast('Site Raw Addr moved to first column.', 'Done', 3);
}

// Helper: Move "Heading" to the second column
function moveHeadingToSecondColumn(sheet) {
  var headingColIndex = getColumnIndexByHeader(sheet, "Heading");

  if (headingColIndex === -1) {
    SpreadsheetApp.getActiveSpreadsheet().toast('Heading column not found.', 'Error', 3);
    return;
  }

  var lastRow = sheet.getLastRow();
  var headingRange = sheet.getRange(1, headingColIndex, lastRow, 1);  // Include the header
  var headingValues = headingRange.getValues();

  // Delete the original "Heading" column
  sheet.deleteColumn(headingColIndex);

  // Insert a new column at the second position
  sheet.insertColumnBefore(2);

  // Set the values of "Heading" in the second column
  sheet.getRange(1, 2, lastRow, 1).setValues(headingValues);

  // Notify that the move is complete
  SpreadsheetApp.getActiveSpreadsheet().toast('Heading moved to second column.', 'Done', 3);
}
