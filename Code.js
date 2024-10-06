// Code.gs

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Create a custom menu
  ui.createMenu('Data Management')
    .addItem('Clear and Import Sheet', 'clearAndImportSheet')  // Clears current sheet and imports a CSV from Drive
    .addItem('Clean Up Imported Data', 'processImportedData')  // This is the main function that orchestrates both stages of the data cleanup
    .addItem('Transfer to Main Sheet', 'copyImportToMain')  // Copies data from the import sheet to the main sheet
    .addToUi();
}