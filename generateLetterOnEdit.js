function onEdit(e) {
    var sheet = e.source.getActiveSheet();
    var range = e.range;
    
    // Check if we're on the "Main" sheet and in the "Letter" column
    if (sheet.getName() === 'Main') {
      var letterColIndex = getColumnIndexByHeader(sheet, "Letter");
      var siteRawAddrColIndex = getColumnIndexByHeader(sheet, "Site Raw Addr");
      
      // If the edit occurred in the "Letter" column
      if (range.getColumn() === letterColIndex) {
        var newValue = range.getValue();
        var oldValue = e.oldValue; // Get the previous value before the change
        
        // Check if the value has changed from "Pending" to "Sent"
        if (oldValue === 'Pending' && newValue === 'Sent') {
          // Get the row and corresponding Site Raw Addr
          var row = range.getRow();
          var siteRawAddr = sheet.getRange(row, siteRawAddrColIndex).getValue();
          
          // Create a Google Doc with the Site Raw Addr as the file name
          createLetterDocument(siteRawAddr);
        }
      }
    }
  }
  
  // Function to create the Google Doc in the specified folder
  function createLetterDocument(fileName) {
    // The folder ID for the "Generated Letters" folder
    var folderId ='1t16S_Sk_GAn_k5Yrp8O5-2cE3DDy0TIl';  // Replace with your actual folder ID
    var folder = DriveApp.getFolderById(folderId);
    
    // Create the Google Doc
    var doc = DocumentApp.create(fileName);
    
    // Move the Google Doc to the specified folder
    var docFile = DriveApp.getFileById(doc.getId());
    folder.addFile(docFile);
    DriveApp.getRootFolder().removeFile(docFile); // Remove from root folder if needed
    
    // Add content to the Google Doc (this can be customized)
    var body = doc.getBody();
    body.appendParagraph('This is an automatically generated letter for ' + fileName);
    
    // Save and close the document
    doc.saveAndClose();
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
  