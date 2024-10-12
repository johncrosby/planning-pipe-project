// generateLetterOnEdit.js
function onSheetEdit(e) {
  var sheet = e.source.getActiveSheet();
  var range = e.range;

  if (sheet.getName() === 'Main') {
    var letterColIndex = getColumnIndexByHeader(sheet, "Letter");
    var siteRawAddrColIndex = getColumnIndexByHeader(sheet, "Site Raw Addr");
    var appContactColIndex = getColumnIndexByHeader(sheet, "App Contact");  
    var appContactTitleColIndex = getColumnIndexByHeader(sheet, "App Contact Title");
    var appContactFirstnameColIndex = getColumnIndexByHeader(sheet, "App Contact Firstname");
    var appContactLastnameColIndex = getColumnIndexByHeader(sheet, "App Contact Lastname");
    var appNameColIndex = getColumnIndexByHeader(sheet, "App Name");
    var appRawAddrColIndex = getColumnIndexByHeader(sheet, "App Raw Addr");
    var headingColIndex = getColumnIndexByHeader(sheet, "Heading"); // Get the column index for Heading
    var agtEmailColIndex = getColumnIndexByHeader(sheet, "Agt Email"); // Column for agent email

    if (range.getColumn() === letterColIndex) {
      var newValue = range.getValue();
      var oldValue = e.oldValue;

      if (oldValue === 'Pending' && newValue === 'Sent') {
        var row = range.getRow();
        var siteRawAddr = sheet.getRange(row, siteRawAddrColIndex).getValue();
        var appContact = sheet.getRange(row, appContactColIndex).getValue();
        var appRawAddr = sheet.getRange(row, appRawAddrColIndex).getValue();
        var heading = sheet.getRange(row, headingColIndex).getValue(); // Get the heading value
        var agtEmail = sheet.getRange(row, agtEmailColIndex).getValue(); // Get agent email

        // Additional fields for the letter creation
        var appContactTitle = sheet.getRange(row, appContactTitleColIndex).getValue();
        var appContactFirstname = sheet.getRange(row, appContactFirstnameColIndex).getValue();
        var appContactLastname = sheet.getRange(row, appContactLastnameColIndex).getValue();
        var appName = sheet.getRange(row, appNameColIndex).getValue();

        // Create the letter and pass a callback to create the draft email after the letter is generated
        createLetterFromTemplate(siteRawAddr, appRawAddr, appContact, heading, appContactFirstname, appContact, appName, function(letterDocId) {
          createDraftEmail(agtEmail, siteRawAddr, appContact, letterDocId);
        });
      }
    }
  }
}

// Function to create the letter from a template
function createLetterFromTemplate(siteRawAddr, appRawAddr, appContact, heading, appContactFirstname, appContact, appName, callback) {
  // var templateId = '1NASG2m4-RoyZnIxFQPqFLsy9Comm1paSDuuqyjVdwjA'; // Template Document ID
  var templateId = '1UAQSoMpYma6gyMqzXt6q3wVpyeHR049lIsUwmRb9DcY';
  var folderId = '1t16S_Sk_GAn_k5Yrp8O5-2cE3DDy0TIl';  // Folder where letters will be saved
  var folder = DriveApp.getFolderById(folderId);

  // Get today's date in the desired format
  var today = formatDate(new Date());

  // Copy the template file
  var templateDoc = DriveApp.getFileById(templateId).makeCopy(today + '-' + siteRawAddr, folder);
  var doc = DocumentApp.openById(templateDoc.getId());
  var body = doc.getBody();

  // Format the address for postage
  var formattedAddress = formatAddress(appContact, appName, appRawAddr);

  // Replace the placeholders with actual values
  body.replaceText('{{SiteRawAddr}}', siteRawAddr);
  body.replaceText('{{AppRawAddr}}', formattedAddress);
  body.replaceText('{{Heading}}', heading);
  body.replaceText('{{Date}}', today);

  // Generate the greeting
  var greeting = generateGreeting(appContactFirstname, appContact, appName);
  body.replaceText('{{greeting}}', greeting);

  // Save and close the document
  doc.saveAndClose();

  // Ensure the document is available before continuing (check every 500ms up to 10 seconds)
  var maxRetries = 20;
  var retryCount = 0;
  var fileAvailable = false;

  while (!fileAvailable && retryCount < maxRetries) {
    try {
      DriveApp.getFileById(templateDoc.getId());
      fileAvailable = true; // File is available
    } catch (e) {
      retryCount++;
      Utilities.sleep(500); // Wait 500ms before retrying
    }
  }

  if (fileAvailable) {
    // Call the callback function with the document ID (to create the email draft)
    callback(templateDoc.getId());
  } else {
    Logger.log('File was not available within the timeout period.');
  }
}

// Function to create a draft email to the agent with PDF attachment
function createDraftEmail(agtEmail, siteRawAddr, appContact, letterDocId) {
  if (agtEmail) {
    var subject = "Regarding Property: " + siteRawAddr;
    var body = "Dear " + appContact + ",\n\nI would like to discuss the property located at " + siteRawAddr + ". Please find the attached letter regarding this matter.\n\nBest regards,\nJohn Crosby";

    // Convert the Google Doc to PDF (in memory)
    var pdfFile = convertDocToPdf(letterDocId, siteRawAddr);

    // Create the draft with the attached PDF
    GmailApp.createDraft(agtEmail, subject, body, {
      attachments: [pdfFile]
    });
    Logger.log("Draft email created with PDF attachment for " + agtEmail);
  } else {
    Logger.log("No agent email found for this entry.");
  }
}

// Function to convert a Google Doc to PDF and return the file object
function convertDocToPdf(docId, siteRawAddr) {
  var docFile = DriveApp.getFileById(docId);
  
  // Convert the document to PDF in memory
  var pdfBlob = docFile.getAs('application/pdf');
  pdfBlob.setName("Letter - " + siteRawAddr + ".pdf");
  
  // Return the PDF blob (not saved to Drive, only for email attachment)
  return pdfBlob;
}

// Helper function to format the address for postage
function formatAddress(appContact, appName, appRawAddr) {
  var addressParts = [];

  if (appContact) {
    addressParts.push(appContact);
  }

  if (appName) {
    addressParts.push(appName);
  }

  if (appRawAddr) {
    // Split the App Raw Addr by commas and add each part on a new line
    var formattedAppRawAddr = appRawAddr.split(",").map(function(part) {
      return part.trim();
    }).join("\n");
    addressParts.push(formattedAppRawAddr);
  }

  return addressParts.join("\n");
}

// Helper function to format today's date as "30th June 2024"
function formatDate(date) {
  var day = date.getDate();
  var month = date.toLocaleString('default', { month: 'long' });
  var year = date.getFullYear();

  // Add "st", "nd", "rd", or "th" to the day
  var suffix = getDaySuffix(day);
  return day + suffix + ' ' + month + ' ' + year;
}

// Helper function to get the day suffix (st, nd, rd, th)
function getDaySuffix(day) {
  if (day > 3 && day < 21) return 'th'; // For 11th to 20th
  switch (day % 10) {
    case 1: return 'st';
    case 2: return 'nd';
    case 3: return 'rd';
    default: return 'th';
  }
}

// Helper function to generate the greeting
function generateGreeting(firstname, contact, name) {
  let greeting = "";

  if (firstname) {
    greeting = `Dear ${firstname},`;
  } else if (contact) {
    greeting = `Dear ${contact},`;
  } else if (name) {
    greeting = `Dear ${name},`;
  } else {
    greeting = "Dear Sir/Madam,";
  }

  return greeting;
}
