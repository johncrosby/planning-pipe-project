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

    if (range.getColumn() === letterColIndex) {
      var newValue = range.getValue();
      var oldValue = e.oldValue;

      if (oldValue === 'Pending' && newValue === 'Sent') {
        var row = range.getRow();
        var siteRawAddr = sheet.getRange(row, siteRawAddrColIndex).getValue();
        var appContact = sheet.getRange(row, appContactColIndex).getValue();
        var appRawAddr = sheet.getRange(row, appRawAddrColIndex).getValue();
        var heading = sheet.getRange(row, headingColIndex).getValue(); // Get the heading value

        // Additional fields for the letter creation
        var appContactTitle = sheet.getRange(row, appContactTitleColIndex).getValue();
        var appContactFirstname = sheet.getRange(row, appContactFirstnameColIndex).getValue();
        var appContactLastname = sheet.getRange(row, appContactLastnameColIndex).getValue();
        var appName = sheet.getRange(row, appNameColIndex).getValue();

        // Pass the new Heading field and other details to createLetterFromTemplate
        createLetterFromTemplate(siteRawAddr, appRawAddr, appContact, heading, appContactFirstname, appContact, appName);
      }
    }
  }
}

// function createLetterFromTemplate(siteRawAddr, appRawAddr, appContact, heading, appContactFirstname, appContact, appName) {
//   var templateId = '1NASG2m4-RoyZnIxFQPqFLsy9Comm1paSDuuqyjVdwjA'; // Template Document ID
//   var folderId = '1t16S_Sk_GAn_k5Yrp8O5-2cE3DDy0TIl';  // Folder where letters will be saved
//   var folder = DriveApp.getFolderById(folderId);

//   // Copy the template file
//   var templateDoc = DriveApp.getFileById(templateId).makeCopy('Letter - ' + siteRawAddr, folder);
//   var doc = DocumentApp.openById(templateDoc.getId());
//   var body = doc.getBody();

//   // Replace the placeholders with actual values
//   body.replaceText('{{SiteRawAddr}}', siteRawAddr);
//   body.replaceText('{{AppRawAddr}}', appRawAddr);
//   body.replaceText('{{Heading}}', heading); // Replace the heading placeholder

//   // Generate the greeting based on the data
//   var greeting = generateGreeting(appContactFirstname, appContact, appName);
//   Logger.log('Generated Greeting: ' + greeting); // Log the greeting to check its value

//   // Replace the {{greeting}} placeholder in the template
//   body.replaceText('{{greeting}}', greeting);

//   // Save and close the document
//   doc.saveAndClose();
// }

function createLetterFromTemplate(siteRawAddr, appRawAddr, appContact, heading, appContactFirstname, appContact, appName) {
  var templateId = '1NASG2m4-RoyZnIxFQPqFLsy9Comm1paSDuuqyjVdwjA'; // Template Document ID
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

  // Get today's date in the desired format
  var today = formatDate(new Date());

  // Replace the placeholders with actual values
  body.replaceText('{{SiteRawAddr}}', siteRawAddr);
  body.replaceText('{{AppRawAddr}}', formattedAddress); // Use the formatted address
  body.replaceText('{{Heading}}', heading); // Replace the heading placeholder
  body.replaceText('{{Date}}', today); // Insert today's formatted date

  // Generate the greeting based on the data
  var greeting = generateGreeting(appContactFirstname, appContact, appName);
  Logger.log('Generated Greeting: ' + greeting); // Log the greeting to check its value

  // Replace the {{greeting}} placeholder in the template
  body.replaceText('{{greeting}}', greeting);

  // Save and close the document
  doc.saveAndClose();
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