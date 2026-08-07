/*
 * =====================================================================================
 * ⚠️ IMPORTANT DEPLOYMENT CONFIGURATION TO PREVENT "REQUEST ACCESS" PROMPTS ⚠️
 * =====================================================================================
 * This Google Apps Script Web App must be deployed with the following specific settings
 * to avoid prompting external homeowners (who do not have Google/Workspace accounts or
 * access to the underlying spreadsheets) to "request access" (Viewer, Editor, etc.):
 *
 * 1. Click "Deploy" > "New deployment"
 * 2. Select "Web app" as the deployment type.
 * 3. Set "Execute as" to: "Me" (your developer/workspace account)
 *    -> This ensures the script runs with your administrative privileges and can read/write
 *       to the Service Master sheet and individual lot files without prompting the homeowner.
 * 4. Set "Who has access" to: "Anyone"
 *    -> This allows anonymous or external homeowners clicking the transactional email
 *       buttons or using the homeowner portal to successfully load and interact.
 *
 * DO NOT select "User accessing the web app" or restrict access to "Anyone with Google account"
 * or "Only myself", as this will break the homeowner portal.
 * =====================================================================================
 */

// --- CONFIG CONSTANTS ---
const SHEET_ID = '1rhHn7mXKpjcGY7D1dSQ3K7aWCTFcjYrF_ls7OHmfp4E';
const SERVICE_MASTER_ID = '1L2-Mbq6Uebqih2qhfhf23A4NRG54aoNR4rQyCEBGCRU'; 
const HOMEOWNER_SHEET_NAME = 'Homeowners'; 
const WARRANTY_ROOT_FOLDER_ID = '1BUt0T6XIs3BgR7Red6EE4L2X9sdk7eHO';
const WARRANTY_TEMPLATE_ID = '1loBscI38L9vtywXvl65tohceruLTCF2jo6A1Px0ppxc';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Georgian Build Connect - Homeowner Portal')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function cleanText(str) {
  return String(str || "").trim().toLowerCase().replace(/\s+/g, ' ');
}

function validateHomeownerLogin(emailAddress, passcode) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName(HOMEOWNER_SHEET_NAME);
  if (!sheet) return { success: false, message: "System Error: Homeowners tab missing." };
  
  var data = sheet.getDataRange().getDisplayValues();
  var inputEmail = cleanText(emailAddress);
  var homeowner = null;
  
  for (var i = 1; i < data.length; i++) {
    var savedEmail = cleanText(data[i][3]); 
    if (savedEmail !== "" && savedEmail === inputEmail) {
      var savedPasscode = String(data[i][4]).trim(); 
      if (savedPasscode !== "" && savedPasscode === String(passcode).trim()) {
        homeowner = { 
          success: true, 
          firstName: data[i][0], 
          lastName: data[i][1],
          email: data[i][3],
          project: data[i][5],
          phase: data[i][6],
          lot: data[i][7],
          address: "Address Not Found",
          model: "Model Not Found",
          lastTask: "Not Available",
          folderUrl: "" 
        }; 
        break;
      }
    }
  }

  // Cross-reference Lots tab to pull Model, Address, Last Task, and Folder URL
  if (homeowner) {
     var lotsSheet = ss.getSheetByName('Lots');
     if (lotsSheet) {
        var lotsData = lotsSheet.getDataRange().getDisplayValues();
        
        var tProj = String(homeowner.project).toLowerCase().trim();
        var tPhase = String(homeowner.phase).toLowerCase().trim();
        var tLot = String(homeowner.lot).toLowerCase().trim();
        
        for (var l = 1; l < lotsData.length; l++) {
           var sheetProj = String(lotsData[l][0]).toLowerCase().trim();
           var sheetPhase = String(lotsData[l][1]).toLowerCase().trim();
           var sheetLot = String(lotsData[l][2]).toLowerCase().trim();
           
           if (sheetProj === tProj && sheetPhase === tPhase && sheetLot === tLot) {
              homeowner.address = lotsData[l][3] || "Address Not Found";
              homeowner.model = lotsData[l][4] || "Model Not Found";
              homeowner.lastTask = lotsData[l][15] || "Not Available"; // Pulls Column P (Index 15)
              homeowner.folderUrl = lotsData[l][17] || ""; // Fetches Column R (Index 17)
              break;
           }
        }
     }
     return homeowner;
  }
  
  return { success: false, message: "Invalid email credentials or passcode. Please try again." };
}

function getHomeownerWarrantyItems(project, phase, lot) {
  var sheet = SpreadsheetApp.openById(SERVICE_MASTER_ID).getSheets()[0];
  var data = sheet.getDataRange().getDisplayValues();
  var items = [];
  
  var targetProject = cleanText(project);
  var targetPhase = cleanText(phase);
  var targetLot = cleanText(lot);
  
  for (var i = 1; i < data.length; i++) {
    if (cleanText(data[i][1]) === targetProject && cleanText(data[i][2]) === targetPhase && cleanText(data[i][3]) === targetLot) {
      items.push({
        formType: data[i][4] || "",
        itemNum: data[i][5] || "",
        location: data[i][6] || "",
        roomArea: data[i][7] || "",
        defectArea: data[i][8] || "",
        description: data[i][9] || "",
        additionalInfo: data[i][10] || "",
        scheduledDate: data[i][13] || "",
        warranted: data[i][15] || "",
        status: data[i][16] || "Unreviewed"
      });
    }
  }
  return items;
}

function getOrCreateFolder(parentFolder, folderName) {
  var safeName = folderName.replace(/'/g, "\\'");
  var folders = parentFolder.searchFolders("title = '" + safeName + "' and trashed = false");
  if (folders.hasNext()) { return folders.next(); }
  return parentFolder.createFolder(folderName);
}

function addManualWarrantyItem(formData) {
  try {
    var projectName = cleanText(formData.project);
    var phase = cleanText(formData.phase).replace(/^0+/, '');
    var lot = cleanText(formData.lot).replace(/^0+/, '');
    var formType = formData.formType;
    var initialStatus = formData.statusOverride ? formData.statusOverride : "Unreviewed";

    var rootFolder = DriveApp.getFolderById(WARRANTY_ROOT_FOLDER_ID);
    
    // Check or build directory path
    var projectFolder = getOrCreateFolder(rootFolder, projectName);
    var phaseFolder = getOrCreateFolder(projectFolder, "Phase " + phase);
    var lotFolder = getOrCreateFolder(phaseFolder, "Lot " + lot);
    var formTypeSubfolder = getOrCreateFolder(lotFolder, formType);

    var sheetName = "Lot " + lot + " - Warranty File";
    var existingFiles = lotFolder.searchFiles("title = '" + sheetName.replace(/'/g, "\\'") + "' and mimeType = 'application/vnd.google-apps.spreadsheet'");
    var ss;

    // Create or grab Spreadsheet
    if (existingFiles.hasNext()) {
      ss = SpreadsheetApp.openById(existingFiles.next().getId());
    } else {
      if (WARRANTY_TEMPLATE_ID && WARRANTY_TEMPLATE_ID.length > 20) {
        var templateFile = DriveApp.getFileById(WARRANTY_TEMPLATE_ID);
        var newFile = templateFile.makeCopy(sheetName, lotFolder);
        ss = SpreadsheetApp.openById(newFile.getId());
        ss.getSheets()[0].clear();
      } else {
        ss = SpreadsheetApp.create(sheetName);
        DriveApp.getFileById(ss.getId()).moveTo(lotFolder);
      }
      
      // Build baseline Homeowner Info tab
      var firstTab = ss.getSheets()[0];
      firstTab.setName("Homeowner Info");
      firstTab.appendRow(["Homeowner Contact Information", ""]);
      firstTab.appendRow(["Name(s):", ""]); 
      firstTab.appendRow(["Email(s):", ""]); 
      firstTab.appendRow(["Phone(s):", ""]); 
      firstTab.appendRow(["", ""]);
      firstTab.appendRow(["Project Details", ""]);
      firstTab.appendRow(["Project Name:", projectName]);
      firstTab.appendRow(["Phase:", phase]);
      firstTab.appendRow(["Lot Number:", lot]);
      firstTab.appendRow(["Enrolment #:", ""]);
      firstTab.appendRow(["Vendor #:", ""]);
      
      firstTab.getRange("A1:B1").setFontWeight("bold");
      firstTab.getRange("A6:B6").setFontWeight("bold");
      firstTab.getRange("A2:A4").setFontWeight("bold");
      firstTab.getRange("A7:A11").setFontWeight("bold");
      firstTab.setColumnWidth(1, 150);
      firstTab.setColumnWidth(2, 300);
    }

    // Build or grab Form tab
    var formTab = ss.getSheetByName(formType);
    if (!formTab) {
      formTab = ss.insertSheet(formType);
      formTab.appendRow(["Item #", "Location", "Room/Area", "Item/Defect Area", "Description", "Additional Information", "Assigned Trade", "Notified Status", "Scheduled Date", "Photo Link", "Warranted", "Status", "", "Case ID #"]);
      formTab.getRange("A1:N1").setFontWeight("bold");
      
      var statusList = ["Unreviewed", "Reviewed", "Assigned", "Completed", "Approved"];
      var statusRule = SpreadsheetApp.newDataValidation().requireValueInList(statusList, true).setAllowInvalid(true).build();
      formTab.getRange("L2:L100").setDataValidation(statusRule);

      var warrantRule = SpreadsheetApp.newDataValidation().requireValueInList(["Warrantable", "Not warrantable"], true).setAllowInvalid(true).build();
      formTab.getRange("K2:K100").setDataValidation(warrantRule);
    }

    // Determine the next Item #
    var fData = formTab.getDataRange().getValues();
    var nextItemNum = 1;
    for (var r = 1; r < fData.length; r++) {
      var val = parseInt(fData[r][0], 10);
      if (!isNaN(val) && val >= nextItemNum) {
        nextItemNum = val + 1;
      }
    }

    // Process and Save Photos using the Standard Naming Convention
    var photoUrls = [];
    if (formData.photos && formData.photos.length > 0) {
      for (var p = 0; p < formData.photos.length; p++) {
        var photoData = formData.photos[p];
        var ext = "";
        if (photoData.fileName.lastIndexOf('.') > -1) {
          ext = photoData.fileName.substring(photoData.fileName.lastIndexOf('.'));
        }
        var picNum = p + 1;
        var newFileName = phase + "-" + lot + "-Warranty for " + formType + "-Item " + nextItemNum + "-Pic " + picNum + ext;

        var blob = Utilities.newBlob(Utilities.base64Decode(photoData.base64), photoData.mimeType, newFileName);
        var savedFile = formTypeSubfolder.createFile(blob);
        photoUrls.push(savedFile.getUrl());
      }
    }
    var photoUrlsString = photoUrls.join("\n\n");

    // Append the new manual entry to the Lot Sheet
    formTab.appendRow([
      nextItemNum, 
      formData.location, 
      formData.roomArea, 
      formData.defectArea, 
      formData.description, 
      "", "", "", "", photoUrlsString, "", initialStatus, "", "Manual"
    ]);

    // Append to Master Sheet so the Homeowner sees it instantly on their portal grid
    try {
      var masterSS = SpreadsheetApp.openById(SERVICE_MASTER_ID);
      var masterSheet = masterSS.getSheets()[0];
      masterSheet.appendRow([
        new Date().toLocaleDateString(),
        projectName, phase, lot, formType, nextItemNum,
        formData.location, formData.roomArea, formData.defectArea, formData.description,
        "", "", "", "", photoUrlsString, "", initialStatus, "Manual"
      ]);
    } catch(masterErr) {
      // Fails silently if master is temporarily locked, it will catch it on the next builder Sync.
    }

    return "Success! Manual warranty item added as Item #" + nextItemNum;

  } catch (e) {
    return "Error: " + e.message;
  }
}

function deleteUnsubmittedWarrantyItem(formData) {
  try {
    var projectName = cleanText(formData.project);
    var nPhase = cleanText(formData.phase).replace(/^0+/, '');
    var nLot = cleanText(formData.lot).replace(/^0+/, '');
    var formType = formData.formType;
    var itemNum = parseInt(formData.itemNum, 10);

    // 1. Locate the Local Lot Spreadsheet
    var rootFolder = DriveApp.getFolderById(WARRANTY_ROOT_FOLDER_ID);
    var pFolder = rootFolder.searchFolders("title = '" + projectName.replace(/'/g, "\\'") + "' and trashed = false");
    if (!pFolder.hasNext()) return "Error: Project folder not found.";
    var phFolder = pFolder.next().searchFolders("title = 'Phase " + nPhase.replace(/'/g, "\\'") + "' and trashed = false");
    if (!phFolder.hasNext()) return "Error: Phase folder not found.";
    var lFolder = phFolder.next().searchFolders("title = 'Lot " + nLot.replace(/'/g, "\\'") + "' and trashed = false");
    if (!lFolder.hasNext()) return "Error: Lot folder not found.";

    var sheetName = "Lot " + nLot + " - Warranty File";
    var existingFiles = lFolder.next().searchFiles("title = '" + sheetName.replace(/'/g, "\\'") + "' and mimeType = 'application/vnd.google-apps.spreadsheet'");
    if (!existingFiles.hasNext()) return "Error: Warranty spreadsheet not found.";

    var ss = SpreadsheetApp.openById(existingFiles.next().getId());
    var formTab = ss.getSheetByName(formType);
    if (!formTab) return "Error: Form tab not found.";

    // 2. Find and delete the row in the lot spreadsheet
    var data = formTab.getDataRange().getValues();
    var rowIndex = -1;
    for (var r = 1; r < data.length; r++) {
      if (data[r][0] == itemNum && data[r][11] === "Unsubmitted") {
        rowIndex = r + 1;
        break;
      }
    }

    if (rowIndex > -1) {
      formTab.deleteRow(rowIndex);
    } else {
      return "Error: Item could not be found or is not Unsubmitted.";
    }

    // 3. Delete from the Master Service Sheet as well
    try {
      var masterSS = SpreadsheetApp.openById(SERVICE_MASTER_ID);
      var masterSheet = masterSS.getSheets()[0];
      var mData = masterSheet.getDataRange().getValues();
      for (var m = 1; m < mData.length; m++) {
        if (cleanText(mData[m][1]) === projectName && cleanText(mData[m][2]) === nPhase && cleanText(mData[m][3]) === nLot && mData[m][4] === formType && mData[m][5] == itemNum && mData[m][16] === "Unsubmitted") {
          masterSheet.deleteRow(m + 1);
          break;
        }
      }
    } catch(masterErr) {
      Logger.log("Error deleting from Master Service Sheet: " + masterErr.message);
    }

    return "Success! Item #" + itemNum + " has been deleted.";
  } catch (e) {
    return "Error: " + e.message;
  }
}

function submitUnsubmittedItems(project, phase, lot, hoName, hoEmail) {
  try {
    var projectName = cleanText(project);
    var nPhase = cleanText(phase).replace(/^0+/, '');
    var nLot = cleanText(lot).replace(/^0+/, '');

    // 1. Update the Local Lot Spreadsheet
    var rootFolder = DriveApp.getFolderById(WARRANTY_ROOT_FOLDER_ID);
    var pFolder = rootFolder.searchFolders("title = '" + projectName.replace(/'/g, "\\'") + "' and trashed = false");
    if (!pFolder.hasNext()) return "Error: Project folder not found.";
    var phFolder = pFolder.next().searchFolders("title = 'Phase " + nPhase.replace(/'/g, "\\'") + "' and trashed = false");
    if (!phFolder.hasNext()) return "Error: Phase folder not found.";
    var lFolder = phFolder.next().searchFolders("title = 'Lot " + nLot.replace(/'/g, "\\'") + "' and trashed = false");
    if (!lFolder.hasNext()) return "Error: Lot folder not found.";

    var sheetName = "Lot " + nLot + " - Warranty File";
    var existingFiles = lFolder.next().searchFiles("title = '" + sheetName.replace(/'/g, "\\'") + "' and mimeType = 'application/vnd.google-apps.spreadsheet'");
    if (!existingFiles.hasNext()) return "Error: Warranty spreadsheet not found.";
    
    var ss = SpreadsheetApp.openById(existingFiles.next().getId());
    var sheets = ss.getSheets();
    var itemsSubmitted = 0;
    var submittedDetails = [];

    for (var i = 0; i < sheets.length; i++) {
      var sName = sheets[i].getName();
      if (sName === "Homeowner Info") continue;
      
      var data = sheets[i].getDataRange().getValues();
      if (data.length <= 1) continue;
      
      var headers = data[0].map(function(h) { return String(h).toLowerCase().trim(); });
      var statusCol = headers.indexOf("status");
      var descCol = headers.indexOf("description");
      if (statusCol === -1) continue;

      for (var r = 1; r < data.length; r++) {
        if (data[r][statusCol] === "Unsubmitted") {
          sheets[i].getRange(r + 1, statusCol + 1).setValue("Unreviewed");
          itemsSubmitted++;
          var descText = (descCol > -1) ? data[r][descCol] : "No Description";
          submittedDetails.push(sName + " - Item #" + data[r][0] + ": " + descText);
        }
      }
    }

    if (itemsSubmitted === 0) return "No unsubmitted items found.";

    // 2. Update the Master Service Sheet so changes reflect instantly
    var masterSS = SpreadsheetApp.openById(SERVICE_MASTER_ID);
    var masterSheet = masterSS.getSheets()[0];
    var mData = masterSheet.getDataRange().getValues();
    for (var m = 1; m < mData.length; m++) {
      if (cleanText(mData[m][1]) === projectName && cleanText(mData[m][2]) === nPhase && cleanText(mData[m][3]) === nLot && mData[m][16] === "Unsubmitted") {
        masterSheet.getRange(m + 1, 17).setValue("Unreviewed");
      }
    }

    // 3. Find Service Coordinator Email in the Users tab and Send Notification
    var settingsSS = SpreadsheetApp.openById(SHEET_ID);
    var usersSheet = settingsSS.getSheetByName('Users');
    var targetEmails = [];

    if (usersSheet) {
      var uData = usersSheet.getDataRange().getDisplayValues();
      var pNameLower = String(project).toLowerCase().trim();
      
      for (var u = 1; u < uData.length; u++) {
        var uJob = String(uData[u][2]).toLowerCase().trim(); // Column C
        var uProjRaw = String(uData[u][3]).toLowerCase();    // Column D
        var uEmail = String(uData[u][4]).trim();             // Column E
        
        // STRICT MATCH: Job title must be exactly "service coordinator"
        if (uJob === "service coordinator" && uEmail !== "") {
          
          // Split the project cell by comma in case there are multiple assigned projects
          var userProjects = uProjRaw.split(',');
          var isMatch = false;
          
          for (var p = 0; p < userProjects.length; p++) {
            var cleanUserProj = userProjects[p].trim();
            // Project must exactly match the homeowner's project OR "all"
            if (cleanUserProj === pNameLower || cleanUserProj === "all") {
              isMatch = true;
              break;
            }
          }

          if (isMatch) {
             if (targetEmails.indexOf(uEmail) === -1) targetEmails.push(uEmail);
          }
        }
      }
    }

    // 4. Send Email (or return a warning if no coordinator was found)
    if (targetEmails.length > 0) {
      var subject = "New Warranty Items Submitted: " + project + " - Lot " + lot;
      var body = "Hello,\n\n" +
                 hoName + " (" + hoEmail + ") has submitted " + itemsSubmitted + " new warranty item(s) for " + project + " Phase " + phase + " Lot " + lot + ".\n\n" +
                 "Items Submitted:\n" + submittedDetails.join("\n") + "\n\n" +
                 "Please review them in the Service Portal.\n\n" +
                 "Thank you,\nGeorgian Build Connect";
      MailApp.sendEmail(targetEmails.join(","), subject, body, {name: "Georgian Build Connect"});
      
      return "Success! " + itemsSubmitted + " item(s) submitted to the Service Coordinator.";
    } else {
      return "Success! " + itemsSubmitted + " item(s) submitted. (Note: No matching 'service coordinator' found for this project in the Users directory).";
    }

  } catch(e) {
    return "Error: " + e.message;
  }
}

// NEW FUNCTION: Updates an existing manual warranty item
function updateManualWarrantyItem(formData) {
  try {
    var projectName = cleanText(formData.project);
    var nPhase = cleanText(formData.phase).replace(/^0+/, '');
    var nLot = cleanText(formData.lot).replace(/^0+/, '');
    var formType = formData.formType;
    var itemNum = parseInt(formData.itemNum, 10);

    // 1. Locate the Local Lot Spreadsheet
    var rootFolder = DriveApp.getFolderById(WARRANTY_ROOT_FOLDER_ID);
    var pFolder = rootFolder.searchFolders("title = '" + projectName.replace(/'/g, "\\'") + "' and trashed = false");
    if (!pFolder.hasNext()) return "Error: Project folder not found.";
    var phFolder = pFolder.next().searchFolders("title = 'Phase " + nPhase.replace(/'/g, "\\'") + "' and trashed = false");
    if (!phFolder.hasNext()) return "Error: Phase folder not found.";
    var lFolder = phFolder.next().searchFolders("title = 'Lot " + nLot.replace(/'/g, "\\'") + "' and trashed = false");
    if (!lFolder.hasNext()) return "Error: Lot folder not found.";

    var sheetName = "Lot " + nLot + " - Warranty File";
    var existingFiles = lFolder.next().searchFiles("title = '" + sheetName.replace(/'/g, "\\'") + "' and mimeType = 'application/vnd.google-apps.spreadsheet'");
    if (!existingFiles.hasNext()) return "Error: Warranty spreadsheet not found.";
    
    var ss = SpreadsheetApp.openById(existingFiles.next().getId());
    var formTab = ss.getSheetByName(formType);
    if (!formTab) return "Error: Form tab not found.";

    // 2. Find and update the exact row in the lot spreadsheet
    var data = formTab.getDataRange().getValues();
    var rowIndex = -1;
    
    for (var r = 1; r < data.length; r++) {
      // Column A (0) is Item #, Column L (11) is Status
      if (data[r][0] == itemNum && data[r][11] === "Unsubmitted") { 
        rowIndex = r + 1;
        break;
      }
    }

    if (rowIndex > -1) {
      formTab.getRange(rowIndex, 2).setValue(formData.location);
      formTab.getRange(rowIndex, 3).setValue(formData.roomArea);
      formTab.getRange(rowIndex, 4).setValue(formData.defectArea);
      formTab.getRange(rowIndex, 5).setValue(formData.description);
    } else {
      return "Error: Item could not be found or has already been submitted.";
    }

    // 3. Update the Master Service Sheet so changes reflect instantly on the homeowner's portal
    try {
      var masterSS = SpreadsheetApp.openById(SERVICE_MASTER_ID);
      var masterSheet = masterSS.getSheets()[0];
      var mData = masterSheet.getDataRange().getValues();
      
      for (var m = 1; m < mData.length; m++) {
        // Col B(1)=Project, Col C(2)=Phase, Col D(3)=Lot, Col E(4)=Form, Col F(5)=Item #
        if (cleanText(mData[m][1]) === projectName && cleanText(mData[m][2]) === nPhase && cleanText(mData[m][3]) === nLot && mData[m][4] === formType && mData[m][5] == itemNum) {
          masterSheet.getRange(m + 1, 7).setValue(formData.location);
          masterSheet.getRange(m + 1, 8).setValue(formData.roomArea);
          masterSheet.getRange(m + 1, 9).setValue(formData.defectArea);
          masterSheet.getRange(m + 1, 10).setValue(formData.description);
          break;
        }
      }
    } catch(masterErr) {
      // Catch in case master is locked temporarily during builder sync
    }

    return "Success! Item #" + itemNum + " has been updated.";

  } catch (e) {
    return "Error: " + e.message;
  }
}