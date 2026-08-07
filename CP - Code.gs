/*
 * =====================================================================================
 * ⚠️ IMPORTANT DEPLOYMENT CONFIGURATION TO PREVENT "REQUEST ACCESS" PROMPTS ⚠️
 * =====================================================================================
 * This Google Apps Script Web App must be deployed with the following specific settings
 * to avoid prompting external users (who do not have Google/Workspace accounts or
 * access to the underlying spreadsheets) to "request access" (Viewer, Editor, etc.):
 *
 * 1. Click "Deploy" > "New deployment"
 * 2. Select "Web app" as the deployment type.
 * 3. Set "Execute as" to: "Me" (your developer/workspace account)
 *    -> This ensures the script runs with your administrative privileges and can read/write
 *       to the Sheets and Drive files without prompting the user.
 * 4. Set "Who has access" to: "Anyone"
 *    -> This allows external/internal users (like contractors/homeowners) to successfully load
 *       and interact with the application.
 *
 * DO NOT select "User accessing the web app" or restrict access to "Anyone with Google account"
 * or "Only myself", as this will break the web application portals.
 * =====================================================================================
 */

// --- CONFIG CONSTANTS (Points to your master sheet) ---
const SHEET_ID = '1rhHn7mXKpjcGY7D1dSQ3K7aWCTFcjYrF_ls7OHmfp4E';
const SERVICE_MASTER_ID = '1L2-Mbq6Uebqih2qhfhf23A4NRG54aoNR4rQyCEBGCRU';
const CONTRACTOR_SHEET_NAME = 'Sheet1'; 
const LOTS_SHEET_NAME = 'Lots'; 
const PO_SHEET_NAME = 'Purchase Orders'; 
const WARRANTY_ROOT_FOLDER_ID = '1BUt0T6XIs3BgR7Red6EE4L2X9sdk7eHO';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Georgian Build Connect - Contractor Portal')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function cleanText(str) {
  return String(str || "").trim().toLowerCase().replace(/\s+/g, ' ');
}

function validateContractorLogin(emailAddress, passcode) {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(CONTRACTOR_SHEET_NAME);
  var data = sheet.getDataRange().getDisplayValues();
  var inputEmail = cleanText(emailAddress);
  
  for (var i = 1; i < data.length; i++) {
    var savedEmail = cleanText(data[i][2]);
    if (savedEmail !== "" && savedEmail === inputEmail) {
      var savedPasscode = String(data[i][9]).trim();
      if (savedPasscode !== "" && savedPasscode === String(passcode).trim()) {
        return { success: true, trade: data[i][0], email: data[i][2] }; 
      }
    }
  }
  return { success: false, message: "Invalid email credentials or passcode. Please try again." };
}

function updateContractorPassword(emailAddress, oldPasscode, newPasscode) {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(CONTRACTOR_SHEET_NAME);
  var range = sheet.getDataRange();
  var data = range.getDisplayValues();
  var targetEmail = cleanText(emailAddress);
  
  for (var i = 1; i < data.length; i++) {
    var savedEmail = cleanText(data[i][2]);
    if (savedEmail !== "" && savedEmail === targetEmail) {
      var savedPasscode = String(data[i][9]).trim();
      if (savedPasscode === String(oldPasscode).trim()) {
        sheet.getRange(i + 1, 10).setValue(String(newPasscode).trim());
        return { success: true, message: "Your access passcode has been updated successfully!" };
      }
    }
  }
  return { success: false, message: "Verification failed. Current passcode entry is invalid." };
}

function getContractorLots(tradeName, serviceItems) {
  var sheetApp = SpreadsheetApp.openById(SHEET_ID);
  var cleanTrade = cleanText(tradeName);
  var allowedPoKeys = []; 
  
  // 1. Gather all Lots the Trade has a PO for
  var poData = sheetApp.getSheetByName(PO_SHEET_NAME).getDataRange().getDisplayValues();
  for (var i = 1; i < poData.length; i++) {
    if (cleanText(poData[i][4]) === cleanTrade) {
      var poProj = cleanText(poData[i][1]);
      var poLotStr = cleanText(poData[i][3]).replace(/^0+/, ''); 
      if (poLotStr !== "") allowedPoKeys.push(poProj + "|" + poLotStr);
    }
  }

  // 2. Filter Master Lots Table (Check POs first, then Fuzzy-Match Service Items)
  var lotsData = sheetApp.getSheetByName(LOTS_SHEET_NAME).getDataRange().getValues();
  var filteredLots = [lotsData[0]]; 
  
  for (var l = 1; l < lotsData.length; l++) {
    var lProj = cleanText(lotsData[l][0]);
    var lPhase = cleanText(lotsData[l][1]);
    var lLotNumRaw = cleanText(lotsData[l][2]).replace(/^0+/, '');
    
    var key1 = lProj + "|" + lLotNumRaw;
    var key2 = lProj + "|" + lPhase + "." + lLotNumRaw;
    var isAllowed = false;
    
    // Check if they have a Purchase Order for this lot
    if (allowedPoKeys.indexOf(key1) > -1 || allowedPoKeys.indexOf(key2) > -1) {
      isAllowed = true;
    }
    
    // Check if they were assigned a Service Item (Uses robust Fuzzy Matching)
    if (!isAllowed && serviceItems && serviceItems.length > 0) {
        for (var s = 0; s < serviceItems.length; s++) {
            if (serviceItems[s].project === "SYSTEM ERROR") continue;
            
            var sProj = cleanText(serviceItems[s].project);
            var sLot = cleanText(serviceItems[s].lot).replace(/^0+/, '');
            
            if (lLotNumRaw === sLot) {
                // If Lot # matches exactly, check if Project Names are similar to prevent Homeowner tab misspellings blocking access
                if (lProj === sProj || lProj.indexOf(sProj) > -1 || sProj.indexOf(lProj) > -1) {
                    isAllowed = true;
                    break;
                }
            }
        }
    }
    
    // Push the allowed lot to their directory
    if (isAllowed) {
      filteredLots.push(lotsData[l]);
    }
  }
  return filteredLots;
}

function getContractorPOs(tradeName) {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(PO_SHEET_NAME);
  var data = sheet.getDataRange().getDisplayValues();
  var filteredPOs = [data[0]];
  var cleanTrade = cleanText(tradeName);
  
  for (var i = 1; i < data.length; i++) {
    if (cleanText(data[i][4]) === cleanTrade) {
      filteredPOs.push(data[i]);
    }
  }
  return filteredPOs;
}

function getContractorServiceItems(tradeName) {
  var sheetApp = SpreadsheetApp.openById(SHEET_ID);
  var cleanTrade = cleanText(tradeName);
  var items = [];
  var errors = [];
  
  var lotsData = sheetApp.getSheetByName(LOTS_SHEET_NAME).getDataRange().getValues();
  var lotCache = {}; 
  for (var l = 1; l < lotsData.length; l++) {
     var cProj = cleanText(lotsData[l][0]);
     var cLot = cleanText(lotsData[l][2]).replace(/^0+/, '');
     lotCache[cProj + "|" + cLot] = {
        project: lotsData[l][0],
        phase: lotsData[l][1],
        lot: lotsData[l][2],
        scheduleUrl: lotsData[l][13]
     };
  }

  // 1. CHOOSE MATCHING WARRANTY FILES VIA SERVICE_MASTER_ID TO AVOID O(N) DRIVE SCAN
  var activeLotsMap = {};
  try {
     var masterSS = SpreadsheetApp.openById(SERVICE_MASTER_ID);
     var masterSheet = masterSS.getSheets()[0];
     var masterData = masterSheet.getDataRange().getDisplayValues();
     for (var i = 1; i < masterData.length; i++) {
        var cellTrade = cleanText(masterData[i][11]); // Column 12: Assigned Trade
        if (cellTrade !== "" && (cellTrade === cleanTrade || cellTrade.indexOf(cleanTrade) > -1)) {
           var mProj = cleanText(masterData[i][1]); // Column 2: Project Name
           var mPhase = cleanText(masterData[i][2]).replace(/^0+/, ''); // Column 3: Phase
           var mLot = cleanText(masterData[i][3]).replace(/^0+/, ''); // Column 4: Lot Number
           activeLotsMap[mProj + "|" + mPhase + "|" + mLot] = true;
        }
     }
  } catch (e) {
     errors.push("Master Sync Fetch Error: " + e.message);
  }

  // Iterate over Lots and process sheets only for those matching activeLotsMap
  for (var l = 1; l < lotsData.length; l++) {
     var lProj = cleanText(lotsData[l][0]);
     var lPhase = cleanText(lotsData[l][1]).replace(/^0+/, '');
     var lLot = cleanText(lotsData[l][2]).replace(/^0+/, '');
     var lotKey = lProj + "|" + lPhase + "|" + lLot;
     
     if (activeLotsMap[lotKey]) {
        var folder = null;
        var folderUrl = lotsData[l][16]; // Service folder URL (Column Q)
        if (folderUrl && String(folderUrl).indexOf("http") === 0) {
           var folderIdMatch = String(folderUrl).match(/[-\w]{25,}/);
           if (folderIdMatch) {
              try {
                 folder = DriveApp.getFolderById(folderIdMatch[0]);
              } catch (e) {
                 Logger.log("Folder access error: " + e.message);
              }
           }
        }
        
        if (!folder) {
           try {
              var rootFolder = DriveApp.getFolderById(WARRANTY_ROOT_FOLDER_ID);
              var pFolders = rootFolder.searchFolders("title = '" + String(lotsData[l][0]).replace(/'/g, "\\'") + "' and trashed = false");
              if (pFolders.hasNext()) {
                  var phFolders = pFolders.next().searchFolders("title = 'Phase " + String(lotsData[l][1]).replace(/'/g, "\\'") + "' and trashed = false");
                  if (phFolders.hasNext()) {
                      var lFolders = phFolders.next().searchFolders("title = 'Lot " + String(lotsData[l][2]).replace(/'/g, "\\'") + "' and trashed = false");
                      if (lFolders.hasNext()) {
                          folder = lFolders.next();
                      }
                  }
              }
           } catch (e) {
              Logger.log("Folder search error: " + e.message);
           }
        }

        if (folder) {
           try {
              var sheetFiles = folder.searchFiles("mimeType = 'application/vnd.google-apps.spreadsheet' and trashed = false");
              while (sheetFiles.hasNext()) {
                 var file = sheetFiles.next();
                 var ss = SpreadsheetApp.openById(file.getId());
                 var sheets = ss.getSheets();

                 var sProj = lotsData[l][0];
                 var sPhase = lotsData[l][1];
                 var sLot = lotsData[l][2];
                 var hasTrade = false;
                 var tempItems = [];

                 for (var s = 0; s < sheets.length; s++) {
                   var tabName = sheets[s].getName();
                   if (tabName === "Homeowner Info") continue;
                   if (tabName.toLowerCase().indexOf('schedule') > -1) continue;

                   var data = sheets[s].getDataRange().getDisplayValues();
                   if (data.length <= 1) continue;

                   var hdrs = data[0].map(function(x){ return String(x).toLowerCase().trim(); });
                   var iTrade = hdrs.indexOf("assigned trade");
                   if (iTrade === -1) continue;

                   var iItem = hdrs.indexOf("item #");
                   var iLoc = hdrs.indexOf("location");
                   var iRoom = hdrs.indexOf("room/area");
                   var iDef = hdrs.indexOf("item/defect area");
                   var iDesc = hdrs.indexOf("description");
                   var iAdd = hdrs.indexOf("additional information") > -1 ? hdrs.indexOf("additional information") : hdrs.indexOf("trade description");
                   var iStat = hdrs.indexOf("status");
                   var iComp = hdrs.indexOf("completion notes");
                   var iSched = hdrs.indexOf("scheduled date");
                   var iPhoto = hdrs.indexOf("photo link");

                   for (var r = 1; r < data.length; r++) {
                      var cellTrade = cleanText(data[r][iTrade]);
                      if (cellTrade !== "" && (cellTrade === cleanTrade || cellTrade.indexOf(cleanTrade) > -1)) {
                          hasTrade = true;
                          var stat = iStat > -1 ? data[r][iStat] : "";
                          var compNotes = iComp > -1 ? data[r][iComp] : "";
                          var schedDate = iSched > -1 ? data[r][iSched] : "";
                          var photo = iPhoto > -1 ? data[r][iPhoto] : "";

                          if(schedDate instanceof Date) schedDate = schedDate.toLocaleDateString();
                          var isRej = (stat.toLowerCase() === "assigned" && compNotes.trim() !== "");

                          tempItems.push({
                            formType: tabName,
                            itemNum: data[r][iItem],
                            location: (iLoc > -1 ? data[r][iLoc] : "") + " / " + (iRoom > -1 ? data[r][iRoom] : ""),
                            defectArea: iDef > -1 ? data[r][iDef] : "",
                            description: iDesc > -1 ? data[r][iDesc] : "",
                            tradeDesc: iAdd > -1 ? data[r][iAdd] : "",
                            status: stat,
                            isRejected: isRej,
                            completionNotes: compNotes,
                            scheduledDate: schedDate,
                            photoLink: photo
                          });
                      }
                   }
                 }

                 if (hasTrade && tempItems.length > 0) {
                     for (var t = 0; t < tempItems.length; t++) {
                         tempItems[t].project = sProj || "Unknown";
                         tempItems[t].phase = sPhase || "";
                         tempItems[t].lot = sLot || "Unknown";
                         items.push(tempItems[t]);
                     }
                 }
              }
           } catch (e) {
              errors.push("Warranty Search Error on Lot " + lLot + ": " + e.message);
           }
        }
     }
  }

  // 2. COMPILE COMBINED ACCESS LIST & CHECK SCHEDULES
  var allowedPoLotKeys = [];
  var poData = sheetApp.getSheetByName(PO_SHEET_NAME).getDataRange().getDisplayValues();
  for (var i = 1; i < poData.length; i++) {
     if (cleanText(poData[i][4]) === cleanTrade) {
         var poProj = cleanText(poData[i][1]);
         var poLotStr = cleanText(poData[i][3]).replace(/^0+/, '');
         allowedPoLotKeys.push(poProj + "|" + poLotStr);
     }
  }

  var warrantyLotKeys = [];
  for (var w = 0; w < items.length; w++) {
      var wKey = cleanText(items[w].project) + "|" + cleanText(items[w].lot).replace(/^0+/, '');
      if (warrantyLotKeys.indexOf(wKey) === -1) warrantyLotKeys.push(wKey);
  }

  var checkedSchedules = [];
  for (var l = 1; l < lotsData.length; l++) {
     var lProj = cleanText(lotsData[l][0]);
     var lPhase = cleanText(lotsData[l][1]);
     var lLotRaw = cleanText(lotsData[l][2]).replace(/^0+/, '');
     var lKey1 = lProj + "|" + lLotRaw;
     var lKey2 = lProj + "|" + lPhase + "." + lLotRaw;
     
     // Include schedules from lots they have a PO for OR a warranty item for
     var isMatch = false;
     if (allowedPoLotKeys.indexOf(lKey1) > -1 || allowedPoLotKeys.indexOf(lKey2) > -1) isMatch = true;
     
     if (!isMatch) {
         for (var wMatch = 0; wMatch < warrantyLotKeys.length; wMatch++) {
             var wParts = warrantyLotKeys[wMatch].split("|");
             if (lLotRaw === wParts[1]) {
                 if (lProj === wParts[0] || lProj.indexOf(wParts[0]) > -1 || wParts[0].indexOf(lProj) > -1) {
                     isMatch = true; break;
                 }
             }
         }
     }

     if (isMatch) {
         var scheduleUrl = lotsData[l][13];
         if (scheduleUrl && checkedSchedules.indexOf(scheduleUrl) === -1 && String(scheduleUrl).indexOf("http") === 0) {
             checkedSchedules.push(scheduleUrl);
             try {
                 var schedIdMatch = String(scheduleUrl).match(/[-\w]{25,}/);
                 if (schedIdMatch) {
                     var schedApp = SpreadsheetApp.openById(schedIdMatch[0]);
                     var sheets = schedApp.getSheets();
                     if (sheets.length > 0 && sheets[0].getName().toLowerCase().indexOf('schedule') > -1) {
                         var schedData = sheets[0].getDataRange().getDisplayValues();
                         for (var r = 1; r < schedData.length; r++) {
                             var cellTrade = cleanText(schedData[r][2]);
                             if (cellTrade !== "" && (cellTrade === cleanTrade || cellTrade.indexOf(cleanTrade) > -1)) {
                                 var taskDesc = (schedData[r][3] && String(schedData[r][3]).trim() !== "") ? schedData[r][3] : (schedData[r][1] !== "" ? schedData[r][1] : "Construction Task");
var duration = schedData[r][4]; // Capture Column E
var startDate = schedData[r][7]; 
var finishDate = schedData[r][8]; 

items.push({
    project: lotsData[l][0], phase: lotsData[l][1], lot: lotsData[l][2],
    formType: "Construction Schedule", itemNum: "-", location: "-", defectArea: "-",
    description: taskDesc, tradeDesc: "", status: "Scheduled", isRejected: false,
    completionNotes: "", scheduledDate: startDate || finishDate, finishDate: finishDate, duration: duration, photoLink: ""
});
                             }
                         }
                     }
                 }
             } catch(e) {
                 errors.push("Schedule Sync Error (" + lotsData[l][2] + "): " + e.message);
             }
         }
     }
  }

  if (items.length === 0 && errors.length > 0) {
     items.push({
       project: "SYSTEM ERROR", phase: "N/A", lot: "N/A", formType: "Diagnostic", itemNum: "-",
       location: "Please screenshot this", defectArea: "Backend Error", description: "Backend Error", tradeDesc: errors.join(" | "), status: "", isRejected: false, completionNotes: "", scheduledDate: "", photoLink: ""
     });
  }
  return items;
}

function submitItemCompletion(tradeName, project, phase, lotNumRaw, formType, itemNum, notes, base64Photo, fileName, mimeType) {
  try {
    var sheetApp = SpreadsheetApp.openById(SHEET_ID);
    var lotsData = sheetApp.getSheetByName(LOTS_SHEET_NAME).getDataRange().getValues();
    
    for (var l = 1; l < lotsData.length; l++) {
      if (cleanText(lotsData[l][0]) == cleanText(project) && cleanText(lotsData[l][1]) == cleanText(phase) && cleanText(lotsData[l][2]) == cleanText(lotNumRaw)) {
        
        var folder = null;
        var folderUrl = lotsData[l][16];
        if (folderUrl && String(folderUrl).indexOf("http") === 0) {
           var folderIdMatch = String(folderUrl).match(/[-\w]{25,}/);
           if (folderIdMatch) folder = DriveApp.getFolderById(folderIdMatch[0]);
        }
        
        if (!folder) {
           var rootFolder = DriveApp.getFolderById(WARRANTY_ROOT_FOLDER_ID);
           var pFolders = rootFolder.searchFolders("title = '" + String(project).replace(/'/g, "\\'") + "' and trashed = false");
           if (pFolders.hasNext()) {
               var phFolders = pFolders.next().searchFolders("title = 'Phase " + String(phase).replace(/'/g, "\\'") + "' and trashed = false");
               if (phFolders.hasNext()) {
                   var lFolders = phFolders.next().searchFolders("title = 'Lot " + String(lotNumRaw).replace(/'/g, "\\'") + "' and trashed = false");
                   if (lFolders.hasNext()) {
                       folder = lFolders.next();
                   }
               }
           }
        }
        
        if (folder) {
             var photoUrl = "";
             if (base64Photo) {
               var safeFormType = formType.replace(/'/g, "\\'");
               var subFolders = folder.searchFolders("title = '" + safeFormType + "' and trashed = false");
               var targetFolder = subFolders.hasNext() ? subFolders.next() : folder.createFolder(formType);
               
               var ext = "";
               if (fileName.lastIndexOf('.') > -1) ext = fileName.substring(fileName.lastIndexOf('.'));
               
               var combinedLot = phase + "-" + lotNumRaw;
               var newFileName = combinedLot + " - " + formType + " - Item " + itemNum + " - CONTRACTOR PHOTO" + ext;
               
               var blob = Utilities.newBlob(Utilities.base64Decode(base64Photo), mimeType, newFileName);
               var savedFile = targetFolder.createFile(blob);
               photoUrl = savedFile.getUrl();
             }

             var files = folder.searchFiles("mimeType = 'application/vnd.google-apps.spreadsheet'");
             if (files.hasNext()) {
               var ss = SpreadsheetApp.openById(files.next().getId());
               var formTab = ss.getSheetByName(formType);
               if (formTab) {
                 var data = formTab.getDataRange().getValues();
                 var hdrs = data[0].map(function(h){ return String(h).toLowerCase().trim(); });
                 
                 var iItem = hdrs.indexOf("item #");
                 var iTrade = hdrs.indexOf("assigned trade");
                 var iComp = hdrs.indexOf("completion notes");
                 var iPhoto = hdrs.indexOf("photo link");
                 var iStat = hdrs.indexOf("status");
                 
                 if (iComp === -1) {
                    iComp = formTab.getLastColumn();
                    formTab.getRange(1, iComp + 1).setValue("Completion Notes").setFontWeight("bold");
                 }
                 if (iPhoto === -1) {
                    iPhoto = formTab.getLastColumn(); 
                    formTab.getRange(1, iPhoto + 1).setValue("Photo Link").setFontWeight("bold");
                 }

                 for (var r = 1; r < data.length; r++) {
                   if (iItem > -1 && iTrade > -1 && String(data[r][iItem]) === String(itemNum) && cleanText(data[r][iTrade]).indexOf(cleanText(tradeName)) > -1) {
                     
                     var existingNotes = data[r][iComp] ? data[r][iComp] + "\n\n" : "";
                     formTab.getRange(r + 1, iComp + 1).setValue(existingNotes + "Completed by " + tradeName + ": " + notes);
                     
                     if (photoUrl !== "") {
                       var currentPhotos = data[r][iPhoto] ? data[r][iPhoto] + "\n\n" : "";
                       formTab.getRange(r + 1, iPhoto + 1).setValue(currentPhotos + "Contractor Photo: " + photoUrl);
                     }
                     
                     var sentCount = 0;
                     var emailDebugMsg = "";
                     try {
                       var usersSheet = sheetApp.getSheetByName('Users');
                       if (usersSheet) {
                         var uData = usersSheet.getDataRange().getValues();
                         var emails = [];
                         var cleanProject = cleanText(project); 
                         
                         for (var u = 1; u < uData.length; u++) {
                           var uTitle = cleanText(uData[u][2]); 
                           var uProjs = cleanText(uData[u][3]); 
                           var uEmail = String(uData[u][4]).trim(); 
                           
                           if (uTitle.indexOf('service coordinator') > -1) {
                             if (uProjs.indexOf(cleanProject) > -1 || uProjs.indexOf('all') > -1) {
                               if (uEmail !== "" && emails.indexOf(uEmail) === -1) {
                                 emails.push(uEmail);
                               }
                             }
                           }
                         }
                         
                         if (emails.length > 0) {
                           var subject = "Trade Completion: " + project + " - Lot " + lotNumRaw + " (Item #" + itemNum + ")";
                           var body = "Hello,\n\nThe following service item has been marked as complete by " + tradeName + ".\n\n" +
                                      "Project: " + project + "\nPhase: " + phase + "\nLot: " + lotNumRaw + "\n" +
                                      "Form: " + formType + "\nItem #: " + itemNum + "\n\n" +
                                      "Completion Notes:\n" + notes + "\n\n" +
                                      (photoUrl !== "" ? "Photo Link: " + photoUrl + "\n\n" : "") +
                                      "Please log into the Service Portal to verify and update the master status.\n\n" +
                                      "Thank you,\nGeorgian Build Connect System";
                           
                           GmailApp.sendEmail(emails.join(","), subject, body, {name: "Georgian Build Connect"});
                           sentCount = emails.length;
                         } else {
                           emailDebugMsg = " (Notice: No matching Service Coordinator found for this project)";
                         }
                       } else {
                           emailDebugMsg = " (Notice: Users tab not found)";
                       }
                     } catch(err) {
                       emailDebugMsg = " (Email Error: " + err.message + ")";
                       sentCount = 0; 
                     }

                     if (iStat > -1) {
                       formTab.getRange(r + 1, iStat + 1).setValue("Completed");
                       var finalMsg = "Item successfully marked as completed!";
                       if (sentCount > 0) {
                         finalMsg += " (Notified " + sentCount + " Coordinator" + (sentCount > 1 ? "s" : "") + ")";
                       } else {
                         finalMsg += emailDebugMsg;
                       }
                       return { success: true, message: finalMsg };
                     } else {
                       return { success: true, message: "Notes & photo saved. Status column not found." };
                     }
                   }
                 }
               }
             }
        }
      }
    }
    return { success: false, message: "Could not locate service document to process completion." };
  } catch (e) {
    return { success: false, message: "Error: " + e.message };
  }
}

function getOrCreateFolder(parentFolder, folderName) {
  var safeName = folderName.replace(/'/g, "\\'");
  var folders = parentFolder.searchFolders("title = '" + safeName + "' and trashed = false");
  if (folders.hasNext()) { return folders.next(); }
  return parentFolder.createFolder(folderName);
}

function generateContractorReportPDF(tradeName, filters, selectedCols) {
  try {
    var items = getContractorServiceItems(tradeName);
    if (items.length === 0 || (items.length === 1 && items[0].project === "SYSTEM ERROR")) {
       return "Error: No service items found.";
    }

    var filteredItems = [];
    for (var i = 0; i < items.length; i++) {
      var it = items[i];
      if (it.formType === "Construction Schedule") continue;

      var isMatch = true;
      var itemStatusDisplay = it.isRejected ? "Item Not Approved" : (it.status || "Unreviewed");

      if (filters.project && cleanText(it.project) !== cleanText(filters.project)) isMatch = false;
      if (filters.phase && cleanText(it.phase) !== cleanText(filters.phase)) isMatch = false;
      if (filters.lot && cleanText(it.lot) !== cleanText(filters.lot)) isMatch = false;
      if (filters.status && cleanText(itemStatusDisplay) !== cleanText(filters.status)) isMatch = false;
      if (isMatch) filteredItems.push(it);
    }

    if (filteredItems.length === 0) return "Error: No data found matching these filters.";

    var html = "<html><head><style>";
    html += "@page { size: landscape; margin: 0.5in; } ";
    html += "body { font-family: sans-serif; font-size: 9px; color: #333; } ";
    html += "table { width: 100%; border-collapse: collapse; margin-top: 15px; table-layout: auto; word-wrap: break-word; } ";
    html += "th, td { border: 1px solid #aaa; padding: 5px; text-align: left; vertical-align: top; } ";
    html += "th { background-color: #1c2d42; color: white; font-weight: bold; } ";
    html += "h2 { color: #1c2d42; margin-bottom: 5px; } ";
    html += "p { color: #666; margin-top: 0; } ";
    html += "</style></head><body>";

    html += "<h2>Service Orders Report - " + tradeName + "</h2>";
    html += "<p>Generated: " + new Date().toLocaleDateString() + " " + new Date().toLocaleTimeString() + "</p>";

    var filterStr = [];
    if(filters.project) filterStr.push("Project: " + filters.project);
    if(filters.phase) filterStr.push("Phase: " + filters.phase);
    if(filters.lot) filterStr.push("Lot: " + filters.lot);
    if(filters.status) filterStr.push("Status: " + filters.status);
    if(filterStr.length > 0) html += "<p><b>Filters:</b> " + filterStr.join(" | ") + "</p>";

    html += "<table><thead><tr>";
    for (var c = 0; c < selectedCols.length; c++) {
        html += "<th>" + selectedCols[c] + "</th>";
    }
    html += "</tr></thead><tbody>";

    var colMap = {
      "project": "project",
      "phase": "phase",
      "lot #": "lot",
      "form type": "formType",
      "item #": "itemNum",
      "location": "location",
      "defect area": "defectArea",
      "description": "description",
      "additional information": "tradeDesc",
      "scheduled date": "scheduledDate",
      "photo link": "photoLink",
      "status": "status"
    };

    for (var r = 0; r < filteredItems.length; r++) {
        html += "<tr>";
        for (var c = 0; c < selectedCols.length; c++) {
            var key = colMap[selectedCols[c].toLowerCase()];
            var val = filteredItems[r][key];
            
            if (key === "status") {
                val = filteredItems[r].isRejected ? "Item Not Approved" : (val || "Unreviewed");
            }
            
            if (val === undefined || val === null) val = "";
            val = String(val).replace(/\n/g, "<br>");

            if (selectedCols[c].toLowerCase() === "photo link" && val.indexOf("http") > -1) {
               val = "[See Portal for Photos]";
            }
            html += "<td>" + val + "</td>";
        }
        html += "</tr>";
    }
    html += "</tbody></table></body></html>";

    var blob = Utilities.newBlob(html, MimeType.HTML).setName("ServiceReport.html");
    var pdfBlob = blob.getAs(MimeType.PDF);
    pdfBlob.setName("Contractor_Report_" + new Date().getTime() + ".pdf");

    var rootFolder = DriveApp.getFolderById(WARRANTY_ROOT_FOLDER_ID); 
    var reportFolder = getOrCreateFolder(rootFolder, "Generated Reports");
    var newFile = reportFolder.createFile(pdfBlob);
    
    return newFile.getUrl();
  } catch (e) {
    return "Error generating PDF: " + e.message;
  }
}

function forceAuth() {
  DriveApp.getRootFolder();
  GmailApp.getAliases();
}