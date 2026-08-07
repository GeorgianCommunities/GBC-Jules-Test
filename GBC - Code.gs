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
 *    -> This allows external/internal users to successfully load and interact.
 *
 * DO NOT select "User accessing the web app" or restrict access to "Anyone with Google account"
 * or "Only myself", as this will break the web application portals.
 * =====================================================================================
 */

// --- IDS & CONSTANTS ---
const SHEET_ID = '1rhHn7mXKpjcGY7D1dSQ3K7aWCTFcjYrF_ls7OHmfp4E';
const SHEET_NAME = 'Sheet1'; 
const LOTS_SHEET_NAME = 'Lots'; 
const PO_SHEET_NAME = 'Purchase Orders'; 
const DRAW_SHEET_NAME = 'Progress Draws'; 

const FOLDER_ID = '1CVrNejsHJwpHhp_IB9PtLkLsW6utMnWE'; 
const COST_CODES_SHEET_ID = '19T_g6GXsefIBTBiCDPWUTSOnCvLFg_KXDuCv0Jphzr0'; // No longer actively used, kept for legacy reference
const MODEL_SHEET_ID = '1KSfwLKiIUc1nXt4IXdqpSFnHMwOn2JjiK8Z3jOzbRm0'; 

// FOLDER CONSTANTS
const LOT_FOLDERS_MASTER_ID = '1zxGK5AqCyYOnqwoQeem7Br0amEALbISX';
const RESTRICTED_LOT_FOLDERS_MASTER_ID = '1jM7UggcrWIE71bLn-NCoHyKvPIlcINNX';
const HOMEOWNER_LOT_FOLDERS_MASTER_ID = '1g_frWrcUFmexU5ci0YvRqrIKCWSOfmRK';
const PO_TEMPLATE_DOC_ID = '1X7qxE4etQOodCjapsh5HlBQ9nlUkY9vOa_hxu82WwMM'; 
const PO_FOLDERS_MASTER_ID = '14Og9DUsIbvdSutd09WZRqSq0maCYATbC';

// SCHEDULE CONSTANTS
const SCHEDULE_TEMPLATE_ID = '1ITV4X3vWUByb8AgKJfGSU4FviQGyMpw2ceOnDMF1viE';
const SCHEDULE_STAGING_FOLDER_ID = '1Vsw2tMqLp7NTwoW3I3EtFofzIQV20WqJ';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Georgian Build Connect')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function cleanText(str) {
  return String(str || "").trim().toLowerCase().replace(/\s+/g, ' ');
}

// --- HELPER: Improved Fuzzy Match Logic ---
function fuzzyMatch(str1, str2) {
  var s1 = String(str1).toLowerCase().replace(/[^a-z0-9]/g, '');
  var s2 = String(str2).toLowerCase().replace(/[^a-z0-9]/g, '');
  if (!s1 || !s2) return false;
  
  if (s1 === s2 || s1.indexOf(s2) > -1 || s2.indexOf(s1) > -1) return true;
  
  if (s1.length > 3 && s2.length > 3 && s1.substring(0, 4) === s2.substring(0, 4)) return true;

  var m = s1.length, n = s2.length, dp = [];
  for (var i = 0; i <= m; i++) dp[i] = [i];
  for (var j = 0; j <= n; j++) dp[0][j] = j;
  for (var i = 1; i <= m; i++) {
    for (var j = 1; j <= n; j++) {
      var cost = (s1.charAt(i - 1) === s2.charAt(j - 1)) ? 0 : 1;
      dp[i][j] = Math.min(dp[i - 1][j] + 1, dp[i][j - 1] + 1, dp[i - 1][j - 1] + cost);
    }
  }
  return ((Math.max(m, n) - dp[m][n]) / Math.max(m, n)) >= 0.50; 
}

// --- HELPER: Generate Random Password ---
function generateRandomPassword(length) {
  var chars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@#$%^&*";
  var pass = "";
  for (var i = 0; i < length; i++) { pass += chars.charAt(Math.floor(Math.random() * chars.length)); }
  return pass;
}

// --- HELPER: Fetch Project Settings ---
function getProjectSettings() {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Settings');
  var data = sheet.getDataRange().getDisplayValues();
  var settings = {};
  for(var i=1; i<data.length; i++) {
    var proj = cleanText(data[i][0]);
    if(proj !== "") {
      settings[proj] = {
        name: data[i][0], 
        billName: data[i][1], 
        billAddress: data[i][2], 
        billPhone: data[i][3],
        billEmail: data[i][4], 
        poPrefix: data[i][5], 
        nextPoNum: parseInt(data[i][6]) || 1,
        taxRate: parseFloat(data[i][7]) || 0, 
        pricingSheetId: data[i][8], 
        trackerSheetId: data[i][11], 
        scheduleTemplateId: data[i][12], 
        rowIndex: i + 1
      };
    }
  }
  return settings;
}

// --- HELPER: Fetch Trade Types from Settings Column K ---
function getTradeTypesData() {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Settings');
  if(!sheet) return [];
  var data = sheet.getRange("K2:K").getDisplayValues();
  var types = [];
  for (var i = 0; i < data.length; i++) {
    var val = data[i][0].trim();
    if (val !== "" && types.indexOf(val) === -1) {
      types.push(val);
    }
  }
  return types.sort();
}

// --- HELPER: Create 3-Tier Nested Lot Folders ---
function getOrCreateNestedLotFolder(masterFolderId, projectName, phase, lotNum) {
  var masterFolder = DriveApp.getFolderById(masterFolderId);
  
  // 1. Project Folder
  var safeProj = projectName.replace(/'/g, "\\'");
  var projSearch = masterFolder.searchFolders("title = '" + safeProj + "' and trashed = false");
  var projFolder = projSearch.hasNext() ? projSearch.next() : masterFolder.createFolder(projectName);
  
  // 2. Phase Folder
  var phaseStr = "Phase " + phase;
  var safePhase = phaseStr.replace(/'/g, "\\'");
  var phaseSearch = projFolder.searchFolders("title = '" + safePhase + "' and trashed = false");
  var phaseFolder = phaseSearch.hasNext() ? phaseSearch.next() : projFolder.createFolder(phaseStr);
  
  // 3. Lot Folder
  var lotStr = "Lot " + lotNum;
  var safeLot = lotStr.replace(/'/g, "\\'");
  var lotSearch = phaseFolder.searchFolders("title = '" + safeLot + "' and trashed = false");
  var lotFolder = lotSearch.hasNext() ? lotSearch.next() : phaseFolder.createFolder(lotStr);
  
  return lotFolder;
}

// --- HELPER: Apply Auto-Formatting to Tracker Sheets ---
function applyTrackerFormatting(sheet) {
  sheet.clearConditionalFormatRules();
  
  var range = sheet.getRange("E2:ZZ1000"); 
  var rules = [];
  
  var ruleRed = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($B2<>"", E2="")')
    .setBackground('#fce8e6') 
    .setFontColor('#cc0000') 
    .setRanges([range])
    .build();
    
  var ruleGreen = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied('=AND($B2<>"", E2<>"")')
    .setBackground('#e6f4ea') 
    .setFontColor('#137333') 
    .setRanges([range])
    .build();
    
  rules.push(ruleRed);
  rules.push(ruleGreen);
  
  sheet.setConditionalFormatRules(rules);
}

// --- TAB 1: CONTRACTOR FUNCTIONS ---
function getContractorData() { 
  return SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME).getDataRange().getDisplayValues(); 
}

function addContractor(formData) {
  var newPassword = generateRandomPassword(10);
  
  SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME).appendRow([
    formData.businessName, formData.contact, formData.email, formData.phone, formData.address, formData.costCodes, formData.activityCode, formData.projects, formData.tradeType, newPassword
  ]);
  
  if (formData.email && formData.email.indexOf("@") > -1) {
    var websiteLink = "Website link not configured yet.";
    try {
      var portalSheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Website Portal Link');
      if (portalSheet) {
        var fetchedLink = portalSheet.getRange("A2").getValue(); 
        if (fetchedLink && String(fetchedLink).trim() !== "") { websiteLink = fetchedLink; }
      }
    } catch(e) { Logger.log("Could not fetch website link: " + e.message); }
    
    var subject = "Welcome to Georgian Build Connect!";
    var body = "Hello " + formData.contact + ",\n\n" +
               "Your company (" + formData.businessName + ") has been added to our Build Connect directory.\n\n" +
               "Here are your login credentials for the portal:\n\n" +
               "Website: " + websiteLink + "\n" +
               "Username: " + formData.email + "\n" +
               "Password: " + newPassword + "\n\n" +
               "Please keep this information secure.\n\n" +
               "Thank you,\nGeorgian Communities";
               
    try {
      GmailApp.sendEmail(formData.email, subject, body, {name: "Georgian Build Connect"});
    } catch(e) { Logger.log("Failed to send welcome email: " + e.message); }
  }

  var projSettings = getProjectSettings();
  var selectedProjects = formData.projects.split(', ');
  var modelApp = SpreadsheetApp.openById(MODEL_SHEET_ID);
  
  for(var p=0; p<selectedProjects.length; p++) {
    var projName = selectedProjects[p].trim();
    var cleanProj = cleanText(projName);
    var pricingId = projSettings[cleanProj] ? projSettings[cleanProj].pricingSheetId : null;
    
    if(pricingId && pricingId.length > 20) {
      try {
        var pricingSpreadsheet = SpreadsheetApp.openById(pricingId);
        var tradeName = String(formData.businessName).trim().substring(0, 100); 
        
        if (!pricingSpreadsheet.getSheetByName(tradeName)) {
          var newSheet = pricingSpreadsheet.insertSheet(tradeName);
          var modelSheet = modelApp.getSheetByName(projName);
          if(modelSheet) {
            var mData = modelSheet.getDataRange().getDisplayValues();
            if(mData.length > 1) {
              var uniqueModels = [];
              var allOptions = ["Base Price"];
              for(var r=1; r<mData.length; r++) {
                 var mod = mData[r][0];
                 if(mod && uniqueModels.indexOf(mod) === -1) uniqueModels.push(mod);
                 for(var c=1; c<=7; c++) { 
                    var opt = mData[r][c];
                    if(opt && allOptions.indexOf(opt) === -1) allOptions.push(opt);
                 }
              }
              var row1 = ["Item / Model"].concat(uniqueModels);
              newSheet.getRange(1, 1, 1, row1.length).setValues([row1]);
              newSheet.getRange(1, 1, 1, row1.length).setFontWeight("bold");
              var colA = [];
              for(var a=0; a<allOptions.length; a++) { colA.push([allOptions[a]]); }
              newSheet.getRange(2, 1, colA.length, 1).setValues(colA);
            }
          }
        }
      } catch (e) { Logger.log("Pricing Tab Creation Error for " + projName + ": " + e.message); }
    }
  }
  return "Success";
}

function deleteContractor(businessName) {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  var data = sheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) { 
    if (data[i][0] === businessName) { sheet.deleteRow(i + 1); return "Deleted"; }
  }
}

function editContractor(formData) {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  var data = sheet.getDataRange().getValues();
  var rowIndex = -1;
  
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === formData.originalBusinessName) {
      rowIndex = i + 1;
      break;
    }
  }
  
  if (rowIndex === -1) {
    return "Error: Contractor not found.";
  }
  
  sheet.getRange(rowIndex, 1, 1, 9).setValues([[
    formData.businessName, 
    formData.contact, 
    formData.email, 
    formData.phone, 
    formData.address, 
    formData.costCodes, 
    formData.activityCode, 
    formData.projects, 
    formData.tradeType
  ]]);

  var projSettings = getProjectSettings();
  var selectedProjects = formData.projects.split(', ');
  var modelApp = SpreadsheetApp.openById(MODEL_SHEET_ID);
  
  for(var p=0; p<selectedProjects.length; p++) {
    var projName = selectedProjects[p].trim();
    if (projName === "") continue;
    
    var cleanProj = cleanText(projName);
    var pricingId = projSettings[cleanProj] ? projSettings[cleanProj].pricingSheetId : null;
    
    if(pricingId && pricingId.length > 20) {
      try {
        var pricingSpreadsheet = SpreadsheetApp.openById(pricingId);
        var newTradeName = String(formData.businessName).trim().substring(0, 100); 
        var oldTradeName = String(formData.originalBusinessName).trim().substring(0, 100);
        
        if (newTradeName !== oldTradeName) {
          var oldSheet = pricingSpreadsheet.getSheetByName(oldTradeName);
          if (oldSheet && !pricingSpreadsheet.getSheetByName(newTradeName)) {
            oldSheet.setName(newTradeName);
          }
        }
        
        if (!pricingSpreadsheet.getSheetByName(newTradeName)) {
          var newSheet = pricingSpreadsheet.insertSheet(newTradeName);
          var modelSheet = modelApp.getSheetByName(projName);
          if(modelSheet) {
            var mData = modelSheet.getDataRange().getDisplayValues();
            if(mData.length > 1) {
              var uniqueModels = [];
              var allOptions = ["Base Price"];
              for(var r=1; r<mData.length; r++) {
                 var mod = mData[r][0];
                 if(mod && uniqueModels.indexOf(mod) === -1) uniqueModels.push(mod);
                 for(var c=1; c<=7; c++) { 
                    var opt = mData[r][c];
                    if(opt && allOptions.indexOf(opt) === -1) allOptions.push(opt);
                 }
              }
              var row1 = ["Item / Model"].concat(uniqueModels);
              newSheet.getRange(1, 1, 1, row1.length).setValues([row1]);
              newSheet.getRange(1, 1, 1, row1.length).setFontWeight("bold");
              var colA = [];
              for(var a=0; a<allOptions.length; a++) { colA.push([allOptions[a]]); }
              newSheet.getRange(2, 1, colA.length, 1).setValues(colA);
            }
          }
        }
      } catch (e) { Logger.log("Pricing Tab Edit Error for " + projName + ": " + e.message); }
    }
  }
  return "Success";
}

function getCostCodesData() {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Settings');
  if(!sheet) return [];
  var data = sheet.getRange("J2:J").getDisplayValues();
  var costCodes = [];
  for (var i = 0; i < data.length; i++) { 
    var val = data[i][0].trim();
    if (val !== "" && costCodes.indexOf(val) === -1) {
      costCodes.push(val); 
    }
  }
  return costCodes;
}

// --- TAB 2: LOT DIRECTORY FUNCTIONS ---
function getLotDropdownData() {
  var settingsSheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Settings');
  var projects = [];
  if (settingsSheet) {
    var pData = settingsSheet.getRange("A2:A").getDisplayValues();
    for (var p = 0; p < pData.length; p++) {
      if (pData[p][0] !== "") projects.push(pData[p][0]);
    }
  }

  var modelSheets = SpreadsheetApp.openById(MODEL_SHEET_ID).getSheets();
  var projectDataMap = {};

  for (var i = 0; i < modelSheets.length; i++) {
    var sheet = modelSheets[i];
    var tabName = sheet.getName();
    var data = sheet.getDataRange().getDisplayValues();
    if(data.length === 0) continue;

    var headers = data[0];
    var optHeaders = [ headers[2] || "", headers[3] || "", headers[4] || "", headers[5] || "", headers[6] || "", headers[7] || "" ];
    var models = [], lotStyles = [], opt1 = [], opt2 = [], opt3 = [], opt4 = [], opt5 = [], opt6 = [];
    
    for (var r = 1; r < data.length; r++) {
      if (data[r][0] && data[r][0] !== "") models.push(data[r][0]);
      if (data[r][1] && data[r][1] !== "") lotStyles.push(data[r][1]);
      if (data[r][2] && data[r][2] !== "") opt1.push(data[r][2]);
      if (data[r][3] && data[r][3] !== "") opt2.push(data[r][3]);
      if (data[r][4] && data[r][4] !== "") opt3.push(data[r][4]);
      if (data[r][5] && data[r][5] !== "") opt4.push(data[r][5]);
      if (data[r][6] && data[r][6] !== "") opt5.push(data[r][6]);
      if (data[r][7] && data[r][7] !== "") opt6.push(data[r][7]);
    }
    projectDataMap[tabName] = { headers: optHeaders, models: models, lotStyles: lotStyles, opt1: opt1, opt2: opt2, opt3: opt3, opt4: opt4, opt5: opt5, opt6: opt6 };
  }

  var lotHierarchy = {};
  var masterSheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Master Lot Mapping');
  var activeLotsSheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(LOTS_SHEET_NAME);
  
  var activeLots = {}; 
  if (activeLotsSheet) {
    var activeData = activeLotsSheet.getDataRange().getDisplayValues();
    for (var a = 1; a < activeData.length; a++) {
      var key = activeData[a][0] + "|" + activeData[a][1] + "|" + activeData[a][2];
      activeLots[key] = true;
    }
  }

  if (masterSheet) {
    var masterData = masterSheet.getDataRange().getDisplayValues();
    for (var m = 1; m < masterData.length; m++) {
      var mProj = masterData[m][0];
      var mPhase = masterData[m][1];
      var mLot = masterData[m][2];
      
      if (!mProj || !mPhase || !mLot) continue;
      
      if (!lotHierarchy[mProj]) lotHierarchy[mProj] = {};
      if (!lotHierarchy[mProj][mPhase]) lotHierarchy[mProj][mPhase] = [];
      
      var lotArray = mLot.split(",");
      for (var l = 0; l < lotArray.length; l++) {
         var cleanLot = lotArray[l].trim();
         var checkKey = mProj + "|" + mPhase + "|" + cleanLot;
         
         if (cleanLot !== "" && lotHierarchy[mProj][mPhase].indexOf(cleanLot) === -1 && !activeLots[checkKey]) {
           lotHierarchy[mProj][mPhase].push(cleanLot);
         }
      }
    }
  }

  return { projects: projects, projectDataMap: projectDataMap, lotHierarchy: lotHierarchy };
}

function getLotsData(filterObj) {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(LOTS_SHEET_NAME);
  if (!sheet) return [["Error: Please create a tab named 'Lots'."]]; 
  
  var data = sheet.getDataRange().getDisplayValues();
  if (data.length <= 1) return data;
  
  var headers = data[0]; 
  if(headers.length < 15) {
     sheet.getRange(1, 15).setValue("Schedule Status");
     headers[14] = "Schedule Status";
  }
  
  var filteredData = [headers]; 
  
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    
    var folderUrl = row[12] || "";
    var schedUrl = row[13] || "";
    var schedStatus = row[14] || "";
    
    if (schedUrl.indexOf("http") === 0 && schedStatus !== "Executed") {
      try {
        var schedIdMatch = schedUrl.match(/[-\w]{25,}/);
        var folderIdMatch = folderUrl.match(/[-\w]{25,}/);
        
        if (schedIdMatch && folderIdMatch) {
          var file = DriveApp.getFileById(schedIdMatch[0]);
          var parents = file.getParents();
          var isMoved = false;
          
          while(parents.hasNext()) {
            if (parents.next().getId() === folderIdMatch[0]) {
              isMoved = true; break;
            }
          }
          
          if (isMoved) {
            schedStatus = "Executed";
            row[14] = "Executed";
            sheet.getRange(i + 1, 15).setValue("Executed"); 
          } else {
            schedStatus = "Staged";
            row[14] = "Staged";
            if (!sheet.getRange(i + 1, 15).getValue()) {
              sheet.getRange(i + 1, 15).setValue("Staged"); 
            }
          }
        }
      } catch(e) { Logger.log("Schedule Check Error: " + e.message); }
    }
    
    var isMatch = true;
    for (var key in filterObj) {
      if (filterObj[key] && filterObj[key] !== "") {
        var colIndex = headers.indexOf(key);
        if (colIndex > -1) {
          if (String(row[colIndex]).toLowerCase().indexOf(String(filterObj[key]).toLowerCase().trim()) === -1) {
            isMatch = false; break; 
          }
        }
      }
    }
    if (isMatch) filteredData.push(row);
  }
  return filteredData.length > 1 ? filteredData : [headers, ["Status", "No lots match your filters."]];
}

function addLot(formData) {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(LOTS_SHEET_NAME);
  
  var existingData = sheet.getDataRange().getValues();
  for (var i = 1; i < existingData.length; i++) {
    if (existingData[i][0] == formData.projectName && 
        existingData[i][1] == formData.phase && 
        existingData[i][2] == formData.lotNum) {
      return "Error: Lot " + formData.phase + "-" + formData.lotNum + " already exists for " + formData.projectName + ".";
    }
  }

  var folderUrl = "";
  var scheduleUrl = "";
  var hoFolderUrl = "";
  var currentDateStr = new Date().toLocaleDateString();
  
  try {
    // --- 1. Generate Standard Lot Folder ---
    var newLotFolder = getOrCreateNestedLotFolder(LOT_FOLDERS_MASTER_ID, formData.projectName, formData.phase, formData.lotNum);
    folderUrl = newLotFolder.getUrl();
    
    // --- 2. Generate Restricted Folder ---
    var newRestrictedLotFolder = getOrCreateNestedLotFolder(RESTRICTED_LOT_FOLDERS_MASTER_ID, formData.projectName, formData.phase, formData.lotNum);

    // --- 3. Generate Homeowner Portal Folder ---
    var newHoLotFolder = getOrCreateNestedLotFolder(HOMEOWNER_LOT_FOLDERS_MASTER_ID, formData.projectName, formData.phase, formData.lotNum);
    hoFolderUrl = newHoLotFolder.getUrl();

    // --- Generate Schedule ---
    var stagingFolder = DriveApp.getFolderById(SCHEDULE_STAGING_FOLDER_ID);
    var safeProjectName = formData.projectName.replace(/'/g, "\\'"); 
    var stagingProjFolders = stagingFolder.searchFolders("title = '" + safeProjectName + "' and trashed = false");
    var stagingProjFolder = stagingProjFolders.hasNext() ? stagingProjFolders.next() : stagingFolder.createFolder(formData.projectName);
    
    var projSettings = getProjectSettings();
    var cleanProjName = cleanText(formData.projectName);
    var targetConfig = projSettings[cleanProjName];
    
    var templateIdToUse = (targetConfig && targetConfig.scheduleTemplateId && targetConfig.scheduleTemplateId.length > 20) 
                          ? targetConfig.scheduleTemplateId 
                          : SCHEDULE_TEMPLATE_ID;
                          
    var templateFile = DriveApp.getFileById(templateIdToUse);
    var scheduleName = formData.projectName + " - Lot " + formData.phase + "-" + formData.lotNum + " - Schedule";
    var newSchedule = templateFile.makeCopy(scheduleName, stagingProjFolder);
    scheduleUrl = newSchedule.getUrl();
    
    try {
      var schedApp = SpreadsheetApp.openById(newSchedule.getId());
      var tempSheet = schedApp.getSheets()[0];
      
      var displayLotString = formData.projectName + " - " + formData.phase + " - " + formData.lotNum;
      tempSheet.getRange("F1").setValue(displayLotString);

      tempSheet.getRange("F2").setValue(currentDateStr);
      tempSheet.getRange("F3").setValue(currentDateStr);

      var lastRow = tempSheet.getLastRow();
      if (lastRow > 0) {
        var rawOptions = [formData.lotStyle, formData.opt1, formData.opt2, formData.opt3, formData.opt4, formData.opt5, formData.opt6];
        var activeOptions = [];
        for (var o = 0; o < rawOptions.length; o++) {
          if (rawOptions[o] && String(rawOptions[o]).trim() !== "") {
            activeOptions.push(cleanText(rawOptions[o])); 
          }
        }

        var universalTags = ["standard", "walkout", "lookout"];
        var optionData = tempSheet.getRange(1, 11, lastRow, 1).getValues();
        
        for (var r = lastRow - 1; r >= 0; r--) {
          var rowOptionType = cleanText(optionData[r][0]);
          if (rowOptionType !== "") {
             var isUniversal = universalTags.indexOf(rowOptionType) > -1;
             var isSelected = activeOptions.indexOf(rowOptionType) > -1;
             if (!isUniversal && !isSelected) {
                tempSheet.deleteRow(r + 1); 
             }
          }
        }
      }
      
      // Write the ID after rows are deleted so it doesn't shift
      tempSheet.getRange("Z100").setValue(newLotFolder.getId()).setFontColor("white");

    } catch(err) {
      Logger.log("Failed writing data to schedule: " + err.message);
    }
  } catch(e) { Logger.log("Failed to create folder or schedule: " + e.message); }
  
  sheet.appendRow([
    formData.projectName, formData.phase, formData.lotNum, formData.civicAddress, 
    formData.model, formData.lotStyle, 
    formData.opt1, formData.opt2, formData.opt3, formData.opt4, formData.opt5, formData.opt6, 
    folderUrl, scheduleUrl, "Staged", "", "", hoFolderUrl
  ]);

  try {
    var projSettings = getProjectSettings();
    var cleanProj = cleanText(formData.projectName);
    var trackerId = projSettings[cleanProj] ? projSettings[cleanProj].trackerSheetId : null;
    
    if (trackerId && trackerId.length > 20) {
      var trackerApp = SpreadsheetApp.openById(trackerId);
      var phaseTabName = "Phase " + formData.phase;
      var phaseSheet = trackerApp.getSheetByName(phaseTabName);
      
      if (phaseSheet) {
        var tData = phaseSheet.getDataRange().getDisplayValues();
        for (var r = 1; r < tData.length; r++) {
          if (String(tData[r][0]).trim() === String(formData.lotNum).trim()) {
            phaseSheet.getRange(r + 1, 2).setValue(currentDateStr); 
            break;
          }
        }
      }
    }
  } catch(trackerErr) {
    Logger.log("Failed to auto-write Lot Sold date onto Tracker Sheet: " + trackerErr.message);
  }

  return "Success";
}

function deleteLot(projectName, phase, lotNum) {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(LOTS_SHEET_NAME);
  var data = sheet.getDataRange().getValues();
  
  for (var i = data.length - 1; i >= 1; i--) { 
    if (data[i][0] == projectName && data[i][1] == phase && data[i][2] == lotNum) {
      try { 
        var folderIdMatch = data[i][12].match(/[-\w]{25,}/);
        if (folderIdMatch) {
          var lotFolder = DriveApp.getFolderById(folderIdMatch[0]);
          var masterFolder = DriveApp.getFolderById(LOT_FOLDERS_MASTER_ID);
          
          var archiveFolders = masterFolder.searchFolders("title = '_ARCHIVED LOTS' and trashed = false");
          var archiveFolder = archiveFolders.hasNext() ? archiveFolders.next() : masterFolder.createFolder('_ARCHIVED LOTS');
          
          lotFolder.moveTo(archiveFolder);
          lotFolder.setName("[ARCHIVED] " + lotFolder.getName());
        }
      } catch(e) { 
        Logger.log("Archiving Error: " + e.message); 
      }
      
      sheet.deleteRow(i + 1); 
      return "Deleted";
    }
  }
}

// --- TAB 3: PURCHASE ORDER FUNCTIONS ---
function getPOFormData() {
  var sheetApp = SpreadsheetApp.openById(SHEET_ID);
  var lotData = sheetApp.getSheetByName(LOTS_SHEET_NAME).getDataRange().getValues();
  var lotMap = {};
  var projects = [];
  
  for(var j=1; j<lotData.length; j++){ 
    var proj = lotData[j][0];
    if(!proj || proj === "Project Name") continue; 
    if(!lotMap[proj]) { lotMap[proj] = []; if(projects.indexOf(proj) === -1) projects.push(proj); }
    lotMap[proj].push({ value: proj + " - " + lotData[j][1] + "." + lotData[j][2], display: "Phase " + lotData[j][1] + " - Lot " + lotData[j][2] });
  } 
  
  var projSettings = getProjectSettings();
  var tradeMap = {};
  for(var k=0; k<projects.length; k++) {
    var pName = projects[k];
    var cProj = cleanText(pName);
    tradeMap[pName] = [];
    if(projSettings[cProj] && projSettings[cProj].pricingSheetId) {
      try {
        var pSheetApp = SpreadsheetApp.openById(projSettings[cProj].pricingSheetId);
        var sheets = pSheetApp.getSheets();
        for(var s=0; s<sheets.length; s++) { tradeMap[pName].push(sheets[s].getName()); }
      } catch(e) { Logger.log("Error loading trades for " + pName + ": " + e.message); }
    }
  }
  return { projects: projects, lotMap: lotMap, tradeMap: tradeMap };
}

function generatePurchaseOrder(lotNumCombined, trade, scope, comments, initialStatus, isManual, manualItems) {
  var sheetApp = SpreadsheetApp.openById(SHEET_ID);
  var lotsData = sheetApp.getSheetByName(LOTS_SHEET_NAME).getDataRange().getValues();
  var lotRecord = null;
  for(var i=1; i<lotsData.length; i++) { 
    if(lotsData[i][0] + " - " + lotsData[i][1] + "." + lotsData[i][2] === lotNumCombined) { lotRecord = lotsData[i]; break; } 
  }
  if(!lotRecord) return "Error: Lot not found.";

  var projectName = lotRecord[0], phase = lotRecord[1], lotNum = lotRecord[2], lotAddress = lotRecord[3], model = lotRecord[4];
  var contractorData = sheetApp.getSheetByName(SHEET_NAME).getDataRange().getDisplayValues();
  var tradeAddress = "Address Not Found in Directory", tradePhone = "Phone Not Found", tradeEmail = "Email Not Found";
  var tradeType = "";
  
  for(var t=1; t<contractorData.length; t++) {
    if(cleanText(contractorData[t][0]) === cleanText(trade)) {
      tradeEmail = contractorData[t][2] || "N/A"; 
      tradePhone = contractorData[t][3] || "N/A"; 
      tradeAddress = contractorData[t][4] || "N/A"; 
      tradeType = contractorData[t][8] || ""; 
      break;
    }
  }

  var projSettings = getProjectSettings();
  var targetConfig = projSettings[cleanText(projectName)];
  if(!targetConfig) return "Error: Project not found in Settings tab.";
  if(!targetConfig.pricingSheetId && !isManual) return "Error: Missing Pricing Sheet ID for this project in Settings.";

  var generatedPONumber = targetConfig.poPrefix + " - " + ("000" + targetConfig.nextPoNum).slice(-3);
  
  var subtotal = 0;
  var itemizedList = [];
  var missingLog = [];

  if (isManual) {
    for (var mIdx = 0; mIdx < manualItems.length; mIdx++) {
      var manualCost = parseFloat(manualItems[mIdx].cost);
      var codeString = manualItems[mIdx].code ? "[" + manualItems[mIdx].code + "] " : ""; 
      subtotal += manualCost;
      itemizedList.push(codeString + manualItems[mIdx].desc + " .................... $" + manualCost.toFixed(2));
    }
  } else {
    try {
      var pricingApp = SpreadsheetApp.openById(targetConfig.pricingSheetId);
      var pricingSheet = pricingApp.getSheetByName(trade);
      if(!pricingSheet) return "Error: Trade tab not found in the Pricing Sheet for " + projectName;
      var pricingData = pricingSheet.getDataRange().getValues();
    } catch(e) { return "Error connecting to Pricing Spreadsheet: " + e.message; }

    var modelCol = -1, targetModelCleaned = cleanText(model);
    for(var c=1; c<pricingData[0].length; c++) { if(cleanText(pricingData[0][c]) === targetModelCleaned) { modelCol = c; break; } }
    if(modelCol === -1) return "Error: Model '" + model + "' column not found on pricing tab.";

    var rawOptions = [lotRecord[5], lotRecord[6], lotRecord[7], lotRecord[8], lotRecord[9], lotRecord[10], lotRecord[11]];
    var activeOptions = [];
    for(var k=0; k<rawOptions.length; k++) { if(rawOptions[k] && String(rawOptions[k]).trim() !== "") activeOptions.push(rawOptions[k]); }

    var cleanOptions = activeOptions.map(function(opt) { 
      var c = cleanText(opt);
      return c === "standard" ? "base price" : c; 
    });
    if (cleanOptions.indexOf("base price") === -1) cleanOptions.push("base price"); 
    var matchedCleanOptions = [];

    for(var r=1; r<pricingData.length; r++) {
      var optionName = pricingData[r][0], cleanOptionName = cleanText(optionName);
      if(cleanOptionName !== "" && cleanOptions.indexOf(cleanOptionName) > -1) {
        var price = parseFloat(String(pricingData[r][modelCol]).replace(/[^0-9.-]+/g,""));
        if(!isNaN(price) && price > 0) {
          subtotal += price; 
          var displayOptionName = (cleanOptionName === "base price") ? "Standard" : optionName; 
          itemizedList.push(displayOptionName + " .................... $" + price.toFixed(2)); 
          matchedCleanOptions.push(cleanOptionName); 
        }
      }
    }

    for(var m=0; m<activeOptions.length; m++) { 
      var expected = cleanText(activeOptions[m]);
      if (expected === "standard") expected = "base price";
      if(matchedCleanOptions.indexOf(expected) === -1) missingLog.push(activeOptions[m]); 
    }
  }

  var colourSelectionCost = 0;
  if (tradeType && tradeType.trim() !== "") {
    try {
      var safeProjectName = projectName.replace(/'/g, "\\'");
      var safePhaseName = ("Phase " + phase).replace(/'/g, "\\'");
      var safeLotName = ("Lot " + lotNum).replace(/'/g, "\\'");
      
      var restrictedMaster = DriveApp.getFolderById(RESTRICTED_LOT_FOLDERS_MASTER_ID);
      var projSearch = restrictedMaster.searchFolders("title = '" + safeProjectName + "' and trashed = false");
      
      if (projSearch.hasNext()) {
        var restrictedProjectFolder = projSearch.next();
        var phaseSearch = restrictedProjectFolder.searchFolders("title = '" + safePhaseName + "' and trashed = false");
        
        if (phaseSearch.hasNext()) {
          var restrictedPhaseFolder = phaseSearch.next();
          var folderSearch = restrictedPhaseFolder.searchFolders("title = '" + safeLotName + "' and trashed = false");
          
          if (folderSearch.hasNext()) {
            var lotFolder = folderSearch.next();
            var files = lotFolder.searchFiles("mimeType='application/vnd.google-apps.spreadsheet' and trashed=false");
            
            while (files.hasNext()) {
              var sheetFile = files.next();
              var colourApp = SpreadsheetApp.openById(sheetFile.getId());
              var sheets = colourApp.getSheets();
              var cvrTab = null;
              
              for(var s = 0; s < sheets.length; s++) {
                if(sheets[s].getName().toLowerCase().replace(/[^a-z]/g, '') === "costvsretail") {
                  cvrTab = sheets[s]; break;
                }
              }
              
              if (cvrTab) {
                var cvrData = cvrTab.getRange("A14:B").getValues();
                for (var c = 0; c < cvrData.length; c++) {
                  var rowTrade = String(cvrData[c][0]).trim();
                  var rowCost = parseFloat(String(cvrData[c][1]).replace(/[^0-9.-]+/g, ""));
                  
                  if (rowTrade !== "" && !isNaN(rowCost) && rowCost !== 0) {
                    if (fuzzyMatch(tradeType, rowTrade)) {
                      colourSelectionCost += rowCost;
                    }
                  }
                }
                break; 
              }
            }
          }
        }
      }
    } catch(e) { Logger.log("Colour Selection Fetch Error: " + e.message); }
  }

  if (colourSelectionCost !== 0) { 
    subtotal += colourSelectionCost;
    itemizedList.push("Colour Selection Items .................... $" + colourSelectionCost.toFixed(2));
  }

  var taxAmount = subtotal * targetConfig.taxRate, grandTotal = subtotal + taxAmount;
  var pdfUrl = "No Template Provided", pdfErrorMsg = "", emailStatus = "Not Sent";

  if(PO_TEMPLATE_DOC_ID.length > 20) { 
    try {
      var templateFile = DriveApp.getFileById(PO_TEMPLATE_DOC_ID);
      var pdfName = generatedPONumber + " - " + trade + " - " + projectName + " - " + phase + "." + lotNum;
      var tempDoc = templateFile.makeCopy(pdfName); 
      var doc = DocumentApp.openById(tempDoc.getId()), body = doc.getBody();

      body.replaceText("{{ProjectName}}", projectName); body.replaceText("{{Phase}}", phase); body.replaceText("{{Lot}}", lotNum); 
      body.replaceText("{{Trade}}", trade); body.replaceText("{{Model}}", model); body.replaceText("{{Items}}", itemizedList.join("\n"));
      body.replaceText("{{Subtotal}}", "$" + subtotal.toFixed(2)); body.replaceText("{{Tax}}", "$" + taxAmount.toFixed(2)); body.replaceText("{{Total}}", "$" + grandTotal.toFixed(2));
      body.replaceText("{{BillName}}", targetConfig.billName); body.replaceText("{{BillAddress}}", targetConfig.billAddress);
      body.replaceText("{{BillPhone}}", targetConfig.billPhone); body.replaceText("{{BillEmail}}", targetConfig.billEmail);
      body.replaceText("{{TradeName}}", trade); body.replaceText("{{TradeAddress}}", tradeAddress);
      body.replaceText("{{TradePhone}}", tradePhone); body.replaceText("{{TradeEmail}}", tradeEmail);
      body.replaceText("{{LotAddress}}", lotAddress); body.replaceText("{{Scope}}", scope || ""); body.replaceText("{{Comments}}", comments || "");
      body.replaceText("{{Date}}", new Date().toLocaleDateString()); body.replaceText("{{PONumber}}", generatedPONumber); body.replaceText("{{RevisionDate}}", ""); 
      doc.saveAndClose();
      
      var poMasterFolder = DriveApp.getFolderById(PO_FOLDERS_MASTER_ID);
      var safeProjectName = projectName.replace(/'/g, "\\'");
      var projectFolders = poMasterFolder.searchFolders("title = '" + safeProjectName + "' and trashed = false");
      var projectFolder = projectFolders.hasNext() ? projectFolders.next() : poMasterFolder.createFolder(projectName);
      var phaseStr = "Phase " + phase, safePhaseStr = phaseStr.replace(/'/g, "\\'");
      var phaseFolders = projectFolder.searchFolders("title = '" + safePhaseStr + "' and trashed = false");
      var phaseFolder = phaseFolders.hasNext() ? phaseFolders.next() : projectFolder.createFolder(phaseStr);
      var lotStr = "Lot " + lotNum, safeLotStr = lotStr.replace(/'/g, "\\'");
      var lotFolders = phaseFolder.searchFolders("title = '" + safeLotStr + "' and trashed = false");
      var destFolder = lotFolders.hasNext() ? lotFolders.next() : phaseFolder.createFolder(lotStr);

      var pdfBlob = doc.getAs('application/pdf');
      pdfUrl = destFolder.createFile(pdfBlob).getUrl();
      tempDoc.setTrashed(true);
      
      sheetApp.getSheetByName('Settings').getRange(targetConfig.rowIndex, 7).setValue(targetConfig.nextPoNum + 1);

      if (tradeEmail && tradeEmail.indexOf("@") > -1) {
        GmailApp.sendEmail(tradeEmail, "New Purchase Order: " + generatedPONumber + " - " + projectName, 
          "Hello " + trade + ",\n\nPlease find attached the new Purchase Order (" + generatedPONumber + ") for " + projectName + ", Lot " + lotNumCombined + " at " + lotAddress + ".\n\nScope of Work: " + (scope ? scope : "As per standard contract") + "\n\n" + (comments ? "Additional Comments: " + comments + "\n\n" : "") + "Thank you,\nGeorgian Communities", 
          { name: "Georgian Build Connect", attachments: [pdfBlob.setName(pdfName + ".pdf")] });
        emailStatus = "SUCCESS";
      } else { emailStatus = "FAILED (No valid email on file)"; }
    } catch(e) { Logger.log("PDF/Email Error: " + e.message); pdfUrl = "PDF Error: " + e.message; pdfErrorMsg = e.message; }
  }

  var poSheet = sheetApp.getSheetByName(PO_SHEET_NAME);
  var statusToSave = initialStatus ? initialStatus : "Open";
  if(poSheet) { poSheet.appendRow([new Date().toLocaleDateString(), projectName, generatedPONumber, lotNum, trade, "$" + subtotal.toFixed(2), pdfUrl, statusToSave, "$0.00"]); }

  try {
    var scheduleUrl = lotRecord[13]; 
    if (scheduleUrl && scheduleUrl.indexOf("http") === 0) {
      var schedIdMatch = scheduleUrl.match(/[-\w]{25,}/);
      if (schedIdMatch) {
        var schedApp = SpreadsheetApp.openById(schedIdMatch[0]);
        var schedSheet = schedApp.getSheets()[0];
        var lastRow = schedSheet.getLastRow();
        
        if (lastRow > 0) {
          var schedData = schedSheet.getRange("B1:C" + lastRow).getValues();
          
          var tradeCostCodes = [];
          for(var t = 1; t < contractorData.length; t++) {
             if(cleanText(contractorData[t][0]) === cleanText(trade)) {
                var codesStr = contractorData[t][5] || ""; 
                var codesArr = codesStr.split(",");
                for(var i = 0; i < codesArr.length; i++) {
                   if(codesArr[i].trim() !== "") {
                     tradeCostCodes.push(cleanText(codesArr[i].trim()));
                   }
                }
                break;
             }
          }

          var updates = [];
          var madeChanges = false;
          for(var r = 0; r < schedData.length; r++) {
             var rowCode = cleanText(String(schedData[r][0]).trim());
             if(rowCode && tradeCostCodes.indexOf(rowCode) > -1) {
                updates.push([trade]); 
                madeChanges = true;
             } else {
                updates.push([schedData[r][1]]); 
             }
          }
          if(madeChanges) {
             schedSheet.getRange(1, 3, updates.length, 1).setValues(updates);
             schedSheet.getRange("F3").setValue(new Date().toLocaleDateString());
          }
        }
      }
    }
  } catch(err) {
    Logger.log("Failed to update schedule with PO trade: " + err.message);
  }

  try {
    var trackerId = targetConfig.trackerSheetId;
    if (trackerId && trackerId.length > 20) {
      var trackerApp = SpreadsheetApp.openById(trackerId);
      var phaseTabName = "Phase " + phase;
      var phaseSheet = trackerApp.getSheetByName(phaseTabName);

      if (!phaseSheet) {
        phaseSheet = trackerApp.insertSheet(phaseTabName);
        var templateHeaders = ["Lots", "Lot Sold", "IFC Drawings", "CS Docs"];
        phaseSheet.getRange(1, 1, 1, templateHeaders.length).setValues([templateHeaders]).setFontWeight("bold");
        phaseSheet.setFrozenRows(1);
        phaseSheet.setFrozenColumns(1);
      }

      var tData = phaseSheet.getDataRange().getDisplayValues();

      var lotRow = -1;
      for (var r = 1; r < tData.length; r++) {
        if (String(tData[r][0]).trim() === String(lotNum).trim()) {
          lotRow = r + 1;
          break;
        }
      }
      
      if (lotRow <= 1) {
        lotRow = Math.max(2, phaseSheet.getLastRow() + 1);
        phaseSheet.getRange(lotRow, 1).setValue(lotNum);
        tData = phaseSheet.getDataRange().getDisplayValues(); 
      }

      var tradeCodes = [];
      for (var t = 1; t < contractorData.length; t++) {
        if (cleanText(contractorData[t][0]) === cleanText(trade)) {
          var cStr = contractorData[t][5] || ""; 
          var cArr = cStr.split(",");
          for (var c = 0; c < cArr.length; c++) {
            if (cArr[c].trim() !== "") {
              tradeCodes.push({
                clean: cleanText(cArr[c].trim()),
                original: cArr[c].trim() 
              });
            }
          }
          break;
        }
      }

      var headers = tData.length > 0 ? tData[0] : ["Lots", "Lot Sold", "IFC Drawings", "CS Docs"];
      var updateValue = trade + " - " + new Date().toLocaleDateString();

      for (var c = 0; c < tradeCodes.length; c++) {
        var targetCode = tradeCodes[c];
        var colIndex = -1;

        for (var h = 4; h < headers.length; h++) {
          if (cleanText(headers[h]) === targetCode.clean) {
            colIndex = h + 1;
            break;
          }
        }

        if (colIndex === -1) {
          colIndex = Math.max(5, headers.length + 1);
          phaseSheet.getRange(1, colIndex).setValue(targetCode.original).setFontWeight("bold");
          headers.push(targetCode.original); 
        }

        phaseSheet.getRange(lotRow, colIndex).setValue(updateValue);
      }
      
      applyTrackerFormatting(phaseSheet);
    }
  } catch (trackerErr) {
    Logger.log("Failed to update PO Tracker for " + projectName + ": " + trackerErr.message);
  }

  if (pdfErrorMsg !== "" || missingLog.length > 0) {
    return "Error: Read Details Below\n\n1. CALCULATED PRE-TAX TOTAL: $" + subtotal.toFixed(2) + "\n\n2. MISSING OPTIONS:\n" + (missingLog.length > 0 ? missingLog.join("\n") : "None!") + "\n\n3. PDF GENERATION: " + (pdfErrorMsg !== "" ? "FAILED" : "SUCCESS") + "\n\n4. EMAIL DISPATCH: " + emailStatus;
  }
  return "Success";
}

function editPurchaseOrder(poNum, scope, comments, overridePrice, additionalItems) {
  var sheetApp = SpreadsheetApp.openById(SHEET_ID);
  var poSheet = sheetApp.getSheetByName(PO_SHEET_NAME);
  var poData = poSheet.getDataRange().getValues();
  var poRow = -1;
  var projectName, lotCombined, trade;

  for(var i=1; i<poData.length; i++) {
    if(poData[i][2] === poNum) {
      poRow = i + 1;
      projectName = poData[i][1];
      lotCombined = poData[i][3]; 
      trade = poData[i][4];
      break;
    }
  }
  if(poRow === -1) return "Error: Original PO not found in database.";

  var match = poNum.match(/^(.*)\s-\s(\d{2})$/);
  var newPoNum = poNum + " - 01"; 
  if (match) {
    var rev = parseInt(match[2], 10) + 1;
    newPoNum = match[1] + " - " + ("0" + rev).slice(-2); 
  }

  var lotsData = sheetApp.getSheetByName(LOTS_SHEET_NAME).getDataRange().getValues();
  var lotRecord = null;
  for(var i=1; i<lotsData.length; i++) {
    if(String(lotsData[i][2]) === String(lotCombined) && lotsData[i][0] === projectName) {
      lotRecord = lotsData[i]; break;
    }
  }
  if(!lotRecord) return "Error: Lot not found in Directory.";

  var phase = lotRecord[1], lotNum = lotRecord[2], lotAddress = lotRecord[3], model = lotRecord[4];

  var contractorData = sheetApp.getSheetByName(SHEET_NAME).getDataRange().getDisplayValues();
  var tradeAddress = "Address Not Found", tradePhone = "Phone Not Found", tradeEmail = "Email Not Found";
  var tradeType = "";
  
  for(var t=1; t<contractorData.length; t++) {
    if(cleanText(contractorData[t][0]) === cleanText(trade)) {
      tradeEmail = contractorData[t][2] || "N/A"; 
      tradePhone = contractorData[t][3] || "N/A"; 
      tradeAddress = contractorData[t][4] || "N/A"; 
      tradeType = contractorData[t][8] || ""; 
      break;
    }
  }

  var projSettings = getProjectSettings();
  var targetConfig = projSettings[cleanText(projectName)];
  if(!targetConfig) return "Error: Project not found in Settings.";

  var subtotal = 0;
  var itemizedList = [];
  var missingLog = [];

  if (overridePrice && String(overridePrice).trim() !== "") {
    var parsedPrice = parseFloat(overridePrice);
    if (!isNaN(parsedPrice) && parsedPrice >= 0) {
      subtotal = parsedPrice;
      itemizedList.push("Revised Contract Amount .................... $" + subtotal.toFixed(2));
    } else {
      return "Error: Invalid override price.";
    }
  } else {
    if(!targetConfig.pricingSheetId) return "Error: Missing Pricing Sheet ID.";
    try {
      var pricingApp = SpreadsheetApp.openById(targetConfig.pricingSheetId);
      var pricingSheet = pricingApp.getSheetByName(trade);
      if(!pricingSheet) return "Error: Trade tab not found.";
      var pricingData = pricingSheet.getDataRange().getValues();
    } catch(e) { return "Error connecting to Pricing: " + e.message; }

    var modelCol = -1, targetModelCleaned = cleanText(model);
    for(var c=1; c<pricingData[0].length; c++) { if(cleanText(pricingData[0][c]) === targetModelCleaned) { modelCol = c; break; } }
    if(modelCol === -1) return "Error: Model '" + model + "' not found on pricing tab.";

    var rawOptions = [lotRecord[5], lotRecord[6], lotRecord[7], lotRecord[8], lotRecord[9], lotRecord[10], lotRecord[11]];
    var activeOptions = [];
    for(var k=0; k<rawOptions.length; k++) { if(rawOptions[k] && String(rawOptions[k]).trim() !== "") activeOptions.push(rawOptions[k]); }

    var cleanOptions = activeOptions.map(function(opt) { 
      var c = cleanText(opt);
      return c === "standard" ? "base price" : c;
    });
    if (cleanOptions.indexOf("base price") === -1) cleanOptions.push("base price");
    var matchedCleanOptions = [];

    for(var r=1; r<pricingData.length; r++) {
      var optionName = pricingData[r][0], cleanOptionName = cleanText(optionName);
      if(cleanOptionName !== "" && cleanOptions.indexOf(cleanOptionName) > -1) {
        var price = parseFloat(String(pricingData[r][modelCol]).replace(/[^0-9.-]+/g,""));
        if(!isNaN(price) && price > 0) {
          subtotal += price; 
          var displayOptionName = (cleanOptionName === "base price") ? "Standard" : optionName;
          itemizedList.push(displayOptionName + " .................... $" + price.toFixed(2)); 
          matchedCleanOptions.push(cleanOptionName);
        }
      }
    }
    for(var m=0; m<activeOptions.length; m++) { 
      var expected = cleanText(activeOptions[m]);
      if (expected === "standard") expected = "base price";
      if(matchedCleanOptions.indexOf(expected) === -1) missingLog.push(activeOptions[m]); 
    }
  }

  var colourSelectionCost = 0;
  if (!overridePrice && tradeType && tradeType.trim() !== "") {
    try {
      var safeProjectName = projectName.replace(/'/g, "\\'");
      var safePhaseName = ("Phase " + phase).replace(/'/g, "\\'");
      var safeLotName = ("Lot " + lotNum).replace(/'/g, "\\'");
      
      var restrictedMaster = DriveApp.getFolderById(RESTRICTED_LOT_FOLDERS_MASTER_ID);
      var projSearch = restrictedMaster.searchFolders("title = '" + safeProjectName + "' and trashed = false");
      
      if (projSearch.hasNext()) {
        var restrictedProjectFolder = projSearch.next();
        var phaseSearch = restrictedProjectFolder.searchFolders("title = '" + safePhaseName + "' and trashed = false");
        
        if (phaseSearch.hasNext()) {
          var restrictedPhaseFolder = phaseSearch.next();
          var folderSearch = restrictedPhaseFolder.searchFolders("title = '" + safeLotName + "' and trashed = false");
          
          if (folderSearch.hasNext()) {
            var lotFolder = folderSearch.next();
            var files = lotFolder.searchFiles("mimeType='application/vnd.google-apps.spreadsheet' and trashed=false");
            
            while (files.hasNext()) {
              var sheetFile = files.next();
              var colourApp = SpreadsheetApp.openById(sheetFile.getId());
              var sheets = colourApp.getSheets();
              var cvrTab = null;
              
              for(var s = 0; s < sheets.length; s++) {
                if(sheets[s].getName().toLowerCase().replace(/[^a-z]/g, '') === "costvsretail") {
                  cvrTab = sheets[s]; break;
                }
              }
              
              if (cvrTab) {
                var cvrData = cvrTab.getRange("A14:B").getValues();
                for (var c = 0; c < cvrData.length; c++) {
                  var rowTrade = String(cvrData[c][0]).trim();
                  var rowCost = parseFloat(String(cvrData[c][1]).replace(/[^0-9.-]+/g, ""));
                  
                  if (rowTrade !== "" && !isNaN(rowCost) && rowCost !== 0) {
                    if (fuzzyMatch(tradeType, rowTrade)) {
                      colourSelectionCost += rowCost;
                    }
                  }
                }
                break; 
              }
            }
          }
        }
      }
    } catch(e) { Logger.log("Colour Selection Fetch Error: " + e.message); }
  }

  if (colourSelectionCost !== 0) { 
    subtotal += colourSelectionCost;
    itemizedList.push("Colour Selection Items .................... $" + colourSelectionCost.toFixed(2));
  }

  if (additionalItems && additionalItems.length > 0) {
    for (var a = 0; a < additionalItems.length; a++) {
      var itemCost = parseFloat(additionalItems[a].cost);
      if (!isNaN(itemCost)) {
        var codeString = additionalItems[a].code ? "[" + additionalItems[a].code + "] " : "";
        subtotal += itemCost;
        itemizedList.push(codeString + additionalItems[a].desc + " .................... $" + itemCost.toFixed(2));
      }
    }
  }

  var taxAmount = subtotal * targetConfig.taxRate, grandTotal = subtotal + taxAmount;
  var pdfUrl = "No Template Provided", pdfErrorMsg = "", emailStatus = "Not Sent";

  if(PO_TEMPLATE_DOC_ID.length > 20) {
    try {
      var templateFile = DriveApp.getFileById(PO_TEMPLATE_DOC_ID);
      var pdfName = newPoNum + " - " + trade + " - " + projectName + " - " + phase + "." + lotNum + " (Rev " + Date.now() + ")";
      var tempDoc = templateFile.makeCopy(pdfName);
      var doc = DocumentApp.openById(tempDoc.getId()), body = doc.getBody();

      body.replaceText("{{ProjectName}}", projectName); body.replaceText("{{Phase}}", phase); body.replaceText("{{Lot}}", lotNum);
      body.replaceText("{{Trade}}", trade); body.replaceText("{{Model}}", model); body.replaceText("{{Items}}", itemizedList.join("\n"));
      body.replaceText("{{Subtotal}}", "$" + subtotal.toFixed(2)); body.replaceText("{{Tax}}", "$" + taxAmount.toFixed(2)); body.replaceText("{{Total}}", "$" + grandTotal.toFixed(2));
      body.replaceText("{{BillName}}", targetConfig.billName); body.replaceText("{{BillAddress}}", targetConfig.billAddress);
      body.replaceText("{{BillPhone}}", targetConfig.billPhone); body.replaceText("{{BillEmail}}", targetConfig.billEmail);
      body.replaceText("{{TradeName}}", trade); body.replaceText("{{TradeAddress}}", tradeAddress);
      body.replaceText("{{TradePhone}}", tradePhone); body.replaceText("{{TradeEmail}}", tradeEmail);
      body.replaceText("{{LotAddress}}", lotAddress); body.replaceText("{{Scope}}", scope || ""); body.replaceText("{{Comments}}", comments || "");
      body.replaceText("{{Date}}", new Date().toLocaleDateString()); body.replaceText("{{PONumber}}", newPoNum);
      body.replaceText("{{RevisionDate}}", "Revised: " + new Date().toLocaleDateString());

      doc.saveAndClose();

      var poMasterFolder = DriveApp.getFolderById(PO_FOLDERS_MASTER_ID);
      var safeProjectName = projectName.replace(/'/g, "\\'");
      var projectFolders = poMasterFolder.searchFolders("title = '" + safeProjectName + "' and trashed = false");
      var projectFolder = projectFolders.hasNext() ? projectFolders.next() : poMasterFolder.createFolder(projectName);
      var phaseStr = "Phase " + phase, safePhaseStr = phaseStr.replace(/'/g, "\\'");
      var phaseFolders = projectFolder.searchFolders("title = '" + safePhaseStr + "' and trashed = false");
      var phaseFolder = phaseFolders.hasNext() ? phaseFolders.next() : projectFolder.createFolder(phaseStr);
      var lotStr = "Lot " + lotNum, safeLotStr = lotStr.replace(/'/g, "\\'");
      var lotFolders = phaseFolder.searchFolders("title = '" + safeLotStr + "' and trashed = false");
      var destFolder = lotFolders.hasNext() ? lotFolders.next() : phaseFolder.createFolder(lotStr);

      var pdfBlob = doc.getAs('application/pdf');
      pdfUrl = destFolder.createFile(pdfBlob).getUrl();
      tempDoc.setTrashed(true);

      if (tradeEmail && tradeEmail.indexOf("@") > -1) {
        GmailApp.sendEmail(tradeEmail, "REVISED Purchase Order: " + newPoNum + " - " + projectName,
          "Hello " + trade + ",\n\nPlease find attached the REVISED Purchase Order (" + newPoNum + ") for " + projectName + ", Lot " + lotCombined + " at " + lotAddress + ".\n\nScope of Work: " + (scope ? scope : "As per standard contract") + "\n\n" + (comments ? "Additional Comments: " + comments + "\n\n" : "") + "Thank you,\nGeorgian Communities",
          { name: "Georgian Build Connect", attachments: [pdfBlob.setName(pdfName + ".pdf")] });
        emailStatus = "SUCCESS";
      } else { emailStatus = "FAILED (No valid email on file)"; }
    } catch(e) { Logger.log("PDF/Email Error: " + e.message); pdfUrl = "PDF Error: " + e.message; pdfErrorMsg = e.message; }
  }

  poSheet.getRange(poRow, 1).setValue(new Date().toLocaleDateString());
  poSheet.getRange(poRow, 3).setValue(newPoNum);
  poSheet.getRange(poRow, 6).setValue("$" + subtotal.toFixed(2));
  poSheet.getRange(poRow, 7).setValue(pdfUrl);

  var drawSheet = sheetApp.getSheetByName(DRAW_SHEET_NAME);
  if(drawSheet) {
    var drawData = drawSheet.getDataRange().getValues();
    for(var d=1; d<drawData.length; d++) {
      if(drawData[d][1] === poNum) {
        drawSheet.getRange(d+1, 2).setValue(newPoNum);
      }
    }
  }

  if (pdfErrorMsg !== "" || missingLog.length > 0) {
    return "Error: Read Details Below\n\n1. REVISED PRE-TAX TOTAL: $" + subtotal.toFixed(2) + "\n\n2. MISSING OPTIONS:\n" + (missingLog.length > 0 ? missingLog.join("\n") : "None!") + "\n\n3. PDF GENERATION: " + (pdfErrorMsg !== "" ? "FAILED" : "SUCCESS") + "\n\n4. EMAIL DISPATCH: " + emailStatus;
  }
  return "Success|" + newPoNum;
}

function getPOsData(filterObj) {
  try {
    var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(PO_SHEET_NAME);
    if (!sheet) return [["Error", "Please create a tab exactly named 'Purchase Orders'."]];
    var data = sheet.getDataRange().getDisplayValues();
    if (!data || data.length === 0 || data[0][0] === "") return [["Status", "No Purchase Orders generated yet."]];
    if (!filterObj || Object.keys(filterObj).length === 0 || data.length <= 1) return data;
    
    var headers = data[0]; var filteredData = [headers];
    for (var i = 1; i < data.length; i++) {
      var row = data[i], isMatch = true;
      for (var key in filterObj) {
        if (filterObj[key] && filterObj[key] !== "") {
          var colIndex = headers.indexOf(key);
          if (colIndex > -1) {
            if (String(row[colIndex]).toLowerCase().indexOf(String(filterObj[key]).toLowerCase().trim()) === -1) { isMatch = false; break; }
          }
        }
      }
      if (isMatch) filteredData.push(row);
    }
    return filteredData.length > 1 ? filteredData : [headers, ["Status", "No POs match your filters."]];
  } catch(e) { return [["CRITICAL BACKEND ERROR", e.message]]; }
}

function getPODetails(poNum) {
  var sheetApp = SpreadsheetApp.openById(SHEET_ID);
  var poSheet = sheetApp.getSheetByName(PO_SHEET_NAME);
  var data = poSheet.getDataRange().getDisplayValues();
  var poDetails = { status: "Open", total: "$0.00", paid: "$0.00", history: [] };
  
  for(var i=1; i<data.length; i++) {
    if(data[i][2] === poNum) {
      poDetails.total = data[i][5] || "$0.00";
      poDetails.status = data[i][7] || "Open";
      poDetails.paid = data[i][8] || "$0.00";
      break;
    }
  }

  var drawSheet = sheetApp.getSheetByName(DRAW_SHEET_NAME);
  if(drawSheet) {
    var drawData = drawSheet.getDataRange().getDisplayValues();
    for(var j=1; j<drawData.length; j++) {
      if(drawData[j][1] === poNum) {
        poDetails.history.push({ date: drawData[j][0], amount: drawData[j][3], notes: drawData[j][4] });
      }
    }
  }
  return poDetails;
}

function updatePOStatus(poNum, newStatus) {
  var sheetApp = SpreadsheetApp.openById(SHEET_ID);
  var poSheet = sheetApp.getSheetByName(PO_SHEET_NAME);
  var data = poSheet.getDataRange().getValues();
  
  for(var i=1; i<data.length; i++) {
    if(data[i][2] == poNum) {
      var currentStatus = data[i][7];
      var emailStatusMsg = ""; 
      
      if (newStatus === "Void" && currentStatus !== "Void") {
        var projectName = data[i][1];
        var lotNum = data[i][3];
        var trade = data[i][4];
        var pdfUrl = data[i][6];
        
        try {
          if (pdfUrl && pdfUrl.indexOf("http") === 0) {
            var fileIdMatch = pdfUrl.match(/[-\w]{25,}/);
            if (fileIdMatch) {
              var file = DriveApp.getFileById(fileIdMatch[0]);
              var currentName = file.getName();
              if (currentName.indexOf("VOID") === -1) { file.setName("VOID - " + currentName); }
            }
          }
        } catch(e) { Logger.log("Error renaming voided file: " + e.message); }

        try {
          var lotsData = sheetApp.getSheetByName(LOTS_SHEET_NAME).getDataRange().getValues();
          var scheduleUrl = null;
          for(var j=1; j<lotsData.length; j++) {
             if(lotsData[j][0] === projectName && String(lotsData[j][2]) === String(lotNum)) {
               scheduleUrl = lotsData[j][13]; 
               break;
             }
          }
          
          if (scheduleUrl && scheduleUrl.indexOf("http") === 0) {
            var schedIdMatch = scheduleUrl.match(/[-\w]{25,}/);
            if (schedIdMatch) {
              var schedApp = SpreadsheetApp.openById(schedIdMatch[0]);
              var schedSheet = schedApp.getSheets()[0];
              var lastRow = schedSheet.getLastRow();
              
              if (lastRow > 0) {
                var schedTrades = schedSheet.getRange(1, 3, lastRow, 1).getValues();
                var updates = [];
                var madeChanges = false;
                
                for (var r = 0; r < schedTrades.length; r++) {
                  if (String(schedTrades[r][0]).trim() === String(trade).trim()) {
                    updates.push([""]); 
                    madeChanges = true;
                  } else {
                    updates.push([schedTrades[r][0]]); 
                  }
                }
                if (madeChanges) {
                  schedSheet.getRange(1, 3, updates.length, 1).setValues(updates);
                  schedSheet.getRange("F3").setValue(new Date().toLocaleDateString());
                }
              }
            }
          }
        } catch(e) { Logger.log("Error updating schedule for voided PO: " + e.message); }

        try {
          var contractorData = sheetApp.getSheetByName(SHEET_NAME).getDataRange().getDisplayValues();
          var tradeEmail = "";
          
          for(var t=1; t<contractorData.length; t++) {
            if(cleanText(contractorData[t][0]) === cleanText(trade)) {
              tradeEmail = contractorData[t][2];
              break;
            }
          }
          
          if (tradeEmail && tradeEmail.indexOf("@") > -1) {
            var subject = "VOIDED Purchase Order: " + poNum + " - " + projectName;
            var body = "Hello " + trade + ",\n\n" +
                       "Please be advised that Purchase Order " + poNum + " for " + projectName + ", Lot " + lotNum + " has been VOIDED.\n\n" +
                       "If you have any questions or require clarification, please contact us.\n\n" +
                       "Thank you,\nGeorgian Communities";
            
            GmailApp.sendEmail(tradeEmail, subject, body, { name: "Georgian Build Connect" });
            emailStatusMsg = " (Void email sent to " + tradeEmail + ")";
          } else {
            emailStatusMsg = " (ERROR: No valid email found for '" + trade + "' in Directory)";
          }
        } catch(e) { 
          emailStatusMsg = " (ERROR sending email: " + e.message + ")";
        }
      }

      poSheet.getRange(i+1, 8).setValue(newStatus);
      return "Status Updated to " + newStatus + "!" + emailStatusMsg;
    }
  }
  return "Error: PO not found.";
}

function addProgressDraw(poNum, trade, amount, notes) {
  var sheetApp = SpreadsheetApp.openById(SHEET_ID);
  var drawSheet = sheetApp.getSheetByName(DRAW_SHEET_NAME);
  if(!drawSheet) return "Error: Please create the 'Progress Draws' tab first.";
  
  var cleanAmount = parseFloat(String(amount).replace(/[^0-9.-]+/g,""));
  if(isNaN(cleanAmount) || cleanAmount === 0) return "Error: Invalid draw amount (must be positive or negative, not zero).";
  
  var poSheet = sheetApp.getSheetByName(PO_SHEET_NAME);
  var data = poSheet.getDataRange().getValues();
  var poRowIndex = -1;
  var poTotal = 0;
  var currentPaid = 0;

  for(var i=1; i<data.length; i++) {
    if(data[i][2] == poNum) {
      poRowIndex = i + 1;
      poTotal = parseFloat(String(data[i][5]).replace(/[^0-9.-]+/g,"")) || 0;
      currentPaid = parseFloat(String(data[i][8]).replace(/[^0-9.-]+/g,"")) || 0;
      break;
    }
  }

  if(poRowIndex === -1) return "Error: PO not found.";

  var newPaid = currentPaid + cleanAmount;

  if (newPaid > (poTotal + 0.01)) {
    var remaining = poTotal - currentPaid;
    return "Error: This draw ($" + cleanAmount.toFixed(2) + ") would exceed the PO Amount. You currently have $" + remaining.toFixed(2) + " remaining to draw.";
  }

  drawSheet.appendRow([new Date().toLocaleDateString(), poNum, trade, "$" + cleanAmount.toFixed(2), notes]);
  poSheet.getRange(poRowIndex, 9).setValue("$" + newPaid.toFixed(2));
  
  return "Success";
}

function scanAndNotifyNewDocuments() {
  var props = PropertiesService.getScriptProperties();
  var lastRunStr = props.getProperty('lastRunTime');
  var now = new Date();
  var lastRunTime = new Date(now.getTime() - 60 * 60 * 1000);
  if (lastRunStr) { var parsedDate = new Date(lastRunStr); if (!isNaN(parsedDate.getTime())) lastRunTime = parsedDate; }
  props.setProperty('lastRunTime', now.toISOString());
  
  var sheetApp = SpreadsheetApp.openById(SHEET_ID);
  var contractorData = sheetApp.getSheetByName(SHEET_NAME).getDataRange().getDisplayValues();
  
  var tradeInfo = {};
  for(var i=1; i<contractorData.length; i++) {
    var tNameClean = cleanText(contractorData[i][0]);
    var tNameDisplay = contractorData[i][0]; 
    var tEmail = contractorData[i][2];
    var tCodesRaw = contractorData[i][5] || "";
    var tCodesClean = tCodesRaw.split(",").map(function(c) { return cleanText(c); }).filter(String);
    
    if(tEmail && tEmail.indexOf("@") > -1) {
      tradeInfo[tNameClean] = { display: tNameDisplay, email: tEmail, codes: tCodesClean };
    }
  }
  
  var poData = sheetApp.getSheetByName(PO_SHEET_NAME).getDataRange().getDisplayValues();
  var lotToTrades = {}; 
  for(var j=1; j<poData.length; j++) {
    var key = cleanText(poData[j][1]) + "|" + poData[j][3];
    if(!lotToTrades[key]) lotToTrades[key] = [];
    var poTradeClean = cleanText(poData[j][4]);
    if(lotToTrades[key].indexOf(poTradeClean) === -1) lotToTrades[key].push(poTradeClean); 
  }
  
  var lotsData = sheetApp.getSheetByName(LOTS_SHEET_NAME).getDataRange().getDisplayValues();
  for(var k=1; k<lotsData.length; k++) {
    var projRaw = lotsData[k][0], phase = lotsData[k][1], lot = lotsData[k][2], civic = lotsData[k][3], folderUrl = lotsData[k][12];
    var key = cleanText(projRaw) + "|" + lot;
    
    if(folderUrl && lotToTrades[key] && lotToTrades[key].length > 0) {
      var folderIdMatch = folderUrl.match(/[-\w]{25,}/);
      if(folderIdMatch) {
        try {
          var folder = DriveApp.getFolderById(folderIdMatch[0]), files = folder.getFiles(), newFiles = [];
          while(files.hasNext()) {
            var file = files.next();
            if(file.getDateCreated().getTime() > lastRunTime.getTime() && !file.isTrashed()) {
               var fileObj = { 
                 id: file.getId(),
                 mimeType: file.getMimeType(),
                 name: file.getName(), 
                 url: file.getUrl(), 
                 blob: null 
               };
               
               try { 
                 if (fileObj.mimeType.indexOf('google-apps') > -1) {
                   fileObj.blob = file.getAs('application/pdf').setName(fileObj.name + ".pdf");
                 } else if (file.getSize() < 25000000) {
                   fileObj.blob = file.getBlob(); 
                 } 
               } catch(e) { 
                 Logger.log("Attachment conversion error for " + fileObj.name + ": " + e.message); 
               }

               newFiles.push(fileObj);
            }
          }
          
          if(newFiles.length > 0) {
            var tradesToNotify = lotToTrades[key];
            var tradePayloads = {};
            
            for(var t=0; t<tradesToNotify.length; t++) {
              var tNameClean = tradesToNotify[t];
              if(tradeInfo[tNameClean]) {
                tradePayloads[tNameClean] = { files: [], attachments: [] };
              }
            }
            
            for(var f=0; f<newFiles.length; f++) {
              var fileObj = newFiles[f];
              var fileNameLower = fileObj.name.toLowerCase();
              var isChangeOrder = fileNameLower.indexOf("change order") > -1;
              var applicableTradesForFile = [];
              
              if (isChangeOrder && fileObj.mimeType === 'application/vnd.google-apps.spreadsheet') {
                try {
                  var coApp = SpreadsheetApp.openById(fileObj.id);
                  var coSheet = coApp.getSheets()[0]; 
                  var bData = coSheet.getRange("B1:B").getDisplayValues();
                  var coCodes = [];
                  
                  for(var b=0; b<bData.length; b++) {
                    var cClean = cleanText(bData[b][0]);
                    if(cClean !== "" && coCodes.indexOf(cClean) === -1) coCodes.push(cClean);
                  }
                  
                  for(var t2=0; t2<tradesToNotify.length; t2++) {
                    var tClean = tradesToNotify[t2];
                    if(tradeInfo[tClean]) {
                      var tCodes = tradeInfo[tClean].codes;
                      var hasMatch = false;
                      for(var tc=0; tc<tCodes.length; tc++) {
                        if(coCodes.indexOf(tCodes[tc]) > -1) { hasMatch = true; break; }
                      }
                      if(hasMatch) applicableTradesForFile.push(tClean);
                    }
                  }
                } catch(coErr) { Logger.log("Error reading CO sheet for routing: " + coErr.message); }
                
              } else {
                applicableTradesForFile = tradesToNotify.filter(function(tr) { return tradeInfo[tr] !== undefined; });
              }
              
              for(var a=0; a<applicableTradesForFile.length; a++) {
                var matchedTrade = applicableTradesForFile[a];
                tradePayloads[matchedTrade].files.push(fileObj);
                if(fileObj.blob) tradePayloads[matchedTrade].attachments.push(fileObj.blob);
              }
            }
            
            for(var tCleanName in tradePayloads) {
              if(tradePayloads[tCleanName].files.length > 0) {
                var payload = tradePayloads[tCleanName];
                var tDisplay = tradeInfo[tCleanName].display;
                var tEmail = tradeInfo[tCleanName].email;
                
                var subject = "New Documents Uploaded: " + projRaw + " - Lot " + phase + "." + lot;
                var body = "Hello " + tDisplay + ",\n\nNew documents have been uploaded to the folder for " + projRaw + " (Lot " + phase + "." + lot + " - " + civic + ").\n\nFiles Included:\n";
                
                for(var fl=0; fl<payload.files.length; fl++) { 
                  body += "- " + payload.files[fl].name + (!payload.files[fl].blob ? " (Link: " + payload.files[fl].url + ")" : "") + "\n"; 
                }
                
                body += "\nThank you,\nGeorgian Build Connect";
                try {
                  GmailApp.sendEmail(tEmail, subject, body, { name: "Georgian Build Connect", attachments: payload.attachments });
                } catch(mailErr) { Logger.log("Failed to email " + tEmail + ": " + mailErr.message); }
              }
            }
            
          }
        } catch(e) { Logger.log("Scan folders error: " + e.message); }
      }
    }
  }
}

function handleAiChat(userQuestion) {
  try {
    var key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (!key) return "Error: GEMINI_API_KEY is not set in Script Properties. Please add it in Project Settings.";

    var dbContext = "";
    var sheetApp = SpreadsheetApp.openById(SHEET_ID);
    
    var poSheet = sheetApp.getSheetByName(PO_SHEET_NAME);
    if (poSheet) {
      var poData = poSheet.getDataRange().getDisplayValues();
      if (poData.length > 1) {
        dbContext += "--- CURRENT PURCHASE ORDERS ---\n";
        for (var i = 1; i < Math.min(poData.length, 50); i++) {
          dbContext += "PO: " + poData[i][2] + " | Project: " + poData[i][1] + " | Lot: " + poData[i][3] + " | Trade: " + poData[i][4] + " | Total: " + poData[i][5] + " | Status: " + poData[i][7] + "\n";
        }
      }
    }

    var driveContext = "";
    var lotsSheet = sheetApp.getSheetByName(LOTS_SHEET_NAME);
    
    if (lotsSheet) {
      var lotsData = lotsSheet.getDataRange().getDisplayValues();
      var targetFolderIds = [];
      
      for (var l = 1; l < lotsData.length; l++) {
        var projName = lotsData[l][0];
        var lotNum = lotsData[l][2];
        var folderUrl = lotsData[l][12];
        
        if (folderUrl && folderUrl.indexOf("http") === 0) {
          if (userQuestion.indexOf(lotNum) > -1 || userQuestion.toLowerCase().indexOf(projName.toLowerCase()) > -1) {
            var match = folderUrl.match(/[-\w]{25,}/);
            if (match && targetFolderIds.indexOf(match[0]) === -1) {
              targetFolderIds.push(match[0]);
            }
          }
        }
      }

      for (var f = 0; f < Math.min(targetFolderIds.length, 2); f++) {
        try {
          var folder = DriveApp.getFolderById(targetFolderIds[f]);
          driveContext += "\n--- CONTENTS OF DRIVE FOLDER: " + folder.getName() + " ---\n";
          
          var files = folder.getFiles();
          while (files.hasNext()) {
            var file = files.next();
            var mime = file.getMimeType();
            
            if (mime === "application/vnd.google-apps.document") {
              var doc = DocumentApp.openById(file.getId());
              driveContext += "[Document: " + file.getName() + " Text Content]:\n" + doc.getBody().getText() + "\n";
            } 
            else if (mime === "application/vnd.google-apps.spreadsheet") {
              var spreadsheet = SpreadsheetApp.openById(file.getId());
              var sheets = spreadsheet.getSheets();
              driveContext += "[Spreadsheet: " + file.getName() + " Data]:\n";
              
              for (var s = 0; s < sheets.length; s++) {
                var sData = sheets[s].getDataRange().getDisplayValues();
                driveContext += "  Tab: " + sheets[s].getName() + ":\n";
                for (var r = 0; r < Math.min(sData.length, 100); r++) { 
                  var rowText = sData[r].filter(String).join(" | ");
                  if (rowText.trim() !== "") {
                    driveContext += "    " + rowText + "\n";
                  }
                }
              }
            }
          }
        } catch(folderErr) {
          Logger.log("Folder read error: " + folderErr.message);
        }
      }
    }

    var systemInstruction = "You are the internal assistant for Georgian Build Connect. " +
                            "You have real-time access to the portal's system records and local project documentation. " +
                            "Answer the user's question clearly, concisely, and accurately using ONLY the provided system logs, " +
                            "schedules, and color selections text context below. Do not mention PDFs. If the information isn't present, " +
                            "politely tell the user you couldn't locate it in the available Docs or Sheets logs.\n\n" +
                            "--- SYSTEM TRACKING DATABASE ---\n" + dbContext + "\n" +
                            "--- ATTACHED DOCUMENTS CONTEXT ---\n" + driveContext;

    var apiResponse = callGeminiApi(key, systemInstruction, userQuestion);
    return apiResponse;

  } catch(e) {
    return "AI Engine Error: " + e.message;
  }
}

function callGeminiApi(apiKey, systemPrompt, userMessage) {
  var url = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent?key=' + apiKey;
  
  var payload = {
    "contents": [
      {
        "role": "user",
        "parts": [
          {"text": systemPrompt + "\n\nUser Question: " + userMessage}
        ]
      }
    ],
    "generationConfig": {
      "temperature": 0.2,
      "maxOutputTokens": 800
    }
  };

  var options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload),
    "muteHttpExceptions": true
  };

  var response = UrlFetchApp.fetch(url, options);
  var responseCode = response.getResponseCode();
  var responseText = response.getContentText();
  var json = JSON.parse(responseText);

  if (responseCode === 200) {
    if (json.candidates && json.candidates[0].content && json.candidates[0].content.parts) {
      return json.candidates[0].content.parts[0].text;
    }
    return "Error: Could not parse response from Gemini structural payload.";
  } else {
    return "Gemini API Error (" + responseCode + "): " + (json.error ? json.error.message : responseText);
  }
}

function forceAuth() {
  UrlFetchApp.fetch("https://google.com");
}

function generateBulkPurchaseOrders(lotsArray, tradesArray, scope, comments, initialStatus, isManual, manualItems) {
  var results = [];
  var successCount = 0;
  var errorCount = 0;

  for (var i = 0; i < lotsArray.length; i++) {
    for (var j = 0; j < tradesArray.length; j++) {
      var lot = lotsArray[i];
      var trade = tradesArray[j];
      
      try {
        var res = generatePurchaseOrder(lot, trade, scope, comments, initialStatus, isManual, manualItems);
        
        if (res === "Success" || res.indexOf("Success") === 0) {
          successCount++;
        } else {
          errorCount++;
          results.push("Error on " + trade + " for lot " + lot + ": " + res);
        }
      } catch(e) {
        errorCount++;
        results.push("Exception on " + trade + " for lot " + lot + ": " + e.message);
      }
    }
  }

  var finalMsg = "Generated " + successCount + " PO(s) successfully.";
  if (errorCount > 0) {
    finalMsg += "\n\nEncountered " + errorCount + " error(s):\n" + results.join("\n");
  }
  
  return finalMsg;
}

function syncScheduleProgress() {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(LOTS_SHEET_NAME);
  var data = sheet.getDataRange().getValues();
  
  if (data[0].length < 16) {
    sheet.getRange(1, 16).setValue("Last Completed Task");
  }

  for (var i = 1; i < data.length; i++) {
    var schedUrl = data[i][13]; 
    var schedStatus = data[i][14]; 
    
    if (schedUrl && schedUrl.indexOf("http") === 0 && schedStatus === "Executed") {
      try {
        var schedIdMatch = schedUrl.match(/[-\w]{25,}/);
        if (schedIdMatch) {
          var schedApp = SpreadsheetApp.openById(schedIdMatch[0]);
          var schedSheet = schedApp.getSheets()[0];
          
          var lastRow = schedSheet.getLastRow();
          if (lastRow > 0) {
            var schedData = schedSheet.getRange("A1:J" + lastRow).getValues();
            var lastCompleted = "";
            
            for (var r = 0; r < schedData.length; r++) {
              var isChecked = schedData[r][9]; 
              
              if (isChecked === true || String(isChecked).toUpperCase() === "TRUE") { 
                var taskName = schedData[r][3]; 
                
                if (taskName && taskName !== "") {
                  lastCompleted = taskName; 
                }
              }
            }
            
            sheet.getRange(i + 1, 16).setValue(lastCompleted);
          }
        }
      } catch(e) {
        Logger.log("Error reading schedule for Lot row " + (i+1) + ": " + e.message);
      }
    }
  }
}

function editLot(formData) {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(LOTS_SHEET_NAME);
  var data = sheet.getDataRange().getValues();
  var rowIndex = -1;
  
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === String(formData.originalProjectName).trim() && 
        String(data[i][1]).trim() === String(formData.originalPhase).trim() && 
        String(data[i][2]).trim() === String(formData.originalLotNum).trim()) {
      rowIndex = i + 1;
      break;
    }
  }
  
  if (rowIndex === -1) return "Error: Lot not found in directory.";

  sheet.getRange(rowIndex, 4).setValue(formData.civicAddress);
  sheet.getRange(rowIndex, 5).setValue(formData.model);
  sheet.getRange(rowIndex, 6).setValue(formData.lotStyle);
  sheet.getRange(rowIndex, 7).setValue(formData.opt1);
  sheet.getRange(rowIndex, 8).setValue(formData.opt2);
  sheet.getRange(rowIndex, 9).setValue(formData.opt3);
  sheet.getRange(rowIndex, 10).setValue(formData.opt4);
  sheet.getRange(rowIndex, 11).setValue(formData.opt5);
  sheet.getRange(rowIndex, 12).setValue(formData.opt6);

  SpreadsheetApp.flush(); 
  var poSheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(PO_SHEET_NAME);
  if (!poSheet) return "Lot updated successfully. No PO tracking sheet found.";
  
  var poData = poSheet.getDataRange().getValues();
  var updatedPOs = 0;
  var poErrors = [];

  for (var p = 1; p < poData.length; p++) {
    var poProject = poData[p][1];
    var poNum = poData[p][2];
    var poLotStr = String(poData[p][3]); 
    var poStatus = poData[p][7]; 
    
    if (String(poProject).trim() === String(formData.originalProjectName).trim() && 
        poLotStr.trim() === String(formData.originalLotNum).trim() && 
        poStatus !== "Void") {
       try {
         var scopeUpdate = "Lot options updated. Please review revised base contract items.";
         var res = editPurchaseOrder(poNum, scopeUpdate, "Auto-generated revision due to Lot option changes.", "", []);
         
         if (res.indexOf("Error") === -1) {
           updatedPOs++;
         } else {
           poErrors.push(poNum + " Failed");
         }
       } catch(e) {
         poErrors.push(poNum + " Error: " + e.message);
       }
    }
  }

  var msg = "Lot updated successfully.";
  if (updatedPOs > 0) msg += " Auto-revised and emailed " + updatedPOs + " existing PO(s).";
  if (poErrors.length > 0) msg += " Errors on some POs: " + poErrors.join(" | ");

  return msg;
}

function initializePOTrackers() {
  var sheetApp = SpreadsheetApp.openById(SHEET_ID);
  
  var settingsSheet = sheetApp.getSheetByName('Settings');
  var mappingSheet = sheetApp.getSheetByName('Master Lot Mapping');
  var activeLotsSheet = sheetApp.getSheetByName(LOTS_SHEET_NAME);
  
  if (!settingsSheet || !mappingSheet) return "Error: Missing required tabs.";

  var activeLotCreationDates = {};
  if (activeLotsSheet) {
    var activeData = activeLotsSheet.getDataRange().getDisplayValues();
  }

  var costCodesRaw = settingsSheet.getRange("J2:J").getDisplayValues();
  var costCodes = [];
  for (var j = 0; j < costCodesRaw.length; j++) {
    var code = String(costCodesRaw[j][0]).trim();
    if (code !== "" && costCodes.indexOf(code) === -1) costCodes.push(code);
  }

  var settingsData = settingsSheet.getDataRange().getDisplayValues();
  var projectTrackers = {};
  for (var s = 1; s < settingsData.length; s++) {
    var pName = cleanText(settingsData[s][0]);
    var tId = settingsData[s][11]; 
    if (pName && tId && tId.length > 20) {
      projectTrackers[pName] = tId;
    }
  }

  var mapData = mappingSheet.getDataRange().getDisplayValues();
  var projectHierarchy = {};
  for (var m = 1; m < mapData.length; m++) {
    var mProj = cleanText(mapData[m][0]);
    var mPhase = String(mapData[m][1]).trim();
    var mLotStr = String(mapData[m][2]).trim();
    
    if (!mProj || !mPhase || !mLotStr) continue;
    
    if (!projectHierarchy[mProj]) projectHierarchy[mProj] = {};
    if (!projectHierarchy[mProj][mPhase]) projectHierarchy[mProj][mPhase] = [];
    
    var lotArray = mLotStr.split(",");
    for (var l = 0; l < lotArray.length; l++) {
      var cleanLot = lotArray[l].trim();
      if (cleanLot !== "" && projectHierarchy[mProj][mPhase].indexOf(cleanLot) === -1) {
        projectHierarchy[mProj][mPhase].push(cleanLot);
      }
    }
  }

  for (var projKey in projectTrackers) {
    var trackerId = projectTrackers[projKey];
    if (projectHierarchy[projKey]) {
      try {
        var trackerApp = SpreadsheetApp.openById(trackerId);
        
        for (var phaseKey in projectHierarchy[projKey]) {
          var phaseTabName = "Phase " + phaseKey;
          var phaseSheet = trackerApp.getSheetByName(phaseTabName);
          
          var retainedDates = {}; 
          if (phaseSheet) {
            var existingData = phaseSheet.getDataRange().getDisplayValues();
            for (var k = 1; k < existingData.length; k++) {
              var eLot = String(existingData[k][0]).trim();
              if (eLot !== "") {
                retainedDates[eLot] = {
                  lotSold: existingData[k][1] || "",
                  ifc: existingData[k][2] || "",
                  csDocs: existingData[k][3] || ""
                };
              }
            }
          } else {
            phaseSheet = trackerApp.insertSheet(phaseTabName);
          }
          
          phaseSheet.clear();
          
          var templateHeaders = ["Lots", "Lot Sold", "IFC Drawings", "CS Docs"];
          var headers = templateHeaders.concat(costCodes);
          
          phaseSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
          phaseSheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
          
          var phaseLots = projectHierarchy[projKey][phaseKey];
          phaseLots.sort(function(a, b) {
             var numA = parseFloat(a), numB = parseFloat(b);
             if (!isNaN(numA) && !isNaN(numB) && numA !== numB) return numA - numB;
             return String(a).localeCompare(String(b));
          });
          
          var matrixRows = [];
          for (var i = 0; i < phaseLots.length; i++) {
            var currentLotStr = phaseLots[i];
            var lotSoldVal = "";
            var ifcVal = "";
            var csDocsVal = "";
            
            if (retainedDates[currentLotStr]) {
              lotSoldVal = retainedDates[currentLotStr].lotSold;
              ifcVal = retainedDates[currentLotStr].ifc;
              csDocsVal = retainedDates[currentLotStr].csDocs;
            }
            
            matrixRows.push([currentLotStr, lotSoldVal, ifcVal, csDocsVal]);
          }
          
          if (matrixRows.length > 0) {
            phaseSheet.getRange(2, 1, matrixRows.length, 4).setValues(matrixRows);
          }
          
          phaseSheet.setFrozenRows(1);
          phaseSheet.setFrozenColumns(1);
          
          applyTrackerFormatting(phaseSheet);
        }
      } catch(e) {
        Logger.log("Could not initialize tracker for " + projKey + ": " + e.message);
      }
    }
  }
}

// ==========================================================================================
// --- USER DIRECTORY & AUTHENTICATION FUNCTIONS ---
// ==========================================================================================
const USERS_SHEET_NAME = 'Users';

function getUsersData() {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(USERS_SHEET_NAME);
  if (!sheet) return [["Error", "Please ensure the tab is named exactly 'Users'."]];
  return sheet.getDataRange().getDisplayValues();
}

function getProjectList() {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Settings');
  if (!sheet) return [];
  var data = sheet.getRange("A2:A").getDisplayValues();
  var projects = [];
  for (var i = 0; i < data.length; i++) {
    var proj = data[i][0].trim();
    if (proj !== "" && projects.indexOf(proj) === -1) {
      projects.push(proj);
    }
  }
  return projects.sort();
}

function addUser(formData) {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(USERS_SHEET_NAME);
  var password = formData.password ? formData.password : generateRandomPassword(10);
  
  sheet.appendRow([
    formData.firstName,
    formData.lastName,
    formData.jobTitle,
    formData.projects,
    formData.email,
    password
  ]);
  return "Success";
}

function deleteUser(email) {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(USERS_SHEET_NAME);
  var data = sheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (data[i][4] === email) { 
      sheet.deleteRow(i + 1); 
      return "Deleted";
    }
  }
  return "Error: User not found.";
}

function authenticateUser(email, password) {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(USERS_SHEET_NAME);
  if (!sheet) return { success: false, message: "Users tab not found." };
  
  var data = sheet.getDataRange().getDisplayValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][4]).trim().toLowerCase() === String(email).trim().toLowerCase()) {
      if (String(data[i][5]).trim() === String(password).trim()) {
        
        var hiddenHeadingsRaw = data[i][7] || "{}"; 
        var hiddenHeadingsObj = {};
        try {
          hiddenHeadingsObj = JSON.parse(hiddenHeadingsRaw);
        } catch(e) {
          hiddenHeadingsObj = {};
        }

        return { 
          success: true, 
          name: data[i][0] + " " + data[i][1], 
          email: data[i][4],          
          projects: data[i][3],       
          defaultTab: data[i][6],     
          hiddenHeadings: hiddenHeadingsObj 
        };
      } else {
        return { success: false, message: "Incorrect password." };
      }
    }
  }
  return { success: false, message: "User email not found." };
}

function updateUserPreferences(email, defaultTab, hiddenHeadings) {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(USERS_SHEET_NAME);
  if (!sheet) return "Error: Users tab not found.";
  
  var data = sheet.getDataRange().getDisplayValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][4]).trim().toLowerCase() === String(email).trim().toLowerCase()) {
      sheet.getRange(i + 1, 7).setValue(defaultTab);
      sheet.getRange(i + 1, 8).setValue(hiddenHeadings);
      return "Success";
    }
  }
  return "Error: User not found.";
}

function changeUserPassword(email, oldPassword, newPassword) {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(USERS_SHEET_NAME);
  if (!sheet) return "Error: Users tab not found.";
  
  var data = sheet.getDataRange().getDisplayValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][4]).trim().toLowerCase() === String(email).trim().toLowerCase()) {
      if (String(data[i][5]).trim() === String(oldPassword).trim()) {
        sheet.getRange(i + 1, 6).setValue(newPassword);
        return "Success";
      } else {
        return "Error: Incorrect current password.";
      }
    }
  }
  return "Error: User not found.";
}