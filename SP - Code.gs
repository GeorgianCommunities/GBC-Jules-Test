// --- IDS & CONSTANTS ---
const SHEET_ID = '1rhHn7mXKpjcGY7D1dSQ3K7aWCTFcjYrF_ls7OHmfp4E'; 
const SERVICE_MASTER_ID = '1L2-Mbq6Uebqih2qhfhf23A4NRG54aoNR4rQyCEBGCRU'; 
const SHEET_NAME = 'Sheet1'; 
const LOTS_SHEET_NAME = 'Lots'; 
const PO_SHEET_NAME = 'Purchase Orders'; 

const PO_TEMPLATE_DOC_ID = '1X7qxE4etQOodCjapsh5HlBQ9nlUkY9vOa_hxu82WwMM'; 
const PO_FOLDERS_MASTER_ID = '14Og9DUsIbvdSutd09WZRqSq0maCYATbC';
const WARRANTY_ROOT_FOLDER_ID = '1BUt0T6XIs3BgR7Red6EE4L2X9sdk7eHO';
const WARRANTY_TEMPLATE_ID = '1loBscI38L9vtywXvl65tohceruLTCF2jo6A1Px0ppxc'; 

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Georgian Build Connect - Service Portal')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function cleanText(str) {
  return String(str || "").trim().toLowerCase().replace(/\s+/g, ' ');
}

function generateRandomPassword(length) {
  var chars = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!@#$%^&*";
  var pass = "";
  for (var i = 0; i < length; i++) { pass += chars.charAt(Math.floor(Math.random() * chars.length)); }
  return pass;
}

function getProjectSettings() {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Settings');
  var data = sheet.getDataRange().getDisplayValues();
  var settings = {};
  for(var i=1; i<data.length; i++) {
    var proj = cleanText(data[i][0]);
    if(proj !== "") {
      settings[proj] = {
        name: data[i][0], billName: data[i][1], billAddress: data[i][2], billPhone: data[i][3],
        billEmail: data[i][4], poPrefix: data[i][5], nextPoNum: parseInt(data[i][6]) || 1,
        taxRate: parseFloat(data[i][7]) || 0, rowIndex: i + 1
      };
    }
  }
  return settings;
}

function getTradeTypesData() {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Settings');
  if(!sheet) return [];
  var data = sheet.getRange("K2:K").getDisplayValues();
  var types = [];
  for (var i = 0; i < data.length; i++) {
    var val = data[i][0].trim();
    if (val !== "" && types.indexOf(val) === -1) types.push(val);
  }
  return types.sort();
}

function getCostCodesData() {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Settings');
  if(!sheet) return [];
  var data = sheet.getRange("J2:J").getDisplayValues();
  var costCodes = [];
  for (var i = 0; i < data.length; i++) { 
    var val = data[i][0].trim();
    if (val !== "" && costCodes.indexOf(val) === -1) costCodes.push(val); 
  }
  return costCodes;
}

function getContractorData() { 
  return SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME).getDataRange().getDisplayValues(); 
}

function addContractor(formData) {
  var newPassword = generateRandomPassword(10);
  SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME).appendRow([
    formData.businessName, formData.contact, formData.email, formData.phone, formData.address, formData.costCodes, formData.activityCode, formData.projects, formData.tradeType, newPassword
  ]);
  return "Success";
}

function deleteContractor(businessName) {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SHEET_NAME);
  var data = sheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 1; i--) { 
    if (data[i][0] === businessName) { sheet.deleteRow(i + 1); return "Deleted"; }
  }
  return "Error: Contractor not found.";
}

function getLotsData(filterObj) {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(LOTS_SHEET_NAME);
  if (!sheet) return [["Error: Please create a tab named 'Lots'."]]; 
  var data = sheet.getDataRange().getDisplayValues();
  if (data.length <= 1) return data;
  
  var headers = data[0]; 
  var filteredData = [headers]; 
  
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
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
  
  var contractorData = sheetApp.getSheetByName(SHEET_NAME).getDataRange().getDisplayValues();
  var trades = [];
  for(var c=1; c<contractorData.length; c++) {
    if(contractorData[c][0]) trades.push(contractorData[c][0]);
  }
  return { projects: projects, lotMap: lotMap, trades: trades };
}

function generatePurchaseOrder(lotNumCombined, trade, scope, comments, initialStatus, manualItems) {
  var sheetApp = SpreadsheetApp.openById(SHEET_ID);
  var lotsData = sheetApp.getSheetByName(LOTS_SHEET_NAME).getDataRange().getValues();
  var lotRecord = null;
  for(var i=1; i<lotsData.length; i++) { 
    if(lotsData[i][0] + " - " + lotsData[i][1] + "." + lotsData[i][2] === lotNumCombined) { lotRecord = lotsData[i]; break; } 
  }
  if(!lotRecord) return "Error: Lot not found.";

  var projectName = lotRecord[0], phase = lotRecord[1], lotNum = lotRecord[2], lotAddress = lotRecord[3], model = lotRecord[4];
  var contractorData = sheetApp.getSheetByName(SHEET_NAME).getDataRange().getDisplayValues();
  var tradeAddress = "Address Not Found", tradePhone = "Phone Not Found", tradeEmail = "Email Not Found";
  
  for(var t=1; t<contractorData.length; t++) {
    if(cleanText(contractorData[t][0]) === cleanText(trade)) {
      tradeEmail = contractorData[t][2] || "N/A"; 
      tradePhone = contractorData[t][3] || "N/A"; 
      tradeAddress = contractorData[t][4] || "N/A"; 
      break;
    }
  }

  var projSettings = getProjectSettings();
  var targetConfig = projSettings[cleanText(projectName)];
  if(!targetConfig) return "Error: Project not found in Settings tab.";

  var generatedPONumber = targetConfig.poPrefix + " - " + ("000" + targetConfig.nextPoNum).slice(-3);
  var subtotal = 0;
  var itemizedList = [];

  for (var mIdx = 0; mIdx < manualItems.length; mIdx++) {
    var manualCost = parseFloat(manualItems[mIdx].cost);
    var codeString = manualItems[mIdx].code ? "[" + manualItems[mIdx].code + "] " : ""; 
    subtotal += manualCost;
    itemizedList.push(codeString + manualItems[mIdx].desc + " .................... $" + manualCost.toFixed(2));
  }

  var taxAmount = subtotal * targetConfig.taxRate, grandTotal = subtotal + taxAmount;
  var pdfUrl = "No Template Provided", pdfErrorMsg = "";

  if(PO_TEMPLATE_DOC_ID.length > 20) { 
    try {
      var templateFile = DriveApp.getFileById(PO_TEMPLATE_DOC_ID);
      var pdfName = generatedPONumber + " - " + trade + " - Lot " + lotNumCombined;
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
      
      var pdfBlob = doc.getAs('application/pdf');
      pdfUrl = projectFolder.createFile(pdfBlob).getUrl();
      tempDoc.setTrashed(true);
      
      sheetApp.getSheetByName('Settings').getRange(targetConfig.rowIndex, 7).setValue(targetConfig.nextPoNum + 1);

    } catch(e) { pdfUrl = "PDF Error: " + e.message; pdfErrorMsg = e.message; }
  }

  var poSheet = sheetApp.getSheetByName(PO_SHEET_NAME);
  var statusToSave = initialStatus ? initialStatus : "Open";
  if(poSheet) { poSheet.appendRow([new Date().toLocaleDateString(), projectName, generatedPONumber, lotNum, trade, "$" + subtotal.toFixed(2), pdfUrl, statusToSave, "$0.00"]); }

  if (pdfErrorMsg !== "") return "Error Generating PDF: " + pdfErrorMsg;
  return "Success";
}

function getPOsData(filterObj) {
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
}

function updatePOStatus(poNum, newStatus) {
  var poSheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(PO_SHEET_NAME);
  var data = poSheet.getDataRange().getValues();
  for(var i=1; i<data.length; i++) {
    if(data[i][2] == poNum) {
      poSheet.getRange(i+1, 8).setValue(newStatus);
      return "Status Updated to " + newStatus + "!";
    }
  }
  return "Error: PO not found.";
}

function getServiceRequestsData(filterObj) {
  return getLotsData(filterObj); 
}

function extractRegex(text, regex) {
  var match = text.match(regex);
  return match ? match[1].trim() : "";
}

function getOrCreateFolder(parentFolder, folderName) {
  var safeName = folderName.replace(/'/g, "\\'");
  var folders = parentFolder.searchFolders("title = '" + safeName + "' and trashed = false");
  if (folders.hasNext()) { return folders.next(); }
  return parentFolder.createFolder(folderName);
}

function parseClaimItems(text, formType) {
  var items = [];
  var blockMatch = text.match(/(?:Claim.Items|Description of Major Structural Defect|The following is a list)[\s\S]*?(?=The items specified|Submitted online|Homeowner's Signature|$)/i);
  if (!blockMatch) return items;
  
  var block = blockMatch[0];
  var lines = block.split('\n');
  var currentItem = null;
  var ignoreHeaders = [
    "item #", "location", "floor/level", "room/area", "description", 
    "item/defect area", "defects", "claim items", "the following is a list of warranty claim items for my home.", "item"
  ];

  for (var i = 0; i < lines.length; i++) {
    var line = lines[i].trim();
    if (!line) continue;
    var lowerLine = line.toLowerCase();
    if (ignoreHeaders.indexOf(lowerLine) > -1) continue;
    if (lowerLine.indexOf("the following is a list") > -1) continue;

    var isNewItem = false;
    var num = "";
    var rest = "";

    var itemStartMatch = line.match(/^(\d{1,3})(?:\s*\|\s*|\s+|$)(.*)/);
    if (itemStartMatch && !isNaN(parseInt(itemStartMatch[1]))) {
        num = itemStartMatch[1];
        rest = itemStartMatch[2];
        if (line === num || line.match(/^\d{1,3}\s*\|/) || line.match(/^\d{1,3}\s+[A-Z]{3,}/)) {
            isNewItem = true;
        } else if (items.length === 0 && num === "1" && currentItem === null) {
            isNewItem = true; 
        }
    }

    if (isNewItem) {
        if (currentItem) items.push(currentItem);
        currentItem = { num: num, rawLines: [] };
        if (rest && rest.replace(/\|/g, '').trim() !== "") { currentItem.rawLines.push(rest); }
        continue;
    }
    if (currentItem) { currentItem.rawLines.push(line); }
  }
  if (currentItem) items.push(currentItem);
  
  var finalItems = [];
  for (var j = 0; j < items.length; j++) {
     var item = items[j];
     var cleanParts = [];
     var hasPipes = item.rawLines.join(" ").indexOf('|') > -1;
     
     if (hasPipes) {
         var combined = item.rawLines.join(" ");
         var parts = combined.split('|');
         for (var p = 0; p < parts.length; p++) {
             var cp = parts[p].trim();
             if (cp) cleanParts.push(cp);
         }
     } else {
         for (var p = 0; p < item.rawLines.length; p++) {
             var cp = item.rawLines[p].trim();
             if (cp) cleanParts.push(cp);
         }
     }
     
     cleanParts = cleanParts.filter(function(cp) {
         var lower = cp.toLowerCase();
         for(var h=0; h < ignoreHeaders.length; h++) {
             if (lower.indexOf(ignoreHeaders[h]) > -1 && cp.length < 25) return false;
         }
         return true;
     });
     
     var location = "", room = "", defect = "", desc = "";
     if ((formType && formType.indexOf("Major Structural") > -1) || cleanParts.length <= 2) {
         desc = cleanParts.join(" ").replace(/\s+/g, ' ');
     } else {
         location = cleanParts[0] || "";
         room = cleanParts[1] || "";
         defect = cleanParts[2] || "";
         desc = cleanParts.slice(3).join(" ").replace(/\s+/g, ' ') || "";
     }
     
     finalItems.push({
        num: item.num,
        location: location.replace(/\s+/g, ' ').substring(0, 70), 
        room: room.replace(/\s+/g, ' ').substring(0, 70),
        defect: defect.replace(/\s+/g, ' ').substring(0, 70),
        desc: desc
     });
  }

  if (finalItems.length === 0 && block.length > 50) {
      finalItems.push({
          num: "*",
          location: "Auto-Parse Failed",
          room: "Raw Text Extracted",
          defect: "See Description",
          desc: block.replace(/\n/g, ' ').replace(/\s+/g, ' ').substring(0, 3000)
      });
  }
  return finalItems;
}

function processWarrantyDocument(base64Data, fileName, mimeType, caseId) {
  try {
    var blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, fileName);
    var resource = { name: fileName, mimeType: MimeType.GOOGLE_DOCS }; 
    var tempFile = Drive.Files.create(resource, blob);
    var doc = DocumentApp.openById(tempFile.id);
    var textContent = doc.getBody().getText();
    
    var rtfBlob = null;
    var ext = "";
    if (fileName.toLowerCase().indexOf('.pdf') > -1) {
      ext = ".pdf";
      try {
        var url = "https://docs.google.com/document/d/" + tempFile.id + "/export?format=rtf";
        var options = { headers: { "Authorization": "Bearer " + ScriptApp.getOAuthToken() }, muteHttpExceptions: true };
        var rtfResponse = UrlFetchApp.fetch(url, options);
        if (rtfResponse.getResponseCode() === 200) {
           rtfBlob = rtfResponse.getBlob().setName(fileName.replace(/\.pdf$/i, '.rtf'));
        }
      } catch(e) {}
    } else if (fileName.lastIndexOf('.') > -1) {
      ext = fileName.substring(fileName.lastIndexOf('.'));
    }
    DriveApp.getFileById(tempFile.id).setTrashed(true);

    var vendorNum = extractRegex(textContent, /\b(B\d{5})\b/i); 
    var enrolmentNum = extractRegex(textContent, /\b(H\d{7})\b/i); 
    var formType = extractRegex(textContent, /(Initial Form|Mid-Year Form|Second-Year Form|30-Day Form|Year-End Form|1 Year Review|Major Structural Defect)/i) || "Warranty Form";
    
    var phase = "Unknown";
    var lot = "Unknown";
    var genericPattern = /(?<![-\/\d])\b([A-Za-z0-9]{1,3})\s*-\s*(\d{1,4}[A-Za-z]?)\b(?![-\/\d])/gi;
    var match;
    var validMatches = [];
    while ((match = genericPattern.exec(textContent)) !== null) { validMatches.push(match); }
    
    if (validMatches.length > 0) {
        var bestMatch = null;
        for (var i = 0; i < validMatches.length; i++) {
            if (/[A-Za-z]/.test(validMatches[i][1])) { bestMatch = validMatches[i]; break; }
        }
        if (!bestMatch) {
            for (var i = 0; i < validMatches.length; i++) {
                var pNum = parseInt(validMatches[i][1], 10);
                var lNum = parseInt(validMatches[i][2], 10);
                if (pNum > 31 || lNum > 31) { bestMatch = validMatches[i]; break; }
            }
        }
        if (!bestMatch) bestMatch = validMatches[0];
        phase = bestMatch[1].toUpperCase().replace(/^0+(?=\d)/, '');
        lot = bestMatch[2].replace(/^0+/, '');
    }

    var homeownerName = "Review PDF for Name";
    var nameBeforeMatch = textContent.match(/([^\n]+)\n[\s|]*Homeowner Name/i);
    if (nameBeforeMatch) {
        var candidate = nameBeforeMatch[1].replace(/\|/g, '').trim();
        if (candidate.length > 2 && candidate.toLowerCase().indexOf("daytime") === -1 && candidate.toLowerCase().indexOf("phone") === -1) {
            homeownerName = candidate;
        }
    }
    if (homeownerName === "Review PDF for Name") {
        var nameAfterMatch = textContent.match(/Homeowner Name(?:(?:\(s\))?[:\s]*)\n([^\n]+)/i);
        if (nameAfterMatch) {
            var candidate = nameAfterMatch[1].replace(/\|/g, '').trim();
            if (candidate.length > 2 && candidate.toLowerCase().indexOf("daytime") === -1 && candidate.toLowerCase().indexOf("phone") === -1) {
                homeownerName = candidate;
            }
        }
    }
    if (homeownerName === "Review PDF for Name") {
        var hoBlock = textContent.match(/Contact Information of Homeowner[\s\S]{1,150}/i);
        if (hoBlock) {
            var lines = hoBlock[0].split('\n');
            for (var i = 1; i < lines.length; i++) {
                var line = lines[i].replace(/\|/g, '').trim();
                var lowerLine = line.toLowerCase();
                if (line.length < 2) continue;
                if (lowerLine.indexOf("homeowner") > -1) continue;
                if (lowerLine.indexOf("name") > -1) continue;
                if (lowerLine.indexOf("contact") > -1) continue;
                if (lowerLine.indexOf("@") > -1) continue; 
                if (/\d{3}[-.\s]?\d{4}/.test(line)) continue; 
                if (lowerLine.indexOf("address") > -1) continue;
                if (lowerLine.indexOf("daytime") > -1) continue;
                if (lowerLine.indexOf("number") > -1) continue; 
                if (lowerLine.indexOf("phone") > -1) continue;
                homeownerName = line;
                break;
            }
        }
    }

    var emails = (textContent.match(/[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}/g) || []).filter(function(v,i,a){ return a.indexOf(v)===i; }).join(", ");
    var phones = (textContent.match(/(?:\+1\s)?\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4}/g) || []).filter(function(v,i,a){ return a.indexOf(v)===i; }).join(", ");

    var projectName = "Unknown Project";
    var ssMaster = SpreadsheetApp.openById(SHEET_ID);
    var settingsData = ssMaster.getSheetByName('Settings').getDataRange().getValues();
    for (var i = 1; i < settingsData.length; i++) {
      if (settingsData[i][13] == vendorNum && vendorNum !== "") { projectName = settingsData[i][0]; break; }
    }

    var rootFolder = DriveApp.getFolderById(WARRANTY_ROOT_FOLDER_ID);
    var projectFolder = getOrCreateFolder(rootFolder, projectName);
    var phaseFolder = getOrCreateFolder(projectFolder, "Phase " + phase);
    var lotFolder = getOrCreateFolder(phaseFolder, "Lot " + lot);

    var formTypeSubfolder = getOrCreateFolder(lotFolder, formType);

    var newFileName = phase + "-" + lot + "-Warranty for " + formType + ext;
    var finalFile = formTypeSubfolder.createFile(blob);
    finalFile.setName(newFileName);
    
    var folderUrl = lotFolder.getUrl(); 
    if (rtfBlob) {
        var rtfFile = formTypeSubfolder.createFile(rtfBlob);
        rtfFile.setName(phase + "-" + lot + "-Warranty for " + formType + ".rtf");
    }

    var sheetName = "Lot " + lot + " - Warranty File";
    var existingFiles = lotFolder.searchFiles("title = '" + sheetName.replace(/'/g, "\\'") + "' and mimeType = 'application/vnd.google-apps.spreadsheet'");
    var ss;
    
    if (existingFiles.hasNext()) {
      ss = SpreadsheetApp.openById(existingFiles.next().getId());
    } else {
      if (WARRANTY_TEMPLATE_ID && WARRANTY_TEMPLATE_ID.length > 20) {
        try {
          var templateFile = DriveApp.getFileById(WARRANTY_TEMPLATE_ID);
          var newFile = templateFile.makeCopy(sheetName, lotFolder);
          ss = SpreadsheetApp.openById(newFile.getId());
          ss.getSheets()[0].clear();
        } catch(e) {
          ss = SpreadsheetApp.create(sheetName);
          DriveApp.getFileById(ss.getId()).moveTo(lotFolder);
        }
      } else {
        ss = SpreadsheetApp.create(sheetName);
        DriveApp.getFileById(ss.getId()).moveTo(lotFolder);
      }
      
      var firstTab = ss.getSheets()[0];
      firstTab.setName("Homeowner Info");
      firstTab.appendRow(["Homeowner Contact Information", ""]);
      firstTab.appendRow(["Name(s):", homeownerName]);
      firstTab.appendRow(["Email(s):", emails]);
      firstTab.appendRow(["Phone(s):", phones]);
      firstTab.appendRow(["", ""]);
      firstTab.appendRow(["Project Details", ""]);
      firstTab.appendRow(["Project Name:", projectName]);
      firstTab.appendRow(["Phase:", phase]);
      firstTab.appendRow(["Lot Number:", lot]);
      firstTab.appendRow(["Enrolment #:", enrolmentNum]);
      firstTab.appendRow(["Vendor #:", vendorNum]);
      
      firstTab.getRange("A1:B1").setFontWeight("bold");
      firstTab.getRange("A6:B6").setFontWeight("bold");
      firstTab.getRange("A2:A4").setFontWeight("bold");
      firstTab.getRange("A7:A11").setFontWeight("bold");
      firstTab.setColumnWidth(1, 150);
      firstTab.setColumnWidth(2, 300);
    }

    var formTab = ss.getSheetByName(formType);
    if (!formTab) {
      formTab = ss.insertSheet(formType);
      
      formTab.appendRow(["Item #", "Location", "Room/Area", "Item/Defect Area", "Description", "Additional Information", "Assigned Trade", "Notified Status", "Scheduled Date", "Photo Link", "Warranted", "Status", "", "Case ID #"]);
      formTab.getRange("A1:N1").setFontWeight("bold");

      var statusList = ["Unreviewed", "Reviewed", "Assigned", "Completed", "Approved"];
      var statusRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(statusList, true)
        .setAllowInvalid(true)
        .build();
      formTab.getRange("L2:L100").setDataValidation(statusRule);

      var warrantRule = SpreadsheetApp.newDataValidation()
        .requireValueInList(["Warrantable", "Not warrantable"], true)
        .setAllowInvalid(true)
        .build();
      formTab.getRange("K2:K100").setDataValidation(warrantRule);

      var poSheet = ssMaster.getSheetByName(PO_SHEET_NAME);
      var contractorSheet = ssMaster.getSheetByName(SHEET_NAME);
      var poTrades = [];
      var otherTrades = [];
      var finalDropdownList = [];
      var normLot = cleanText(lot).replace(/^0+/, '');

      if (poSheet) {
        var poData = poSheet.getDataRange().getValues();
        for (var r = 1; r < poData.length; r++) {
          var poProject = cleanText(poData[r][1]);
          var poLot = cleanText(poData[r][3]).replace(/^0+/, '');
          if (poProject === cleanText(projectName) && poLot === normLot) {
            var tradeName = poData[r][4];
            if (tradeName && poTrades.indexOf(tradeName) === -1) {
              poTrades.push(tradeName);
            }
          }
        }
      }

      if (contractorSheet) {
        var contractorData = contractorSheet.getDataRange().getValues();
        for (var c = 1; c < contractorData.length; c++) {
          var dirTradeName = contractorData[c][0];
          if (dirTradeName && poTrades.indexOf(dirTradeName) === -1 && otherTrades.indexOf(dirTradeName) === -1) {
            otherTrades.push(dirTradeName);
          }
        }
      }

      if (poTrades.length > 0) {
        finalDropdownList = finalDropdownList.concat(poTrades);
      }
      if (otherTrades.length > 0) {
        if (finalDropdownList.length > 0) {
           finalDropdownList.push("------------------------"); 
        }
        finalDropdownList = finalDropdownList.concat(otherTrades.sort());
      }

      if (finalDropdownList.length > 0) {
        var rule = SpreadsheetApp.newDataValidation()
          .requireValueInList(finalDropdownList, true)
          .setAllowInvalid(true) 
          .build();
        formTab.getRange("G2:G100").setDataValidation(rule);
      }
    } else {
      formTab.getRange(1, 14).setValue("Case ID #").setFontWeight("bold");
    }

    var claims = parseClaimItems(textContent, formType);
    if (claims.length === 0) {
      formTab.appendRow(["*", "See attached Document", "Data Extraction requires Document AI", "Manual entry required", "", "", "", "", "", "", "", "Unreviewed", "", caseId || ""]);
    } else {
      for (var c = 0; c < claims.length; c++) {
        formTab.appendRow([claims[c].num, claims[c].location, claims[c].room, claims[c].defect, claims[c].desc, "", "", "", "", "", "", "Unreviewed", "", caseId || ""]);
      }
    }

    var lotsSheet = ssMaster.getSheetByName(LOTS_SHEET_NAME);
    var lotsData = lotsSheet.getDataRange().getValues();
    var normExtractedPhase = cleanText(phase).replace(/^0+/, '');
    var normExtractedLot = cleanText(lot).replace(/^0+/, '');
    
    for (var row = 1; row < lotsData.length; row++) {
      var sheetProject = cleanText(lotsData[row][0]);
      var sheetPhase = cleanText(lotsData[row][1]).replace(/^0+/, '');
      var sheetLot = cleanText(lotsData[row][2]).replace(/^0+/, '');

      if (sheetProject === cleanText(projectName) && sheetPhase === normExtractedPhase && sheetLot === normExtractedLot) {
        lotsSheet.getRange(row + 1, 17).setValue(folderUrl); 
        break;
      }
    }

    return "Success! Document saved and processed. Service Folder linked to Lot " + lot + " in directory.";

  } catch (e) {
    return "Error Processing File: " + e.message;
  }
}

function uploadServicePhotosBatch(fileDataArray, project, lotCombined, formType) {
  try {
    var splitIdx = lotCombined.lastIndexOf(' - ');
    var phaseLot = lotCombined.substring(splitIdx + 3).trim().split('.');
    var phase = phaseLot[0].replace(/^0+(?=\d)/, ''); 
    var lot = phaseLot[1].replace(/^0+(?=\d)/, '');   
    
    var rootFolder = DriveApp.getFolderById(WARRANTY_ROOT_FOLDER_ID);
    var projectFolder = getOrCreateFolder(rootFolder, project);
    var phaseFolder = getOrCreateFolder(projectFolder, "Phase " + phase);
    var lotFolder = getOrCreateFolder(phaseFolder, "Lot " + lot);
    var subFolder = getOrCreateFolder(lotFolder, formType);
    
    var groupedFiles = {};
    for (var i = 0; i < fileDataArray.length; i++) {
      var f = fileDataArray[i];
      if (!groupedFiles[f.itemNum]) groupedFiles[f.itemNum] = [];
      groupedFiles[f.itemNum].push(f);
    }
    
    var sheetName = "Lot " + lot + " - Warranty File";
    var existingFiles = lotFolder.searchFiles("title = '" + sheetName.replace(/'/g, "\\'") + "' and mimeType = 'application/vnd.google-apps.spreadsheet'");
    if (!existingFiles.hasNext()) return "Error: Warranty spreadsheet not found.";
    
    var ss = SpreadsheetApp.openById(existingFiles.next().getId());
    var formTab = ss.getSheetByName(formType);
    if (!formTab) return "Error: Form tab not found.";
    
    var data = formTab.getDataRange().getValues();
    var headers = data[0].map(function(h) { return String(h).toLowerCase().trim(); });
    var iPhoto = headers.indexOf("photo link");
    
    if (iPhoto === -1) {
       formTab.getRange(1, data[0].length + 1).setValue("Photo Link").setFontWeight("bold");
       iPhoto = data[0].length;
    }
    
    var itemsUpdated = 0;
    
    for (var iNum in groupedFiles) {
      var startingPicNum = 1;
      var rowIndex = -1;
      var currentPhotos = "";
      
      for (var r = 1; r < data.length; r++) {
        if (String(data[r][0]).trim() === String(iNum).trim()) {
          rowIndex = r + 1;
          currentPhotos = formTab.getRange(rowIndex, iPhoto + 1).getValue();
          
          if (currentPhotos && String(currentPhotos).trim() !== "") {
            var links = String(currentPhotos).split(/\n+/);
            var validLinks = 0;
            for (var l = 0; l < links.length; l++) {
              if (links[l].indexOf("http") > -1) validLinks++;
            }
            startingPicNum = validLinks + 1;
          }
          break; 
        }
      }
      
      if (rowIndex === -1) continue; 

      var savedUrls = [];
      for (var j = 0; j < groupedFiles[iNum].length; j++) {
        var fileData = groupedFiles[iNum][j];
        var ext = "";
        if (fileData.fileName.lastIndexOf('.') > -1) {
          ext = fileData.fileName.substring(fileData.fileName.lastIndexOf('.'));
        }
        
        var picNum = startingPicNum + j; 
        var newFileName = phase + "-" + lot + "-Warranty for " + formType + "-Item " + iNum + "-Pic " + picNum + ext;
        
        var blob = Utilities.newBlob(Utilities.base64Decode(fileData.base64Data), fileData.mimeType, newFileName);
        var savedFile = subFolder.createFile(blob);
        savedUrls.push(savedFile.getUrl());
      }
      
      var newUrlsString = savedUrls.join("\n\n");
      var updatedPhotos = currentPhotos ? currentPhotos + "\n\n" + newUrlsString : newUrlsString;
      formTab.getRange(rowIndex, iPhoto + 1).setValue(updatedPhotos);
      itemsUpdated++;
    }
    
    return "Success! " + fileDataArray.length + " photo(s) mapped to " + itemsUpdated + " item(s).";
  } catch (e) {
    return "Error Uploading Photos: " + e.message;
  }
}

function getLotWarrantyMetadata(lotCombined) {
  try {
    var splitIdx = lotCombined.lastIndexOf(' - ');
    var project = lotCombined.substring(0, splitIdx).trim();
    var phaseLot = lotCombined.substring(splitIdx + 3).trim().split('.');
    var phase = phaseLot[0].replace(/^0+(?=\d)/, ''); 
    var lot = phaseLot[1].replace(/^0+(?=\d)/, '');

    var rootFolder = DriveApp.getFolderById(WARRANTY_ROOT_FOLDER_ID);
    
    var pFolders = rootFolder.searchFolders("title = '" + project.replace(/'/g, "\\'") + "' and trashed = false");
    if (!pFolders.hasNext()) return { error: "Project folder not found." };
    var projectFolder = pFolders.next();

    var phFolders = projectFolder.searchFolders("title = 'Phase " + phase.replace(/'/g, "\\'") + "' and trashed = false");
    if (!phFolders.hasNext()) return { error: "Phase folder not found." };
    var phaseFolder = phFolders.next();

    var lFolders = phaseFolder.searchFolders("title = 'Lot " + lot.replace(/'/g, "\\'") + "' and trashed = false");
    if (!lFolders.hasNext()) return { error: "Lot folder not found." };
    var lotFolder = lFolders.next();
    
    var sheetName = "Lot " + lot + " - Warranty File";
    var files = lotFolder.searchFiles("title = '" + sheetName.replace(/'/g, "\\'") + "' and mimeType = 'application/vnd.google-apps.spreadsheet'");
    
    if (!files.hasNext()) return { error: "No warranty spreadsheet found for this lot. Has it been generated?" };
    
    var ss = SpreadsheetApp.openById(files.next().getId());
    var sheets = ss.getSheets();
    var metadata = { forms: {} };
    
    for (var i = 0; i < sheets.length; i++) {
      var sName = sheets[i].getName();
      if (sName === "Homeowner Info") continue;
      
      var sheetData = sheets[i].getDataRange().getDisplayValues();
      if (sheetData.length <= 1) continue;
      
      var hdrs = sheetData[0].map(function(h) { return String(h).toLowerCase().trim(); });
      var iItem = hdrs.indexOf("item #");
      if (iItem === -1) iItem = 0; 
      
      var items = [];
      for (var r = 1; r < sheetData.length; r++) {
         var val = String(sheetData[r][iItem]).trim();
         if (val !== "" && val.toLowerCase().indexOf("item") === -1 && val !== "*") {
             items.push(val);
         }
      }
      
      if (items.length > 0) {
         metadata.forms[sName] = items;
      }
    }
    
    if (Object.keys(metadata.forms).length === 0) {
       return { error: "No valid items found inside the spreadsheet." };
    }
    
    return metadata;
  } catch (e) {
    return { error: "System Error: " + e.message };
  }
}

function syncServiceItemsToMaster() {
  try {
    var targetSS = SpreadsheetApp.openById(SERVICE_MASTER_ID);
    var targetSheet = targetSS.getSheets()[0]; 
    
    var masterHeaders = targetSheet.getRange(1, 1, 1, targetSheet.getMaxColumns()).getDisplayValues()[0];
    if (masterHeaders[17] !== "Case ID #") {
       targetSheet.getRange(1, 18).setValue("Case ID #").setFontWeight("bold");
    }

    var lastRow = targetSheet.getLastRow();
    if (lastRow > 1) {
      targetSheet.getRange(2, 1, lastRow - 1, targetSheet.getLastColumn()).clearContent();
    }
    
    var files = DriveApp.searchFiles("title contains 'Warranty' and mimeType = 'application/vnd.google-apps.spreadsheet' and trashed = false");
    var allData = [];
    var filesFound = 0;
    var tabsProcessed = 0;
    var errorLog = "";
    
    while (files.hasNext()) {
      var file = files.next();
      filesFound++;
      try {
        var ss = SpreadsheetApp.openById(file.getId());
        var homeTab = ss.getSheetByName("Homeowner Info");
        
        if (!homeTab) continue; 
        
        var homeData = homeTab.getDataRange().getValues();
        var projName = "Unknown", phase = "Unknown", lotNum = "Unknown";
        
        for (var h = 0; h < homeData.length; h++) {
          var label = String(homeData[h][0]).trim();
          if (label === "Project Name:") projName = String(homeData[h][1]).trim();
          else if (label === "Phase:") phase = String(homeData[h][1]).replace(/^0+(?=\d)/, '').trim();
          else if (label === "Lot Number:") lotNum = String(homeData[h][1]).replace(/^0+(?=\d)/, '').trim();
        }
        
        var sheets = ss.getSheets();
        for (var s = 0; s < sheets.length; s++) {
          var formName = sheets[s].getName(); 
          if (formName === "Homeowner Info") continue; 
          
          tabsProcessed++;
          var data = sheets[s].getDataRange().getDisplayValues();
          if (data.length <= 1) continue;
          
          var hdrs = data[0].map(function(x){ return String(x).toLowerCase().trim(); });
          var iItem = hdrs.indexOf("item #");
          var iLoc = hdrs.indexOf("location");
          var iRoom = hdrs.indexOf("room/area");
          var iDef = hdrs.indexOf("item/defect area");
          var iDesc = hdrs.indexOf("description");
          var iAdd = hdrs.indexOf("additional information") > -1 ? hdrs.indexOf("additional information") : hdrs.indexOf("trade description");
          var iTrade = hdrs.indexOf("assigned trade");
          var iNotif = hdrs.indexOf("notified status");
          var iSched = hdrs.indexOf("scheduled date");
          var iPhoto = hdrs.indexOf("photo link");
          var iWarr = hdrs.indexOf("warranted");
          var iStat = hdrs.indexOf("status");
          
          var iCase = hdrs.indexOf("case id #");
          if (iCase === -1) iCase = 13; 

          for (var r = 1; r < data.length; r++) {
            var itemNum = iItem > -1 ? data[r][iItem] : "";
            if (!itemNum || itemNum === "" || String(itemNum).toLowerCase().indexOf("item") > -1) continue; 
            
            var location = iLoc > -1 ? data[r][iLoc] : "";
            var roomArea = iRoom > -1 ? data[r][iRoom] : "";
            var defectArea = iDef > -1 ? data[r][iDef] : "";
            var desc = iDesc > -1 ? data[r][iDesc] : "";
            var tradeDesc = iAdd > -1 ? data[r][iAdd] : ""; 
            var trade = iTrade > -1 ? data[r][iTrade] : "";
            var notified = iNotif > -1 ? data[r][iNotif] : "";
            var sched = iSched > -1 ? data[r][iSched] : "";
            var photo = iPhoto > -1 ? data[r][iPhoto] : "";
            var warr = iWarr > -1 ? data[r][iWarr] : "";
            var status = iStat > -1 ? data[r][iStat] : "";
            var caseId = iCase > -1 ? data[r][iCase] : ""; 
            
            if (!status || String(status).trim() === "") {
               if (trade !== "" && notified !== "") status = "Assigned";
               else if (tradeDesc !== "") status = "Reviewed";
               else status = "Unreviewed";
            }
            
            allData.push([
              new Date().toLocaleDateString(), 
              projName, phase, lotNum, formName, itemNum, location, roomArea, defectArea, desc, tradeDesc, trade, notified, sched, photo, warr, status, caseId                            
            ]);
          }
        }
      } catch(e) {
        errorLog += "[" + file.getName() + ": " + e.message + "] ";
      }
    }
    
    if (allData.length > 0) {
      targetSheet.getRange(2, 1, allData.length, allData[0].length).setValues(allData);
      return "Sync Complete! Found " + filesFound + " warranty files and compiled " + allData.length + " items.";
    } else {
      return "Diagnostic Result: Found " + filesFound + " warranty files, but compiled 0 items. Errors: " + (errorLog ? errorLog : "None. Please ensure items have an Item # listed in Column A.");
    }
  } catch (e) {
    return "Critical Error Syncing: " + e.message;
  }
}

function getServiceOrderFilterOptions(selectedFilters) {
  try {
    var sheet = SpreadsheetApp.openById(SERVICE_MASTER_ID).getSheets()[0];
    if (!sheet) throw new Error("Could not access Master Service Items sheet.");
    
    var data = sheet.getDataRange().getDisplayValues();
    var res = { projects: [], phases: [], lots: [], forms: [], trades: [], statuses: [] };
    
    selectedFilters = selectedFilters || {};
    
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var proj = row[1] ? row[1].trim() : "";
      var phase = row[2] ? row[2].trim() : "";
      var lot = row[3] ? row[3].trim() : "";
      var form = row[4] ? row[4].trim() : "";
      var trade = row[11] ? row[11].trim() : "";
      var status = row[16] ? row[16].trim() : "";

      if (proj && res.projects.indexOf(proj) === -1) res.projects.push(proj);

      var matchesProject = !selectedFilters.project || proj === selectedFilters.project;
      var matchesPhase = !selectedFilters.phase || phase === selectedFilters.phase;
      var matchesLot = !selectedFilters.lot || lot === selectedFilters.lot;

      if (matchesProject) {
        if (phase && res.phases.indexOf(phase) === -1) res.phases.push(phase);
        
        if (matchesPhase) {
          if (lot && res.lots.indexOf(lot) === -1) res.lots.push(lot);
          
          if (matchesLot) {
            if (form && res.forms.indexOf(form) === -1) res.forms.push(form);
            if (status && res.statuses.indexOf(status) === -1) res.statuses.push(status);
            
            if (trade) {
              var tSplit = trade.split(',');
              for(var t = 0; t < tSplit.length; t++) {
                var tClean = tSplit[t].trim();
                if (tClean && res.trades.indexOf(tClean) === -1) res.trades.push(tClean);
              }
            }
          }
        }
      }
    }
    return res;
  } catch (e) {
    return { error: e.message };
  }
}

function getFilteredServiceOrders(filterObj) {
  var sheet = SpreadsheetApp.openById(SERVICE_MASTER_ID).getSheets()[0];
  var data = sheet.getDataRange().getDisplayValues();
  if (data.length <= 1) return [["Status"], ["No service items found in the database. Try running a Sync."]];
  
  var headers = data[0];
  var filtered = [headers];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var isMatch = true;
    
    if (filterObj.project && String(row[1]).trim() !== filterObj.project) isMatch = false;
    if (filterObj.phase && String(row[2]).trim() !== filterObj.phase) isMatch = false;
    if (filterObj.lot && String(row[3]).trim() !== filterObj.lot) isMatch = false;
    if (filterObj.form && String(row[4]).trim() !== filterObj.form) isMatch = false;
    if (filterObj.itemNum && String(row[5]).toLowerCase().trim() !== String(filterObj.itemNum).toLowerCase().trim()) isMatch = false;
    if (filterObj.trade && String(row[11]).indexOf(filterObj.trade) === -1) isMatch = false;
    if (filterObj.status && String(row[16]).trim() !== filterObj.status) isMatch = false; 
    
    if (isMatch) filtered.push(row);
  }
  
  return filtered.length > 1 ? filtered : [headers, ["Status", "No items match your specific filters.", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""]];
}

function updateServiceItemDetails(project, phase, lot, formName, itemNum, targetTrade, newStatus, scheduledDate, warrantedValue, rejectionReason, rejectionFile, rescheduleReason) {
  try {
    var masterSS = SpreadsheetApp.openById(SERVICE_MASTER_ID);
    var masterSheet = masterSS.getSheets()[0];
    var mData = masterSheet.getDataRange().getValues();
    
    var cProj = cleanText(project), cPhase = cleanText(phase), cLot = cleanText(lot), cForm = cleanText(formName), cItem = cleanText(itemNum), cTrade = cleanText(targetTrade);
    var assignedTrade = "";
    var finalSchedDate = scheduledDate; 
    
    var masterSheetMsg = "Item not found on Master Sheet.";
    var lotSheetMsg = "Lot Sheet file not found.";

    var savedFileUrl = "";
    var emailBlob = null;
    var rootFolder = DriveApp.getFolderById(WARRANTY_ROOT_FOLDER_ID);
    var projectFolder = getOrCreateFolder(rootFolder, project); 
    var phaseFolder = getOrCreateFolder(projectFolder, "Phase " + phase);
    var lotFolder = getOrCreateFolder(phaseFolder, "Lot " + lot);

    if (rejectionFile && rejectionFile.base64) {
        try {
            var blob = Utilities.newBlob(Utilities.base64Decode(rejectionFile.base64), rejectionFile.mimeType, rejectionFile.fileName);
            emailBlob = blob; 
            var formFolder = getOrCreateFolder(lotFolder, formName);
            var savedFile = formFolder.createFile(blob);
            savedFileUrl = savedFile.getUrl();
        } catch(e) {
            Logger.log("File save error: " + e.message);
        }
    }
    
    for (var i = 1; i < mData.length; i++) {
      if (cleanText(mData[i][1]) === cProj && cleanText(mData[i][2]) === cPhase && cleanText(mData[i][3]) === cLot && cleanText(mData[i][4]) === cForm && cleanText(mData[i][5]) === cItem && cleanText(mData[i][11]) === cTrade) {
         assignedTrade = mData[i][11]; 
         
         var currentSchedCell = String(mData[i][13]).trim(); 
         var primaryOldDate = currentSchedCell.split('\n')[0].trim();
         var historyStr = currentSchedCell.substring(primaryOldDate.length).trim();
         
         if (scheduledDate !== "" && scheduledDate !== primaryOldDate) {
             if (primaryOldDate !== "") {
                 var todayStr = new Date().toLocaleDateString();
                 var rsn = (rescheduleReason && rescheduleReason.trim() !== "") ? rescheduleReason : "No reason provided";
                 finalSchedDate = scheduledDate + "\n[Rescheduled " + todayStr + " - Reason: " + rsn + " - Prev: " + primaryOldDate + "]";
                 if (historyStr !== "") finalSchedDate += "\n" + historyStr;
             }
         } else if (scheduledDate === primaryOldDate || (scheduledDate === "" && primaryOldDate === "")) {
             finalSchedDate = currentSchedCell; 
         } else if (scheduledDate === "" && primaryOldDate !== "") {
             var todayStr = new Date().toLocaleDateString();
             var rsn = (rescheduleReason && rescheduleReason.trim() !== "") ? rescheduleReason : "Removed manually";
             finalSchedDate = "[Removed " + todayStr + " - Reason: " + rsn + " - Prev: " + primaryOldDate + "]";
             if (historyStr !== "") finalSchedDate += "\n" + historyStr;
         }
         
         if(newStatus !== "") masterSheet.getRange(i + 1, 17).setValue(newStatus); 
         masterSheet.getRange(i + 1, 14).setValue(finalSchedDate); 
         if(warrantedValue !== "") masterSheet.getRange(i + 1, 16).setValue(warrantedValue); 
         
         if (savedFileUrl !== "") {
             var currentMasterPhoto = String(mData[i][14] || "").trim();
             var newMasterPhoto = currentMasterPhoto ? currentMasterPhoto + "\n\n" + savedFileUrl : savedFileUrl;
             masterSheet.getRange(i + 1, 15).setValue(newMasterPhoto);
         }

         masterSheetMsg = "Master Sheet updated.";
         break;
      }
    }
    
    var sheetName = "Lot " + lot + " - Warranty File";
    var existingFiles = lotFolder.searchFiles("title = '" + sheetName.replace(/'/g, "\\'") + "' and mimeType = 'application/vnd.google-apps.spreadsheet'");

    if (existingFiles.hasNext()) {
      var ss = SpreadsheetApp.openById(existingFiles.next().getId());
      var formTab = ss.getSheetByName(formName);
      if (formTab) {
        var fData = formTab.getDataRange().getValues();
        var hdrs = fData[0].map(function(x){ return String(x).toLowerCase().trim(); });
        var iSched = hdrs.indexOf("scheduled date");
        var iWarr = hdrs.indexOf("warranted");
        var iStat = hdrs.indexOf("status");
        var iTrade = hdrs.indexOf("assigned trade");
        var iPhoto = hdrs.indexOf("photo link");

        for (var r = 1; r < fData.length; r++) {
           if (cleanText(fData[r][0]) === cItem && cleanText(fData[r][iTrade]) === cTrade) {
              if(iTrade > -1 && assignedTrade === "") assignedTrade = String(fData[r][iTrade]).trim(); 
              if(iStat > -1 && newStatus !== "") formTab.getRange(r + 1, iStat + 1).setValue(newStatus); 
              if(iSched > -1) formTab.getRange(r + 1, iSched + 1).setValue(finalSchedDate); 
              if(iWarr > -1 && warrantedValue !== "") formTab.getRange(r + 1, iWarr + 1).setValue(warrantedValue); 
              
              if(savedFileUrl !== "" && iPhoto > -1) {
                  var currentLotPhoto = String(fData[r][iPhoto] || "").trim();
                  var newLotPhoto = currentLotPhoto ? currentLotPhoto + "\n\n" + savedFileUrl : savedFileUrl;
                  formTab.getRange(r + 1, iPhoto + 1).setValue(newLotPhoto);
              }

              lotSheetMsg = "Lot Sheet updated.";
              break;
           }
        }
      } else {
        lotSheetMsg = "Form Tab not found in Lot Sheet.";
      }
    }

    var emailDebugMsg = "";
    if (rejectionReason && rejectionReason.trim() !== "") {
        if (assignedTrade !== "") {
            
            var civicAddress = "Unknown Address";
            var lotsSheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(LOTS_SHEET_NAME);
            if (lotsSheet) {
               var lotsData = lotsSheet.getDataRange().getDisplayValues();
               for (var ld = 1; ld < lotsData.length; ld++) {
                  if (cleanText(lotsData[ld][0]) === cProj && cleanText(lotsData[ld][1]).replace(/^0+/, '') === cPhase && cleanText(lotsData[ld][2]).replace(/^0+/, '') === cLot) {
                     civicAddress = lotsData[ld][3];
                     break;
                  }
               }
            }

            var ssSettings = SpreadsheetApp.openById(SHEET_ID);
            var contractorData = ssSettings.getSheetByName(SHEET_NAME).getDataRange().getDisplayValues();
            var trades = assignedTrade.split(',');
            var emailsSent = 0;
            
            for(var t = 0; t < trades.length; t++) {
                var singleTrade = cleanText(trades[t]);
                for (var c = 1; c < contractorData.length; c++) {
                    if (cleanText(contractorData[c][0]) === singleTrade) {
                        var tradeEmail = contractorData[c][2];
                        if (tradeEmail && tradeEmail.indexOf("@") > -1) {
                            var subject = "Service Item Not Approved: " + project + " - Lot " + lot + " (Item #" + itemNum + ")";
                            var body = "Hello " + contractorData[c][0] + ",\n\n" +
                                       "A service item that was previously marked as complete has been reviewed and was not approved by the service coordinator.\n\n" +
                                       "Project: " + project + "\n" +
                                       "Phase: " + phase + "\n" +
                                       "Lot: " + lot + "\n" +
                                       "Civic Address: " + civicAddress + "\n" + 
                                       "Form: " + formName + "\n" +
                                       "Item #: " + itemNum + "\n\n" +
                                       "Reason for Rejection:\n" + rejectionReason + "\n\n" +
                                       "The status has been reverted to 'Assigned'. Please log in to your portal to review and address the issue.\n\n" +
                                       "Thank you,\nGeorgian Build Connect";
                            
                            var emailOptions = { name: "Georgian Build Connect" };
                            
                            if (emailBlob) {
                                emailOptions.attachments = [emailBlob];
                            }
                            
                            try { MailApp.sendEmail(tradeEmail, subject, body, emailOptions); emailsSent++; emailDebugMsg += "\n(Rejection emailed to " + contractorData[c][0] + ")"; } catch(e) { emailDebugMsg += "\n(Email failed for " + contractorData[c][0] + ": " + e.message + ")"; }
                        } else {
                            emailDebugMsg += "\n(No valid email for " + contractorData[c][0] + ")";
                        }
                        break;
                    }
                }
            }
            if (emailsSent === 0 && emailDebugMsg === "") emailDebugMsg = "\n(Could not match trade name '" + assignedTrade + "' in directory)";
        } else {
            emailDebugMsg = "\n(Warning: No trade assigned to this item, email not sent)";
        }
    }
    
    return masterSheetMsg + " " + lotSheetMsg + emailDebugMsg;
  } catch (e) {
    return "Error updating details: " + e.message;
  }
}

function generateServiceReportPDF(filters, selectedCols) {
  try {
      var data = getFilteredServiceOrders(filters);
      if (data.length <= 1 || data[0][0] === "Status") return "Error: No data found matching these filters.";
      
      var headers = data[0].map(function(h) { return String(h).toLowerCase().trim(); });
      var colIndices = [];
      var colNames = [];
      
      for (var i = 0; i < selectedCols.length; i++) {
          var idx = headers.indexOf(selectedCols[i]);
          if (idx > -1) {
              colIndices.push(idx);
              colNames.push(data[0][idx]);
          }
      }
      
      var html = "<html><head><style>";
      html += "@page { size: landscape; margin: 0.5in; } ";
      html += "body { font-family: sans-serif; font-size: 9px; color: #333; } ";
      html += "table { width: 100%; border-collapse: collapse; margin-top: 15px; table-layout: auto; word-wrap: break-word; } ";
      html += "th, td { border: 1px solid #aaa; padding: 5px; text-align: left; vertical-align: top; } ";
      html += "th { background-color: #1c2d42; color: white; font-weight: bold; } ";
      html += "h2 { color: #1c2d42; margin-bottom: 5px; } ";
      html += "p { color: #666; margin-top: 0; } ";
      html += "</style></head><body>";
      
      html += "<h2>Service Orders Report</h2>";
      html += "<p>Generated: " + new Date().toLocaleDateString() + " " + new Date().toLocaleTimeString() + "</p>";
      
      var filterStr = [];
      if(filters.project) filterStr.push("Project: " + filters.project);
      if(filters.phase) filterStr.push("Phase: " + filters.phase);
      if(filters.lot) filterStr.push("Lot: " + filters.lot);
      if(filters.trade) filterStr.push("Trade: " + filters.trade);
      if(filters.status) filterStr.push("Status: " + filters.status);
      if(filterStr.length > 0) html += "<p><b>Filters:</b> " + filterStr.join(" | ") + "</p>";
      
      html += "<table><thead><tr>";
      for (var c = 0; c < colNames.length; c++) {
          html += "<th>" + colNames[c] + "</th>";
      }
      html += "</tr></thead><tbody>";
      
      for (var r = 1; r < data.length; r++) {
          html += "<tr>";
          for (var c = 0; c < colIndices.length; c++) {
              var val = data[r][colIndices[c]];
              if (val === undefined || val === null) val = "";
              val = String(val).replace(/\n/g, "<br>");
              if (colNames[c].toLowerCase() === "photo link" && val.indexOf("http") > -1) {
                 val = "[See Portal for Photos]";
              }
              html += "<td>" + val + "</td>";
          }
          html += "</tr>";
      }
      html += "</tbody></table></body></html>";
      
      var blob = Utilities.newBlob(html, MimeType.HTML).setName("ServiceReport.html");
      var pdfBlob = blob.getAs(MimeType.PDF);
      pdfBlob.setName("Service_Report_" + new Date().getTime() + ".pdf");
      
      var rootFolder = DriveApp.getFolderById(WARRANTY_ROOT_FOLDER_ID);
      var reportFolder = getOrCreateFolder(rootFolder, "Generated Reports");
      var newFile = reportFolder.createFile(pdfBlob);
      
      return newFile.getUrl();
  } catch (e) {
      return "Error generating PDF: " + e.message;
  }
}

function forceAuth() { GmailApp.getAliases(); }

// =====================================================================================
// --- USER DIRECTORY & AUTHENTICATION FUNCTIONS ---
// =====================================================================================
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

// =====================================================================================
// --- SCHEDULE EVENT GENERATOR ---
// =====================================================================================
function getScheduleEvents(allowedProjects) {
  var sheet = SpreadsheetApp.openById(SERVICE_MASTER_ID).getSheets()[0];
  var data = sheet.getDataRange().getDisplayValues();
  if (data.length <= 1) return [];

  var events = [];
  var colorPalette = ['#3498db', '#e74c3c', '#2ecc71', '#f39c12', '#9b59b6', '#34495e', '#16a085', '#d35400', '#27ae60', '#2980b9'];
  var projectColors = {};
  var colorIndex = 0;

  var allowAll = (!allowedProjects || allowedProjects.toLowerCase().indexOf("all") > -1);

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var proj = String(row[1]).trim();
    
    if (!allowAll && allowedProjects.indexOf(proj) === -1) continue;

    var rawDate = String(row[13]).trim();
    if (!rawDate) continue;

    var primaryDateStr = rawDate.split('\n')[0].trim();
    if (primaryDateStr.indexOf('[') > -1) continue; 

    var phase = String(row[2]).trim();
    var lot = String(row[3]).trim();
    var form = String(row[4]).trim();
    var itemNum = String(row[5]).trim();
    var desc = String(row[9]).trim(); 
    var trade = String(row[11]).trim();
    var status = String(row[16]).trim();

    if (!projectColors[proj]) {
      projectColors[proj] = colorPalette[colorIndex % colorPalette.length];
      colorIndex++;
    }

    var d = new Date(primaryDateStr);
    if (isNaN(d.getTime())) continue;

    var yyyy = d.getFullYear();
    var mm = ("0" + (d.getMonth() + 1)).slice(-2);
    var dd = ("0" + d.getDate()).slice(-2);
    var isoDate = yyyy + "-" + mm + "-" + dd;

    events.push({
      title: (trade ? trade : "Unassigned") + " | Lot " + lot + " | Item " + itemNum,
      start: isoDate,
      color: projectColors[proj],
      extendedProps: {
        project: proj,
        phase: phase,
        lot: lot,
        form: form,
        item: itemNum,
        desc: desc, 
        trade: trade,
        status: status,
        originalDateStr: primaryDateStr
      }
    });
  }
  return events;
}

// =====================================================================================
// --- AI CHATBOT HANDLERS ---
// =====================================================================================
function handleAiChat(userQuestion) {
  try {
    var key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
    if (!key) return "Error: GEMINI_API_KEY is not set in Script Properties. Please add it in Project Settings.";

    var dbContext = "";
    var sheetApp = SpreadsheetApp.openById(SHEET_ID);
    
    try {
      var masterApp = SpreadsheetApp.openById(SERVICE_MASTER_ID);
      var masterSheet = masterApp.getSheets()[0];
      var masterData = masterSheet.getDataRange().getDisplayValues();
      if (masterData.length > 1) {
        dbContext += "--- CURRENT SERVICE ORDERS ---\n";
        for (var i = 1; i < Math.min(masterData.length, 50); i++) {
          dbContext += "Item #: " + masterData[i][5] + " | Project: " + masterData[i][1] + " | Lot: " + masterData[i][3] + " | Trade: " + masterData[i][11] + " | Status: " + masterData[i][16] + "\n";
        }
      }
    } catch(e) {
      dbContext += "[Could not load Master Service Items: " + e.message + "]\n";
    }

    var driveContext = "";
    var lotsSheet = sheetApp.getSheetByName(LOTS_SHEET_NAME);
    
    if (lotsSheet) {
      var lotsData = lotsSheet.getDataRange().getDisplayValues();
      var targetFolderIds = [];
      
      for (var l = 1; l < lotsData.length; l++) {
        var projName = lotsData[l][0];
        var lotNum = lotsData[l][2];
        var folderUrl = lotsData[l][16]; 
        
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
          driveContext += "\n--- CONTENTS OF SERVICE FOLDER: " + folder.getName() + " ---\n";
          
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

    var systemInstruction = "You are the internal assistant for the Georgian Build Connect Service Portal. " +
                            "You have real-time access to the portal's system records and local project documentation. " +
                            "Answer the user's question clearly, concisely, and accurately using ONLY the provided system logs, " +
                            "schedules, and text context below. Do not mention PDFs. If the information isn't present, " +
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

// =====================================================================================
// --- HOMEOWNER DIRECTORY INTEGRATION ---
// =====================================================================================
function getHomeownerData() {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Homeowners');
  if (!sheet) return [["Error", "Homeowners tab missing."]];
  return sheet.getDataRange().getDisplayValues();
}

function editHomeowner(formData) {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Homeowners');
  if (!sheet) return "Error: Homeowners tab missing.";
  
  var data = sheet.getDataRange().getDisplayValues();
  var rowIndex = -1;

  for (var i = 1; i < data.length; i++) {
    if (String(data[i][3]).trim().toLowerCase() === String(formData.originalEmail).trim().toLowerCase() &&
        String(data[i][5]).trim() === String(formData.originalProject).trim() &&
        String(data[i][6]).trim() === String(formData.originalPhase).trim() &&
        String(data[i][7]).trim() === String(formData.originalLot).trim()) {
      rowIndex = i + 1;
      break;
    }
  }

  if (rowIndex === -1) return "Error: Specific homeowner lot record not found.";

  sheet.getRange(rowIndex, 1).setValue(formData.firstName);
  sheet.getRange(rowIndex, 2).setValue(formData.lastName);
  sheet.getRange(rowIndex, 3).setValue(formData.phone);
  sheet.getRange(rowIndex, 4).setValue(formData.email);
  sheet.getRange(rowIndex, 6).setValue(formData.project);
  sheet.getRange(rowIndex, 7).setValue(formData.phase);
  sheet.getRange(rowIndex, 8).setValue(formData.lot);

  return "Success";
}

function deleteHomeowner(email, project, phase, lot) {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Homeowners');
  if (!sheet) return "Error: Homeowners tab missing.";
  
  var data = sheet.getDataRange().getDisplayValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (String(data[i][3]).trim().toLowerCase() === String(email).trim().toLowerCase() &&
        String(data[i][5]).trim() === String(project).trim() &&
        String(data[i][6]).trim() === String(phase).trim() &&
        String(data[i][7]).trim() === String(lot).trim()) {
      sheet.deleteRow(i + 1);
      return "Deleted";
    }
  }
  return "Error: Record not found.";
}

function getLotsHierarchy() {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(LOTS_SHEET_NAME);
  if (!sheet) return { projects: [], hierarchy: {} };
  
  var data = sheet.getDataRange().getDisplayValues();
  var hierarchy = {};
  var projects = [];
  
  for (var i = 1; i < data.length; i++) {
    var proj = data[i][0].trim();
    var phase = data[i][1].trim();
    var lot = data[i][2].trim();
    
    if (!proj || proj === "Project Name" || proj === "Project") continue; 
    
    if (projects.indexOf(proj) === -1) projects.push(proj);
    if (!hierarchy[proj]) hierarchy[proj] = {};
    if (!hierarchy[proj][phase]) hierarchy[proj][phase] = [];
    if (hierarchy[proj][phase].indexOf(lot) === -1) hierarchy[proj][phase].push(lot);
  }
  return { projects: projects.sort(), hierarchy: hierarchy };
}

function addHomeowner(formData) {
  var sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName('Homeowners');
  if (!sheet) return "Error: Please create a tab named 'Homeowners' in your Master Sheet.";
  
  var newPassword = generateRandomPassword(10);
  
  sheet.appendRow([
    formData.firstName, 
    formData.lastName, 
    formData.phone, 
    formData.email, 
    newPassword, 
    formData.project, 
    formData.phase, 
    formData.lot
  ]);
  
  try {
    var subject = "Welcome to your Homeowner Portal";
    var body = "Hello " + formData.firstName + ",\n\n" +
               "Your Georgian Build Connect Homeowner Portal account has been created.\n\n" +
               "Property Details:\n" +
               "Project: " + formData.project + "\n" +
               "Phase: " + formData.phase + "\n" +
               "Lot: " + formData.lot + "\n\n" +
               "You can log in to view your property details and warranty items using this email address and the following passcode:\n" +
               "Passcode: " + newPassword + "\n\n" +
               "Portal Link: https://sites.google.com/georgiancommunitiesconstruction.ca/georgian-pm-program-portal/home\n\n" +
               "Thank you,\nGeorgian Build Connect";
    
    MailApp.sendEmail(formData.email, subject, body, {name: "Georgian Build Connect"});
  } catch(e) {
    return "Success, but the welcome email failed to send: " + e.message;
  }
  
  return "Success";
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

    // 3. Find Service Coordinator Email and Send Notification
    var settingsSS = SpreadsheetApp.openById(SHEET_ID);
    var usersSheet = settingsSS.getSheetByName('Users');
    var targetEmails = [];

    if (usersSheet) {
      var uData = usersSheet.getDataRange().getValues();
      for (var u = 1; u < uData.length; u++) {
        var uJob = String(uData[u][2]).toLowerCase();
        var uProj = String(uData[u][3]).toLowerCase();
        var uEmail = String(uData[u][4]);
        
        if ((uProj.indexOf(projectName) > -1 || uProj.indexOf("all") > -1) && (uJob.indexOf("service") > -1 || uJob.indexOf("coordinator") > -1) && uEmail !== "") {
          if (targetEmails.indexOf(uEmail) === -1) targetEmails.push(uEmail);
        }
      }
    }
    
    // Fallback if no coordinator is mapped in the Users tab
    if (targetEmails.length === 0) {
      var setSheet = settingsSS.getSheetByName('Settings');
      if (setSheet) {
        var sData = setSheet.getDataRange().getValues();
        for (var s = 1; s < sData.length; s++) {
          if (cleanText(sData[s][0]) === projectName && sData[s][4] !== "") {
            targetEmails.push(sData[s][4]);
            break;
          }
        }
      }
    }

    if (targetEmails.length > 0) {
      var subject = "New Warranty Items Submitted: " + project + " - Lot " + lot;
      var body = "Hello,\n\n" +
                 hoName + " (" + hoEmail + ") has submitted " + itemsSubmitted + " new warranty item(s) for " + project + " Phase " + phase + " Lot " + lot + ".\n\n" +
                 "Items Submitted:\n" + submittedDetails.join("\n") + "\n\n" +
                 "Please review them in the Service Portal.\n\n" +
                 "Thank you,\nGeorgian Build Connect";
      MailApp.sendEmail(targetEmails.join(","), subject, body, {name: "Georgian Build Connect"});
    }

    return "Success! " + itemsSubmitted + " item(s) submitted to the Service Coordinator.";

  } catch(e) {
    return "Error: " + e.message;
  }
}

// Utility Function: Ensure this exists in your Homeowner Code.gs as well.
function getOrCreateFolder(parentFolder, folderName) {
  var safeName = folderName.replace(/'/g, "\\'");
  var folders = parentFolder.searchFolders("title = '" + safeName + "' and trashed = false");
  if (folders.hasNext()) { return folders.next(); }
  return parentFolder.createFolder(folderName);
}