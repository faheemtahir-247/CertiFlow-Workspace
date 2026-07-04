// CONFIGURATION PARAMETERS (SECURE PLACEHOLDER)

const OUTPUT_FOLDER_ID = 'YOUR_GOOGLE_DRIVE_FOLDER_ID_HERE';

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('CertiFlow Workspace')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function extractIdFromUrl(url) {
  if (!url) return "";
  const matches = url.match(/\/d\/([a-zA-Z0-9-_]+)/);
  return (matches && matches[1]) ? matches[1] : url;
}

function getRemoteTabs(sheetUrl) {
  try {
    const sheetId = extractIdFromUrl(sheetUrl);
    const ss = SpreadsheetApp.openById(sheetId);
    return ss.getSheets().map(s => s.getName());
  } catch(e) {
    throw new Error("Target Spreadsheet URL is invalid or inaccessible! Please check share settings.");
  }
}

function processGoogleSheetPipeline(config) {
  const masterSS = SpreadsheetApp.getActiveSpreadsheet();
  let configSheet = masterSS.getSheetByName("Master_Config");
  if (!configSheet) {
    configSheet = masterSS.insertSheet("Master_Config");
    configSheet.appendRow(["Tab Name", "Slide URL", "Slide Number", "Release Date", "Source", "Status"]);
  }
  
  const targetSheetId = extractIdFromUrl(config.targetSheetUrl);
  const slideId = extractIdFromUrl(config.slideUrl);
  
  try {
    const targetSS = SpreadsheetApp.openById(targetSheetId);
    const tSheet = targetSS.getSheetByName(config.tabName);
    if (!tSheet) throw new Error("The selected Tab Name does not exist in that Spreadsheet.");
    DriveApp.getFileById(slideId);
  } catch(err) {
    throw new Error("Validation Failed: " + err.message);
  }
  
  if (config.deliveryType === 'instant') {
    const targetSS = SpreadsheetApp.openById(targetSheetId);
    const targetSheet = targetSS.getSheetByName(config.tabName);
    const success = runGenerationLogic(targetSheet, slideId, parseInt(config.slideNumber, 10));
    
    configSheet.appendRow([config.tabName, config.slideUrl, config.slideNumber, new Date(), "Google Sheet URL (Instant): " + config.targetSheetUrl, success ? "Processed" : "Partial Error"]);
    
    if (success) return "SUCCESS: Certificates processed and emails dispatched instantly!";
    else throw new Error("Data match schema mismatch or partial row error inside the sheet.");
    
  } else {
    const formattedDate = new Date(config.releaseDate + "T" + config.releaseTime);
    configSheet.appendRow([config.tabName, config.slideUrl, config.slideNumber, formattedDate, "Google Sheet URL (Scheduled): " + config.targetSheetUrl, "Scheduled"]);
    return "SUCCESS: Pipeline successfully scheduled. The automation engine will release it at the designated time.";
  }
}

function processExcelUploadPipeline(excelBlobData, slideUrl, slideNumber) {
  let tempFileId = null;
  try {
    const slideId = extractIdFromUrl(slideUrl);
    try {
      DriveApp.getFileById(slideId);
    } catch(e) {
      throw new Error("Google Slide Template URL is invalid or inaccessible.");
    }
    
    const fileBlob = Utilities.newBlob(Utilities.base64Decode(excelBlobData), MimeType.MICROSOFT_EXCEL, "temp_uploaded.xlsx");
    const resource = {
      name: "Temp_Converted_Data",
      mimeType: "application/vnd.google-apps.spreadsheet"
    };
    const response = Drive.Files.create(resource, fileBlob);
    tempFileId = response.id;
    
    const tempSS = SpreadsheetApp.openById(tempFileId);
    const targetSheet = tempSS.getSheets()[0]; 
    
    const success = runGenerationLogic(targetSheet, slideId, parseInt(slideNumber, 10));
    DriveApp.getFileById(tempFileId).setTrashed(true);
    
    if (success) {
      return "SUCCESS: Excel matrix processed and emails dispatched instantly!";
    } else {
      throw new Error("Columns mapping error. Standard columns: 'Full Name', 'Roll Number', 'Email Address' not found in Excel.");
    }
  } catch(e) {
    if (tempFileId) DriveApp.getFileById(tempFileId).setTrashed(true);
    throw new Error(e.message);
  }
}

function processScheduledCertificates() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName("Master_Config");
  if (!configSheet) return;
  
  const configData = configSheet.getDataRange().getValues();
  const currentTime = new Date();
  
  for (let i = 1; i < configData.length; i++) {
    const [tabName, slideUrl, slideNumber, releaseDate, sourceInfo, status] = configData[i];
    if (status !== "Scheduled") continue;
    
    if (releaseDate instanceof Date && currentTime >= releaseDate) {
      try {
        const urlMatch = sourceInfo.match(/https:\/\/docs\.google\.com\/spreadsheets\/d\/[a-zA-Z0-9-_]+/);
        if (!urlMatch) throw new Error("Metadata extraction fault");
        
        const targetSS = SpreadsheetApp.openByUrl(urlMatch[0]);
        const targetSheet = targetSS.getSheetByName(tabName);
        const slideId = extractIdFromUrl(slideUrl);
        
        const success = runGenerationLogic(targetSheet, slideId, parseInt(slideNumber, 10));
        configSheet.getRange(i + 1, 6).setValue(success ? "Processed" : "Partial Error");
      } catch (e) {
        configSheet.getRange(i + 1, 6).setValue("Execution Failure");
      }
    }
  }
}

function runGenerationLogic(sheet, templateSlideId, slideNumber) {
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return false;
  
  const headers = data[0];
  const nameIdx = headers.indexOf("Full Name");
  const rollIdx = headers.indexOf("Roll Number");
  const emailIdx = headers.indexOf("Email Address");
  const statusIdx = headers.indexOf("Status");
  
  if (nameIdx === -1 || rollIdx === -1 || emailIdx === -1) return false;
  
  const targetFolder = DriveApp.getFolderById(OUTPUT_FOLDER_ID);
  let overallSuccess = true;
  
  for (let i = 1; i < data.length; i++) {
    const name = data[i][nameIdx];
    const rollNo = data[i][rollIdx];
    const email = data[i][emailIdx];
    const status = statusIdx !== -1 ? data[i][statusIdx] : "";
    
    if (status === "Sent" || !email || !name) continue;
    
    try {
      const copyFile = DriveApp.getFileById(templateSlideId).makeCopy(`${name} - Certificate`, targetFolder);
      const copyPresentation = SlidesApp.openById(copyFile.getId());
      const slides = copyPresentation.getSlides();
      const totalSlides = slides.length;
      
      if (slideNumber > totalSlides || slideNumber < 1) {
        throw new Error("Selected slide number is out of bounds.");
      }
      
      const targetIndex = slideNumber - 1;
      for (let k = totalSlides - 1; k >= 0; k--) {
        if (k !== targetIndex) {
          copyPresentation.getSlides()[k].remove();
        }
      }
      
      const activeSlide = copyPresentation.getSlides()[0];
      activeSlide.replaceAllText("{{Name}}", name);
      activeSlide.replaceAllText("{{RollNo}}", rollNo);
      
      const qrData = `Verified Certificate\nName: ${name}\nVerification ID: ${rollNo}`;
      const qrUrl = `https://api.qrserver.com/v1/create-qr-code/?size=300x300&margin=4&data=${encodeURIComponent(qrData)}`;
      
      const pageElements = activeSlide.getPageElements();
      for (let j = 0; j < pageElements.length; j++) {
        const element = pageElements[j];
        if (element.getPageElementType() == SlidesApp.PageElementType.SHAPE && element.asShape().getDescription() === "QR_PLACEHOLDER") {
          activeSlide.insertImage(qrUrl, element.getLeft(), element.getTop(), element.getWidth(), element.getHeight());
          element.remove();
          break;
        }
      }
      
      copyPresentation.saveAndClose();
      
      const pdfBlob = copyFile.getAs(MimeType.PDF);
      const pdfFile = targetFolder.createFile(pdfBlob).setName(`${name}_Certificate.pdf`);
      copyFile.setTrashed(true);
      
      GmailApp.sendEmail(email, "Your Verified Certificate Attached", `Dear ${name},\n\nPlease find attached your verified certificate.\n\nVerification ID: ${rollNo}`, {
        attachments: [pdfFile.getAs(MimeType.PDF)],
        name: "CertiFlow Automation Suite"
      });
      
      if (statusIdx !== -1) {
        sheet.getRange(i + 1, statusIdx + 1).setValue("Sent");
      }
      Utilities.sleep(1000);
    } catch (e) {
      overallSuccess = false;
      if (statusIdx !== -1) sheet.getRange(i + 1, statusIdx + 1).setValue("Error");
    }
  }
  return overallSuccess;
}
