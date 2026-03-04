function doGet(e) {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('ระบบบริหารจัดการบิล Patties Bill')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Function to get the settings sheet or create it if not exists
function getSettingsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let settingsSheet = ss.getSheetByName('Settings');
  if (!settingsSheet) {
    settingsSheet = ss.insertSheet('Settings');
    settingsSheet.appendRow(['Key', 'Value']);
    settingsSheet.appendRow(['FolderID', '']);
  }
  return settingsSheet;
}

function processAction(action, data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheets()[0]; // Assume first sheet is data
  let headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const settingsSheet = getSettingsSheet();
  
  // --- Data CRUD ---
  if (action === 'read') {
    const rows = sheet.getDataRange().getDisplayValues();
    const result = [];
    for (let i = 1; i < rows.length; i++) {
        if(rows[i].join('').trim() === '') continue;
        let obj = {};
        for (let j = 0; j < headers.length; j++) {
            obj[headers[j]] = rows[i][j];
        }
        obj._rowIndex = i + 1;
        result.push(obj);
    }
    
    // Read Settings
    let folderId = '';
    const settingsData = settingsSheet.getDataRange().getValues();
    for (let i = 1; i < settingsData.length; i++) {
      if (settingsData[i][0] === 'FolderID') folderId = settingsData[i][1];
    }
    
    return JSON.stringify({status: 'success', data: result.reverse(), settings: { folderId }});
  }
  
  if (action === 'create') {
    const newRow = [];
    const timestampId = new Date().getTime().toString();
    for (let j = 0; j < headers.length; j++) {
      let header = headers[j];
      if (header === 'ID') newRow.push(timestampId);
      else newRow.push(data[header] !== undefined ? data[header] : '');
    }
    sheet.appendRow(newRow);
    return JSON.stringify({status: 'success', message: 'สร้างข้อมูลบิลเรียบร้อยแล้ว'});
  }
  
  if (action === 'update') {
    const rowIndex = parseInt(data._rowIndex);
    if (!rowIndex || rowIndex < 2) throw new Error("Invalid Row Index");
    for (let j = 0; j < headers.length; j++) {
      let header = headers[j];
      if (header !== 'ID' && header !== '_rowIndex' && data[header] !== undefined) {
         sheet.getRange(rowIndex, j + 1).setValue(data[header]);
      }
    }
    return JSON.stringify({status: 'success', message: 'อัปเดตข้อมูลเรียบร้อยแล้ว'});
  }
  
  if (action === 'delete') {
    const rowIndex = parseInt(data._rowIndex);
    if (!rowIndex || rowIndex < 2) throw new Error("Invalid Row Index");
    sheet.deleteRow(rowIndex);
    return JSON.stringify({status: 'success', message: 'ลบข้อมูลบิลเรียบร้อยแล้ว'});
  }

  // --- Settings ---
  if (action === 'saveSettings') {
    const settingsData = settingsSheet.getDataRange().getValues();
    let found = false;
    for (let i = 1; i < settingsData.length; i++) {
      if (settingsData[i][0] === 'FolderID') {
        settingsSheet.getRange(i + 1, 2).setValue(data.folderId);
        found = true;
        break;
      }
    }
    if (!found) settingsSheet.appendRow(['FolderID', data.folderId]);
    return JSON.stringify({status: 'success', message: 'บันทึกการตั้งค่าแล้ว'});
  }

  // --- Drive File Management ---
  if (action === 'listDriveFiles') {
    const folderId = data.folderId;
    if (!folderId) throw new Error("ยังไม่ได้กำหนด Folder ID ในตั้งค่า");
    
    const folder = DriveApp.getFolderById(folderId);
    const files = folder.getFilesByType(MimeType.PDF);
    const fileList = [];
    
    while (files.hasNext()) {
      const file = files.next();
      fileList.push({
        id: file.getId(),
        name: file.getName(),
        url: file.getUrl(),
        downloadUrl: file.getDownloadUrl(),
        dateCreated: Utilities.formatDate(file.getDateCreated(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm")
      });
    }
    // Sort by newest first based on dateCreated inside AppScript environment
    fileList.sort((a,b) => {
        let da = parseInt(a.dateCreated.substring(6,10))*10000 + parseInt(a.dateCreated.substring(3,5))*100 + parseInt(a.dateCreated.substring(0,2));
        let db = parseInt(b.dateCreated.substring(6,10))*10000 + parseInt(b.dateCreated.substring(3,5))*100 + parseInt(b.dateCreated.substring(0,2));
        return db - da; 
    });
    
    return JSON.stringify({status: 'success', data: fileList});
  }

  if (action === 'deleteDriveFile') {
    const fileId = data.fileId;
    if (!fileId) throw new Error("ไม่พบรหัสไฟล์ที่ต้องการลบ");
    DriveApp.getFileById(fileId).setTrashed(true);
    return JSON.stringify({status: 'success', message: 'ลบไฟล์ PDF ออกจากระบบเรียบร้อยแล้ว'});
  }

  if (action === 'savePdf') {
    const folderId = data.folderId;
    if (!folderId) throw new Error("กรุณาตั้งค่า Folder ID ก่อนสรา้งไฟล์");

    const htmlContent = data.htmlContent; 
    const fileName = data.fileName || ('Invoice_' + new Date().getTime() + '.pdf');

    // Create a Blob from HTML content
    const blob = Utilities.newBlob(htmlContent, MimeType.HTML, fileName).getAs(MimeType.PDF);
    
    const folder = DriveApp.getFolderById(folderId);
    const file = folder.createFile(blob);
    
    return JSON.stringify({
      status: 'success', 
      message: 'สร้างไฟล์ PDF ลง Google Drive สำเร็จ', 
      fileUrl: file.getUrl(),
      fileId: file.getId()
    });
  }
}

function apiRequest(action, data) {
  try {
    return processAction(action, data);
  } catch(e) {
    return JSON.stringify({status: 'error', message: e.toString()});
  }
}
