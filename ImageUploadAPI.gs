// API สำหรับรับไฟล์รูปภาพ Base64 ผ่าน POST Request และบันทึกลง Google Drive + Google Sheets

function doPost(e) {
  try {
    // กำหนด Folder ID ปลายทางที่ต้องการบันทึกรูปภาพ (เอาไว้รับค่า JSON หรือ Fix ไว้ในโค้ด)
    const FOLDER_ID = 'ใส่_FOLDER_ID_ที่ต้องการ_ที่นี่'; 

    // 1. รับข้อมูล JSON POST จากหน้าเว็บ
    const data = JSON.parse(e.postData.contents);
    
    // ดึงข้อมูล Base64 Image
    const base64Data = data.base64Image; 
    
    // ชื่อไฟล์รูปภาพ ถ้าไม่มีให้ใช้วันที่เวลาตั้งชื่อ
    const fileName = data.fileName || ('Image_' + new Date().getTime() + '.png');
    
    // จัดการ String Base64 กรณีมี Header Data URI ติดมาด้วย เช่น "data:image/png;base64,xxxx"
    let imageStr = base64Data;
    if (imageStr.indexOf('base64,') !== -1) {
      imageStr = imageStr.split('base64,')[1];
    }
    
    // 2. แปลง Base64 เป็นไฟล์รูปภาพ (Blob) และบันทึกลง Google Drive
    // สกัดนามสกุลไฟล์เพื่อกำหนด MimeType
    const ext = fileName.split('.').pop().toLowerCase();
    let mimeType = MimeType.PNG; // Default เป็น PNG
    if(ext === 'jpg' || ext === 'jpeg') mimeType = MimeType.JPEG;
    if(ext === 'gif') mimeType = MimeType.GIF;

    const decodedImage = Utilities.base64Decode(imageStr);
    const blob = Utilities.newBlob(decodedImage, mimeType, fileName);
    
    const folder = DriveApp.getFolderById(FOLDER_ID);
    const file = folder.createFile(blob); // สร้างไฟล์ลง Drive
    
    const fileUrl = file.getUrl();
    const uploadDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
    
    // 3. บันทึกข้อมูล (ชื่อไฟล์, วันที่อัปโหลด, ลิงก์ไฟล์) ลง Google Sheets แผ่นที่ชื่อว่า 'Bill_Logs'
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('Bill_Logs');
    
    // ตรวจสอบว่าแผ่นปลายทาง (Sheet) ชื่อนี้มีอยู่แล้วหรือไม่ ถ้าไม่มีให้สร้างขึ้นมาใหม่พร้อม Header Column
    if (!sheet) {
      sheet = ss.insertSheet('Bill_Logs');
      sheet.appendRow(['File Name', 'Upload Date', 'File URL']);
      sheet.getRange("A1:C1").setFontWeight("bold"); // สร้างตัวหนาให้ Header
    }
    
    // เพิ่มแถวข้อมูลใหม่
    sheet.appendRow([fileName, uploadDate, fileUrl]);
    
    // 4. ส่งผลลัพธ์กลับไปเป็น JSON บอกสถานะการอัปโหลด + URL ของไฟล์
    const responseData = {
      status: 'success',
      message: 'อัปโหลดรูปภาพลง Google Drive สาเร็จ',
      fileName: fileName,
      fileUrl: fileUrl
    };

    return ContentService.createTextOutput(JSON.stringify(responseData))
      .setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    // กรณีมี Error ดักจับแล้วส่งข้อความ Error กลับไป
    return ContentService.createTextOutput(JSON.stringify({
      status: 'error',
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// **เสริม:** ฟังก์ชันจำลองรองรับ CORS Options (Preflight Request)
// สิ่งนี้จำเป็นมากๆ เมื่อส่งข้อมูลแบบ POST JSON ข้ามโดเมนจากหน้าเว็บ HTML ทั่วไป มายัง Google Apps Script (GAS) 
function doOptions(e) {
  return ContentService.createTextOutput("")
    .setMimeType(ContentService.MimeType.JSON);
}
