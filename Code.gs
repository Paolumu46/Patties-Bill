function doGet(e) {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('ระบบบริหารจัดการบิล Patties Bill')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function processAction(action, data) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  let headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
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
    return JSON.stringify({status: 'success', data: result.reverse()});
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
}

function apiRequest(action, data) {
  try {
    return processAction(action, data);
  } catch(e) {
    return JSON.stringify({status: 'error', message: e.toString()});
  }
}
