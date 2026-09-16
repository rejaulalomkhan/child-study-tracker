/**
 * Arians Study - Syllabus Tracking Backend
 * Google Apps Script for Google Spreadsheet
 */

function doPost(e) {
  try {
    const sheetName = "Syllabus_Progress";
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(sheetName);
    
    // শিট না থাকলে স্বয়ংক্রিয়ভাবে কলামসহ তৈরি হবে
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.appendRow([
        "ক্রমিক", 
        "বিষয়", 
        "পাঠ্যসূচি / বিষয়বস্তু", 
        "শেখানো হয়েছে", 
        "বলতে পারে", 
        "লিখতে পারে", 
        "সর্বশেষ আপডেট"
      ]);
      sheet.getRange(1, 1, 1, 7).setFontWeight("bold").setBackground("#152C5B").setFontColor("#FFFFFF");
      sheet.setFrozenRows(1);
    }
    
    const requestData = JSON.parse(e.postData.contents);
    const rows = requestData.data;
    
    // পুরাতন ডাটা মুছে নতুন ডাটা রিফ্রেশ করা
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, 7).clearContent();
    }
    
    // নতুন ডাটা ইনসার্ট করা
    const dataToInsert = rows.map(item => [
      item.id,
      item.subject,
      item.topic,
      item.taught,
      item.oral,
      item.written,
      item.updatedAt
    ]);
    
    if (dataToInsert.length > 0) {
      sheet.getRange(2, 1, dataToInsert.length, 7).setValues(dataToInsert);
    }
    
    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      message: "Data saved successfully!"
    })).setMimeType(ContentService.MimeType.JSON);
    
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}