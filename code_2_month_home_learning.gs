/**
 * Arian's 2-Month Core Learning Plan
 * Google Apps Script backend for a separate spreadsheet.
 */

const PROGRESS_SHEET = 'Core_Learning_Progress';
const REPORT_SHEET = 'Study_Reports';
const PROGRESS_HEADERS = [
  'Date', 'Week', 'Day', 'Subject', 'Goal', 'Teacher Activity',
  'Home Practice', 'Status', 'Last Updated'
];
const REPORT_HEADERS = [
  'Report Type', 'Period', 'Planned Lessons', 'Completed',
  'Partial', 'Pending', 'Generated At'
];

function doGet() {
  return jsonResponse({ status: 'ok', message: 'Arian core learning tracker is ready.' });
}

function doPost(e) {
  try {
    const request = JSON.parse(e.postData.contents || '{}');
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const progressRows = Array.isArray(request.progress) ? request.progress : [];
    const reports = Array.isArray(request.reports) ? request.reports : [];

    writeTable(spreadsheet, PROGRESS_SHEET, PROGRESS_HEADERS, progressRows);
    writeTable(spreadsheet, REPORT_SHEET, REPORT_HEADERS, reports);

    return jsonResponse({
      status: 'success',
      progressRows: progressRows.length,
      reportRows: reports.length,
      message: 'Arian study progress saved successfully.'
    });
  } catch (error) {
    return jsonResponse({ status: 'error', message: error.toString() });
  }
}

function writeTable(spreadsheet, sheetName, headers, rows) {
  let sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) sheet = spreadsheet.insertSheet(sheetName);

  sheet.clearContents();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground('#173B57')
    .setFontColor('#FFFFFF');
  sheet.setFrozenRows(1);

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }

  sheet.autoResizeColumns(1, headers.length);
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
