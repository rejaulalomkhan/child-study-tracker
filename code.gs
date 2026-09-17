/**
 * Arians Study - multi-student Google Sheets backend.
 * Deploy as a Web App: Execute as owner, access for anyone with the link.
 */
const REGISTRY_SHEET = 'Students';
const STUDENT_HEADERS = ['ID', 'Subject', 'Topic', 'Taught', 'Oral', 'Written', 'Updated At'];

function doGet(e) {
  try {
    const params = e.parameter || {};
    if (params.action === 'students') {
      return jsonResponse({ status: 'success', students: findStudents_(params.phone) });
    }
    if (params.action === 'load') {
      return jsonResponse(loadStudent_(params.studentId, params.phone));
    }
    return jsonResponse({ status: 'success', message: 'Arians Study API is running.' });
  } catch (error) {
    return jsonResponse({ status: 'error', message: error.message });
  }
}

function doPost(e) {
  try {
    const request = JSON.parse(e.postData.contents || '{}');
    if (request.action === 'register') return jsonResponse(registerStudent_(request));
    if (request.action === 'save') return jsonResponse(saveStudent_(request));
    return jsonResponse({ status: 'error', message: 'Unknown action.' });
  } catch (error) {
    return jsonResponse({ status: 'error', message: error.message });
  }
}

function registerStudent_(request) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
  const name = clean_(request.name);
  const phone = normalizePhone_(request.phone);
  const startMonth = clean_(request.startMonth);
  const endMonth = clean_(request.endMonth);
  if (!name || !phone || !startMonth || !endMonth) throw new Error('Name, phone, and teaching period are required.');
  if (endMonth < startMonth) throw new Error('End month must be after start month.');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const registry = getRegistrySheet_(ss);
  const studentId = Utilities.getUuid();
  const sheetName = uniqueSheetName_(ss, name);
  const now = new Date();
  registry.appendRow([studentId, name, phone, startMonth, endMonth, sheetName, now, now]);

  const studentSheet = ss.insertSheet(sheetName);
  studentSheet.getRange(1, 1, 1, STUDENT_HEADERS.length).setValues([STUDENT_HEADERS]);
  formatHeader_(studentSheet);
  return { status: 'success', student: { studentId, name, phone, startMonth, endMonth } };
  } finally {
    lock.releaseLock();
  }
}

function findStudents_(phone) {
  const normalizedPhone = normalizePhone_(phone);
  if (!normalizedPhone) throw new Error('Guardian phone is required.');
  const registry = getRegistrySheet_(SpreadsheetApp.getActiveSpreadsheet());
  const values = registry.getDataRange().getValues();
  return values.slice(1).filter(row => normalizePhone_(row[2]) === normalizedPhone).map(studentFromRow_);
}

function loadStudent_(studentId, phone) {
  const student = findStudent_(studentId, phone);
  if (!student) throw new Error('Student not found or phone number does not match.');
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(student.sheetName);
  const values = sheet && sheet.getLastRow() > 1 ? sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues() : [];
  const data = values.map(row => ({
    id: Number(row[0]),
    subject: row[1],
    topic: row[2],
    taught: row[3],
    oral: row[4],
    written: row[5],
    updatedAt: row[6]
  }));
  return { status: 'success', student, data };
}

function saveStudent_(request) {
  const student = findStudent_(request.studentId, request.phone);
  if (!student) throw new Error('Student not found or phone number does not match.');
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(student.sheetName);
  const rows = Array.isArray(request.rows) ? request.rows : [];
  if (sheet.getLastRow() > 1) sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).clearContent();
  if (rows.length) {
    sheet.getRange(2, 1, rows.length, 7).setValues(rows.map(item => [
      item.id, item.subject, item.topic, item.taught, item.oral, item.written, item.updatedAt
    ]));
  }
  const registry = getRegistrySheet_(SpreadsheetApp.getActiveSpreadsheet());
  const rowIndex = findRegistryRow_(registry, student.studentId);
  registry.getRange(rowIndex, 8).setValue(new Date());
  return { status: 'success', message: 'Student data saved.' };
}

function findStudent_(studentId, phone) {
  return findStudents_(phone).find(student => student.studentId === studentId) || null;
}

function getRegistrySheet_(ss) {
  let sheet = ss.getSheetByName(REGISTRY_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(REGISTRY_SHEET, 0);
    sheet.getRange(1, 1, 1, 8).setValues([[
      'Student ID', 'Student Name', 'Guardian Phone', 'Start Month', 'End Month', 'Sheet Name', 'Created At', 'Updated At'
    ]]);
    formatHeader_(sheet);
  }
  sheet.getRange('C:E').setNumberFormat('@');
  return sheet;
}

function studentFromRow_(row) {
  return {
    studentId: String(row[0]),
    name: String(row[1]),
    phone: String(row[2]),
    startMonth: monthValue_(row[3]),
    endMonth: monthValue_(row[4]),
    sheetName: String(row[5])
  };
}

function monthValue_(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM');
  }
  return String(value);
}

function findRegistryRow_(sheet, studentId) {
  const values = sheet.getRange(2, 1, Math.max(sheet.getLastRow() - 1, 1), 1).getValues();
  const index = values.findIndex(row => String(row[0]) === String(studentId));
  if (index < 0) throw new Error('Student registry entry not found.');
  return index + 2;
}

function uniqueSheetName_(ss, name) {
  const base = name.replace(/[\\/?*\[\]:]/g, '').trim().slice(0, 80) || 'Student';
  let candidate = base;
  let suffix = 2;
  while (ss.getSheetByName(candidate)) candidate = `${base.slice(0, 76)} (${suffix++})`;
  return candidate;
}

function normalizePhone_(phone) {
  return String(phone || '').replace(/[\s-]/g, '');
}

function clean_(value) {
  return String(value || '').trim();
}

function formatHeader_(sheet) {
  sheet.getRange(1, 1, 1, sheet.getLastColumn()).setFontWeight('bold').setBackground('#152C5B').setFontColor('#FFFFFF');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, sheet.getLastColumn());
}

function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}
