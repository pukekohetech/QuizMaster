/**
 * PHS Safety Submission Gateway
 *
 * Folder structure:
 *   PHS Safety Submissions / Teacher / US #### / Student ID - Student Name /
 *
 * Each deliberate resubmission:
 * - replaces the PDF for the SAME assessment component
 * - replaces the student's single latest .puk backup for the unit standard
 * - updates one register row for that student + teacher + unit standard
 * - increments the submission counter
 *
 * Exact duplicate requests (same submissionId) do not create extra files or
 * increment the register.
 */

const CONFIG = {
  ROOT_FOLDER_NAME: 'PHS Safety Submissions',
  ROOT_FOLDER_ID: '', // Optional: paste an existing Drive folder ID. Leave blank to auto-create.
  REGISTER_FILE_NAME: 'PHS Safety Submission Register',
  REGISTER_SPREADSHEET_ID: '', // Optional: paste an existing Sheet ID. Leave blank to auto-create.
  REGISTER_SHEET_NAME: 'Submissions',
  NOTIFY_TEACHER: false, // OFF by default. Change to true for link-only teacher notifications.
  EXPECTED_APP_PREFIX: 'pukekohetech-',
  MAX_PDF_BYTES: 25 * 1024 * 1024,
  MAX_PUK_CHARS: 5 * 1024 * 1024,
};

const HEADERS = [
  'Record Key',
  'Student',
  'Student ID',
  'Teacher',
  'Teacher ID',
  'Unit Standard',
  'Last Assessment',
  'Result',
  'Last Submitted',
  'Submissions',
  'Student Folder',
  'Latest PDF',
  'PUK',
  'Last Submission ID',
  'App Version',
  'Status',
];

function setup() {
  const root = getRootFolder_();
  const spreadsheet = getRegisterSpreadsheet_(root);
  const sheet = getRegisterSheet_(spreadsheet);
  return {
    ok: true,
    rootFolder: root.getUrl(),
    register: spreadsheet.getUrl(),
    sheet: sheet.getName(),
    notifyTeacher: CONFIG.NOTIFY_TEACHER,
  };
}

function doGet(e) {
  const params = (e && e.parameter) || {};
  const callback = params.callback || '';
  try {
    const action = String(params.action || 'health').toLowerCase();
    if (action === 'status') {
      return output_(getStatus_(String(params.submissionId || '')), callback);
    }
    const info = setup();
    return output_({
      ok: true,
      state: 'ready',
      message: 'PHS Safety submission gateway is ready.',
      notifyTeacher: info.notifyTeacher,
      rootFolder: info.rootFolder,
      register: info.register,
    }, callback);
  } catch (error) {
    return output_({ ok: false, state: 'error', message: String(error && error.message || error) }, callback);
  }
}

function doPost(e) {
  let submissionId = '';
  try {
    if (!e || !e.postData || !e.postData.contents) throw new Error('Missing request body.');
    const payload = JSON.parse(e.postData.contents);
    submissionId = cleanText_(payload.submissionId, 160);
    if (!submissionId) throw new Error('Missing submissionId.');
    setStatus_(submissionId, { state: 'processing', message: 'Submission is being saved.' });

    validatePayload_(payload);
    const result = saveSubmission_(payload);
    setStatus_(submissionId, result);
    return output_(result, '');
  } catch (error) {
    const result = { state: 'error', message: String(error && error.message || error) };
    if (submissionId) setStatus_(submissionId, result);
    return output_(result, '');
  }
}

function saveSubmission_(payload) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const root = getRootFolder_();
    const spreadsheet = getRegisterSpreadsheet_(root);
    const sheet = getRegisterSheet_(spreadsheet);

    const teacherId = cleanText_(payload.teacherId, 40) || 'Teacher';
    const teacherName = cleanText_(payload.teacherName, 100) || teacherId;
    const studentId = cleanText_(payload.studentId, 60);
    const studentName = cleanText_(payload.studentName, 120);
    const standard = normaliseStandard_(payload.unitStandard);
    const assessmentId = safeFilePart_(payload.assessmentId || payload.assessmentTitle || 'assessment', 90);
    const assessmentTitle = cleanText_(payload.assessmentTitle, 180) || assessmentId;
    const submissionId = cleanText_(payload.submissionId, 160);
    const recordKey = [teacherId.toLowerCase(), standard.toLowerCase(), studentId.toLowerCase()].join('|');

    const existing = findRegisterRow_(sheet, recordKey);
    if (existing && String(existing.values[13] || '') === submissionId) {
      return {
        state: 'duplicate',
        confirmed: true,
        message: 'This exact submission was already received.',
        submissionId,
        studentFolderUrl: String(existing.values[10] || ''),
        pdfUrl: String(existing.values[11] || ''),
        pukUrl: String(existing.values[12] || ''),
      };
    }

    const teacherFolder = getOrCreateFolder_(root, safeFolderPart_(`${teacherId} - ${teacherName}`, 120));
    const standardFolder = getOrCreateFolder_(teacherFolder, safeFolderPart_(standard, 80));
    const studentFolder = getOrCreateFolder_(standardFolder, safeFolderPart_(`${studentId} - ${studentName}`, 140));

    const base = `${safeFilePart_(studentId, 50)}_${safeFilePart_(studentName, 80)}_${safeFilePart_(standard.replace(/\s+/g, ''), 40)}`;
    const pdfName = `${base}_${assessmentId}.pdf`;
    const pukName = `${base}.puk`;

    const pdfBytes = Utilities.base64Decode(String(payload.pdfBase64 || ''));
    if (!pdfBytes.length) throw new Error('The PDF file was empty.');
    if (pdfBytes.length > CONFIG.MAX_PDF_BYTES) throw new Error('The PDF is too large for this submission gateway.');
    const pukText = String(payload.pukText || '');
    if (!pukText) throw new Error('The .puk backup was empty.');
    if (pukText.length > CONFIG.MAX_PUK_CHARS) throw new Error('The .puk backup is too large for this submission gateway.');

    const pdfFile = replaceFile_(studentFolder, pdfName, Utilities.newBlob(pdfBytes, 'application/pdf', pdfName));
    const pukFile = replaceFile_(studentFolder, pukName, Utilities.newBlob(pukText, 'application/json', pukName));

    const previousCount = existing ? Number(existing.values[9] || 0) : 0;
    const now = new Date();
    const resultText = `${Number(payload.score || 0)}/${Number(payload.totalMarks || 0)} (${Number(payload.percentage || 0)}%)`;
    const row = [
      recordKey,
      studentName,
      studentId,
      teacherName,
      teacherId,
      standard,
      assessmentTitle,
      resultText,
      now,
      previousCount + 1,
      studentFolder.getUrl(),
      pdfFile.getUrl(),
      pukFile.getUrl(),
      submissionId,
      cleanText_(payload.appVersion, 80),
      'Confirmed',
    ];

    if (existing) sheet.getRange(existing.row, 1, 1, HEADERS.length).setValues([row]);
    else sheet.appendRow(row);

    sheet.getRange(2, 9, Math.max(1, sheet.getLastRow() - 1), 1).setNumberFormat('dd/mm/yyyy h:mm am/pm');

    if (CONFIG.NOTIFY_TEACHER) notifyTeacher_(payload, standard, studentFolder, pdfFile, pukFile);

    return {
      state: 'confirmed',
      confirmed: true,
      message: 'PDF and .puk saved and the register was updated.',
      submissionId,
      studentFolderUrl: studentFolder.getUrl(),
      pdfUrl: pdfFile.getUrl(),
      pukUrl: pukFile.getUrl(),
      submissionCount: previousCount + 1,
    };
  } finally {
    lock.releaseLock();
  }
}

function validatePayload_(payload) {
  const appId = cleanText_(payload.appId, 100);
  if (CONFIG.EXPECTED_APP_PREFIX && !appId.startsWith(CONFIG.EXPECTED_APP_PREFIX)) throw new Error('Unrecognised assessment app.');
  ['submissionId', 'studentName', 'studentId', 'teacherName', 'unitStandard', 'assessmentId', 'assessmentTitle', 'pdfBase64', 'pukText']
    .forEach(key => {
      if (!String(payload[key] || '').trim()) throw new Error(`Missing ${key}.`);
    });
  if (!/^\d{3,6}$/.test(String(payload.studentId || '').trim())) throw new Error('Student ID is invalid.');
  const pct = Number(payload.percentage);
  if (!Number.isFinite(pct) || pct < 0 || pct > 100) throw new Error('Percentage is invalid.');
}

function getRootFolder_() {
  if (CONFIG.ROOT_FOLDER_ID) return DriveApp.getFolderById(CONFIG.ROOT_FOLDER_ID);
  const props = PropertiesService.getScriptProperties();
  const savedId = props.getProperty('ROOT_FOLDER_ID');
  if (savedId) {
    try { return DriveApp.getFolderById(savedId); } catch (_) {}
  }
  const existing = DriveApp.getFoldersByName(CONFIG.ROOT_FOLDER_NAME);
  const folder = existing.hasNext() ? existing.next() : DriveApp.createFolder(CONFIG.ROOT_FOLDER_NAME);
  props.setProperty('ROOT_FOLDER_ID', folder.getId());
  return folder;
}

function getRegisterSpreadsheet_(root) {
  if (CONFIG.REGISTER_SPREADSHEET_ID) return SpreadsheetApp.openById(CONFIG.REGISTER_SPREADSHEET_ID);
  const props = PropertiesService.getScriptProperties();
  const savedId = props.getProperty('REGISTER_SPREADSHEET_ID');
  if (savedId) {
    try { return SpreadsheetApp.openById(savedId); } catch (_) {}
  }
  const ss = SpreadsheetApp.create(CONFIG.REGISTER_FILE_NAME);
  try { DriveApp.getFileById(ss.getId()).moveTo(root); } catch (_) {}
  props.setProperty('REGISTER_SPREADSHEET_ID', ss.getId());
  return ss;
}

function getRegisterSheet_(spreadsheet) {
  let sheet = spreadsheet.getSheetByName(CONFIG.REGISTER_SHEET_NAME);
  if (!sheet) {
    sheet = spreadsheet.getSheets()[0];
    sheet.setName(CONFIG.REGISTER_SHEET_NAME);
  }
  const first = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0];
  if (HEADERS.some((header, index) => first[index] !== header)) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, HEADERS.length);
    try { sheet.hideColumns(1); } catch (_) {}
    try { sheet.hideColumns(14); } catch (_) {}
  }
  return sheet;
}

function findRegisterRow_(sheet, recordKey) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  const match = sheet.getRange(2, 1, lastRow - 1, 1)
    .createTextFinder(recordKey)
    .matchEntireCell(true)
    .findNext();
  if (!match) return null;
  const row = match.getRow();
  return { row, values: sheet.getRange(row, 1, 1, HEADERS.length).getValues()[0] };
}

function getOrCreateFolder_(parent, name) {
  const folders = parent.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : parent.createFolder(name);
}

function replaceFile_(folder, name, blob) {
  const existing = folder.getFilesByName(name);
  while (existing.hasNext()) existing.next().setTrashed(true);
  return folder.createFile(blob).setName(name);
}

function normaliseStandard_(value) {
  const match = String(value || '').match(/\d{3,6}/);
  if (!match) throw new Error('Unit standard is invalid.');
  return `US ${match[0]}`;
}

function cleanText_(value, maxLength) {
  return String(value == null ? '' : value).trim().replace(/[\u0000-\u001F\u007F]/g, ' ').replace(/\s+/g, ' ').slice(0, maxLength || 200);
}

function safeFolderPart_(value, maxLength) {
  return cleanText_(value, maxLength || 120).replace(/[<>:"/\\|?*]/g, ' ').replace(/\s+/g, ' ').replace(/[. ]+$/g, '') || 'Folder';
}

function safeFilePart_(value, maxLength) {
  return cleanText_(value, maxLength || 100).replace(/[<>:"/\\|?*]/g, '_').replace(/\s+/g, '_').replace(/_+/g, '_').replace(/^[_ .]+|[_ .]+$/g, '') || 'file';
}

function statusCacheKey_(submissionId) {
  return `status:${submissionId}`;
}

function setStatus_(submissionId, status) {
  CacheService.getScriptCache().put(statusCacheKey_(submissionId), JSON.stringify(status), 21600);
}

function getStatus_(submissionId) {
  if (!submissionId) return { state: 'error', message: 'Missing submissionId.' };
  const cached = CacheService.getScriptCache().get(statusCacheKey_(submissionId));
  if (cached) return JSON.parse(cached);
  try {
    const root = getRootFolder_();
    const ss = getRegisterSpreadsheet_(root);
    const sheet = getRegisterSheet_(ss);
    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      const match = sheet.getRange(2, 14, lastRow - 1, 1).createTextFinder(submissionId).matchEntireCell(true).findNext();
      if (match) return { state: 'confirmed', confirmed: true, submissionId, message: 'Submission is recorded.' };
    }
  } catch (_) {}
  return { state: 'pending', confirmed: false, submissionId };
}

function notifyTeacher_(payload, standard, studentFolder, pdfFile, pukFile) {
  const email = cleanText_(payload.teacherEmail, 180);
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return;
  const subject = `${standard} submission - ${cleanText_(payload.studentName, 120)}`;
  const body = [
    `${cleanText_(payload.studentName, 120)} (${cleanText_(payload.studentId, 60)}) submitted ${standard}.`,
    `Assessment: ${cleanText_(payload.assessmentTitle, 180)}`,
    `Result: ${Number(payload.percentage || 0)}%`,
    `Student folder: ${studentFolder.getUrl()}`,
    `PDF: ${pdfFile.getUrl()}`,
    `PUK: ${pukFile.getUrl()}`,
  ].join('\n');
  MailApp.sendEmail(email, subject, body);
}

function output_(payload, callback) {
  const json = JSON.stringify(payload);
  const safeCallback = /^[A-Za-z_$][0-9A-Za-z_$.]*$/.test(callback || '') ? callback : '';
  const content = safeCallback ? `${safeCallback}(${json});` : json;
  return ContentService.createTextOutput(content)
    .setMimeType(safeCallback ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON);
}
