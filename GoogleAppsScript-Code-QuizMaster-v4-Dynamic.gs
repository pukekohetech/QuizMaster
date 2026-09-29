/**
 * PHS QuizMaster Submission Gateway - v4 Dynamic Destinations
 *
 * The QuizMaster front end supplies a base evidence collection name from
 * submission-settings.json, for example:
 *   PHS Building Evidence
 *
 * The gateway adds the Auckland calendar year automatically and stores work as:
 *   PHS Building Evidence - 2026 /
 *     RNR - Mr Ranson /
 *       123456 /
 *         123456_Student_Name_US24352_learner-self-reflection.pdf
 *         123456_Student_Name_US24352.puk
 *
 * The root folder and register are created lazily on the first real submission.
 * Changing storage.rootName in submission-settings.json automatically selects a
 * different root/register without editing this Apps Script project.
 *
 * Register identity:
 *   Teacher + Unit Standard + Student + Assessment ID
 *
 * Exact duplicate submission IDs do not create duplicate register counts.
 * Expensive Drive work is outside the global script lock so a class can submit
 * concurrently. Only short claim/register operations are serialised.
 */

const CONFIG = {
  REGISTER_FILE_PREFIX: 'Submission Register',
  REGISTER_SHEET_NAME: 'Submissions',
  NOTIFY_TEACHER: false,
  EXPECTED_APP_PREFIX: 'pukekohetech-',
  ALLOWED_ROOT_PREFIX: 'PHS ',
  TIME_ZONE: 'Pacific/Auckland',
  MAX_PDF_BYTES: 25 * 1024 * 1024,
  MAX_PUK_CHARS: 5 * 1024 * 1024,
  CACHE_TTL_SECONDS: 21600,
  DESTINATION_PROPERTY_PREFIX: 'QM_DEST_',
};

const HEADERS = [
  'Record Key',
  'Student',
  'Student ID',
  'Teacher',
  'Teacher ID',
  'Unit Standard',
  'Assessment',
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

/**
 * Run once from the Apps Script editor after replacing Code.gs.
 * This authorises the Drive/Sheets services but intentionally creates no
 * evidence folders or registers. Those are created by the first submission for
 * each storage.rootName + year destination.
 */
function setup() {
  DriveApp.getRootFolder().getId();
  SpreadsheetApp.flush();
  return {
    ok: true,
    state: 'ready',
    message: 'QuizMaster gateway authorised. Evidence folders will be created on first submission.',
    year: currentYear_(),
    notifyTeacher: CONFIG.NOTIFY_TEACHER,
  };
}

function doGet(e) {
  const params = (e && e.parameter) || {};
  const callback = params.callback || '';
  try {
    const action = String(params.action || 'health').toLowerCase();
    if (action === 'status') {
      const fastOnly = String(params.fast || '') === '1';
      const rootName = cleanText_(params.rootName, 100);
      return output_(getStatus_(String(params.submissionId || ''), fastOnly, rootName), callback);
    }

    return output_({
      ok: true,
      state: 'ready',
      message: 'PHS QuizMaster submission gateway is ready.',
      year: currentYear_(),
      notifyTeacher: CONFIG.NOTIFY_TEACHER,
    }, callback);
  } catch (error) {
    return output_({ ok: false, state: 'error', message: String(error && error.message || error) }, callback);
  }
}

function doPost(e) {
  let submissionId = '';
  let claimed = false;
  try {
    if (!e || !e.postData || !e.postData.contents) throw new Error('Missing request body.');
    const payload = JSON.parse(e.postData.contents);
    submissionId = cleanText_(payload.submissionId, 160);
    if (!submissionId) throw new Error('Missing submissionId.');

    const claim = claimSubmission_(submissionId);
    if (!claim.claimed) return output_(claim.status, '');
    claimed = true;

    validatePayload_(payload);
    const result = saveSubmission_(payload);
    setStatus_(submissionId, result);
    return output_(result, '');
  } catch (error) {
    const result = { state: 'error', confirmed: false, message: String(error && error.message || error) };
    if (submissionId && claimed) setStatus_(submissionId, result);
    return output_(result, '');
  }
}

function claimSubmission_(submissionId) {
  const lock = LockService.getScriptLock();
  lock.waitLock(5000);
  try {
    const cache = CacheService.getScriptCache();
    const key = statusCacheKey_(submissionId);
    const raw = cache.get(key);
    if (raw) {
      try {
        const status = JSON.parse(raw);
        if (status && (status.state === 'confirmed' || status.state === 'duplicate')) {
          return {
            claimed: false,
            status: {
              ...status,
              state: 'duplicate',
              confirmed: true,
              message: 'This exact submission was already received.',
            },
          };
        }
        if (status && status.state === 'processing') {
          const ageMs = Date.now() - Number(status.startedAtMs || 0);
          if (ageMs >= 0 && ageMs < 120000) return { claimed: false, status };
        }
      } catch (_) {}
    }

    const status = {
      state: 'processing',
      confirmed: false,
      submissionId,
      startedAtMs: Date.now(),
      message: 'Submission is being saved.',
    };
    cache.put(key, JSON.stringify(status), CONFIG.CACHE_TTL_SECONDS);
    return { claimed: true, status };
  } finally {
    lock.releaseLock();
  }
}

function saveSubmission_(payload) {
  const collectionName = normaliseCollectionName_(payload.storageRootName);
  const destination = getOrCreateDestination_(collectionName);
  const root = destination.root;
  const spreadsheet = destination.spreadsheet;
  const sheet = getRegisterSheet_(spreadsheet);

  const teacherId = cleanText_(payload.teacherId, 40) || 'Teacher';
  const teacherName = cleanText_(payload.teacherName, 100) || teacherId;
  const studentId = cleanText_(payload.studentId, 60);
  const studentName = cleanText_(payload.studentName, 120);
  const standard = normaliseStandard_(payload.unitStandard);
  const assessmentId = safeFilePart_(payload.assessmentId || payload.assessmentTitle || 'assessment', 90);
  const assessmentTitle = cleanText_(payload.assessmentTitle, 180) || assessmentId;
  const submissionId = cleanText_(payload.submissionId, 160);
  const recordKey = [
    teacherId.toLowerCase(),
    standard.toLowerCase(),
    studentId.toLowerCase(),
    assessmentId.toLowerCase(),
  ].join('|');

  const pdfBytes = Utilities.base64Decode(String(payload.pdfBase64 || ''));
  if (!pdfBytes.length) throw new Error('The PDF file was empty.');
  if (pdfBytes.length > CONFIG.MAX_PDF_BYTES) throw new Error('The PDF is too large for this submission gateway.');

  const pukText = String(payload.pukText || '');
  if (!pukText) throw new Error('The .puk backup was empty.');
  if (pukText.length > CONFIG.MAX_PUK_CHARS) throw new Error('The .puk backup is too large for this submission gateway.');

  // Required hierarchy: Collection - Year / Teacher / Student ID
  const teacherFolder = getOrCreateFolder_(root, safeFolderPart_(`${teacherId} - ${teacherName}`, 120));
  const studentFolder = getOrCreateFolder_(teacherFolder, safeFolderPart_(studentId, 60));

  const base = `${safeFilePart_(studentId, 50)}_${safeFilePart_(studentName, 80)}_${safeFilePart_(standard.replace(/\s+/g, ''), 40)}`;
  const pdfName = `${base}_${assessmentId}.pdf`;
  const pukName = `${base}.puk`;

  // File work stays outside the register lock so different students can save in parallel.
  const pdfFile = replaceFile_(studentFolder, pdfName, Utilities.newBlob(pdfBytes, 'application/pdf', pdfName));
  const pukFile = replaceFile_(studentFolder, pukName, Utilities.newBlob(pukText, 'application/json', pukName));

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  let previousCount = 0;
  let exactDuplicate = false;
  let submissionCount = 1;
  try {
    const existing = findRegisterRow_(sheet, recordKey);
    previousCount = existing ? Number(existing.values[9] || 0) : 0;
    exactDuplicate = !!existing && String(existing.values[13] || '') === submissionId;
    submissionCount = exactDuplicate ? Math.max(1, previousCount) : previousCount + 1;

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
      submissionCount,
      studentFolder.getUrl(),
      pdfFile.getUrl(),
      pukFile.getUrl(),
      submissionId,
      cleanText_(payload.appVersion, 80),
      'Confirmed',
    ];

    if (existing) {
      sheet.getRange(existing.row, 1, 1, HEADERS.length).setValues([row]);
      cacheRegisterRow_(sheet, recordKey, existing.row);
    } else {
      sheet.appendRow(row);
      cacheRegisterRow_(sheet, recordKey, sheet.getLastRow());
    }
  } finally {
    lock.releaseLock();
  }

  if (CONFIG.NOTIFY_TEACHER && !exactDuplicate) {
    notifyTeacher_(payload, standard, studentFolder, pdfFile, pukFile);
  }

  return {
    state: exactDuplicate ? 'duplicate' : 'confirmed',
    confirmed: true,
    message: exactDuplicate
      ? 'This exact submission was already received.'
      : 'PDF and .puk saved and the register was updated.',
    submissionId,
    collectionName,
    year: destination.year,
    rootFolderName: destination.rootFolderName,
    rootFolderUrl: root.getUrl(),
    registerUrl: spreadsheet.getUrl(),
    studentFolderUrl: studentFolder.getUrl(),
    pdfUrl: pdfFile.getUrl(),
    pukUrl: pukFile.getUrl(),
    submissionCount,
  };
}

function validatePayload_(payload) {
  const appId = cleanText_(payload.appId, 100);
  if (CONFIG.EXPECTED_APP_PREFIX && !appId.startsWith(CONFIG.EXPECTED_APP_PREFIX)) {
    throw new Error('Unrecognised assessment app.');
  }

  [
    'submissionId',
    'studentName',
    'studentId',
    'teacherName',
    'unitStandard',
    'assessmentId',
    'assessmentTitle',
    'pdfBase64',
    'pukText',
    'storageRootName',
  ].forEach(key => {
    if (!String(payload[key] || '').trim()) throw new Error(`Missing ${key}.`);
  });

  normaliseCollectionName_(payload.storageRootName);

  if (!/^\d{3,6}$/.test(String(payload.studentId || '').trim())) {
    throw new Error('Student ID is invalid.');
  }

  const pct = Number(payload.percentage);
  if (!Number.isFinite(pct) || pct < 0 || pct > 100) throw new Error('Percentage is invalid.');
}

function currentYear_() {
  return Utilities.formatDate(new Date(), CONFIG.TIME_ZONE, 'yyyy');
}

function normaliseCollectionName_(value) {
  let name = cleanText_(value, 100);
  name = name.replace(/\s+-\s+\d{4}\s*$/, '').trim();
  name = safeFolderPart_(name, 90);
  if (!name) throw new Error('Evidence collection name is missing.');
  if (CONFIG.ALLOWED_ROOT_PREFIX && !name.startsWith(CONFIG.ALLOWED_ROOT_PREFIX)) {
    throw new Error(`Evidence collection name must begin with "${CONFIG.ALLOWED_ROOT_PREFIX}".`);
  }
  return name;
}

function destinationPropertyKey_(collectionName, year) {
  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    `${collectionName.toLowerCase()}|${year}`
  );
  const token = Utilities.base64EncodeWebSafe(digest).replace(/=+$/g, '').slice(0, 40);
  return `${CONFIG.DESTINATION_PROPERTY_PREFIX}${token}`;
}

function readDestinationRecord_(collectionName, year) {
  const raw = PropertiesService.getScriptProperties().getProperty(destinationPropertyKey_(collectionName, year));
  if (!raw) return null;
  try {
    const record = JSON.parse(raw);
    if (!record || !record.rootFolderId || !record.registerSpreadsheetId) return null;
    return record;
  } catch (_) {
    return null;
  }
}

function openDestinationRecord_(record, collectionName, year) {
  if (!record) return null;
  try {
    const root = DriveApp.getFolderById(record.rootFolderId);
    const spreadsheet = SpreadsheetApp.openById(record.registerSpreadsheetId);
    return {
      root,
      spreadsheet,
      collectionName,
      year,
      rootFolderName: `${collectionName} - ${year}`,
      registerFileName: `${CONFIG.REGISTER_FILE_PREFIX} - ${year}`,
    };
  } catch (_) {
    return null;
  }
}

function getExistingDestination_(collectionName) {
  const normalised = normaliseCollectionName_(collectionName);
  const year = currentYear_();
  return openDestinationRecord_(readDestinationRecord_(normalised, year), normalised, year);
}

function getOrCreateDestination_(collectionName) {
  const normalised = normaliseCollectionName_(collectionName);
  const year = currentYear_();

  const existing = openDestinationRecord_(readDestinationRecord_(normalised, year), normalised, year);
  if (existing) return existing;

  // Only the rare first-use/repair path is locked. Routine submissions skip this.
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const afterLock = openDestinationRecord_(readDestinationRecord_(normalised, year), normalised, year);
    if (afterLock) return afterLock;

    const driveRoot = DriveApp.getRootFolder();
    const rootFolderName = `${normalised} - ${year}`;
    const registerFileName = `${CONFIG.REGISTER_FILE_PREFIX} - ${year}`;

    const roots = driveRoot.getFoldersByName(rootFolderName);
    const root = roots.hasNext() ? roots.next() : driveRoot.createFolder(rootFolderName);

    let spreadsheet = null;
    const registerFiles = root.getFilesByName(registerFileName);
    while (registerFiles.hasNext()) {
      const file = registerFiles.next();
      if (file.getMimeType() === MimeType.GOOGLE_SHEETS) {
        try {
          spreadsheet = SpreadsheetApp.openById(file.getId());
          break;
        } catch (_) {}
      }
    }

    if (!spreadsheet) {
      spreadsheet = SpreadsheetApp.create(registerFileName);
      DriveApp.getFileById(spreadsheet.getId()).moveTo(root);
    }

    getRegisterSheet_(spreadsheet);

    const record = {
      collectionName: normalised,
      year,
      rootFolderId: root.getId(),
      registerSpreadsheetId: spreadsheet.getId(),
      rootFolderName,
      registerFileName,
    };
    PropertiesService.getScriptProperties().setProperty(
      destinationPropertyKey_(normalised, year),
      JSON.stringify(record)
    );

    return { root, spreadsheet, collectionName: normalised, year, rootFolderName, registerFileName };
  } finally {
    lock.releaseLock();
  }
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
    sheet.getRange('I:I').setNumberFormat('dd/mm/yyyy h:mm am/pm');
    try { sheet.hideColumns(1); } catch (_) {}
    try { sheet.hideColumns(14); } catch (_) {}
  }
  return sheet;
}

function registerRowCacheKey_(sheet, recordKey) {
  const spreadsheetId = sheet.getParent().getId();
  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.MD5,
    `${spreadsheetId}|${recordKey}`
  );
  return `row:${Utilities.base64EncodeWebSafe(digest).replace(/=+$/g, '')}`;
}

function cacheRegisterRow_(sheet, recordKey, row) {
  CacheService.getScriptCache().put(
    registerRowCacheKey_(sheet, recordKey),
    String(row),
    CONFIG.CACHE_TTL_SECONDS
  );
}

function findRegisterRow_(sheet, recordKey) {
  const cache = CacheService.getScriptCache();
  const cachedRow = Number(cache.get(registerRowCacheKey_(sheet, recordKey)) || 0);
  if (cachedRow >= 2 && cachedRow <= sheet.getLastRow()) {
    const values = sheet.getRange(cachedRow, 1, 1, HEADERS.length).getValues()[0];
    if (String(values[0] || '') === recordKey) return { row: cachedRow, values };
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  const match = sheet.getRange(2, 1, lastRow - 1, 1)
    .createTextFinder(recordKey)
    .matchEntireCell(true)
    .findNext();
  if (!match) return null;

  const row = match.getRow();
  cacheRegisterRow_(sheet, recordKey, row);
  return { row, values: sheet.getRange(row, 1, 1, HEADERS.length).getValues()[0] };
}

function folderCacheKey_(parentId, name) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, `${parentId}|${name}`);
  return `folder:${Utilities.base64EncodeWebSafe(digest).replace(/=+$/g, '')}`;
}

function getOrCreateFolder_(parent, name) {
  const cache = CacheService.getScriptCache();
  const key = folderCacheKey_(parent.getId(), name);
  const cachedId = cache.get(key);
  if (cachedId) {
    try { return DriveApp.getFolderById(cachedId); } catch (_) {}
  }

  const folders = parent.getFoldersByName(name);
  const folder = folders.hasNext() ? folders.next() : parent.createFolder(name);
  cache.put(key, folder.getId(), CONFIG.CACHE_TTL_SECONDS);
  return folder;
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
  return String(value == null ? '' : value)
    .trim()
    .replace(/[\u0000-\u001F\u007F]/g, ' ')
    .replace(/\s+/g, ' ')
    .slice(0, maxLength || 200);
}

function safeFolderPart_(value, maxLength) {
  return cleanText_(value, maxLength || 120)
    .replace(/[<>:"/\\|?*]/g, ' ')
    .replace(/\s+/g, ' ')
    .replace(/[. ]+$/g, '') || 'Folder';
}

function safeFilePart_(value, maxLength) {
  return cleanText_(value, maxLength || 100)
    .replace(/[<>:"/\\|?*]/g, '_')
    .replace(/\s+/g, '_')
    .replace(/_+/g, '_')
    .replace(/^[_ .]+|[_ .]+$/g, '') || 'file';
}

function statusCacheKey_(submissionId) {
  return `status:${submissionId}`;
}

function setStatus_(submissionId, status) {
  CacheService.getScriptCache().put(
    statusCacheKey_(submissionId),
    JSON.stringify(status),
    CONFIG.CACHE_TTL_SECONDS
  );
}

function getStatus_(submissionId, fastOnly, rootName) {
  if (!submissionId) return { state: 'error', message: 'Missing submissionId.' };

  const cached = CacheService.getScriptCache().get(statusCacheKey_(submissionId));
  if (cached) return JSON.parse(cached);

  if (fastOnly) return { state: 'pending', confirmed: false, submissionId };

  // Recovery lookup after a cache miss. The front end supplies the same rootName
  // used for the submission so only that register needs to be searched.
  if (!rootName) return { state: 'pending', confirmed: false, submissionId };

  try {
    const destination = getExistingDestination_(rootName);
    if (!destination) return { state: 'pending', confirmed: false, submissionId };
    const sheet = getRegisterSheet_(destination.spreadsheet);
    const lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      const match = sheet.getRange(2, 14, lastRow - 1, 1)
        .createTextFinder(submissionId)
        .matchEntireCell(true)
        .findNext();
      if (match) {
        return {
          state: 'confirmed',
          confirmed: true,
          submissionId,
          collectionName: destination.collectionName,
          year: destination.year,
          message: 'Submission is recorded.',
        };
      }
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
