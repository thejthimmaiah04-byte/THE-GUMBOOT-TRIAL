// ===== TLT GUMBOOT TRIAL — Google Apps Script Backend =====
// Bind this script to a Google Sheet. Deploy as a web app.
// Sheets "Trials", "Boots", "Snakes" are auto-created on first use.
// A Google Drive folder "TLT_Gumboot_Trial_Photos" is auto-created for images.
//
// The front-end is a static PWA hosted on GitHub Pages, not served from here.
// Apps Script web apps don't send Access-Control-Allow-Origin headers, so a
// cross-origin fetch() from the GitHub Pages app would be blocked by CORS.
// Instead:
//   - Reads (doGet) support JSONP: ?action=...&callback=fn — <script src>
//     loads aren't subject to CORS, so this works cleanly.
//   - Writes (doPost) are submitted via a hidden iframe + HTML <form> POST
//     (also exempt from CORS). The response is a tiny HTML page whose inline
//     script calls parent.postMessage(...) to hand the result back to the
//     static page — postMessage is explicitly designed to cross origins.

// ---------- Reads (JSONP) ----------

function doGet(e) {
  var action = e.parameter.action;
  var callback = e.parameter.callback;
  var result;
  try {
    if (action === 'getBoots') result = getBoots();
    else if (action === 'getSnakes') result = getSnakes();
    else if (action === 'getAllTrials') result = getAllTrials();
    else result = { error: 'Unknown action: ' + action };
  } catch (err) {
    result = { error: friendlyError_(err) };
  }
  if (callback) {
    return ContentService.createTextOutput(callback + '(' + JSON.stringify(result) + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---------- Writes (hidden-iframe form POST + postMessage) ----------

function doPost(e) {
  var reqId = e.parameter.reqId || '';
  var action = e.parameter.action;
  var data;
  try {
    data = JSON.parse(e.parameter.data || '{}');
  } catch (err) {
    return postMessageResponse_(reqId, { success: false, message: 'Invalid request data' });
  }
  var result;
  try {
    if (action === 'registerBoot') result = registerBoot(data);
    else if (action === 'registerSnake') result = registerSnake(data);
    else if (action === 'submitTrial') result = submitTrial(data);
    else if (action === 'updateBoot') result = updateBoot(data);
    else if (action === 'updateSnake') result = updateSnake(data);
    else result = { success: false, message: 'Unknown action: ' + action };
  } catch (err) {
    result = { success: false, message: friendlyError_(err) };
  }
  return postMessageResponse_(reqId, result);
}

// Apps Script's own error text for lock timeouts / quota limits is cryptic
// ("Exception: Lock timeout...", raw quota errors). Translate the ones a
// field researcher might actually hit into something actionable.
function friendlyError_(err) {
  var msg = (err && err.message) || String(err);
  if (/lock/i.test(msg)) {
    return 'Server was busy handling another submission — please try again in a few seconds.';
  }
  if (/quota|limit/i.test(msg)) {
    return 'Google service limit reached — please wait a minute and try again.';
  }
  return msg;
}

function postMessageResponse_(reqId, result) {
  var payload = JSON.stringify({ reqId: reqId, result: result });
  // Apps Script nests this response inside its own wrapper + sandboxFrame,
  // so top (not parent) is needed to reach the actual embedding page.
  var html = '<script>top.postMessage(' + payload + ', "*");<\/script>';
  return HtmlService.createHtmlOutput(html)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ---------- Sheet helpers ----------

var TRIAL_HEADERS = [
  'Trial_ID','Date','Session','Boot_ID','Boot_Brand','Boot_Size',
  'Boot_Thick_T','Boot_Thick_LM','Boot_Thick_I',
  'Snake_ID','Species_Code','Age_Class','Sex',
  'TL_mm','SVL_mm','HL_mm','HW_mm','BM_g',
  'FL_L_mm','FL_R_mm','Dentition',
  'Outcome','Strikes','Region_Struck',
  'Temp_C','Humidity_pct','Photo_Links','Notes','Timestamp',
  'Trial_Start_Time','Trial_Stop_Time','Trial_Duration_Sec','Recorded_By',
  'Trial_Time','Bites','Region_Outcomes'
];

// Boot_ID format: Brand_Abbr + Sole_Color + Rubber_Color + Pair_Number +
// Side (L/R) — each physical boot is its own registered record, not the
// pair, so a pair produces two rows (...-01-L and ...-01-R).
var BOOT_HEADERS = [
  'Boot_ID','Brand_Abbr','Brand_Full','Model','Boot_Size','Session',
  'IS_Standard','Mfg_Date','Batch_No',
  'Thick_T_mm','Thick_LM_mm','Thick_I_mm',
  'Photo_Links','Registered_Date','Recorded_By','Model_Abbr',
  'Last_Edited_By','Last_Edited_Date',
  'Cost_Per_Pair','Bill_Photo_Link','Store_Name','Store_Location',
  'Sole_Color','Rubber_Color','Pair_Number','Side'
];

var SNAKE_HEADERS = [
  'Snake_ID','Species_Code','Common_Name','Age_Class','Sex',
  'TL_mm','SVL_mm','HL_mm','HW_mm','BM_g',
  'FL_L_mm','FL_R_mm','Dentition',
  'Body_Condition','Last_Feed_Date','Registered_Date',
  'Iso_In_Time','Iso_Out_Time','Time_To_Unconscious_Sec','Recorded_By',
  'Last_Edited_By','Last_Edited_Date','Origin'
];

// Single source of truth for species metadata — used to build both the name
// lookup and the dentition lookup, so there's only one place to edit when a
// species is added. NOTE: the <select id="s-species"> options in index.html
// must be kept in sync with these codes by hand (a static HTML page can't
// import this file directly) — if you add a species here, add the matching
// <option> there too.
var SPECIES_DEFS = [
  { code: 'DR', name: "Russell's viper", dentition: 'S' },
  { code: 'NN', name: 'Spectacled cobra', dentition: 'P' },
  { code: 'BC', name: 'Common krait', dentition: 'P' },
  { code: 'EC', name: 'Saw-scaled viper', dentition: 'S' }
];

function speciesName_(code) {
  for (var i = 0; i < SPECIES_DEFS.length; i++) {
    if (SPECIES_DEFS[i].code === code) return SPECIES_DEFS[i].name;
  }
  return '';
}

function speciesDentition_(code) {
  for (var i = 0; i < SPECIES_DEFS.length; i++) {
    if (SPECIES_DEFS[i].code === code) return SPECIES_DEFS[i].dentition;
  }
  return '';
}

// Auto-heals an existing sheet's header row against the current HEADERS
// array whenever a schema change adds a new column, instead of requiring
// someone to manually retype it into the live Sheet after every redeploy
// (the recurring step this project kept needing). Only ever APPENDS
// headers the sheet doesn't already have — it never reorders or removes
// an existing one, since every read/write in this file addresses columns
// by fixed position and a reorder would silently corrupt every row.
function getOrCreateSheet_(name, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
    return sheet;
  }

  var lastCol = sheet.getLastColumn();
  var existingHeaders = lastCol > 0 ? sheet.getRange(1, 1, 1, lastCol).getValues()[0] : [];
  var missingHeaders = headers.filter(function(h) { return existingHeaders.indexOf(h) === -1; });
  if (missingHeaders.length > 0) {
    var startCol = existingHeaders.length + 1;
    var range = sheet.getRange(1, startCol, 1, missingHeaders.length);
    range.setValues([missingHeaders]);
    range.setFontWeight('bold');
  }
  return sheet;
}

// Google Sheets treats a leading =, +, -, or @ as the start of a formula.
// Free-text fields (notes, brand names, batch numbers...) come straight from
// user input, so a value like "-5 boots left" would otherwise silently turn
// into a broken formula. A leading apostrophe forces plain text.
function sanitizeCell_(val) {
  if (typeof val === 'string' && /^[=+\-@]/.test(val)) {
    return "'" + val;
  }
  return val;
}

function missingFields_(data, required) {
  var missing = [];
  required.forEach(function(f) {
    var v = data[f.key];
    if (v === undefined || v === null || String(v).trim() === '') missing.push(f.label);
  });
  return missing;
}

// ---------- Data retrieval ----------

function getBoots() {
  var sheet = getOrCreateSheet_('Boots', BOOT_HEADERS);
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  var result = [];
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    result.push({
      bootId: r[0], brandAbbr: r[1], brandFull: r[2], model: r[3],
      size: r[4], session: r[5], isStandard: r[6], mfgDate: r[7], batch: r[8],
      thickT: r[9], thickLM: r[10], thickI: r[11], photoLinks: r[12],
      recordedBy: r[14], modelAbbr: r[15],
      costPerPair: r[18], billPhotoLink: r[19], storeName: r[20], storeLocation: r[21],
      soleColor: r[22], rubberColor: r[23], pairNumber: r[24], side: r[25]
    });
  }
  return result;
}

function getSnakes() {
  var sheet = getOrCreateSheet_('Snakes', SNAKE_HEADERS);
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  var result = [];
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    result.push({
      snakeId: r[0], speciesCode: r[1], commonName: r[2], ageClass: r[3],
      sex: r[4], tl: r[5], svl: r[6], hl: r[7], hw: r[8], bm: r[9],
      flL: r[10], flR: r[11], dentition: r[12], bodyCondition: r[13],
      lastFeedDate: r[14],
      isoInTime: r[16], isoOutTime: r[17], timeToUnconsciousSec: r[18],
      recordedBy: r[19], origin: r[22]
    });
  }
  return result;
}

function formatDateStr_(val) {
  // Duck-type instead of `instanceof Date` — values read back from Sheets
  // cells don't reliably pass instanceof checks, which was silently falling
  // through to a raw String(date) like "Mon Aug 03 2026 00:00:00 GMT+0530...".
  if (val && typeof val.getFullYear === 'function') {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(val || '');
}

function todayDateStr_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd');
}

function getAllTrials() {
  var sheet = getOrCreateSheet_('Trials', TRIAL_HEADERS);
  var data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  var result = [];
  for (var i = 1; i < data.length; i++) {
    var r = data[i];
    if (!r[0]) continue;
    result.push({
      trialId: r[0], date: formatDateStr_(r[1]), session: r[2],
      bootId: r[3], bootBrand: r[4], bootSize: r[5],
      bootThickT: r[6], bootThickLM: r[7], bootThickI: r[8],
      snakeId: r[9], speciesCode: r[10], ageClass: r[11], sex: r[12],
      tl: r[13], svl: r[14], hl: r[15], hw: r[16], bm: r[17],
      flL: r[18], flR: r[19], dentition: r[20],
      outcome: r[21], strikes: r[22], regionStruck: r[23],
      tempC: r[24], humidityPct: r[25], photoLinks: r[26], notes: r[27],
      trialStartTime: r[29], trialStopTime: r[30], trialDurationSec: r[31],
      recordedBy: r[32], trialTime: r[33], bites: r[34], regionOutcomes: r[35]
    });
  }
  return result;
}

// ---------- Drive / image upload ----------
// Everything lives under one main study folder, organized as:
//   <Study Folder>/Boot Registration/<Boot_ID>/...   (sole, rubber, bill photos)
//   <Study Folder>/Trial Images/<Trial_ID>/...        (balloon, region, additional photos)
var STUDY_FOLDER_NAME = 'efficacy study of gumboots against snakebite';
var BOOT_REG_SUBFOLDER_NAME = 'Boot Registration';
var TRIAL_IMAGES_SUBFOLDER_NAME = 'Trial Images';

// Scoped to a specific parent folder (or Drive root when parent is null),
// so an unrelated folder elsewhere in Drive that happens to share a name
// is never mistaken for the right one — unlike a bare
// DriveApp.getFoldersByName, which searches the whole Drive regardless of
// location.
function getOrCreateChildFolder_(parentFolder, name) {
  var existing = parentFolder ? parentFolder.getFoldersByName(name) : DriveApp.getFoldersByName(name);
  if (existing.hasNext()) return existing.next();
  return parentFolder ? parentFolder.createFolder(name) : DriveApp.createFolder(name);
}

function getStudyRootFolder_() {
  return getOrCreateChildFolder_(null, STUDY_FOLDER_NAME);
}

// Boot Registration/<Boot_ID> — every photo for that boot (sole, rubber,
// bill/receipt) lives together in its own folder.
function getBootFolder_(bootId) {
  var bootRegFolder = getOrCreateChildFolder_(getStudyRootFolder_(), BOOT_REG_SUBFOLDER_NAME);
  return getOrCreateChildFolder_(bootRegFolder, bootId);
}

// Trial Images/<Trial_ID> — every photo for that trial (balloon, per-region,
// additional) lives together in its own folder.
function getTrialFolder_(trialId) {
  var trialImagesFolder = getOrCreateChildFolder_(getStudyRootFolder_(), TRIAL_IMAGES_SUBFOLDER_NAME);
  return getOrCreateChildFolder_(trialImagesFolder, trialId);
}

function uploadImageToFolder_(folder, base64Data, fileName, mimeType) {
  var decoded = Utilities.base64Decode(base64Data);
  var blob = Utilities.newBlob(decoded, mimeType || 'image/jpeg', fileName);
  var file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}

function uploadBootPhotos_(photos, bootId) {
  var links = [];
  if (!photos || photos.length === 0) return links;
  var folder = getBootFolder_(bootId);
  for (var i = 0; i < photos.length; i++) {
    if (photos[i] && photos[i].data) {
      try {
        var fileName = bootId + '_' + photos[i].label + '.jpg';
        var url = uploadImageToFolder_(folder, photos[i].data, fileName, 'image/jpeg');
        links.push(photos[i].label + ': ' + url);
      } catch (e) {
        Logger.log('Photo upload failed for ' + bootId + ': ' + e.message);
        links.push(photos[i].label + ': UPLOAD_FAILED');
      }
    }
  }
  return links;
}

// Bill/receipt photo lives in the same per-boot folder as the rest of that
// boot's registration photos — it's tracked in its own sheet column rather
// than mixed into Photo_Links, since it documents the purchase rather than
// the gear itself, but there's no reason for it to live in a separate part
// of Drive.
function uploadBootBillPhoto_(photo, bootId) {
  if (!photo || !photo.data) return '';
  try {
    return uploadImageToFolder_(getBootFolder_(bootId), photo.data, bootId + '_bill.jpg', 'image/jpeg');
  } catch (e) {
    Logger.log('Bill photo upload failed for ' + bootId + ': ' + e.message);
    return 'UPLOAD_FAILED';
  }
}

// Balloon gets a fixed descriptive name; a region photo (label
// "region_<Name>", e.g. "region_Heel") is named with the trial ID and
// struck region rather than just "additional", since the point is to be
// unambiguous about which region it documents even outside the per-trial
// folder; any further photos are all tagged "additional" and numbered here
// in upload order (1, 2, 3...) since a trial can have more than one.
function uploadTrialPhotos_(photos, trialId) {
  var links = [];
  if (!photos || photos.length === 0) return links;
  var folder = getTrialFolder_(trialId);
  var additionalCount = 0;
  for (var i = 0; i < photos.length; i++) {
    var p = photos[i];
    if (!p || !p.data) continue;
    var niceName;
    if (p.label === 'balloon') niceName = trialId + '_Balloon photo';
    else if (p.label && p.label.indexOf('region_') === 0) {
      niceName = trialId + '_' + p.label.slice('region_'.length);
    }
    else { additionalCount++; niceName = trialId + '_Additional photo ' + additionalCount; }
    try {
      var url = uploadImageToFolder_(folder, p.data, niceName + '.jpg', 'image/jpeg');
      links.push(niceName + ': ' + url);
    } catch (e) {
      Logger.log('Trial photo upload failed for ' + trialId + ': ' + e.message);
      links.push(niceName + ': UPLOAD_FAILED');
    }
  }
  return links;
}

// ---------- Registration ----------
// Each of these wraps its duplicate-ID check + append in a script lock.
// Without this, two near-simultaneous submissions (two researchers, or a
// double-tap) can both read "not a duplicate" before either has written its
// row, producing two rows with the same ID. The lock makes check-then-append
// atomic across concurrent requests. Everything after the lock is acquired
// runs exactly as it did before — this only serializes access, it doesn't
// change what a single request does.

function registerBoot(data) {
  var missing = missingFields_(data, [
    { key: 'bootId', label: 'Boot ID' },
    { key: 'brandAbbr', label: 'Brand Abbr' },
    { key: 'soleColor', label: 'Sole Color' },
    { key: 'rubberColor', label: 'Rubber Color' },
    { key: 'pairNumber', label: 'Pair Number' },
    { key: 'side', label: 'Left/Right' },
    { key: 'brandFull', label: 'Brand Name' }
  ]);
  if (missing.length) {
    return { success: false, message: 'Missing: ' + missing.join(', ') };
  }

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    return { success: false, message: 'Server busy — please try again in a moment.' };
  }
  try {
    var sheet = getOrCreateSheet_('Boots', BOOT_HEADERS);
    var existing = sheet.getDataRange().getValues();
    for (var i = 1; i < existing.length; i++) {
      if (existing[i][0] === data.bootId) {
        return { success: false, message: 'Boot ID already exists: ' + data.bootId };
      }
    }

    var bootPhotos = (data.photos || []).filter(function(p) { return p && p.label !== 'bill'; });
    var billPhoto = (data.photos || []).filter(function(p) { return p && p.label === 'bill'; })[0];
    var photoLinks = uploadBootPhotos_(bootPhotos, data.bootId);
    var billLink = uploadBootBillPhoto_(billPhoto, data.bootId);

    sheet.appendRow([
      data.bootId, data.brandAbbr, sanitizeCell_(data.brandFull), sanitizeCell_(data.model),
      data.size, data.session, sanitizeCell_(data.isStandard), sanitizeCell_(data.mfgDate), sanitizeCell_(data.batch),
      data.thickT, data.thickLM, data.thickI,
      photoLinks.join('\n'), todayDateStr_(), sanitizeCell_(data.recordedBy), data.modelAbbr,
      '', '',
      data.costPerPair, billLink, sanitizeCell_(data.storeName), sanitizeCell_(data.storeLocation),
      data.soleColor, data.rubberColor, data.pairNumber, data.side
    ]);

    return { success: true, message: 'Boot registered: ' + data.bootId };
  } finally {
    lock.releaseLock();
  }
}

// Only Brand_Abbr is immutable on edit (the front-end disables that field)
// since it's the true anchor of a boot's identity — changing it would mean
// renaming the row's key, not editing its data. Everything else, including
// the newer ID components (Sole_Color/Rubber_Color/Pair_Number/Side) and
// Model_Abbr/Boot_Size, stays editable even for an already-registered
// boot: several of these fields were added after some boots already
// existed, so a legacy record needs to be completable without that
// meaning a rename (Boot_ID itself is looked up by the pre-existing value
// and never recomputed here, regardless of what the other fields say).
// Photos/bill photo are only touched if a new one is attached, so editing
// other fields never wipes out previously uploaded links.
function updateBoot(data) {
  var missing = missingFields_(data, [
    { key: 'bootId', label: 'Boot ID' },
    { key: 'brandFull', label: 'Brand Name' }
  ]);
  if (missing.length) {
    return { success: false, message: 'Missing: ' + missing.join(', ') };
  }

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    return { success: false, message: 'Server busy — please try again in a moment.' };
  }
  try {
    var sheet = getOrCreateSheet_('Boots', BOOT_HEADERS);
    var values = sheet.getDataRange().getValues();
    var rowIdx = -1;
    for (var i = 1; i < values.length; i++) {
      if (values[i][0] === data.bootId) { rowIdx = i; break; }
    }
    if (rowIdx === -1) {
      return { success: false, message: 'Boot ID not found (may have been removed): ' + data.bootId };
    }

    var newRow = values[rowIdx].slice();
    while (newRow.length < BOOT_HEADERS.length) newRow.push('');

    var bootPhotos = (data.photos || []).filter(function(p) { return p && p.label !== 'bill'; });
    var billPhoto = (data.photos || []).filter(function(p) { return p && p.label === 'bill'; })[0];
    if (bootPhotos.length) {
      var uploaded = uploadBootPhotos_(bootPhotos, data.bootId);
      if (uploaded.length) newRow[12] = uploaded.join('\n');
    }
    if (billPhoto) {
      var billLink = uploadBootBillPhoto_(billPhoto, data.bootId);
      if (billLink) newRow[19] = billLink;
    }
    newRow[2] = sanitizeCell_(data.brandFull);
    newRow[3] = sanitizeCell_(data.model);
    newRow[6] = sanitizeCell_(data.isStandard);
    newRow[7] = sanitizeCell_(data.mfgDate);
    newRow[8] = sanitizeCell_(data.batch);
    newRow[9] = data.thickT;
    newRow[10] = data.thickLM;
    newRow[11] = data.thickI;
    if (data.modelAbbr) newRow[15] = data.modelAbbr;
    newRow[16] = sanitizeCell_(data.recordedBy);
    newRow[17] = todayDateStr_();
    newRow[18] = data.costPerPair;
    newRow[20] = sanitizeCell_(data.storeName);
    newRow[21] = sanitizeCell_(data.storeLocation);
    if (data.size) newRow[4] = data.size;
    if (data.soleColor) newRow[22] = data.soleColor;
    if (data.rubberColor) newRow[23] = data.rubberColor;
    if (data.pairNumber) newRow[24] = data.pairNumber;
    if (data.side) newRow[25] = data.side;

    sheet.getRange(rowIdx + 1, 1, 1, BOOT_HEADERS.length).setValues([newRow]);
    return { success: true, message: 'Boot updated: ' + data.bootId };
  } finally {
    lock.releaseLock();
  }
}

function registerSnake(data) {
  var missing = missingFields_(data, [
    { key: 'snakeId', label: 'Snake ID' },
    { key: 'speciesCode', label: 'Species' },
    { key: 'ageClass', label: 'Age Class' }
  ]);
  if (missing.length) {
    return { success: false, message: 'Missing: ' + missing.join(', ') };
  }

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    return { success: false, message: 'Server busy — please try again in a moment.' };
  }
  try {
    var sheet = getOrCreateSheet_('Snakes', SNAKE_HEADERS);
    var existing = sheet.getDataRange().getValues();
    for (var i = 1; i < existing.length; i++) {
      if (existing[i][0] === data.snakeId) {
        return { success: false, message: 'Snake ID already exists: ' + data.snakeId };
      }
    }

    // A leading apostrophe forces Sheets to store these as plain text instead
    // of auto-detecting "yyyy-MM-dd HH:mm:ss" as a real date/time and silently
    // converting the cell — which would turn it back into a raw ISO string
    // (with T/Z/milliseconds) the next time it's read back via the API.
    sheet.appendRow([
      data.snakeId, data.speciesCode, speciesName_(data.speciesCode),
      data.ageClass, data.sex,
      data.tl, data.svl, data.hl, data.hw, data.bm,
      data.flL, data.flR, speciesDentition_(data.speciesCode),
      data.bodyCondition, sanitizeCell_(data.lastFeedDate),
      todayDateStr_(),
      data.isoInTime ? "'" + data.isoInTime : '',
      data.isoOutTime ? "'" + data.isoOutTime : '',
      data.timeToUnconsciousSec || '',
      sanitizeCell_(data.recordedBy),
      '', '',
      data.origin || ''
    ]);

    return { success: true, message: 'Snake registered: ' + data.snakeId };
  } finally {
    lock.releaseLock();
  }
}

// Snake_ID, Species_Code, and Age_Class are immutable on edit (the front-end
// disables those fields) since they compose the ID.
function updateSnake(data) {
  var missing = missingFields_(data, [
    { key: 'snakeId', label: 'Snake ID' }
  ]);
  if (missing.length) {
    return { success: false, message: 'Missing: ' + missing.join(', ') };
  }

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    return { success: false, message: 'Server busy — please try again in a moment.' };
  }
  try {
    var sheet = getOrCreateSheet_('Snakes', SNAKE_HEADERS);
    var values = sheet.getDataRange().getValues();
    var rowIdx = -1;
    for (var i = 1; i < values.length; i++) {
      if (values[i][0] === data.snakeId) { rowIdx = i; break; }
    }
    if (rowIdx === -1) {
      return { success: false, message: 'Snake ID not found (may have been removed): ' + data.snakeId };
    }

    var newRow = values[rowIdx].slice();
    while (newRow.length < SNAKE_HEADERS.length) newRow.push('');

    newRow[4] = data.sex;
    newRow[5] = data.tl;
    newRow[6] = data.svl;
    newRow[7] = data.hl;
    newRow[8] = data.hw;
    newRow[9] = data.bm;
    newRow[10] = data.flL;
    newRow[11] = data.flR;
    newRow[13] = data.bodyCondition;
    newRow[14] = sanitizeCell_(data.lastFeedDate);
    if (data.isoInTime) newRow[16] = "'" + data.isoInTime;
    if (data.isoOutTime) newRow[17] = "'" + data.isoOutTime;
    if (data.timeToUnconsciousSec) newRow[18] = data.timeToUnconsciousSec;
    newRow[20] = sanitizeCell_(data.recordedBy);
    newRow[21] = todayDateStr_();
    if (data.origin) newRow[22] = data.origin;

    sheet.getRange(rowIdx + 1, 1, 1, SNAKE_HEADERS.length).setValues([newRow]);
    return { success: true, message: 'Snake updated: ' + data.snakeId };
  } finally {
    lock.releaseLock();
  }
}

// ---------- Trial submission ----------

function submitTrial(data) {
  var missing = missingFields_(data, [
    { key: 'trialId', label: 'Trial ID' },
    { key: 'date', label: 'Date' },
    { key: 'bootId', label: 'Boot' },
    { key: 'snakeId', label: 'Snake' },
    { key: 'outcome', label: 'Outcome' }
  ]);
  if (missing.length) {
    return { success: false, message: 'Missing: ' + missing.join(', ') };
  }

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    return { success: false, message: 'Server busy — please try again in a moment.' };
  }
  try {
    var sheet = getOrCreateSheet_('Trials', TRIAL_HEADERS);
    var existing = sheet.getDataRange().getValues();
    for (var i = 1; i < existing.length; i++) {
      if (existing[i][0] === data.trialId) {
        return { success: false, message: 'Trial already recorded: ' + data.trialId + '. Use a unique combination.' };
      }
    }

    var photoLinks = uploadTrialPhotos_(data.photos, data.trialId);

    // Leading apostrophe forces plain text so Sheets doesn't auto-detect
    // "HH:MM:SS" as a real time value and corrupt it on the next read-back
    // (same issue fixed earlier for the isoflurane in/out timestamps).
    sheet.appendRow([
      data.trialId, data.date, data.session,
      data.bootId, sanitizeCell_(data.bootBrand), data.bootSize,
      data.bootThickT, data.bootThickLM, data.bootThickI,
      data.snakeId, data.speciesCode, data.ageClass, data.sex,
      data.tl, data.svl, data.hl, data.hw, data.bm,
      data.flL, data.flR, data.dentition,
      data.outcome, data.strikes, data.regionStruck,
      data.tempC, data.humidityPct,
      photoLinks.join('\n'), sanitizeCell_(data.notes),
      new Date().toISOString(),
      data.trialStartTime ? "'" + data.trialStartTime : '',
      data.trialStopTime ? "'" + data.trialStopTime : '',
      data.trialDurationSec || '',
      sanitizeCell_(data.recordedBy),
      // Same leading-apostrophe guard as isoInTime itself — this is a copy
      // of that same plain-text "yyyy-MM-dd HH:mm:ss" string, and would
      // otherwise get silently reinterpreted as a real date/time cell.
      data.trialTime ? "'" + data.trialTime : '',
      data.bites,
      sanitizeCell_(data.regionOutcomes)
    ]);

    return { success: true, message: 'Trial recorded: ' + data.trialId };
  } finally {
    lock.releaseLock();
  }
}

// ---------- Backup (optional, opt-in) ----------
// Run installDailyBackupTrigger() ONCE from the Apps Script editor (select
// it in the function dropdown, click Run) to schedule an automatic daily
// snapshot. It copies the whole spreadsheet into a "TLT_Gumboot_Backups"
// Drive folder with a dated filename — a safety net beyond Sheets' own
// version history. Safe to skip; nothing else depends on it.

function installDailyBackupTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'dailyBackupSnapshot') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('dailyBackupSnapshot').timeBased().everyDays(1).atHour(2).create();
}

function dailyBackupSnapshot() {
  var folders = DriveApp.getFoldersByName('TLT_Gumboot_Backups');
  var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder('TLT_Gumboot_Backups');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var file = DriveApp.getFileById(ss.getId());
  var dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  file.makeCopy(ss.getName() + ' — backup ' + dateStr, folder);
}
