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
    result = { error: err.message };
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
    else result = { success: false, message: 'Unknown action: ' + action };
  } catch (err) {
    result = { success: false, message: err.message };
  }
  return postMessageResponse_(reqId, result);
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
  'TL_cm','SVL_cm','HL_mm','HW_mm','BM_g',
  'FL_L_mm','FL_R_mm','FBW_mm','Dentition',
  'Outcome','Strikes','Region_Struck',
  'Temp_C','Humidity_pct','Photo_Links','Notes','Timestamp'
];

var BOOT_HEADERS = [
  'Boot_ID','Brand_Abbr','Brand_Full','Model','Boot_Size','Session',
  'IS_Standard','Mfg_Date','Batch_No',
  'Thick_T_mm','Thick_LM_mm','Thick_I_mm',
  'Photo_Links','Registered_Date'
];

var SNAKE_HEADERS = [
  'Snake_ID','Species_Code','Common_Name','Age_Class','Sex',
  'TL_cm','SVL_cm','HL_mm','HW_mm','BM_g',
  'FL_L_mm','FL_R_mm','FBW_mm','Dentition',
  'Body_Condition','Last_Feed_Date','Photo_Links','Registered_Date'
];

function getOrCreateSheet_(name, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
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
      size: r[4], session: r[5], isStandard: r[6],
      thickT: r[9], thickLM: r[10], thickI: r[11]
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
      flL: r[10], flR: r[11], fbw: r[12], dentition: r[13], bodyCondition: r[14]
    });
  }
  return result;
}

function formatDateStr_(val) {
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(val || '');
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
      flL: r[18], flR: r[19], fbw: r[20], dentition: r[21],
      outcome: r[22], strikes: r[23], regionStruck: r[24],
      tempC: r[25], humidityPct: r[26], photoLinks: r[27], notes: r[28]
    });
  }
  return result;
}

// ---------- Drive / image upload ----------

function getOrCreateFolder_() {
  var folders = DriveApp.getFoldersByName('TLT_Gumboot_Trial_Photos');
  if (folders.hasNext()) return folders.next();
  return DriveApp.createFolder('TLT_Gumboot_Trial_Photos');
}

function uploadImage_(base64Data, fileName, mimeType) {
  var folder = getOrCreateFolder_();
  var decoded = Utilities.base64Decode(base64Data);
  var blob = Utilities.newBlob(decoded, mimeType || 'image/jpeg', fileName);
  var file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return file.getUrl();
}

function uploadPhotos_(photos, idPrefix) {
  var links = [];
  if (!photos || photos.length === 0) return links;
  for (var i = 0; i < photos.length; i++) {
    if (photos[i] && photos[i].data) {
      try {
        var fileName = idPrefix + '_' + photos[i].label + '.jpg';
        var url = uploadImage_(photos[i].data, fileName, 'image/jpeg');
        links.push(photos[i].label + ': ' + url);
      } catch (e) {
        Logger.log('Photo upload failed for ' + idPrefix + ': ' + e.message);
        links.push(photos[i].label + ': UPLOAD_FAILED');
      }
    }
  }
  return links;
}

// ---------- Registration ----------

function registerBoot(data) {
  var sheet = getOrCreateSheet_('Boots', BOOT_HEADERS);
  var existing = sheet.getDataRange().getValues();
  for (var i = 1; i < existing.length; i++) {
    if (existing[i][0] === data.bootId) {
      return { success: false, message: 'Boot ID already exists: ' + data.bootId };
    }
  }

  var photoLinks = uploadPhotos_(data.photos, data.bootId);

  sheet.appendRow([
    data.bootId, data.brandAbbr, data.brandFull, data.model,
    data.size, data.session, data.isStandard, data.mfgDate, data.batch,
    data.thickT, data.thickLM, data.thickI,
    photoLinks.join('\n'), new Date().toISOString()
  ]);

  return { success: true, message: 'Boot registered: ' + data.bootId };
}

function registerSnake(data) {
  var SPECIES = {
    DR: "Russell's viper", NN: 'Spectacled cobra',
    BC: 'Common krait', EC: 'Saw-scaled viper'
  };
  var DENT = { DR: 'S', NN: 'P', BC: 'P', EC: 'S' };

  var sheet = getOrCreateSheet_('Snakes', SNAKE_HEADERS);
  var existing = sheet.getDataRange().getValues();
  for (var i = 1; i < existing.length; i++) {
    if (existing[i][0] === data.snakeId) {
      return { success: false, message: 'Snake ID already exists: ' + data.snakeId };
    }
  }

  var photoLinks = uploadPhotos_(data.photos, data.snakeId);

  sheet.appendRow([
    data.snakeId, data.speciesCode, SPECIES[data.speciesCode] || '',
    data.ageClass, data.sex,
    data.tl, data.svl, data.hl, data.hw, data.bm,
    data.flL, data.flR, data.fbw, DENT[data.speciesCode] || '',
    data.bodyCondition, data.lastFeedDate,
    photoLinks.join('\n'), new Date().toISOString()
  ]);

  return { success: true, message: 'Snake registered: ' + data.snakeId };
}

// ---------- Trial submission ----------

function submitTrial(data) {
  var sheet = getOrCreateSheet_('Trials', TRIAL_HEADERS);

  var existing = sheet.getDataRange().getValues();
  for (var i = 1; i < existing.length; i++) {
    if (existing[i][0] === data.trialId) {
      return { success: false, message: 'Trial already recorded: ' + data.trialId + '. Use a unique combination.' };
    }
  }

  var photoLinks = uploadPhotos_(data.photos, data.trialId);

  sheet.appendRow([
    data.trialId, data.date, data.session,
    data.bootId, data.bootBrand, data.bootSize,
    data.bootThickT, data.bootThickLM, data.bootThickI,
    data.snakeId, data.speciesCode, data.ageClass, data.sex,
    data.tl, data.svl, data.hl, data.hw, data.bm,
    data.flL, data.flR, data.fbw, data.dentition,
    data.outcome, data.strikes, data.regionStruck,
    data.tempC, data.humidityPct,
    photoLinks.join('\n'), data.notes,
    new Date().toISOString()
  ]);

  return { success: true, message: 'Trial recorded: ' + data.trialId };
}
