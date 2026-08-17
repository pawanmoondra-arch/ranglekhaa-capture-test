const CONFIG = {
  SPREADSHEET_ID: '1ItA0QbFAdAhF3kzebMVPWFs2o2c_T2kgM_0i_MbgdS4',
  ARTICLE_SHEET: 'ARTICLES',
  LOG_SHEET: 'CAPTURE_LOG',
  ROOT_FOLDER: 'RANGLEKHAA_TEST',
  ARTICLES_FOLDER: '01_ARTICLES',
  DEVELOPMENT_FOLDER: '01_DEVELOPMENT_REFERENCES',
  FINISHED_FOLDER: '02_FINISHED_PRODUCT',
  ARCHIVE_FOLDER: '99_ARCHIVE'
};

function doGet(e) {
  try {
    const action = e && e.parameter ? e.parameter.action : '';

    if (action === 'articles') {
      return jsonResponse_({success:true, articles:getArticles_()});
    }

    if (action === 'captures') {
      const article = String(e.parameter.article || '').trim();
      if (!article) throw new Error('Article Number is missing.');
      return jsonResponse_({success:true, captures:getCaptures_(article)});
    }

    return jsonResponse_({
      success:true,
      app:'Ranglekhaa Capture Backend',
      status:'online'
    });
  } catch (error) {
    return jsonResponse_({success:false,error:error.message});
  }
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error('No request data received.');
    }

    const data = JSON.parse(e.postData.contents);

    if (data.action === 'REPLACE' || data.action === 'CREATE') {
      const result = savePhoto(data);
      return jsonResponse_({success:true,result:result});
    }

    throw new Error('Invalid action.');
  } catch (error) {
    return jsonResponse_({success:false,error:error.message});
  }
}

function getArticles_() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.ARTICLE_SHEET);
  if (!sheet) throw new Error('ARTICLES sheet not found.');

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  const headers = data[0];
  const articleCol = headers.indexOf('Article No');
  const statusCol = headers.indexOf('Overall Status');

  if (articleCol === -1) throw new Error('Article No column not found.');

  return data.slice(1)
    .filter(row => {
      if (!row[articleCol]) return false;
      if (statusCol === -1) return true;
      return String(row[statusCol]).trim().toUpperCase() !== 'CANCELLED';
    })
    .map(row => ({
      articleNo:String(row[articleCol]).trim(),
      status:statusCol === -1 ? '' : String(row[statusCol]).trim()
    }));
}

function savePhoto(data) {
  if (!data.articleNo) throw new Error('Article Number is missing.');
  if (!data.stage) throw new Error('Capture stage is missing.');
  if (!data.photoType) throw new Error('Photo type is missing.');
  if (!data.fileData) throw new Error('Photo data is missing.');

  validateArticle_(data.articleNo);

  const root = getFolderByName_(CONFIG.ROOT_FOLDER);
  const articlesFolder = getOrCreateFolder_(root, CONFIG.ARTICLES_FOLDER);
  const articleFolder = getOrCreateFolder_(articlesFolder, data.articleNo);
  getOrCreateFolder_(articleFolder, '01_COSTING');
  const photosFolder = getOrCreateFolder_(articleFolder, '02_PHOTOS');
  getOrCreateFolder_(articleFolder, '03_DIGITAL_CONTENT');

  let targetFolder;
  if (data.stage === 'DEVELOPMENT') {
    targetFolder = getOrCreateFolder_(photosFolder, CONFIG.DEVELOPMENT_FOLDER);
  } else if (data.stage === 'FINISHED') {
    targetFolder = getOrCreateFolder_(photosFolder, CONFIG.FINISHED_FOLDER);
  } else {
    throw new Error('Invalid capture stage: ' + data.stage);
  }

  const archiveFolder = getOrCreateFolder_(photosFolder, CONFIG.ARCHIVE_FOLDER);

  // Replacement: archive the current active capture(s) for this article/stage/type.
  if (data.action === 'REPLACE') {
    const oldCaptures = getCaptures_(data.articleNo)
      .filter(x => x.stage === data.stage && x.photoType === data.photoType && x.active);

    oldCaptures.forEach(x => {
      try {
        const oldFile = DriveApp.getFileById(x.fileId);
        oldFile.moveTo(archiveFolder);
        logCapture_({
          articleNo:data.articleNo,
          stage:data.stage,
          photoType:data.photoType,
          fileName:oldFile.getName(),
          fileId:oldFile.getId(),
          action:'ARCHIVED'
        });
      } catch (err) {
        // If an old file was already moved/deleted, continue with replacement.
      }
    });
  }

  const parts = String(data.fileData).split(',');
  if (parts.length < 2) throw new Error('Invalid image data.');

  const decoded = Utilities.base64Decode(parts[1]);
  const extension = getExtension_(data.fileName);
  const safeArticle = String(data.articleNo).replace(/[^A-Za-z0-9_-]/g,'_');
  const safeType = String(data.photoType).toUpperCase().replace(/[^A-Z0-9_-]/g,'_');
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');

  const fileName = safeArticle + '_' + safeType + '_' + timestamp + '.' + extension;
  const blob = Utilities.newBlob(decoded, data.mimeType || 'image/jpeg', fileName);
  const file = targetFolder.createFile(blob);

  logCapture_({
    articleNo:data.articleNo,
    stage:data.stage,
    photoType:data.photoType,
    fileName:file.getName(),
    fileId:file.getId(),
    action:data.action === 'REPLACE' ? 'REPLACED' : 'CREATED'
  });

  return {
    fileName:file.getName(),
    fileId:file.getId(),
    url:file.getUrl(),
    action:data.action === 'REPLACE' ? 'REPLACED' : 'CREATED'
  };
}

function getCaptures_(articleNo) {
  validateArticle_(articleNo);

  const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID)
    .getSheetByName(CONFIG.LOG_SHEET);

  if (!sheet) throw new Error('CAPTURE_LOG sheet not found.');

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];

  // Expected columns:
  // Timestamp, User, Article No, Stage, Photo Type, File Name, File ID, Action
  const captures = {};
  data.slice(1).forEach(row => {
    const article = String(row[2] || '').trim();
    if (article !== articleNo) return;

    const stage = String(row[3] || '').trim();
    const photoType = String(row[4] || '').trim();
    const fileName = String(row[5] || '').trim();
    const fileId = String(row[6] || '').trim();
    const action = String(row[7] || '').trim().toUpperCase();

    if (!stage || !photoType || !fileId) return;

    const key = stage + '|' + photoType;

    if (action === 'ARCHIVED') {
      if (captures[key] && captures[key].fileId === fileId) {
        captures[key].active = false;
      }
      return;
    }

    if (action === 'CREATED' || action === 'REPLACED') {
      // A replacement event identifies the new active file.
      captures[key] = {
        articleNo:articleNo,
        stage:stage,
        photoType:photoType,
        fileName:fileName,
        fileId:fileId,
        active:true
      };
    }
  });

  return Object.keys(captures)
    .map(k => captures[k])
    .filter(x => x.active)
    .map(x => {
      let url = '';
      try { url = DriveApp.getFileById(x.fileId).getUrl(); } catch(e) {}
      x.url = url;
      return x;
    });
}

function validateArticle_(articleNo) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONFIG.ARTICLE_SHEET);
  if (!sheet) throw new Error('ARTICLES sheet not found.');

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) throw new Error('No articles found in Master Sheet.');

  const headers = data[0];
  const articleCol = headers.indexOf('Article No');
  const statusCol = headers.indexOf('Overall Status');

  if (articleCol === -1) throw new Error('Article No column not found.');

  const requested = String(articleNo).trim();
  const found = data.slice(1).some(row => {
    if (String(row[articleCol]).trim() !== requested) return false;
    if (statusCol === -1) return true;
    return String(row[statusCol]).trim().toUpperCase() !== 'CANCELLED';
  });

  if (!found) throw new Error('Article Number ' + requested + ' is not registered.');
}

function logCapture_(entry) {
  const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID)
    .getSheetByName(CONFIG.LOG_SHEET);

  if (!sheet) throw new Error('CAPTURE_LOG sheet not found.');

  sheet.appendRow([
    new Date(),
    Session.getActiveUser().getEmail(),
    entry.articleNo,
    entry.stage,
    entry.photoType,
    entry.fileName,
    entry.fileId,
    entry.action
  ]);
}

function getFolderByName_(name) {
  const folders = DriveApp.getFoldersByName(name);
  if (!folders.hasNext()) throw new Error('Folder not found: ' + name);
  return folders.next();
}

function getOrCreateFolder_(parent, name) {
  const folders = parent.getFoldersByName(name);
  if (folders.hasNext()) return folders.next();
  return parent.createFolder(name);
}

function getExtension_(fileName) {
  const parts = String(fileName || '').split('.');
  if (parts.length < 2) return 'jpg';
  const ext = parts.pop().toLowerCase().replace(/[^a-z0-9]/g,'');
  return ext || 'jpg';
}

function jsonResponse_(object) {
  return ContentService.createTextOutput(JSON.stringify(object))
    .setMimeType(ContentService.MimeType.JSON);
}

function TEST_validateArticle() {
  validateArticle_('RKL26-001');
  Logger.log('SUCCESS — RKL26-001 is registered.');
}
