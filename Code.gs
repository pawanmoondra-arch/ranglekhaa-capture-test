const CONFIG = {
  SPREADSHEET_ID: '1ItA0QbFAdAhF3kzebMVPWFs2o2c_T2kgM_0i_MbgdS4',
  ARTICLE_SHEET: 'ARTICLES',
  LOG_SHEET: 'CAPTURE_LOG',
  ROOT_FOLDER: 'RANGLEKHAA_TEST',
  ARTICLES_FOLDER: '01_ARTICLES',
  DEVELOPMENT_FOLDER: '01_DEVELOPMENT_REFERENCES',
  FINISHED_FOLDER: '02_FINISHED_PRODUCT',
  ARCHIVE_FOLDER: '99_ARCHIVE',
  PARTNER_EMAILS: ['nsoni9068@gmail.com'],
  ADMIN_EMAILS: ['pawanmoondra@gmail.com']
};

const CAPTURE_RULES = {
  DEVELOPMENT: [
    {type:'COLOUR_OPTIONS', required:true},
    {type:'EMBROIDERY_SAMPLE', required:true}
  ],
  FINISHED: [
    {type:'FINISHED_01', required:true},
    {type:'FINISHED_02', required:true},
    {type:'FINISHED_03', required:true},
    {type:'FINISHED_04', required:true},
    {type:'FINISHED_05', required:false},
    {type:'FINISHED_06', required:false}
  ]
};

function doGet(e) {
  try {
    const action = e && e.parameter ? String(e.parameter.action || '') : '';

    if (action === 'articles') {
      return jsonResponse_({success:true, articles:getArticles_()});
    }

    if (action === 'captures') {
      const article = String(e.parameter.article || '').trim();
      if (!article) throw new Error('Article Number is missing.');
      return jsonResponse_({success:true, captures:getCaptures_(article)});
    }

    if (action === 'status') {
      const article = String(e.parameter.article || '').trim();
      if (!article) throw new Error('Article Number is missing.');
      return jsonResponse_({success:true, status:getArticleStatus_(article)});
    }

    return jsonResponse_({
      success:true,
      app:'Ranglekhaa Capture Backend',
      status:'online',
      version:'V2.2'
    });
  } catch (error) {
    return jsonResponse_({success:false,error:error.message,version:'V2.2'});
  }
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error('No request data received.');
    }

    const data = JSON.parse(e.postData.contents);

    if (data.action === 'CREATE' || data.action === 'REPLACE') {
      return jsonResponse_({success:true,result:savePhoto_(data)});
    }

    if (data.action === 'LOCK') {
      return jsonResponse_({success:true,result:lockStage_(data)});
    }

    if (data.action === 'UNLOCK') {
      return jsonResponse_({success:true,result:unlockStage_(data)});
    }

    throw new Error('Invalid action.');
  } catch (error) {
    return jsonResponse_({success:false,error:error.message,version:'V2.2'});
  }
}

function getArticles_() {
  const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.ARTICLE_SHEET);
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

function getArticleStatus_(articleNo) {
  validateArticle_(articleNo);

  const captures = getCaptures_(articleNo);
  const byKey = {};
  captures.photos.forEach(x => byKey[x.stage + '|' + x.photoType] = x);

  return {
    articleNo:articleNo,
    DEVELOPMENT:buildStageStatus_('DEVELOPMENT',byKey,captures.workflow),
    FINISHED:buildStageStatus_('FINISHED',byKey,captures.workflow)
  };
}

function buildStageStatus_(stage, byKey, workflow) {
  const rules = CAPTURE_RULES[stage];
  const items = rules.map(rule => ({
    photoType:rule.type,
    required:rule.required,
    captured:!!byKey[stage + '|' + rule.type],
    fileName:byKey[stage + '|' + rule.type] ? byKey[stage + '|' + rule.type].fileName : ''
  }));

  const requiredTotal = items.filter(x => x.required).length;
  const requiredCaptured = items.filter(x => x.required && x.captured).length;
  const state = workflow[stage] || 'DRAFT';

  return {
    items:items,
    requiredTotal:requiredTotal,
    requiredCaptured:requiredCaptured,
    complete:requiredCaptured === requiredTotal,
    state:state,
    locked:state === 'LOCKED'
  };
}

function lockStage_(data) {
  const articleNo = String(data.articleNo || '').trim();
  const stage = String(data.stage || '').trim().toUpperCase();
  const userEmail = String(data.userEmail || '').trim().toLowerCase();

  if (!articleNo || !stage || !userEmail) {
    throw new Error('Article Number, stage and Partner email are required.');
  }

  requirePartner_(userEmail);
  validateArticle_(articleNo);

  const status = getArticleStatus_(articleNo)[stage];
  if (!status) throw new Error('Invalid stage.');
  if (!status.complete) {
    throw new Error(
      stage + ' cannot be locked. Required captures complete: ' +
      status.requiredCaptured + '/' + status.requiredTotal
    );
  }
  if (status.locked) throw new Error(stage + ' is already locked.');

  appendWorkflow_(articleNo, stage, 'LOCKED', userEmail);

  return {
    articleNo:articleNo,
    stage:stage,
    state:'LOCKED'
  };
}

function unlockStage_(data) {
  const articleNo = String(data.articleNo || '').trim();
  const stage = String(data.stage || '').trim().toUpperCase();
  const userEmail = String(data.userEmail || '').trim().toLowerCase();

  if (!articleNo || !stage || !userEmail) {
    throw new Error('Article Number, stage and Partner & Admin email are required.');
  }

  requireAdmin_(userEmail);
  validateArticle_(articleNo);

  const status = getArticleStatus_(articleNo)[stage];
  if (!status) throw new Error('Invalid stage.');
  if (!status.locked) throw new Error(stage + ' is not locked.');

  appendWorkflow_(articleNo, stage, 'UNLOCKED', userEmail);

  return {
    articleNo:articleNo,
    stage:stage,
    state:'DRAFT'
  };
}

function appendWorkflow_(articleNo, stage, action, userEmail) {
  const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.LOG_SHEET);
  if (!sheet) throw new Error('CAPTURE_LOG sheet not found.');

  sheet.appendRow([
    new Date(),
    userEmail,
    articleNo,
    stage,
    '',
    '',
    '',
    action
  ]);
}

function requirePartner_(email) {
  const allowed = CONFIG.PARTNER_EMAILS.concat(CONFIG.ADMIN_EMAILS)
    .map(x => x.toLowerCase());
  if (allowed.indexOf(email) === -1) {
    throw new Error('Not authorized. Partner access required.');
  }
}

function requireAdmin_(email) {
  if (CONFIG.ADMIN_EMAILS.map(x => x.toLowerCase()).indexOf(email) === -1) {
    throw new Error('Not authorized. Partner & Admin access required.');
  }
}

function savePhoto_(data) {
  const articleNo = String(data.articleNo || '').trim();
  const stage = String(data.stage || '').trim().toUpperCase();
  const photoType = String(data.photoType || '').trim().toUpperCase();
  const userEmail = String(data.userEmail || '').trim().toLowerCase();

  if (!articleNo || !stage || !photoType || !userEmail || !data.fileData) {
    throw new Error('Article Number, stage, photo type, Partner email and photo are required.');
  }

  requirePartner_(userEmail);
  validateArticle_(articleNo);

  const status = getArticleStatus_(articleNo)[stage];
  if (!status) throw new Error('Invalid stage.');
  if (status.locked) throw new Error(stage + ' is locked. Pawan must unlock it before changes.');

  const root = getFolderByName_(CONFIG.ROOT_FOLDER);
  const articlesFolder = getOrCreateFolder_(root,CONFIG.ARTICLES_FOLDER);
  const articleFolder = getOrCreateFolder_(articlesFolder,articleNo);

  getOrCreateFolder_(articleFolder,'01_COSTING');
  getOrCreateFolder_(articleFolder,'03_DIGITAL_CONTENT');

  const photosFolder = getOrCreateFolder_(articleFolder,'02_PHOTOS');
  const archiveFolder = getOrCreateFolder_(photosFolder,CONFIG.ARCHIVE_FOLDER);

  let targetFolder;
  if (stage === 'DEVELOPMENT') {
    targetFolder = getOrCreateFolder_(photosFolder,CONFIG.DEVELOPMENT_FOLDER);
  } else if (stage === 'FINISHED') {
    targetFolder = getOrCreateFolder_(photosFolder,CONFIG.FINISHED_FOLDER);
  } else {
    throw new Error('Invalid capture stage: ' + stage);
  }

  if (data.action === 'REPLACE') {
    getCaptures_(articleNo).photos
      .filter(x => x.stage === stage && x.photoType === photoType && x.active)
      .forEach(x => {
        try {
          const oldFile = DriveApp.getFileById(x.fileId);
          oldFile.moveTo(archiveFolder);
          logPhoto_(articleNo,stage,photoType,oldFile.getName(),oldFile.getId(),'ARCHIVED',userEmail);
        } catch (err) {}
      });
  }

  const parts = String(data.fileData).split(',');
  if (parts.length < 2) throw new Error('Invalid image data.');

  const decoded = Utilities.base64Decode(parts[1]);
  const extension = getExtension_(data.fileName);
  const safeArticle = articleNo.replace(/[^A-Za-z0-9_-]/g,'_');
  const safeType = photoType.replace(/[^A-Z0-9_-]/g,'_');
  const timestamp = Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyyMMdd_HHmmss');

  const fileName = safeArticle + '_' + safeType + '_' + timestamp + '.' + extension;
  const blob = Utilities.newBlob(decoded,data.mimeType || 'image/jpeg',fileName);
  const file = targetFolder.createFile(blob);

  logPhoto_(
    articleNo,stage,photoType,file.getName(),file.getId(),
    data.action === 'REPLACE' ? 'REPLACED' : 'CREATED',
    userEmail
  );

  return {
    fileName:file.getName(),
    fileId:file.getId(),
    url:file.getUrl(),
    action:data.action === 'REPLACE' ? 'REPLACED' : 'CREATED'
  };
}

function getCaptures_(articleNo) {
  validateArticle_(articleNo);

  const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.LOG_SHEET);
  if (!sheet) throw new Error('CAPTURE_LOG sheet not found.');

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return {photos:[],workflow:{}};

  const photos = {};
  const workflow = {};

  data.slice(1).forEach(row => {
    const timestamp = row[0];
    const user = String(row[1] || '').trim().toLowerCase();
    const article = String(row[2] || '').trim();
    const stage = String(row[3] || '').trim().toUpperCase();
    const photoType = String(row[4] || '').trim().toUpperCase();
    const fileName = String(row[5] || '').trim();
    const fileId = String(row[6] || '').trim();
    const action = String(row[7] || '').trim().toUpperCase();

    if (article !== articleNo) return;

    if (action === 'LOCKED' || action === 'UNLOCKED') {
      workflow[stage] = action === 'LOCKED' ? 'LOCKED' : 'DRAFT';
      return;
    }

    if (!stage || !photoType || !fileId) return;

    const key = stage + '|' + photoType;

    if (action === 'ARCHIVED') {
      if (photos[key] && photos[key].fileId === fileId) {
        photos[key].active = false;
      }
      return;
    }

    if (action === 'CREATED' || action === 'REPLACED') {
      photos[key] = {
        articleNo:articleNo,
        stage:stage,
        photoType:photoType,
        fileName:fileName,
        fileId:fileId,
        active:true
      };
    }
  });

  const activePhotos = Object.keys(photos)
    .map(k => photos[k])
    .filter(x => x.active)
    .map(x => {
      let url = '';
      try { url = DriveApp.getFileById(x.fileId).getUrl(); } catch(e) {}
      x.url = url;
      return x;
    });

  return {photos:activePhotos,workflow:workflow};
}

function validateArticle_(articleNo) {
  const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.ARTICLE_SHEET);
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

function logPhoto_(articleNo,stage,photoType,fileName,fileId,action,userEmail) {
  const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.LOG_SHEET);
  if (!sheet) throw new Error('CAPTURE_LOG sheet not found.');

  sheet.appendRow([
    new Date(),userEmail,articleNo,stage,photoType,fileName,fileId,action
  ]);
}

function getFolderByName_(name) {
  const folders = DriveApp.getFoldersByName(name);
  if (!folders.hasNext()) throw new Error('Folder not found: ' + name);
  return folders.next();
}

function getOrCreateFolder_(parent,name) {
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
