const CONFIG = {
  SPREADSHEET_ID: '1ItA0QbFAdAhF3kzebMVPWFs2o2c_T2kgM_0i_MbgdS4',
  ARTICLE_SHEET: 'ARTICLES',
  LOG_SHEET: 'CAPTURE_LOG',
  ROOT_FOLDER: 'RANGLEKHAA_TEST',
  ARTICLES_FOLDER: '01_ARTICLES',
  DEVELOPMENT_FOLDER: '01_DEVELOPMENT_REFERENCES',
  FINISHED_FOLDER: '02_FINISHED_PRODUCT',
  ARCHIVE_FOLDER: '99_ARCHIVE',
  STAFF_EMAILS: ['nsoni9068@gmail.com'],
  APPROVER_EMAILS: ['pawanmoondra@gmail.com']
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
    const action = e && e.parameter ? e.parameter.action : '';

    if (action === 'articles') {
      return jsonResponse_({success:true, articles:getArticles_()});
    }

    if (action === 'status') {
      const article = String(e.parameter.article || '').trim();
      if (!article) throw new Error('Article Number is missing.');
      return jsonResponse_({success:true, status:getArticleStatus_(article)});
    }

    if (action === 'captures') {
      const article = String(e.parameter.article || '').trim();
      if (!article) throw new Error('Article Number is missing.');
      return jsonResponse_({success:true, captures:getCaptures_(article)});
    }

    return jsonResponse_({
      success:true,
      app:'Ranglekhaa Capture Backend',
      status:'online',
      version:'V2.1'
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

    if (data.action === 'CREATE' || data.action === 'REPLACE') {
      return jsonResponse_({success:true,result:savePhoto(data)});
    }

    if (data.action === 'SUBMIT_APPROVAL') {
      return jsonResponse_({success:true,result:submitForApproval_(data)});
    }

    if (data.action === 'APPROVE_LOCK') {
      return jsonResponse_({success:true,result:approveAndLock_(data)});
    }

    if (data.action === 'UNLOCK') {
      return jsonResponse_({success:true,result:unlockStage_(data)});
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

function getArticleStatus_(articleNo) {
  validateArticle_(articleNo);

  const captures = getCaptures_(articleNo);
  const byKey = {};
  captures.forEach(x => byKey[x.stage + '|' + x.photoType] = x);

  return {
    articleNo:articleNo,
    DEVELOPMENT:buildStageStatus_('DEVELOPMENT',byKey,captures),
    FINISHED:buildStageStatus_('FINISHED',byKey,captures)
  };
}

function buildStageStatus_(stage, byKey, captures) {
  const rules = CAPTURE_RULES[stage];

  const items = rules.map(rule => ({
    photoType:rule.type,
    required:rule.required,
    captured:!!byKey[stage + '|' + rule.type],
    fileName:byKey[stage + '|' + rule.type] ? byKey[stage + '|' + rule.type].fileName : ''
  }));

  const requiredTotal = items.filter(x => x.required).length;
  const requiredCaptured = items.filter(x => x.required && x.captured).length;

  const workflow = getLatestWorkflowState_(captures,stage);

  return {
    items:items,
    requiredTotal:requiredTotal,
    requiredCaptured:requiredCaptured,
    complete:requiredCaptured === requiredTotal,
    state:workflow.state,
    submittedAt:workflow.submittedAt,
    submittedBy:workflow.submittedBy,
    approvedAt:workflow.approvedAt,
    approvedBy:workflow.approvedBy,
    unlockedAt:workflow.unlockedAt,
    unlockedBy:workflow.unlockedBy,
    locked:workflow.state === 'LOCKED'
  };
}

function getLatestWorkflowState_(captures,stage) {
  // captures includes internal workflow state attached by getCaptures_.
  let state = 'DRAFT';
  let submittedAt = '', submittedBy = '', approvedAt = '', approvedBy = '', unlockedAt = '', unlockedBy = '';

  captures.filter(x => x.stage === stage && x._workflow).forEach(x => {
    const action = x.action;
    if (action === 'SUBMITTED_FOR_APPROVAL') {
      state = 'PENDING_APPROVAL';
      submittedAt = x.timestamp;
      submittedBy = x.user;
    } else if (action === 'APPROVED_LOCKED') {
      state = 'LOCKED';
      approvedAt = x.timestamp;
      approvedBy = x.user;
    } else if (action === 'UNLOCKED') {
      state = 'DRAFT';
      unlockedAt = x.timestamp;
      unlockedBy = x.user;
    }
  });

  return {state,submittedAt,submittedBy,approvedAt,approvedBy,unlockedAt,unlockedBy};
}

function getUserRole_(email) {
  const e = String(email || '').trim().toLowerCase();
  if (CONFIG.APPROVER_EMAILS.indexOf(e) !== -1) return 'APPROVER';
  if (CONFIG.STAFF_EMAILS.indexOf(e) !== -1) return 'STAFF';
  return 'UNKNOWN';
}

function requireRole_(email,role) {
  const actual = getUserRole_(email);
  if (actual !== role) {
    throw new Error('Not authorized for this action.');
  }
}

function submitForApproval_(data) {
  if (!data.articleNo || !data.stage || !data.userEmail) {
    throw new Error('Article Number, stage and user email are required.');
  }

  requireRole_(data.userEmail,'STAFF');
  validateArticle_(data.articleNo);

  const status = getArticleStatus_(data.articleNo)[data.stage];
  if (!status.complete) {
    throw new Error(
      data.stage + ' cannot be submitted. Required captures complete: ' +
      status.requiredCaptured + '/' + status.requiredTotal
    );
  }

  if (status.locked) throw new Error('Stage is already locked.');
  if (status.state === 'PENDING_APPROVAL') throw new Error('Stage is already pending approval.');

  logWorkflow_({
    articleNo:data.articleNo,
    stage:data.stage,
    action:'SUBMITTED_FOR_APPROVAL',
    user:data.userEmail
  });

  return {articleNo:data.articleNo,stage:data.stage,state:'PENDING_APPROVAL'};
}

function approveAndLock_(data) {
  if (!data.articleNo || !data.stage || !data.userEmail) {
    throw new Error('Article Number, stage and approver email are required.');
  }

  requireRole_(data.userEmail,'APPROVER');
  validateArticle_(data.articleNo);

  const status = getArticleStatus_(data.articleNo)[data.stage];

  if (!status.complete) throw new Error('Stage is not complete.');
  if (status.locked) throw new Error('Stage is already locked.');
  if (status.state !== 'PENDING_APPROVAL') {
    throw new Error('Stage has not been submitted for approval.');
  }

  logWorkflow_({
    articleNo:data.articleNo,
    stage:data.stage,
    action:'APPROVED_LOCKED',
    user:data.userEmail
  });

  return {articleNo:data.articleNo,stage:data.stage,state:'LOCKED'};
}

function unlockStage_(data) {
  if (!data.articleNo || !data.stage || !data.userEmail) {
    throw new Error('Article Number, stage and admin email are required.');
  }

  requireRole_(data.userEmail,'APPROVER');
  validateArticle_(data.articleNo);

  const status = getArticleStatus_(data.articleNo)[data.stage];
  if (!status.locked) throw new Error('Stage is not locked.');

  logWorkflow_({
    articleNo:data.articleNo,
    stage:data.stage,
    action:'UNLOCKED',
    user:data.userEmail
  });

  return {articleNo:data.articleNo,stage:data.stage,state:'DRAFT'};
}

function logWorkflow_(entry) {
  const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID)
    .getSheetByName(CONFIG.LOG_SHEET);

  if (!sheet) throw new Error('CAPTURE_LOG sheet not found.');

  sheet.appendRow([
    new Date(),
    entry.user,
    entry.articleNo,
    entry.stage,
    '',
    '',
    '',
    entry.action
  ]);
}

function savePhoto(data) {
  if (!data.articleNo) throw new Error('Article Number is missing.');
  if (!data.stage) throw new Error('Capture stage is missing.');
  if (!data.photoType) throw new Error('Photo type is missing.');
  if (!data.fileData) throw new Error('Photo data is missing.');

  validateArticle_(data.articleNo);

  const status = getArticleStatus_(data.articleNo)[data.stage];
  if (status.locked) throw new Error(data.stage + ' is locked. Unlock required before changes.');

  const root = getFolderByName_(CONFIG.ROOT_FOLDER);
  const articlesFolder = getOrCreateFolder_(root,CONFIG.ARTICLES_FOLDER);
  const articleFolder = getOrCreateFolder_(articlesFolder,data.articleNo);
  getOrCreateFolder_(articleFolder,'01_COSTING');
  const photosFolder = getOrCreateFolder_(articleFolder,'02_PHOTOS');
  getOrCreateFolder_(articleFolder,'03_DIGITAL_CONTENT');

  let targetFolder;
  if (data.stage === 'DEVELOPMENT') targetFolder=getOrCreateFolder_(photosFolder,CONFIG.DEVELOPMENT_FOLDER);
  else if (data.stage === 'FINISHED') targetFolder=getOrCreateFolder_(photosFolder,CONFIG.FINISHED_FOLDER);
  else throw new Error('Invalid capture stage: ' + data.stage);

  const archiveFolder=getOrCreateFolder_(photosFolder,CONFIG.ARCHIVE_FOLDER);

  if (data.action === 'REPLACE') {
    getCaptures_(data.articleNo)
      .filter(x=>x.stage===data.stage && x.photoType===data.photoType && x.active)
      .forEach(x=>{
        try{
          const oldFile=DriveApp.getFileById(x.fileId);
          oldFile.moveTo(archiveFolder);
          logCapture_({
            articleNo:data.articleNo,stage:data.stage,photoType:data.photoType,
            fileName:oldFile.getName(),fileId:oldFile.getId(),action:'ARCHIVED'
          });
        }catch(err){}
      });
  }

  const parts=String(data.fileData).split(',');
  if(parts.length<2) throw new Error('Invalid image data.');

  const decoded=Utilities.base64Decode(parts[1]);
  const extension=getExtension_(data.fileName);
  const safeArticle=String(data.articleNo).replace(/[^A-Za-z0-9_-]/g,'_');
  const safeType=String(data.photoType).toUpperCase().replace(/[^A-Z0-9_-]/g,'_');
  const timestamp=Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyyMMdd_HHmmss');
  const fileName=safeArticle+'_'+safeType+'_'+timestamp+'.'+extension;

  const blob=Utilities.newBlob(decoded,data.mimeType||'image/jpeg',fileName);
  const file=targetFolder.createFile(blob);

  logCapture_({
    articleNo:data.articleNo,stage:data.stage,photoType:data.photoType,
    fileName:file.getName(),fileId:file.getId(),
    action:data.action==='REPLACE'?'REPLACED':'CREATED'
  });

  return {fileName:file.getName(),fileId:file.getId(),url:file.getUrl(),action:data.action==='REPLACE'?'REPLACED':'CREATED'};
}

function getCaptures_(articleNo) {
  validateArticle_(articleNo);

  const sheet=SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.LOG_SHEET);
  if(!sheet) throw new Error('CAPTURE_LOG sheet not found.');

  const data=sheet.getDataRange().getValues();
  if(data.length<=1) return [];

  const captures={};
  const workflow=[];

  data.slice(1).forEach(row=>{
    const timestamp=row[0];
    const user=String(row[1]||'').trim();
    const article=String(row[2]||'').trim();
    if(article!==articleNo)return;

    const stage=String(row[3]||'').trim();
    const photoType=String(row[4]||'').trim();
    const fileName=String(row[5]||'').trim();
    const fileId=String(row[6]||'').trim();
    const action=String(row[7]||'').trim().toUpperCase();

    if(['SUBMITTED_FOR_APPROVAL','APPROVED_LOCKED','UNLOCKED'].indexOf(action)!==-1){
      workflow.push({
        stage:stage,action:action,user:user,
        timestamp:timestamp ? new Date(timestamp).toISOString() : '',
        _workflow:true
      });
      return;
    }

    if(!stage||!photoType||!fileId)return;

    const key=stage+'|'+photoType;

    if(action==='ARCHIVED'){
      if(captures[key]&&captures[key].fileId===fileId)captures[key].active=false;
      return;
    }

    if(action==='CREATED'||action==='REPLACED'){
      captures[key]={
        articleNo:articleNo,stage:stage,photoType:photoType,
        fileName:fileName,fileId:fileId,active:true
      };
    }
  });

  const result=Object.keys(captures).map(k=>captures[k]).filter(x=>x.active).map(x=>{
    let url='';
    try{url=DriveApp.getFileById(x.fileId).getUrl();}catch(e){}
    x.url=url;
    return x;
  });

  return result.concat(workflow);
}

function validateArticle_(articleNo) {
  const ss=SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const sheet=ss.getSheetByName(CONFIG.ARTICLE_SHEET);
  if(!sheet)throw new Error('ARTICLES sheet not found.');

  const data=sheet.getDataRange().getValues();
  if(data.length<=1)throw new Error('No articles found in Master Sheet.');

  const headers=data[0];
  const articleCol=headers.indexOf('Article No');
  const statusCol=headers.indexOf('Overall Status');
  if(articleCol===-1)throw new Error('Article No column not found.');

  const requested=String(articleNo).trim();
  const found=data.slice(1).some(row=>{
    if(String(row[articleCol]).trim()!==requested)return false;
    if(statusCol===-1)return true;
    return String(row[statusCol]).trim().toUpperCase()!=='CANCELLED';
  });

  if(!found)throw new Error('Article Number '+requested+' is not registered.');
}

function logCapture_(entry) {
  const sheet=SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.LOG_SHEET);
  if(!sheet)throw new Error('CAPTURE_LOG sheet not found.');

  sheet.appendRow([
    new Date(),Session.getActiveUser().getEmail(),
    entry.articleNo,entry.stage,entry.photoType,entry.fileName,entry.fileId,entry.action
  ]);
}

function getFolderByName_(name) {
  const folders=DriveApp.getFoldersByName(name);
  if(!folders.hasNext())throw new Error('Folder not found: '+name);
  return folders.next();
}

function getOrCreateFolder_(parent,name) {
  const folders=parent.getFoldersByName(name);
  if(folders.hasNext())return folders.next();
  return parent.createFolder(name);
}

function getExtension_(fileName) {
  const parts=String(fileName||'').split('.');
  if(parts.length<2)return 'jpg';
  const ext=parts.pop().toLowerCase().replace(/[^a-z0-9]/g,'');
  return ext||'jpg';
}

function jsonResponse_(object) {
  return ContentService.createTextOutput(JSON.stringify(object))
    .setMimeType(ContentService.MimeType.JSON);
}

function TEST_validateArticle() {
  validateArticle_('RKL26-001');
  Logger.log('SUCCESS — RKL26-001 is registered.');
}
