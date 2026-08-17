const CONFIG = {
  SPREADSHEET_ID: '1ItA0QbFAdAhF3kzebMVPWFs2o2c_T2kgM_0i_MbgdS4',
  ARTICLE_SHEET: 'ARTICLES',
  CAPTURE_LOG: 'CAPTURE_LOG',
  WORKFLOW_SHEET: 'WORKFLOW',

  ROOT_FOLDER: 'RANGLEKHAA_TEST',
  ARTICLES_FOLDER: '01_ARTICLES',
  DEVELOPMENT_FOLDER: '01_DEVELOPMENT_REFERENCES',
  FINISHED_FOLDER: '02_FINISHED_PRODUCT',
  ARCHIVE_FOLDER: '99_ARCHIVE',

  STAFF_EMAIL: 'nsoni9068@gmail.com',
  ADMIN_EMAIL: 'pawanmoondra@gmail.com'
};

const RULES = {
  DEVELOPMENT: [
    {type:'COLOUR_OPTIONS', title:'Colour Options', required:true},
    {type:'EMBROIDERY_SAMPLE', title:'Approved Embroidery Sample', required:true}
  ],
  FINISHED: [
    {type:'FINISHED_01', title:'Full Product — Front', required:true},
    {type:'FINISHED_02', title:'Back / Overall', required:true},
    {type:'FINISHED_03', title:'Embroidery / Work Detail', required:true},
    {type:'FINISHED_04', title:'Fabric / Texture Close-up', required:true},
    {type:'FINISHED_05', title:'Important Component', required:false},
    {type:'FINISHED_06', title:'Colour Options / Assortment', required:false}
  ]
};

function doGet(e) {
  try {
    setup_();
    const action = e && e.parameter ? String(e.parameter.action || '') : '';

    if (action === 'articles') {
      return out_({success:true, articles:getArticles_()});
    }

    if (action === 'status') {
      const article = clean_(e.parameter.article);
      if (!article) throw new Error('Article Number is required.');
      return out_({success:true, status:getStatus_(article)});
    }

    if (action === 'captures') {
      const article = clean_(e.parameter.article);
      if (!article) throw new Error('Article Number is required.');
      return out_({success:true, captures:getCaptures_(article)});
    }

    return out_({
      success:true,
      app:'Ranglekhaa Product Capture',
      version:'FRESH-V1',
      status:'online'
    });
  } catch (err) {
    return out_({success:false,error:err.message});
  }
}

function doPost(e) {
  try {
    setup_();
    const d = JSON.parse(e.postData.contents || '{}');
    const action = String(d.action || '').toUpperCase();

    if (action === 'CREATE' || action === 'REPLACE') {
      return out_({success:true,result:savePhoto_(d)});
    }

    if (action === 'LOCK') {
      return out_({success:true,result:lockStage_(d)});
    }

    if (action === 'UNLOCK') {
      return out_({success:true,result:unlockStage_(d)});
    }

    throw new Error('Unknown action.');
  } catch (err) {
    return out_({success:false,error:err.message});
  }
}

function setup_() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  let ws = ss.getSheetByName(CONFIG.WORKFLOW_SHEET);
  if (!ws) {
    ws = ss.insertSheet(CONFIG.WORKFLOW_SHEET);
    ws.appendRow(['Timestamp','User Email','Article No','Stage','Action']);
  }
}

function getArticles_() {
  const sh = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.ARTICLE_SHEET);
  if (!sh) throw new Error('ARTICLES sheet not found.');

  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];

  const h = values[0].map(x => String(x).trim());
  const a = h.indexOf('Article No');
  const s = h.indexOf('Overall Status');
  if (a < 0) throw new Error('Article No column not found.');

  return values.slice(1).filter(r => {
    if (!r[a]) return false;
    if (s < 0) return true;
    return String(r[s]).trim().toUpperCase() !== 'CANCELLED';
  }).map(r => ({
    articleNo:String(r[a]).trim(),
    status:s < 0 ? '' : String(r[s]).trim()
  }));
}

function validateArticle_(article) {
  const found = getArticles_().some(x => x.articleNo === article);
  if (!found) throw new Error('Article Number ' + article + ' is not registered.');
}

function getCaptures_(article) {
  const sh = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.CAPTURE_LOG);
  if (!sh) throw new Error('CAPTURE_LOG sheet not found.');

  const v = sh.getDataRange().getValues();
  if (v.length < 2) return [];

  const current = {};
  v.slice(1).forEach(r => {
    const a = clean_(r[2]);
    if (a !== article) return;

    const stage = clean_(r[3]);
    const type = clean_(r[4]);
    const fileName = clean_(r[5]);
    const fileId = clean_(r[6]);
    const action = clean_(r[7]).toUpperCase();
    if (!stage || !type || !fileId) return;

    const key = stage + '|' + type;

    if (action === 'ARCHIVED') {
      if (current[key] && current[key].fileId === fileId) delete current[key];
      return;
    }

    if (action === 'CREATED' || action === 'REPLACED') {
      current[key] = {
        articleNo:article,
        stage:stage,
        photoType:type,
        fileName:fileName,
        fileId:fileId,
        active:true
      };
    }
  });

  return Object.values(current).map(x => {
    try { x.url = DriveApp.getFileById(x.fileId).getUrl(); } catch(e) { x.url=''; }
    return x;
  });
}

function getWorkflow_(article, stage) {
  const sh = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.WORKFLOW_SHEET);
  if (!sh) return {state:'DRAFT'};

  const v = sh.getDataRange().getValues();
  let state = 'DRAFT';
  let last = null;

  v.slice(1).forEach(r => {
    if (clean_(r[2]) !== article || clean_(r[3]) !== stage) return;
    const action = clean_(r[4]).toUpperCase();
    if (action !== 'LOCK' && action !== 'UNLOCK') return;
    const ts = r[0] ? new Date(r[0]).getTime() : 0;
    if (!last || ts >= last.ts) last = {ts:ts,action:action,user:clean_(r[1]),timestamp:r[0]};
  });

  if (last && last.action === 'LOCK') {
    state = 'LOCKED';
  }

  return {
    state:state,
    locked:state === 'LOCKED',
    lastAction:last ? last.action : '',
    lastUser:last ? last.user : '',
    lastTimestamp:last ? String(last.timestamp) : ''
  };
}

function getStatus_(article) {
  validateArticle_(article);
  const captures = getCaptures_(article);
  const map = {};
  captures.forEach(x => map[x.stage+'|'+x.photoType] = x);

  return {
    articleNo:article,
    DEVELOPMENT:stageStatus_('DEVELOPMENT',map,captures),
    FINISHED:stageStatus_('FINISHED',map,captures)
  };
}

function stageStatus_(stage,map,captures) {
  const items = RULES[stage].map(r => {
    const x = map[stage+'|'+r.type];
    return {
      photoType:r.type,
      title:r.title,
      required:r.required,
      captured:!!x,
      fileName:x ? x.fileName : '',
      url:x ? x.url : ''
    };
  });

  const requiredTotal = items.filter(x => x.required).length;
  const requiredCaptured = items.filter(x => x.required && x.captured).length;
  const wf = getWorkflow_(captures.length ? captures[0].articleNo : '',stage);

  return {
    items:items,
    requiredTotal:requiredTotal,
    requiredCaptured:requiredCaptured,
    complete:requiredCaptured === requiredTotal,
    state:wf.state,
    locked:wf.locked,
    lastAction:wf.lastAction,
    lastUser:wf.lastUser,
    lastTimestamp:wf.lastTimestamp
  };
}

function savePhoto_(d) {
  const article = clean_(d.articleNo);
  const stage = clean_(d.stage).toUpperCase();
  const type = clean_(d.photoType).toUpperCase();

  if (!article || !stage || !type || !d.fileData) throw new Error('Missing photo information.');
  validateArticle_(article);

  const status = getStatus_(article)[stage];
  if (status.locked) throw new Error(stage + ' is locked. Pawan must unlock it before changes.');

  const root = getFolder_(CONFIG.ROOT_FOLDER);
  const articles = getOrCreate_(root,CONFIG.ARTICLES_FOLDER);
  const articleFolder = getOrCreate_(articles,article);
  const photos = getOrCreate_(articleFolder,'02_PHOTOS');
  const target = getOrCreate_(photos,stage === 'DEVELOPMENT' ? CONFIG.DEVELOPMENT_FOLDER : CONFIG.FINISHED_FOLDER);
  const archive = getOrCreate_(photos,CONFIG.ARCHIVE_FOLDER);

  if (String(d.action).toUpperCase() === 'REPLACE') {
    getCaptures_(article).filter(x => x.stage === stage && x.photoType === type).forEach(x => {
      try {
        const old = DriveApp.getFileById(x.fileId);
        old.moveTo(archive);
        appendCapture_(article,stage,type,old.getName(),old.getId(),'ARCHIVED');
      } catch(e) {}
    });
  }

  const comma = String(d.fileData).indexOf(',');
  if (comma < 0) throw new Error('Invalid image data.');
  const bytes = Utilities.base64Decode(String(d.fileData).slice(comma+1));
  const ext = extension_(d.fileName);
  const stamp = Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyyMMdd_HHmmss');
  const name = article+'_'+type+'_'+stamp+'.'+ext;
  const file = target.createFile(Utilities.newBlob(bytes,d.mimeType || 'image/jpeg',name));

  appendCapture_(article,stage,type,file.getName(),file.getId(),String(d.action).toUpperCase() === 'REPLACE' ? 'REPLACED' : 'CREATED');

  return {fileName:file.getName(),fileId:file.getId(),url:file.getUrl()};
}

function appendCapture_(article,stage,type,fileName,fileId,action) {
  const sh = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName(CONFIG.CAPTURE_LOG);
  sh.appendRow([new Date(),Session.getActiveUser().getEmail(),article,stage,type,fileName,fileId,action]);
}

function lockStage_(d) {
  const article = clean_(d.articleNo);
  const stage = clean_(d.stage).toUpperCase();
  const user = clean_(d.userEmail).toLowerCase();

  if (user !== CONFIG.STAFF_EMAIL) throw new Error('Only Naveen can lock a stage in this test.');
  validateArticle_(article);

  const status = getStatus_(article)[stage];
  if (!status.complete) throw new Error('Complete required captures first: '+status.requiredCaptured+'/'+status.requiredTotal);
  if (status.locked) throw new Error('Stage is already locked.');

  workflow_(article,stage,'LOCK',user);
  return {articleNo:article,stage:stage,state:'LOCKED'};
}

function unlockStage_(d) {
  const article = clean_(d.articleNo);
  const stage = clean_(d.stage).toUpperCase();
  const user = clean_(d.userEmail).toLowerCase();

  if (user !== CONFIG.ADMIN_EMAIL) throw new Error('Only Pawan can unlock a locked stage in this test.');
  validateArticle_(article);

  const status = getStatus_(article)[stage];
  if (!status.locked) throw new Error('Stage is not locked.');

  workflow_(article,stage,'UNLOCK',user);
  return {articleNo:article,stage:stage,state:'DRAFT'};
}

function workflow_(article,stage,action,user) {
  SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID)
    .getSheetByName(CONFIG.WORKFLOW_SHEET)
    .appendRow([new Date(),user,article,stage,action]);
}

function getFolder_(name) {
  const it = DriveApp.getFoldersByName(name);
  if (!it.hasNext()) throw new Error('Folder not found: '+name);
  return it.next();
}

function getOrCreate_(parent,name) {
  const it = parent.getFoldersByName(name);
  return it.hasNext() ? it.next() : parent.createFolder(name);
}

function extension_(name) {
  const p = String(name || '').split('.');
  const e = (p.length > 1 ? p.pop() : 'jpg').toLowerCase().replace(/[^a-z0-9]/g,'');
  return e || 'jpg';
}

function clean_(x) { return String(x == null ? '' : x).trim(); }

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function TEST_SETUP() {
  setup_();
  Logger.log('FRESH-V1 ready.');
}
