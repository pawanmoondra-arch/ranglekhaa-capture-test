/**
 * RANGLEKHAA CAPTURE BACKEND V3.2
 * Bound to RANGLEKHAA_MASTER.
 * Manual sheet required: ARTICLES
 * Required headers: Article No, Article Name, Category, Status
 *
 * Roles:
 *   Partner: nsoni9068@gmail.com
 *   Partner & Admin: pawanmoondra@gmail.com
 *
 * State: DRAFT -> LOCKED; Admin can UNLOCK -> DRAFT.
 * No approval workflow.
 */

const CFG={
  VERSION:"V3.2",
  ARTICLE_SHEET:"ARTICLES",
  LOG_SHEET:"CAPTURE_LOG",
  PARTNER_EMAILS:["nsoni9068@gmail.com"],
  ADMIN_EMAILS:["pawanmoondra@gmail.com"],
  ROOT_FOLDER:"RANGLEKHAA_CAPTURE",
  ARTICLES_FOLDER:"01_ARTICLES",
  DEVELOPMENT_FOLDER:"01_DEVELOPMENT_REFERENCES",
  FINISHED_FOLDER:"02_FINISHED_PRODUCT",
  ARCHIVE_FOLDER:"99_ARCHIVE",
  DIGITAL_FOLDER:"03_DIGITAL_CONTENT",
  COSTING_FOLDER:"01_COSTING"
};

const RULES={
  DEVELOPMENT:[
    {type:"COLOUR_OPTIONS",label:"Colour Options",required:true,help:"All approved colour options together."},
    {type:"EMBROIDERY_SAMPLE",label:"Approved Embroidery Sample",required:true,help:"Approved embroidery/value-addition sample."}
  ],
  FINISHED:[
    {type:"FINISHED_01",label:"Full Product — Front",required:true,help:"Clear full-product photograph."},
    {type:"FINISHED_02",label:"Back / Overall",required:true,help:"Back or overall product view."},
    {type:"FINISHED_03",label:"Embroidery / Work Detail",required:true,help:"Clear workmanship/detail close-up."},
    {type:"FINISHED_04",label:"Fabric / Texture Close-up",required:true,help:"Fabric texture, print or finish."},
    {type:"FINISHED_05",label:"Dupatta / Important Component",required:false,help:"Optional significant component."},
    {type:"FINISHED_06",label:"Colour Options / Assortment",required:false,help:"Optional colour assortment."}
  ]
};

/**
 * PUBLIC SETUP BUTTON.
 * Apps Script hides function names ending in "_" from the Run menu.
 */
function setupRanglekhaa() {
  return setupRanglekhaa_();
}

function doGet(e){
  try{
    var a=String(e&&e.parameter&&e.parameter.action||"").toLowerCase();
    if(a==="health"||!a) return out_({success:true,app:"Ranglekhaa Capture Backend",status:"online",version:CFG.VERSION});
    if(a==="articles") return out_({success:true,articles:getArticles_()});
    if(a==="status") return out_({success:true,status:getStatus_(req_(e.parameter.article,"Article Number"))});
    if(a==="captures") return out_({success:true,captures:getCaptures_(req_(e.parameter.article,"Article Number"))});
    throw new Error("Unknown action: "+a);
  }catch(err){return out_({success:false,error:String(err.message||err),version:CFG.VERSION});}
}

function doPost(e){
  try{
    if(!e||!e.postData||!e.postData.contents) throw new Error("POST body missing.");
    var d=JSON.parse(e.postData.contents),a=String(d.action||"").toUpperCase();
    if(a==="CREATE"||a==="REPLACE") return out_({success:true,result:savePhoto_(d)});
    if(a==="LOCK") return out_({success:true,result:lockStage_(d)});
    if(a==="UNLOCK") return out_({success:true,result:unlockStage_(d)});
    throw new Error("Unknown POST action: "+a);
  }catch(err){return out_({success:false,error:String(err.message||err),version:CFG.VERSION});}
}

/** Run once after pasting. Creates CAPTURE_LOG and base Drive structure. */
function setupRanglekhaa_(){
  var ss=SpreadsheetApp.getActiveSpreadsheet();
  if(!ss) throw new Error("This script must be bound to RANGLEKHAA_MASTER.");
  var sh=ss.getSheetByName(CFG.ARTICLE_SHEET);
  if(!sh) throw new Error('Sheet "ARTICLES" not found.');
  var h=sh.getRange(1,1,1,Math.max(1,sh.getLastColumn())).getValues()[0].map(function(x){return String(x).trim();});
  if(h.indexOf("Article No")<0) throw new Error('ARTICLES needs "Article No".');
  if(h.indexOf("Status")<0) throw new Error('ARTICLES needs "Status".');

  var log=ss.getSheetByName(CFG.LOG_SHEET);
  if(!log){
    log=ss.insertSheet(CFG.LOG_SHEET);
    log.getRange(1,1,1,8).setValues([["Timestamp","User Email","Article No","Stage","Photo Type","File Name","File ID","Action"]]);
    log.setFrozenRows(1);
  }

  var root=getRoot_();
  folder_(root,CFG.ARTICLES_FOLDER);
  folder_(root,"02_MASTER_DATA");
  folder_(root,CFG.DIGITAL_FOLDER);
  folder_(root,CFG.ARCHIVE_FOLDER);

  return {success:true,version:CFG.VERSION,spreadsheetId:ss.getId(),rootFolderId:root.getId(),logSheet:CFG.LOG_SHEET};
}

function getArticles_(){
  var sh=sheet_(CFG.ARTICLE_SHEET),v=sh.getDataRange().getValues();
  if(v.length<2)return [];
  var h=v[0].map(function(x){return String(x).trim();}),ai=h.indexOf("Article No"),si=h.indexOf("Status");
  if(ai<0||si<0)throw new Error("Master headers missing.");
  return v.slice(1).filter(function(r){
    return String(r[ai]||"").trim() && String(r[si]||"").trim().toUpperCase()!=="CANCELLED";
  }).map(function(r){return {articleNo:String(r[ai]).trim(),status:String(r[si]).trim()};});
}

function validateArticle_(articleNo){
  if(!getArticles_().some(function(x){return x.articleNo===articleNo;}))
    throw new Error("Article Number "+articleNo+" is not registered.");
}

function getStatus_(articleNo){
  validateArticle_(articleNo);
  var c=getCaptures_(articleNo),map={};
  c.photos.forEach(function(x){map[x.stage+"|"+x.photoType]=x;});
  return {
    articleNo:articleNo,
    DEVELOPMENT:stageStatus_("DEVELOPMENT",map,c.workflow),
    FINISHED:stageStatus_("FINISHED",map,c.workflow)
  };
}

function stageStatus_(stage,map,workflow){
  var items=RULES[stage].map(function(r){
    var x=map[stage+"|"+r.type];
    return {photoType:r.type,label:r.label,required:r.required,captured:!!x,fileName:x?x.fileName:""};
  });
  var total=items.filter(function(x){return x.required;}).length;
  var got=items.filter(function(x){return x.required&&x.captured;}).length;
  var state=workflow[stage]||"DRAFT";
  return {items:items,requiredTotal:total,requiredCaptured:got,complete:got===total,state:state,locked:state==="LOCKED"};
}

function lockStage_(d){
  var a=req_(d.articleNo,"Article Number"),s=stage_(d.stage),u=partner_(d.userEmail);
  var st=getStatus_(a)[s];
  if(st.locked)throw new Error(s+" is already locked.");
  if(!st.complete)throw new Error(s+" is incomplete: "+st.requiredCaptured+"/"+st.requiredTotal+" required captures.");
  writeWorkflow_(a,s,"LOCKED",u);
  return {articleNo:a,stage:s,state:"LOCKED",user:u};
}

function unlockStage_(d){
  var a=req_(d.articleNo,"Article Number"),s=stage_(d.stage),u=admin_(d.userEmail);
  var st=getStatus_(a)[s];
  if(!st.locked)throw new Error(s+" is not locked.");
  writeWorkflow_(a,s,"UNLOCKED",u);
  return {articleNo:a,stage:s,state:"DRAFT",user:u};
}

function savePhoto_(d){
  var a=req_(d.articleNo,"Article Number"),s=stage_(d.stage),t=req_(d.photoType,"Photo Type"),u=partner_(d.userEmail);
  validateArticle_(a);
  if(!RULES[s].some(function(r){return r.type===t;}))throw new Error("Invalid photo type for "+s+".");
  if(getStatus_(a)[s].locked)throw new Error(s+" is locked. Partner & Admin must unlock before changes.");
  if(!d.fileData)throw new Error("Photo data missing.");

  var root=getRoot_(),articles=folder_(root,CFG.ARTICLES_FOLDER),article=folder_(articles,a);
  folder_(article,CFG.COSTING_FOLDER);
  folder_(article,CFG.DIGITAL_FOLDER);
  var photos=folder_(article,"02_PHOTOS"),archive=folder_(photos,CFG.ARCHIVE_FOLDER);
  var dest=folder_(photos,s==="DEVELOPMENT"?CFG.DEVELOPMENT_FOLDER:CFG.FINISHED_FOLDER);

  if(String(d.action).toUpperCase()==="REPLACE"){
    getCaptures_(a).photos.filter(function(x){return x.stage===s&&x.photoType===t&&x.active;}).forEach(function(old){
      try{
        var f=DriveApp.getFileById(old.fileId);
        f.moveTo(archive);
        writePhotoLog_(a,s,t,f.getName(),f.getId(),"ARCHIVED",u);
      }catch(ignore){}
    });
  }

  var raw=String(d.fileData),comma=raw.indexOf(",");
  if(comma<0)throw new Error("Invalid image data.");
  var bytes=Utilities.base64Decode(raw.substring(comma+1));
  var ext=extension_(d.fileName);
  var stamp=Utilities.formatDate(new Date(),Session.getScriptTimeZone(),"yyyyMMdd_HHmmss");
  var name=a.replace(/[^A-Za-z0-9_-]/g,"_")+"_"+t.replace(/[^A-Z0-9_-]/g,"_")+"_"+stamp+"."+ext;
  var file=dest.createFile(Utilities.newBlob(bytes,d.mimeType||"image/jpeg",name));

  writePhotoLog_(a,s,t,file.getName(),file.getId(),String(d.action).toUpperCase()==="REPLACE"?"REPLACED":"CREATED",u);
  return {fileName:file.getName(),fileId:file.getId(),url:file.getUrl()};
}

function getCaptures_(articleNo){
  validateArticle_(articleNo);
  var v=sheet_(CFG.LOG_SHEET).getDataRange().getValues(),photos={},workflow={};
  v.slice(1).forEach(function(r){
    if(String(r[2]||"").trim()!==articleNo)return;
    var s=String(r[3]||"").toUpperCase(),t=String(r[4]||"").toUpperCase(),id=String(r[6]||""),act=String(r[7]||"").toUpperCase();
    if(act==="LOCKED"||act==="UNLOCKED"){workflow[s]=act==="LOCKED"?"LOCKED":"DRAFT";return;}
    if(!s||!t||!id)return;
    var k=s+"|"+t;
    if(act==="ARCHIVED"){if(photos[k]&&photos[k].fileId===id)photos[k].active=false;return;}
    if(act==="CREATED"||act==="REPLACED")photos[k]={articleNo:articleNo,stage:s,photoType:t,fileName:String(r[5]),fileId:id,active:true};
  });
  var active=Object.keys(photos).map(function(k){return photos[k];}).filter(function(x){return x.active;}).map(function(x){
    try{x.url=DriveApp.getFileById(x.fileId).getUrl();}catch(ignore){x.url="";} return x;
  });
  return {photos:active,workflow:workflow};
}

function writePhotoLog_(a,s,t,n,id,act,u){
  sheet_(CFG.LOG_SHEET).appendRow([new Date(),u,a,s,t,n,id,act]);
}
function writeWorkflow_(a,s,act,u){
  sheet_(CFG.LOG_SHEET).appendRow([new Date(),u,a,s,"","","",act]);
}
function partner_(e){
  var x=String(e||"").trim().toLowerCase();
  if(CFG.PARTNER_EMAILS.concat(CFG.ADMIN_EMAILS).map(function(v){return v.toLowerCase();}).indexOf(x)<0)
    throw new Error("Partner access required.");
  return x;
}
function admin_(e){
  var x=String(e||"").trim().toLowerCase();
  if(CFG.ADMIN_EMAILS.map(function(v){return v.toLowerCase();}).indexOf(x)<0)
    throw new Error("Partner & Admin access required.");
  return x;
}
function stage_(s){
  var x=String(s||"").trim().toUpperCase();
  if(!RULES[x])throw new Error("Invalid stage.");
  return x;
}
function req_(v,n){var x=String(v||"").trim();if(!x)throw new Error(n+" is required.");return x;}
function sheet_(n){var s=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(n);if(!s)throw new Error(n+" sheet not found.");return s;}
function getRoot_(){
  var sf=DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId()),p=sf.getParents();
  var parent=p.hasNext()?p.next():DriveApp.getRootFolder();
  return folder_(parent,CFG.ROOT_FOLDER);
}
function folder_(parent,name){var f=parent.getFoldersByName(name);return f.hasNext()?f.next():parent.createFolder(name);}
function extension_(n){var p=String(n||"").split(".");if(p.length<2)return"jpg";var e=p.pop().toLowerCase().replace(/[^a-z0-9]/g,"");return e||"jpg";}
function out_(x){return ContentService.createTextOutput(JSON.stringify(x)).setMimeType(ContentService.MimeType.JSON);}
