// ================================================================
// VENDOR MANAGEMENT — Code.gs
// ================================================================
const CONFIG = {
  SHEET_ID     : '1yDLx7Yb0dn2MNBBto22oPXKRLDG0SIvicix73d5emho',
  ALLOWED_DOMAIN: 'theporter.in'
};

const SN = {
  HIRING:'Hiring', REQUIREMENT:'Requirement', ACCESS:'Access',
  DV:'Data Validation', USERBASE:'User Base', LOG:'Log'
};

const HEADERS = {
  HIRING:['ID','Batch ID','Req ID','Vendor Name','Hiring Date','Full Name','First Name','Last Name',
    'Contact Number','Alt Contact Number','Personal Gmail ID','Work Experience','Languages',
    'Interview Date','Interview By','Interview Status','Rejected Reason','Selected Process',
    'Training Date','Training By','Training Status','Dropout Reason',
    'Certification By','Certification Status','Not Certified Reason','Certified Date',
    'Date of Joining','Final Status','Added By','Added On','Updated By','Updated On'],
  REQUIREMENT:['Req ID','Process','Required Language','Head count','Requirement Skill','Required TAT','Required By','Created By','Created On'],
  ACCESS:['Email','Name','Role','Company','Added By','Added On'],
  DV:['Languages','Vendors','Interview Status','Selected Process','Training Status','Certification Status'],
  USERBASE:['Email','Name','Role','Company','Session Date','Session Time'],
  LOG:['Timestamp','Email','Name','Action','Details','Record ID']
};

const DV_DEFAULTS = {
  languages:['Assamese','Bengali','Bodo','Dogri','Gujarati','Hindi','Kannada','Kashmiri',
    'Konkani','Maithili','Malayalam','Manipuri','Marathi','Nepali','Odia','Punjabi',
    'Sanskrit','Santali','Sindhi','Tamil','Telugu','Urdu','English'],
  vendors:['Degitide','Essencea'],
  interviewStatus:['Selected','Rejected','Call not answer','Not Available Today'],
  selectedProcess:['Research & survey','PTL','Kam Process'],
  trainingStatus:['Completed','Dropout'],
  certStatus:['Certified','Not Certified']
};

function _s(v) {
  if (v === null || v === undefined || v === '') return '';
  if (v instanceof Date) { try { return v.toISOString(); } catch(e) { return ''; } }
  return String(v).trim();
}

function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate().setTitle('Vendor Management')
    .addMetaTag('viewport','width=device-width,initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function initApp() {
  try {
    const email = Session.getActiveUser().getEmail();
    if (!email) return {ok:false, msg:'Not authenticated.'};
    if (email.split('@')[1] !== CONFIG.ALLOWED_DOMAIN)
      return {ok:false, msg:'Access restricted to @'+CONFIG.ALLOWED_DOMAIN+' only.'};

    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    ensureAllHeaders(ss);
    let name = email.split('@')[0];
    try { const p=People.People.get('people/me',{personFields:'names'}); if(p.names&&p.names[0]) name=p.names[0].displayName||name;
    } catch(e){}
    
    const acc = _lookupAccess(ss, email);
    const dvData = _getDVData(ss);
    dvData.reqIds = _getReqIds(ss); // Fetch existing Req IDs for dropdown

    if (!acc.found) {
      const sh = ss.getSheetByName(SN.ACCESS);
      const rows = sh.getDataRange().getValues().filter((_,i)=>i>0&&_[0]);
      if (rows.length===0) {
        sh.appendRow([email,name,'Admin','Porter','system',new Date().toISOString()]);
        _logSession(ss,email,name,'Admin','Porter');
        _writeLog(ss,email,name,'FIRST_LOGIN','Auto Admin/Porter','');
        return {ok:true,email,name,role:'Admin',company:'Porter',dv:dvData};
      }
      return {ok:false, msg:'No access. Contact your Admin.'};
    }
    _logSession(ss,email,name,acc.role,acc.company);
    _writeLog(ss,email,name,'LOGIN','Logged in','');
    return {ok:true,email,name,role:acc.role,company:acc.company,dv:dvData};
  } catch(e) { return {ok:false, msg:e.message}; }
}

function _lookupAccess(ss, email) {
  const rows = ss.getSheetByName(SN.ACCESS).getDataRange().getValues();
  for (let i=1;i<rows.length;i++)
    if (rows[i][0]&&_s(rows[i][0]).toLowerCase()===email.toLowerCase())
      return {found:true, role:_s(rows[i][2]), company:_s(rows[i][3])};
  return {found:false};
}

function ensureAllHeaders(ss) {
  Object.keys(SN).forEach(k=>{
    const sh = ss.getSheetByName(SN[k]) || ss.insertSheet(SN[k]);
    const hdr = HEADERS[k]; if (!hdr) return;
    if (!sh.getRange(1,1).getValue()) {
      sh.getRange(1,1,1,hdr.length).setValues([hdr])
        .setBackground('#1E293B').setFontColor('#FFFFFF').setFontWeight('bold');
      sh.setFrozenRows(1);
    }
  });
  _ensureDVDefaults(ss);
}

function _ensureDVDefaults(ss) {
  const sh = ss.getSheetByName(SN.DV);
  const d  = sh.getDataRange().getValues();
  if (d.length>1&&d[1][0]) return;
  const maxL = Math.max(...Object.values(DV_DEFAULTS).map(a=>a.length));
  const rows = [];
  for (let i=0;i<maxL;i++) rows.push([DV_DEFAULTS.languages[i]||'',DV_DEFAULTS.vendors[i]||'',
    DV_DEFAULTS.interviewStatus[i]||'',DV_DEFAULTS.selectedProcess[i]||'',
    DV_DEFAULTS.trainingStatus[i]||'',DV_DEFAULTS.certStatus[i]||'']);
  if (rows.length) sh.getRange(2,1,rows.length,6).setValues(rows);
}

function _getReqIds(ss) {
  const sh = ss.getSheetByName(SN.REQUIREMENT);
  const d = sh.getDataRange().getValues();
  const ids = [];
  for(let i=1; i<d.length; i++) if(d[i][0]) ids.push(_s(d[i][0]));
  return ids;
}

// ── REQUIREMENT ──
function getRequirementData() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    
    // Get candidate pipeline counts for dynamic mapping
    const shHir = ss.getSheetByName(SN.HIRING);
    const dHir = shHir.getDataRange().getValues();
    const hirHdr = dHir[0] || [];
    const idIdx = hirHdr.indexOf('Req ID'), intIdx = hirHdr.indexOf('Interview Status'), 
          trnIdx = hirHdr.indexOf('Training Status'), crtIdx = hirHdr.indexOf('Certification Status'),
          dojIdx = hirHdr.indexOf('Date of Joining');
    
    const counts = {};
    if (idIdx >= 0) {
      for (let i=1; i<dHir.length; i++) {
        const reqId = _s(dHir[i][idIdx]);
        if (!reqId) continue;
        if (!counts[reqId]) counts[reqId] = { intT:0, intS:0, intR:0, trnT:0, trnC:0, trnD:0, crtT:0, crtC:0, crtN:0, wf:0 };
        
        const iSt = _s(dHir[i][intIdx]), tSt = _s(dHir[i][trnIdx]), cSt = _s(dHir[i][crtIdx]), wf = _s(dHir[i][dojIdx]);
        
        if (iSt) { counts[reqId].intT++; if(iSt==='Selected') counts[reqId].intS++; else if(iSt==='Rejected') counts[reqId].intR++; }
        if (tSt) { counts[reqId].trnT++; if(tSt==='Completed') counts[reqId].trnC++; else if(tSt==='Dropout') counts[reqId].trnD++; }
        if (cSt) { counts[reqId].crtT++; if(cSt==='Certified') counts[reqId].crtC++; else if(cSt==='Not Certified') counts[reqId].crtN++; }
        if (wf)  { counts[reqId].wf++; }
      }
    }

    const shReq = ss.getSheetByName(SN.REQUIREMENT);
    const all = shReq.getDataRange().getValues();
    if (all.length<=1) return {ok:true, data:[]};
    const hdr = all[0];
    const rows = all.slice(1).filter(r => r[0]);

    return {ok:true, data: rows.map(r => {
      const o = {}; hdr.forEach((h,i) => { o[h] = _s(r[i]); });
      o['metrics'] = counts[o['Req ID']] || { intT:0, intS:0, intR:0, trnT:0, trnC:0, trnD:0, crtT:0, crtC:0, crtN:0, wf:0 };
      return o;
    })};
  } catch(e) { return {ok:false, msg:e.message}; }
}

function saveRequirement(rec, ctx) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const sh = ss.getSheetByName(SN.REQUIREMENT);
    const hdr = HEADERS.REQUIREMENT;
    const now = new Date().toISOString();

    if (!rec['Req ID'] || rec['Req ID'] === '') {
      const d = sh.getDataRange().getValues();
      let maxId = 10000;
      for (let i = 1; i < d.length; i++) {
        const val = parseInt(_s(d[i][0]));
        if (!isNaN(val) && val > maxId) maxId = val;
      }
      rec['Req ID'] = (maxId + 1).toString();
      rec['Created By'] = Session.getActiveUser().getEmail();
      rec['Created On'] = now;
      
      const row = hdr.map(h => rec[h] !== undefined ? rec[h] : '');
      sh.appendRow(row);
      _writeLog(ss, Session.getActiveUser().getEmail(), ctx.name, 'ADD_REQUIREMENT', 'Added Req: ' + rec['Req ID'], rec['Req ID']);
      return { ok: true, msg: 'Requirement Added!', data: rec };
    }
    return {ok: false, msg: 'Updating requirements is not supported yet.'};
  } catch(e) { return { ok: false, msg: e.message }; }
}

// ── CANDIDATE PIPELINE (HIRING) ──
function getHiringData(ctx) {
  try {
    const ss  = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const sh  = ss.getSheetByName(SN.HIRING);
    const all = sh.getDataRange().getValues();
    if (all.length<=1) return {ok:true,data:[]};
    const hdr = all[0];
    let rows  = all.slice(1).filter(r=>r[0]!==''&&r[0]!==null&&r[0]!==undefined);
    if (ctx&&ctx.role==='Supervisor') {
      const vIdx = hdr.indexOf('Vendor Name');
      rows = rows.filter(r=>_s(r[vIdx])===ctx.company);
    }
    return {ok:true, data:rows.map(r=>{ const o={}; hdr.forEach((h,i)=>{o[h]=_s(r[i]);}); return o; })};
  } catch(e) { return {ok:false,msg:e.message}; }
}

function saveHiringRecord(rec, ctx) {
  try {
    const ss    = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const sh    = ss.getSheetByName(SN.HIRING);
    const hdr   = HEADERS.HIRING;
    const email = Session.getActiveUser().getEmail();
    const now   = new Date().toISOString();
    rec['Final Status'] = _finalStatus(rec);
    if (!rec['ID']||rec['ID']==='') {
      rec['ID']        = _genID(sh,rec['Vendor Name']);
      rec['Added By']  = email; rec['Added On']  = now;
      rec['Updated By']= email; rec['Updated On']= now;
      sh.appendRow(hdr.map(h=>rec[h]!==undefined?rec[h]:''));
      _writeLog(ss,email,ctx.name||'','ADD_HIRING','Added: '+rec['Full Name'],rec['ID']);
      return {ok:true,id:rec['ID'],msg:'Candidate added!'};
    } else {
      const d = sh.getDataRange().getValues();
      let ri=-1;
      for (let i=1;i<d.length;i++) if(_s(d[i][0])===rec['ID']){ri=i+1;break;}
      if (ri<0) return {ok:false,msg:'Record not found'};
      rec['Updated By']=email; rec['Updated On']=now;
      sh.getRange(ri,1,1,hdr.length).setValues([hdr.map((h,i)=>rec[h]!==undefined?rec[h]:_s(d[ri-1][i]))]);
      _writeLog(ss,email,ctx.name||'','UPDATE_HIRING','Updated: '+rec['Full Name'],rec['ID']);
      return {ok:true,id:rec['ID'],msg:'Candidate updated!'};
    }
  } catch(e) { return {ok:false,msg:e.message}; }
}

function deleteHiringRecord(id, ctx) {
  try {
    const ss=SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const sh=ss.getSheetByName(SN.HIRING);
    const d=sh.getDataRange().getValues();
    for (let i=1;i<d.length;i++) {
      if (_s(d[i][0])===id) {
        sh.deleteRow(i+1);
        _writeLog(ss,Session.getActiveUser().getEmail(),ctx.name||'','DELETE_HIRING','Deleted: '+id,id);
        return {ok:true};
      }
    }
    return {ok:false,msg:'Not found'};
  } catch(e){return{ok:false,msg:e.message};}
}

function _genID(sh,vendor) {
  const map={'Degitide':'D','Essencea':'E'};
  const prefix=map[vendor]||vendor.charAt(0).toUpperCase();
  const d=sh.getDataRange().getValues();
  let max=0;
  for (let i=1;i<d.length;i++){const id=_s(d[i][0]);if(id.startsWith(prefix)){const n=parseInt(id.slice(prefix.length));if(!isNaN(n)&&n>max)max=n;}}
  return prefix+(max+1);
}

function _finalStatus(r){
  if(r['Date of Joining']) return 'Working from';
  if(r['Certification Status'])return r['Certification Status'];
  if(r['Training Status'])return r['Training Status'];
  if(r['Interview Status'])return r['Interview Status'];
  return '';
}

// ── DATA VALIDATION ──
function getDVData(){try{const ss=SpreadsheetApp.openById(CONFIG.SHEET_ID);return{ok:true,data:_getDVData(ss)};}catch(e){return{ok:false,msg:e.message};}}
function _getDVData(ss){
  const sh=ss.getSheetByName(SN.DV);const d=sh.getDataRange().getValues();
  const res={languages:[],vendors:[],interviewStatus:[],selectedProcess:[],trainingStatus:[],certStatus:[]};
  for(let i=1;i<d.length;i++){
    if(d[i][0])res.languages.push(_s(d[i][0]));if(d[i][1])res.vendors.push(_s(d[i][1]));
    if(d[i][2])res.interviewStatus.push(_s(d[i][2]));if(d[i][3])res.selectedProcess.push(_s(d[i][3]));
    if(d[i][4])res.trainingStatus.push(_s(d[i][4]));if(d[i][5])res.certStatus.push(_s(d[i][5]));
  }
  return res;
}
function addDVOption(colKey,value,ctx){
  try{
    const colMap={languages:1,vendors:2,interviewStatus:3,selectedProcess:4,trainingStatus:5,certStatus:6};
    const col=colMap[colKey];if(!col)return{ok:false,msg:'Invalid'};
    const ss=SpreadsheetApp.openById(CONFIG.SHEET_ID);const sh=ss.getSheetByName(SN.DV);
    const lr=Math.max(sh.getLastRow()-1,1);
    const existing=sh.getRange(2,col,lr,1).getValues().flat().filter(v=>v).map(v=>v.toString().toLowerCase());
    if(existing.includes(value.toLowerCase()))return{ok:false,msg:'Already exists'};
    sh.getRange(existing.length+2,col).setValue(value);
    _writeLog(ss,Session.getActiveUser().getEmail(),ctx.name||'','ADD_DV','Added "'+value+'" to '+colKey,'');
    return{ok:true};
  }catch(e){return{ok:false,msg:e.message};}
}
function saveDVAll(newData,ctx){
  try{
    const ss=SpreadsheetApp.openById(CONFIG.SHEET_ID);const sh=ss.getSheetByName(SN.DV);
    const lr=sh.getLastRow();if(lr>1)sh.getRange(2,1,lr-1,6).clearContent();
    const cols=['languages','vendors','interviewStatus','selectedProcess','trainingStatus','certStatus'];
    const maxL=Math.max(...cols.map(c=>(newData[c]||[]).length));
    if(maxL>0){const rows=[];for(let i=0;i<maxL;i++)rows.push(cols.map(c=>(newData[c]&&newData[c][i])?newData[c][i]:''));sh.getRange(2,1,rows.length,6).setValues(rows);}
    _writeLog(ss,Session.getActiveUser().getEmail(),ctx.name||'','SAVE_DV','Saved all DV','');
    return{ok:true};
  }catch(e){return{ok:false,msg:e.message};}
}

// ── PEOPLE ──
function searchPeople(query){
  try{
    const resp=People.People.searchDirectoryPeople({query,readMask:'names,emailAddresses',sources:['DIRECTORY_SOURCE_TYPE_DOMAIN_PROFILE'],pageSize:8});
    const results=[];
    if(resp.people)resp.people.forEach(p=>{const email=(p.emailAddresses||[])[0]?.value||'';const name=(p.names||[])[0]?.displayName||email;if(email)results.push({email,name});});
    return{ok:true,results};
  }catch(e){return{ok:false,results:[],msg:e.message};}
}

// ── ACCESS & USER BASE & LOG ──
function getAccessList(){try{const d=SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheetByName(SN.ACCESS).getDataRange().getValues();return{ok:true,data:d.slice(1).filter(r=>r[0]).map(r=>({email:_s(r[0]),name:_s(r[1]),role:_s(r[2]),company:_s(r[3]),addedBy:_s(r[4]),addedOn:_s(r[5])}))}}catch(e){return{ok:false,msg:e.message};}}
function saveAccess(rec,ctx){try{const ss=SpreadsheetApp.openById(CONFIG.SHEET_ID);const sh=ss.getSheetByName(SN.ACCESS);const d=sh.getDataRange().getValues();const email=Session.getActiveUser().getEmail();let ri=-1;for(let i=1;i<d.length;i++)if(d[i][0]&&_s(d[i][0]).toLowerCase()===rec.email.toLowerCase()){ri=i+1;break;}if(ri>0){sh.getRange(ri,1,1,6).setValues([[rec.email,rec.name,rec.role,rec.company,email,_s(d[ri-1][5])]]);_writeLog(ss,email,ctx.name||'','UPDATE_ACCESS','Updated '+rec.email,'');}else{sh.appendRow([rec.email,rec.name,rec.role,rec.company,email,new Date().toISOString()]);_writeLog(ss,email,ctx.name||'','ADD_ACCESS','Added '+rec.email,'');}return{ok:true};}catch(e){return{ok:false,msg:e.message};}}
function removeAccess(targetEmail,ctx){try{const ss=SpreadsheetApp.openById(CONFIG.SHEET_ID);const sh=ss.getSheetByName(SN.ACCESS);const d=sh.getDataRange().getValues();for(let i=1;i<d.length;i++)if(d[i][0]&&_s(d[i][0]).toLowerCase()===targetEmail.toLowerCase()){sh.deleteRow(i+1);_writeLog(ss,Session.getActiveUser().getEmail(),ctx.name||'','REMOVE_ACCESS','Removed '+targetEmail,'');return{ok:true};}return{ok:false,msg:'Not found'};}catch(e){return{ok:false,msg:e.message};}}
function _logSession(ss,email,name,role,company){try{const tz=Session.getScriptTimeZone();const now=new Date();ss.getSheetByName(SN.USERBASE).appendRow([email,name,role,company,Utilities.formatDate(now,tz,'yyyy-MM-dd'),Utilities.formatDate(now,tz,'HH:mm:ss')]);}catch(e){}}
function getUserBaseData(){try{const d=SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheetByName(SN.USERBASE).getDataRange().getValues();if(d.length<=1)return{ok:true,data:[],summary:{total:0,admin:0,supervisor:0}};const logs=d.slice(1).filter(r=>r[0]).map(r=>({email:_s(r[0]),name:_s(r[1]),role:_s(r[2]),company:_s(r[3]),date:_s(r[4]),time:_s(r[5])}));const unique=[...new Set(logs.map(l=>l.email))];const roleOf={};logs.forEach(l=>{if(!roleOf[l.email])roleOf[l.email]=l.role;});return{ok:true,data:[...logs].reverse(),summary:{total:unique.length,admin:unique.filter(e=>roleOf[e]==='Admin').length,supervisor:unique.filter(e=>roleOf[e]==='Supervisor').length}};}catch(e){return{ok:false,msg:e.message};}}
function _writeLog(ss,email,name,action,details,recId){try{ss.getSheetByName(SN.LOG).appendRow([new Date().toISOString(),email,name||'',action,details||'',recId||'']);}catch(e){}}
function getLogData(){try{const d=SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheetByName(SN.LOG).getDataRange().getValues();if(d.length<=1)return{ok:true,data:[]};return{ok:true,data:d.slice(1).filter(r=>r[0]).map(r=>({timestamp:_s(r[0]),email:_s(r[1]),name:_s(r[2]),action:_s(r[3]),details:_s(r[4]),recordId:_s(r[5])})).reverse()};}catch(e){return{ok:false,msg:e.message};}}
