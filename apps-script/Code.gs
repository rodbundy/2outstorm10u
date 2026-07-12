const PUBLIC_SHEETS = {
  'Website Players': 'players',
  'Website Calendar': 'calendar',
  'Website Videos': 'videos',
  'Website Announcements': 'announcements',
  'Website Shoutouts': 'shoutouts',
  'Website Picture of the Week': 'picture',
  'Website Gallery': 'gallery'
};

const MANAGED_SHEETS = [
  'Website Settings',
  'Website Players',
  'Website Calendar',
  'Website Videos',
  'Website Announcements',
  'Website Shoutouts',
  'Website Picture of the Week',
  'Website Gallery',
  'Parent Codes',
  'Availability',
  'Admin Log'
];

const SHEET_HEADERS = {
  'Website Settings': ['Key','Value','Notes'],
  'Website Players': ['Show','Approved','PlayerID','Jersey','FirstName','Positions','BatsThrows','ClassYear','PhotoURL','BackgroundURL','StrongestPart','SeasonGoal','FavoriteColor','FavoritePlayer','Excitement','Quote','YouTubePlaylist','SortOrder'],
  'Website Calendar': ['Show','Approved','EventID','Date','Time','EndTime','Type','Title','Opponent','Location','Field','Link','Status','Result','ScoreUs','ScoreThem','Notes','SortOrder'],
  'Website Videos': ['Show','Approved','VideoID','Category','Title','Description','YouTubeID','ThumbnailURL','Date','Opponent','Event','Score','Featured','PlayerID','SortOrder','VideoURL'],
  'Website Announcements': ['Show','Approved','AlertLevel','Title','Message','StartDate','EndDate','ButtonText','ButtonURL','SortOrder'],
  'Website Shoutouts': ['Show','Approved','Category','Title','Player','Message','Date','SortOrder'],
  'Website Picture of the Week': ['Show','Approved','Title','Week','Caption','ImageURL','SubmittedBy'],
  'Website Gallery': ['Show','Approved','PhotoID','Title','Caption','ImageURL','Date','PlayerID','Category','SortOrder'],
  'Parent Codes': ['PlayerID','PlayerName','ParentCode','Active','LastUpdated'],
  'Availability': ['EventID','EventDate','EventTitle','PlayerID','PlayerName','Status','Notes','UpdatedAt'],
  'Admin Log': ['Timestamp','Action','Sheet','Row','Details']
};

const ADMIN_EDIT_SHEETS = MANAGED_SHEETS.filter(name => name !== 'Admin Log');
const SESSION_SECONDS = 21600;

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Storm Website')
    .addItem('Set Coach Password', 'setCoachPasswordFromMenu')
    .addItem('Generate Missing Parent Codes', 'generateMissingParentCodesFromMenu')
    .addItem('Create / Repair Website Tabs', 'setupStormWorkbook')
    .addSeparator()
    .addItem('Show Web App URL', 'showWebAppUrl')
    .addToUi();
}

function doGet(e) {
  const page = String((e && e.parameter && e.parameter.page) || '').toLowerCase();
  if (page === 'admin') return serveHtml_('Admin', 'Coach Edit Board');
  if (page === 'availability') return serveHtml_('Availability', 'Storm Availability');
  const action = String((e && e.parameter && e.parameter.action) || 'public').toLowerCase();
  if (action === 'public') return jsonOutput_({ok:true,data:getPublicData_()}, e);
  return jsonOutput_({ok:false,error:'Unknown request.'}, e);
}

function serveHtml_(file, title) {
  const template = HtmlService.createTemplateFromFile(file);
  template.webAppUrl = ScriptApp.getService().getUrl() || '';
  template.teamName = getSettingsObject_().teamName || '2 Out Storm 10U';
  return template.evaluate().setTitle(title + ' | 2 Out Storm 10U')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport','width=device-width, initial-scale=1');
}

function include(filename) { return HtmlService.createHtmlOutputFromFile(filename).getContent(); }

function jsonOutput_(payload, e) {
  const json = JSON.stringify(payload);
  const callback = e && e.parameter && e.parameter.callback;
  if (callback && /^[A-Za-z_$][\w$\.]*$/.test(callback)) {
    return ContentService.createTextOutput(callback + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function setupStormWorkbook() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  MANAGED_SHEETS.forEach(name => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) sheet = ss.insertSheet(name);
    ensureHeaders_(sheet, SHEET_HEADERS[name]);
    styleSheet_(sheet, name);
  });
  seedSettings_();
  seedPublicContent_();
  SpreadsheetApp.flush();
  try { SpreadsheetApp.getUi().alert('Storm website tabs are ready.'); } catch (e) {}
  return true;
}

function ensureHeaders_(sheet, required) {
  const scanRows=Math.min(Math.max(sheet.getLastRow(),1),5);
  const scanWidth=Math.max(sheet.getLastColumn(),required.length,1);
  const scan=sheet.getRange(1,1,scanRows,scanWidth).getDisplayValues();
  let headerRow=1,bestMatches=-1;
  scan.forEach((row,index)=>{
    const values=row.map(String).map(s=>s.trim()).filter(Boolean);
    const matches=required.filter(h=>values.indexOf(h)!==-1).length;
    if(matches>bestMatches){bestMatches=matches;headerRow=index+1}
  });
  // Older decorated templates placed headers on row 4. Move the database header to row 1 automatically.
  if(headerRow>1&&bestMatches>=Math.min(3,required.length))sheet.deleteRows(1,headerRow-1);
  const width=Math.max(sheet.getLastColumn(),1);
  let current=sheet.getRange(1,1,1,width).getDisplayValues()[0].map(String).map(s=>s.trim()).filter(Boolean);
  if(!current.length||bestMatches<=0){
    sheet.getRange(1,1,1,required.length).setValues([required]);
    return required;
  }
  const missing=required.filter(h=>current.indexOf(h)===-1);
  if(missing.length)sheet.getRange(1,current.length+1,1,missing.length).setValues([missing]);
  return current.concat(missing);
}

function styleSheet_(sheet, name) {
  const headers = SHEET_HEADERS[name];
  sheet.setFrozenRows(1);
  sheet.getRange(1,1,1,headers.length)
    .setBackground('#32174f').setFontColor('#ffffff').setFontWeight('bold')
    .setHorizontalAlignment('center').setVerticalAlignment('middle');
  sheet.setRowHeight(1, 34);
  sheet.getRange(1,1,Math.max(sheet.getMaxRows(),2),headers.length).setWrap(true);
  headers.forEach((h,i) => sheet.setColumnWidth(i+1, ['Message','Notes','Description','Caption','StrongestPart','SeasonGoal','Quote'].includes(h) ? 250 : ['PhotoURL','BackgroundURL','ThumbnailURL','ImageURL','Link','ButtonURL'].includes(h) ? 220 : 120));
  if (name === 'Parent Codes' || name === 'Availability' || name === 'Admin Log') sheet.setTabColor('#ff6b25'); else sheet.setTabColor('#8e49d7');
  if (headers.includes('Show')) addYesNoValidation_(sheet, headers.indexOf('Show')+1);
  if (headers.includes('Approved')) addYesNoValidation_(sheet, headers.indexOf('Approved')+1);
  if (headers.includes('Featured')) addYesNoValidation_(sheet, headers.indexOf('Featured')+1);
  if (name === 'Parent Codes') addListValidation_(sheet, headers.indexOf('Active')+1, ['YES','NO']);
  if (name === 'Availability') addListValidation_(sheet, headers.indexOf('Status')+1, ['READY FOR THE STORM','NOT AVAILABLE','FORECAST UNCERTAIN']);
}

function addYesNoValidation_(sheet, col) { addListValidation_(sheet,col,['YES','NO']); }
function addListValidation_(sheet,col,values) {
  const rule=SpreadsheetApp.newDataValidation().requireValueInList(values,true).setAllowInvalid(false).build();
  sheet.getRange(2,col,Math.max(sheet.getMaxRows()-1,1),1).setDataValidation(rule);
}

function seedSettings_() {
  const sheet=getSheet_('Website Settings');
  if (sheet.getLastRow()>1) return;
  const rows=[
    ['teamName','2 Out Storm 10U','Main public team name'],
    ['tagline','Together. Tougher.','Main team tagline'],
    ['coach','Coach Rodney Bundy','Coach contact name'],
    ['phone','(859) 414-4622','Public phone display'],
    ['smsPhone','+18594144622','Digits used by text-message buttons'],
    ['email','rodbundy@yahoo.com','Coach email'],
    ['location','CAP, 999 Speedway Drive, Lawrenceburg, IN 47025','Home field / contact location'],
    ['interest','https://forms.gle/4PgA83dnnVBXMo2h8','Existing Player Interest form'],
    ['message','','Paste Private Parent Message form URL'],
    ['profile','','Paste Accepted Player Profile / Update Athlete Info form URL'],
    ['photo','','Paste Team Photo / Picture of the Week form URL'],
    ['gamechanger','','Paste GameChanger team URL'],
    ['facebook','','Optional Facebook page URL'],
    ['instagram','','Optional Instagram URL']
  ];
  sheet.getRange(2,1,rows.length,3).setValues(rows);
}

function seedPublicContent_() {
  const players=getSheet_('Website Players');
  if (players.getLastRow()===1) {
    const rows=[];
    for(let i=1;i<=12;i++) rows.push(['YES','YES','player-'+String(i).padStart(2,'0'),i,'Player '+i,i%3===1?'OF / INF':i%3===2?'C / Utility':'P / INF','R/R','2035','assets/img/players/player-'+String(i).padStart(2,'0')+'.svg','','Energy, effort, and coachability','Grow stronger every practice','','','10','The storm is building.','',i]);
    players.getRange(2,1,rows.length,rows[0].length).setValues(rows);
  }
  const cal=getSheet_('Website Calendar');
  if (cal.getLastRow()===1) cal.getRange(2,1,3,18).setValues([
    ['YES','YES','practice-jul22',new Date(2026,6,22),'6:00 PM','8:00 PM','Practice','Final Open Practice','','CAP Complex','','https://forms.gle/4PgA83dnnVBXMo2h8','Storm Forming','','','','Open to all positions.',1],
    ['YES','YES','tryout-jul25',new Date(2026,6,25),'10:00 AM','','Tryout','Saturday Tryout','','CAP Complex','','https://forms.gle/4PgA83dnnVBXMo2h8','Storm Watch','','','','Registration begins at 9:30 AM.',2],
    ['YES','YES','tryout-jul29',new Date(2026,6,29),'6:00 PM','','Tryout','Wednesday Tryout','','CAP Complex','','https://forms.gle/4PgA83dnnVBXMo2h8','Storm Watch','','','','Registration begins at 5:30 PM.',3]
  ]);
  const vids=getSheet_('Website Videos');
  if (vids.getLastRow()===1) vids.getRange(2,1,8,16).setValues([
    ['YES','YES','throw-1','Control the Wind','Throwing Progression 1','Ready feet · Glove side · Finish the throw','_psW-80nllQ','','','','Development','','YES','',1,''],
    ['YES','YES','throw-2','Control the Wind','Throwing Progression 2','Build a repeatable throwing pattern','53JNddN4KXk','','','','Development','','NO','',2,''],
    ['YES','YES','hit-1','Lightning Strikes','Hitting Work 1','Balance · Contact point · Finish strong','OEkDXHoqtcE','','','','Development','','NO','',3,''],
    ['YES','YES','hit-2','Lightning Strikes','Hitting Work 2','Strong lower half and confident swings','kvvCsUGYDVM','','','','Development','','NO','',4,''],
    ['YES','YES','field-1','Own the Ground','Fielding Footwork 1','Ready position · Glove out front · Field through','AJMacRtbx9s','','','','Development','','NO','',5,''],
    ['YES','YES','field-2','Own the Ground','Fielding Footwork 2','Move through the ball with purpose','LqN39dIf0s8','','','','Development','','NO','',6,''],
    ['YES','YES','catch-1','Protect the Eye','Catching and Receiving 1','Strong stance · Quiet glove · Stay low','BT2ULwoAWlY','','','','Development','','NO','',7,''],
    ['YES','YES','catch-2','Protect the Eye','Catching and Receiving 2','Receive, frame, and recover','Pe39wZrSkEs','','','','Development','','NO','',8,'']
  ]);
  const announcements=getSheet_('Website Announcements');
  if (announcements.getLastRow()===1) announcements.getRange(2,1,1,10).setValues([['YES','YES','Storm Warning','The Storm Is Building','Schedules, roster updates, videos, and family information will all live here.',new Date(2026,6,1),new Date(2026,11,31),'Join the Storm','https://forms.gle/4PgA83dnnVBXMo2h8',1]]);
  const shout=getSheet_('Website Shoutouts');
  if (shout.getLastRow()===1) shout.getRange(2,1,1,8).setValues([['YES','YES','Ground Shaker','Maximum Effort','The Entire Storm','Hot weather did not stop the effort. Every athlete competed and got quality reps.',new Date(2026,6,8),1]]);
  const pic=getSheet_('Website Picture of the Week');
  if (pic.getLastRow()===1) pic.getRange(2,1,1,7).setValues([['YES','YES','Picture of the Week','Storm Season','Replace this with a parent-approved team photo from the Coach Edit Board.','assets/img/storm-logo.svg','Coach Rodney']]);
}

function getSheet_(name) {
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  let sheet=ss.getSheetByName(name);
  if(!sheet){sheet=ss.insertSheet(name);ensureHeaders_(sheet,SHEET_HEADERS[name]);styleSheet_(sheet,name)}
  return sheet;
}

function getSettingsObject_() {
  const rows=readRows_('Website Settings',false);
  const out={}; rows.forEach(r=>{if(r.Key)out[String(r.Key).trim()]=r.Value});
  return out;
}

function getPublicData_() {
  setupIfNeeded_();
  const data={settings:getSettingsObject_()};
  Object.keys(PUBLIC_SHEETS).forEach(name=>data[PUBLIC_SHEETS[name]]=readRows_(name,false,true));
  return data;
}

function setupIfNeeded_() {
  if(!SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Website Settings')) setupStormWorkbook();
}

function readRows_(sheetName, includeRowNumber, publicOnly) {
  const sheet=getSheet_(sheetName);
  const values=sheet.getDataRange().getValues();
  if(values.length<2)return [];
  const headers=values[0].map(h=>String(h).trim());
  return values.slice(1).map((row,index)=>{
    const obj={}; headers.forEach((h,i)=>obj[h]=serializeValue_(row[i]));
    if(includeRowNumber)obj._rowNumber=index+2;
    return obj;
  }).filter(obj=>{
    const hasValue=Object.keys(obj).some(k=>k!=='_rowNumber'&&String(obj[k]??'').trim()!=='');
    if(!hasValue)return false;
    if(!publicOnly)return true;
    if(headers.indexOf('Show')<0||headers.indexOf('Approved')<0)return true;
    return String(obj.Show).toUpperCase()==='YES'&&String(obj.Approved).toUpperCase()==='YES';
  }).sort((a,b)=>(Number(a.SortOrder)||999)-(Number(b.SortOrder)||999));
}

function serializeValue_(v) {
  if(Object.prototype.toString.call(v)==='[object Date]') return Utilities.formatDate(v,Session.getScriptTimeZone(),'yyyy-MM-dd');
  return v;
}

function setCoachPasswordFromMenu() {
  const ui=SpreadsheetApp.getUi();
  const response=ui.prompt('Set Coach Password','Enter a strong password for the Coach Edit Board.',ui.ButtonSet.OK_CANCEL);
  if(response.getSelectedButton()!==ui.Button.OK)return;
  const password=response.getResponseText().trim();
  if(password.length<8){ui.alert('Use at least 8 characters.');return}
  PropertiesService.getScriptProperties().setProperty('COACH_PASSWORD_HASH',hash_(password));
  ui.alert('Coach password saved.');
}

function adminLogin(password) {
  const stored=PropertiesService.getScriptProperties().getProperty('COACH_PASSWORD_HASH');
  if(!stored)return {ok:false,error:'Coach password has not been set. Use the Storm Website menu in the spreadsheet.'};
  if(hash_(String(password||''))!==stored)return {ok:false,error:'Incorrect password.'};
  const token=Utilities.getUuid();
  CacheService.getScriptCache().put('admin:'+token,'1',SESSION_SECONDS);
  log_('LOGIN','Coach Edit Board','', 'Successful login');
  return {ok:true,token:token,expiresIn:SESSION_SECONDS};
}

function requireAdmin_(token) {
  if(!token||CacheService.getScriptCache().get('admin:'+token)!=='1')throw new Error('Your coach session expired. Sign in again.');
}

function adminGetData(token) {
  requireAdmin_(token); setupIfNeeded_();
  const data={},schema={};
  ADMIN_EDIT_SHEETS.forEach(name=>{data[name]=readRows_(name,true,false);schema[name]=SHEET_HEADERS[name]});
  return {ok:true,data:data,schema:schema,webAppUrl:ScriptApp.getService().getUrl()||''};
}

function adminSaveRow(token,sheetName,rowObj,rowNumber) {
  requireAdmin_(token);
  if(ADMIN_EDIT_SHEETS.indexOf(sheetName)<0)throw new Error('This sheet cannot be edited here.');
  const lock=LockService.getScriptLock();lock.waitLock(10000);
  try{
    const sheet=getSheet_(sheetName),headers=SHEET_HEADERS[sheetName];
    let row=Number(rowNumber||0);
    if(!row||row<2)row=sheet.getLastRow()+1;
    const values=headers.map(h=>coerceForSheet_(h,rowObj[h]));
    sheet.getRange(row,1,1,headers.length).setValues([values]);
    log_('SAVE',sheetName,row,JSON.stringify(rowObj).slice(0,500));
    return {ok:true,rowNumber:row,row:objectFromRow_(headers,values,row)};
  }finally{lock.releaseLock()}
}

function adminDeleteRow(token,sheetName,rowNumber) {
  requireAdmin_(token);
  if(ADMIN_EDIT_SHEETS.indexOf(sheetName)<0)throw new Error('This sheet cannot be edited here.');
  const row=Number(rowNumber); if(row<2)throw new Error('Invalid row.');
  getSheet_(sheetName).deleteRow(row); log_('DELETE',sheetName,row,'Row deleted');
  return {ok:true};
}

function adminUploadImage(token,sheetName,rowNumber,fieldName,fileData,fileName,mimeType) {
  return adminUploadMedia(token,sheetName,rowNumber,fieldName,fileData,fileName,mimeType);
}

function adminUploadMedia(token,sheetName,rowNumber,fieldName,fileData,fileName,mimeType) {
  requireAdmin_(token);
  const allowed={
    'Website Players':{PhotoURL:'image',BackgroundURL:'image'},
    'Website Picture of the Week':{ImageURL:'image'},
    'Website Gallery':{ImageURL:'image'},
    'Website Videos':{ThumbnailURL:'image',VideoURL:'video'}
  };
  if(!allowed[sheetName]||!allowed[sheetName][fieldName])throw new Error('Upload is not allowed for this field.');
  const kind=allowed[sheetName][fieldName];
  const type=String(mimeType||'').toLowerCase();
  if(kind==='image'&&!type.startsWith('image/'))throw new Error('Choose an image file for this field.');
  if(kind==='video'&&!type.startsWith('video/'))throw new Error('Choose a video file for this field.');
  const raw=String(fileData||'');
  const approxBytes=Math.floor((raw.length-(raw.indexOf(',')+1))*0.75);
  if(approxBytes>45*1024*1024)throw new Error('This upload is over 45 MB. Upload the video to Google Drive manually, share it with anyone who has the link, and paste the link into VideoURL.');
  const bytes=Utilities.base64Decode(raw.replace(/^data:[^;]+;base64,/,''));
  const blob=Utilities.newBlob(bytes,mimeType||(kind==='video'?'video/mp4':'image/jpeg'),fileName||(kind==='video'?'storm-video.mp4':'storm-photo.jpg'));
  const folder=getOrCreateFolder_(kind==='video'?'2 Out Storm Website Videos':'2 Out Storm Website Photos');
  const file=folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK,DriveApp.Permission.VIEW);
  const id=file.getId();
  const url=kind==='video'?'https://drive.google.com/file/d/'+id+'/preview':'https://drive.google.com/uc?export=view&id='+id;
  const sheet=getSheet_(sheetName),headers=SHEET_HEADERS[sheetName],row=Number(rowNumber);
  const col=headers.indexOf(fieldName)+1; if(row<2||col<1)throw new Error('Save the row before uploading the file.');
  sheet.getRange(row,col).setValue(url); log_('UPLOAD',sheetName,row,fieldName+': '+file.getName());
  return {ok:true,url:url,rowNumber:row,fileId:id,kind:kind};
}

function getOrCreateFolder_(name) {
  const folders=DriveApp.getFoldersByName(name); return folders.hasNext()?folders.next():DriveApp.createFolder(name);
}

function generateMissingParentCodesFromMenu(){const result=generateMissingParentCodes_();SpreadsheetApp.getUi().alert(result.created+' parent code(s) created.');}
function adminGenerateParentCodes(token){requireAdmin_(token);return generateMissingParentCodes_();}
function generateMissingParentCodes_(){
  setupIfNeeded_();
  const players=readRows_('Website Players',false,false).filter(p=>p.PlayerID&&String(p.Show).toUpperCase()!=='NO');
  const sheet=getSheet_('Parent Codes'),existing=readRows_('Parent Codes',false,false),byId={};
  existing.forEach(r=>byId[r.PlayerID]=r); let created=0;
  players.forEach(p=>{
    if(!byId[p.PlayerID]){sheet.appendRow([p.PlayerID,p.FirstName,uniqueCode_(),'YES',new Date()]);created++}
    else if(byId[p.PlayerID].PlayerName!==p.FirstName){const row=findRowByField_(sheet,'PlayerID',p.PlayerID);if(row)sheet.getRange(row,2).setValue(p.FirstName)}
  });
  log_('GENERATE','Parent Codes','',created+' created');return {ok:true,created:created};
}

function uniqueCode_(){
  const existing=new Set(readRows_('Parent Codes',false,false).map(r=>String(r.ParentCode)));
  let code; do{code=String(Math.floor(100000+Math.random()*900000))}while(existing.has(code)); return code;
}

function parentLookup(code) {
  setupIfNeeded_();
  const clean=String(code||'').replace(/\D/g,'');
  const rec=readRows_('Parent Codes',false,false).find(r=>String(r.ParentCode)===clean&&String(r.Active).toUpperCase()!=='NO');
  if(!rec)return {ok:false,error:'That player code was not found.'};
  const today=Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM-dd');
  const events=readRows_('Website Calendar',false,true).filter(e=>!e.Date||String(e.Date)>=today).slice(0,12);
  const availability=readRows_('Availability',false,false).filter(a=>a.PlayerID===rec.PlayerID);
  const current={};availability.forEach(a=>current[a.EventID]=a);
  return {ok:true,player:{PlayerID:rec.PlayerID,PlayerName:rec.PlayerName},events:events,current:current};
}

function parentUpdateAvailability(code,eventId,status,notes) {
  setupIfNeeded_();
  const lookup=parentLookup(code); if(!lookup.ok)return lookup;
  const allowed=['READY FOR THE STORM','NOT AVAILABLE','FORECAST UNCERTAIN'];
  if(allowed.indexOf(status)<0)return {ok:false,error:'Choose a valid availability status.'};
  const event=lookup.events.find(e=>e.EventID===eventId)||readRows_('Website Calendar',false,true).find(e=>e.EventID===eventId);
  if(!event)return {ok:false,error:'That event was not found.'};
  const sheet=getSheet_('Availability'),rows=readRows_('Availability',true,false);
  const existing=rows.find(r=>r.PlayerID===lookup.player.PlayerID&&r.EventID===eventId);
  const values=[eventId,event.Date,event.Title,lookup.player.PlayerID,lookup.player.PlayerName,status,String(notes||'').slice(0,500),new Date()];
  if(existing)sheet.getRange(existing._rowNumber,1,1,values.length).setValues([values]); else sheet.appendRow(values);
  log_('AVAILABILITY','Availability',existing?existing._rowNumber:sheet.getLastRow(),lookup.player.PlayerName+' - '+status);
  return {ok:true,message:'Availability updated for '+lookup.player.PlayerName+'.'};
}

function findRowByField_(sheet,field,value){const headers=SHEET_HEADERS[sheet.getName()],col=headers.indexOf(field)+1;if(col<1)return 0;const vals=sheet.getRange(2,col,Math.max(sheet.getLastRow()-1,1),1).getDisplayValues();for(let i=0;i<vals.length;i++)if(String(vals[i][0])===String(value))return i+2;return 0}
function objectFromRow_(headers,values,row){const obj={_rowNumber:row};headers.forEach((h,i)=>obj[h]=serializeValue_(values[i]));return obj}
function coerceForSheet_(header,value){if(['Date','StartDate','EndDate','UpdatedAt','LastUpdated','Timestamp','EventDate'].includes(header)&&value){const d=new Date(value);if(!isNaN(d))return d}return value===undefined?'':value}
function log_(action,sheet,row,details){try{getSheet_('Admin Log').appendRow([new Date(),action,sheet,row,details])}catch(e){}}
function hash_(text){const digest=Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,text,Utilities.Charset.UTF_8);return digest.map(b=>(b+256)%256).map(b=>('0'+b.toString(16)).slice(-2)).join('')}
function showWebAppUrl(){SpreadsheetApp.getUi().alert('Web App URL',ScriptApp.getService().getUrl()||'Deploy the script as a web app first.',SpreadsheetApp.getUi().ButtonSet.OK)}
