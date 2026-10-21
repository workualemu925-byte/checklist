// Google Sheet > Extensions > Apps Script > ይህን ለጥፉ > Deploy > Web app (Execute as: Me, Access: Anyone)
const KEY=''; // ከፈለጉ ሚስጥር ቃል ያስገቡ (በሲስተሙ ክላውድ ገጽ ላይም ተመሳሳዩን ያስገቡ)
function out(o){return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON)}
function sh(n){const s=SpreadsheetApp.getActive();return s.getSheetByName(n)||s.insertSheet(n)}
function doGet(e){
  if(KEY&&e.parameter.key!==KEY)return out({error:'የሚስጥር ቃል ስህተት'});
  if(e.parameter.action=='pull'){const s=sh('DB'),n=s.getLastRow();
    const t=n?s.getRange(1,1,n,1).getValues().map(r=>r[0]).join(''):'';return out({db:t?JSON.parse(t):null})}
  return out({ok:1})}
function doPost(e){
  const b=JSON.parse(e.postData.contents);
  if(KEY&&b.key!==KEY)return out({error:'የሚስጥር ቃል ስህተት'});
  const l=LockService.getScriptLock();l.waitLock(20000);
  const t=JSON.stringify(b.db),s=sh('DB');s.clear();
  const ch=t.match(/[\s\S]{1,40000}/g)||[];
  s.getRange(1,1,ch.length,1).setValues(ch.map(c=>[c]));
  const r=sh('Records');r.clear();
  const rows=[['ቀን','ተቋም','ወረዳ/ከተማ','ክፍል','ነጥብ','%','ውጤት','ሃላፊ']].concat(b.db.rec.map(x=>[x.date,x.subject,x.area,x.st,x.sc+'/'+x.mx,x.pct,x.res,x.officer]));
  r.getRange(1,1,rows.length,8).setValues(rows);
  const q=sh('Letters');q.clear();
  const lr=[['ቁጥር','ቀን','ተቋም']].concat(b.db.let.map(x=>[x.ref,x.date,x.org]));
  q.getRange(1,1,lr.length,3).setValues(lr);
  l.releaseLock();return out({message:'saved '+b.db.rec.length})}
