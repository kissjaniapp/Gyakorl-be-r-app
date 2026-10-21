// Eredmény beíró – Google Táblázat fogadó
// Másold ezt a teljes kódot a Google Apps Script szerkesztőbe.
// Fontos: a SECRET értéke egyezzen az Android alkalmazásban megadott közös kulccsal.
const SECRET = 'CSERELD_LE_EGY_SAJAT_HOSSZU_VELETLEN_KULCSRA';
const RESULTS_SHEET = 'Versenyeredmények';
const RANKING_SHEET = 'Összesített ranglista';

function doGet() {
  return ContentService.createTextOutput('Az Eredmény beíró Google Táblázat fogadó működik.');
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const p = JSON.parse(e && e.postData && e.postData.contents || '{}');
    if (!SECRET || SECRET.indexOf('CSERELD_LE_') === 0 || p.secret !== SECRET) {
      return out_({ok:false, error:'Hibás vagy be nem állított közös kulcs.'});
    }
    if (!p.eventId || !Array.isArray(p.results) || !p.results.length) {
      return out_({ok:false, error:'Hiányzó versenyazonosító vagy eredmények.'});
    }
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = getSheet_(ss, RESULTS_SHEET, ['Versenyazonosító','Dátum','Futam','Versenyző','Pont','Össz idő','Végeredmény','Mód','Idő használata','Rögzítés ideje']);
    const old = sh.getDataRange().getValues();
    const headers = old.length ? old[0] : ['Versenyazonosító','Dátum','Futam','Versenyző','Pont','Össz idő','Végeredmény','Mód','Idő használata','Rögzítés ideje'];
    const kept = old.slice(1).filter(r => String(r[0]) !== String(p.eventId));
    const stamp = new Date();
    const added = p.results.map(r => [String(p.eventId),String(p.date || ''),Number(r.heat)||1,String(r.name||'').trim(),Number(r.points)||0,Number(r.totalTime)||0,Number(r.result)||0,String(p.mode||''),p.usedTime ? 'Igen' : 'Nem',stamp]);
    sh.clearContents();
    sh.getRange(1,1,1,headers.length).setValues([headers]);
    const rows = kept.concat(added);
    if (rows.length) sh.getRange(2,1,rows.length,headers.length).setValues(rows);
    styleHeader_(sh, headers.length);
    sh.setFrozenRows(1);
    if (rows.length) sh.getRange(2,5,rows.length,3).setNumberFormat('0.00');
    updateRanking_(ss, rows);
    return out_({ok:true,eventId:p.eventId,results:added.length});
  } catch (err) {
    return out_({ok:false,error:String(err)});
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

function updateRanking_(ss, rows) {
  const sh = getSheet_(ss, RANKING_SHEET, ['Helyezés','Versenyző','Legjobb végeredmény','Pont','Össz idő','Dátum','Futam']);
  const best = new Map();
  rows.forEach(r => {
    const name = String(r[3] || '').trim();
    if (!name) return;
    const key = name.toLocaleLowerCase();
    const item = {name, result:Number(r[6])||0, points:Number(r[4])||0, time:Number(r[5])||0, date:String(r[1]||''), heat:Number(r[2])||1};
    const prev = best.get(key);
    // A program pontozása szerint a nagyobb végeredmény a jobb; döntetlennél a nagyobb összidő kerül előre.
    if (!prev || item.result > prev.result || (item.result === prev.result && item.time > prev.time)) best.set(key,item);
  });
  const list = Array.from(best.values()).sort((a,b)=>b.result-a.result || b.time-a.time || a.name.localeCompare(b.name,'hu'));
  const data = [['Helyezés','Versenyző','Legjobb végeredmény','Pont','Össz idő','Dátum','Futam']];
  list.forEach((r,i)=>data.push([i+1,r.name,r.result,r.points,r.time,r.date,r.heat]));
  sh.clearContents();
  sh.getRange(1,1,data.length,data[0].length).setValues(data);
  styleHeader_(sh,data[0].length);
  sh.setFrozenRows(1);
  if (list.length) sh.getRange(2,3,list.length,3).setNumberFormat('0.00');
  sh.autoResizeColumns(1,data[0].length);
}

function getSheet_(ss,name,headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) sh.getRange(1,1,1,headers.length).setValues([headers]);
  return sh;
}
function styleHeader_(sh,cols) {
  sh.getRange(1,1,1,cols).setFontWeight('bold').setBackground('#252540').setFontColor('#ffffff');
  sh.autoResizeColumns(1,cols);
}
function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
