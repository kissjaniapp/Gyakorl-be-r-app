// Eredmény beíró – Google Táblázat fogadó
// Másold ezt a teljes kódot a Google Apps Script szerkesztőbe.
// Fontos: a SECRET értéke egyezzen az Android alkalmazásban megadott közös kulccsal.
const SECRET = 'CSERELD_LE_EGY_SAJAT_HOSSZU_VELETLEN_KULCSRA';
const RESULTS_SHEET = 'Versenyeredmények';
const RANKING_SHEET = 'Összesített ranglista';
const ROUNDS_SHEET = 'Körök részletei';

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
    const formulas = sh.getDataRange().getFormulas();
    const headers = old[0].slice();
    // Új oszlopok a meglévők végére: nem változtatjuk meg a korábbi oszlopok helyét.
    ['Helyszín','Esemény típusa'].forEach(h => { if (!headers.includes(h)) headers.push(h); });
    const locationColumn = headers.indexOf('Helyszín');
    const typeColumn = headers.indexOf('Esemény típusa');
    const keptIndexes = old.slice(1).map((r,i) => i+1).filter(i => String(old[i][0]) !== String(p.eventId));
    const kept = keptIndexes.map(i => old[i]);
    // Őrizzük meg a felhasználó esetleges további oszlopait és képleteit.
    const keptForWrite = keptIndexes.map(i => old[i].map((v,j) => formulas[i][j] || v));
    const stamp = new Date();
    const added = p.results.map(r => {
      const row = Array(headers.length).fill('');
      [String(p.eventId),String(p.date || ''),Number(r.heat)||1,String(r.name||'').trim(),Number(r.points)||0,Number(r.totalTime)||0,Number(r.result)||0,String(p.mode||''),p.usedTime ? 'Igen' : 'Nem',stamp].forEach((v,i) => { row[i]=v; });
      row[locationColumn] = String(p.location || '');
      row[typeColumn] = p.eventType === 'gyakorlas' ? 'Gyakorlás' : p.eventType === 'verseny' ? 'Verseny' : '';
      return row;
    });
    sh.clearContents();
    sh.getRange(1,1,1,headers.length).setValues([headers]);
    const rows = kept.concat(added);
    const writeRows = keptForWrite.concat(added).map(r => Array.from({length:headers.length}, (_,i) => i < r.length ? r[i] : ''));
    if (writeRows.length) sh.getRange(2,1,writeRows.length,headers.length).setValues(writeRows);
    styleHeader_(sh, headers.length);
    sh.setFrozenRows(1);
    if (rows.length) sh.getRange(2,5,rows.length,3).setNumberFormat('0.00');
    updateRanking_(ss, rows, typeColumn);
    // Régi kliens nem küld rounds mezőt: ne töröljük a korábbi részletes sorokat.
    if (p.results.some(r => Array.isArray(r.rounds))) {
      updateRounds_(ss, p, stamp);
      updateEventSheet_(ss, p);
    }
    // A kért sorrend: ranglista, majd az összesített eredményeket tartalmazó lap.
    ss.setActiveSheet(ss.getSheetByName(RANKING_SHEET));
    ss.moveActiveSheet(1);
    ss.setActiveSheet(sh);
    ss.moveActiveSheet(2);
    return out_({ok:true,eventId:p.eventId,results:added.length,schemaVersion:2});
  } catch (err) {
    return out_({ok:false,error:String(err)});
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

function updateRounds_(ss, p, stamp) {
  const headers = ['Versenyazonosító','Dátum','Helyszín','Esemény típusa','Futam','Versenyző','Szakasz','Kör','Pont','Idő','Eredmény','Érvényes','Rögzítés ideje'];
  const sh = getSheet_(ss, ROUNDS_SHEET, headers);
  const values = sh.getDataRange().getValues();
  const formulas = sh.getDataRange().getFormulas();
  const existingHeaders = values[0].slice();
  const width = Math.max(existingHeaders.length, headers.length);
  for (let i=0; i<headers.length; i++) if (!existingHeaders[i]) existingHeaders[i]=headers[i];
  const kept = values.slice(1).map((r,i) => ({r,i:i+1})).filter(x => String(x.r[0]) !== String(p.eventId))
    .map(x => Array.from({length:width}, (_,j) => formulas[x.i][j] || (x.r[j] == null ? '' : x.r[j])));
  const added = [];
  p.results.forEach(r => (Array.isArray(r.rounds) ? r.rounds : []).forEach(round => {
    const row = [String(p.eventId),String(p.date||''),String(p.location||''),p.eventType === 'gyakorlas' ? 'Gyakorlás' : p.eventType === 'verseny' ? 'Verseny' : '',Number(r.heat)||1,String(r.name||'').trim(),round.phase === 'gyak' ? 'Gyakorló' : round.phase === 'eles' ? 'Éles' : String(round.phase||''),Number(round.round)||1,round.points == null ? '' : round.points,round.time == null ? '' : round.time,round.result == null ? '' : round.result,round.valid ? 'Igen' : 'Nem',stamp];
    added.push(Array.from({length:width}, (_,i) => i < row.length ? row[i] : ''));
  }));
  sh.clearContents();
  sh.getRange(1,1,1,width).setValues([Array.from({length:width}, (_,i) => existingHeaders[i] || '')]);
  const rows = kept.concat(added);
  if (rows.length) sh.getRange(2,1,rows.length,width).setValues(rows);
  styleHeader_(sh,width);
  sh.setFrozenRows(1);
}

function updateRanking_(ss, rows, typeColumn) {
  const sh = getSheet_(ss, RANKING_SHEET, ['Helyezés','Versenyző','Legjobb végeredmény','Pont','Össz idő','Dátum','Futam']);
  const best = new Map();
  rows.forEach(r => {
    // A régi sorokban nem szerepelt típus; ezeket versenyként kezeljük.
    if (String(r[typeColumn] || '').toLocaleLowerCase() === 'gyakorlás') return;
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

// Az egyedi lap neve eseményazonosítóból származik, a lap elején lévő jelölés
// ellenőrzése megakadályozza, hogy egy hasonló nevű felhasználói lapot felülírjunk.
function updateEventSheet_(ss, p) {
  const id=String(p.eventId);
  let hash=2166136261;
  for (let i=0;i<id.length;i++) hash=Math.imul(hash ^ id.charCodeAt(i),16777619);
  const base='Esemény '+id.replace(/[^A-Za-z0-9_-]/g,'_').slice(0,50)+' '+(hash>>>0).toString(16);
  let sh;
  for (let n=0;n<1000;n++) {
    const name=base+(n ? ' '+n : '');
    sh=ss.getSheetByName(name);
    if (!sh) { sh=ss.insertSheet(name); break; }
    const marker=sh.getRange(1,1,1,2).getValues()[0];
    if (marker[0]==='Eredmény beíró esemény' && String(marker[1])===id) break;
    sh=null;
  }
  if (!sh) throw new Error('Nem hozható létre eseménylap.');
  const header=['Futam','Versenyző','Összes pont','Összes idő','Összes eredmény','Szakasz','Kör','Pont','Idő','Eredmény','Érvényes'];
  const data=[['Eredmény beíró esemény',id,'Dátum',String(p.date||''),'Helyszín',String(p.location||''),'Típus',p.eventType==='gyakorlas'?'Gyakorlás':p.eventType==='verseny'?'Verseny':'','','',''],header];
  p.results.forEach(r => {
    const rounds=Array.isArray(r.rounds) ? r.rounds : [];
    if (!rounds.length) data.push([Number(r.heat)||1,String(r.name||''),r.points,r.totalTime,r.result,'','','','','','']);
    rounds.forEach(q => data.push([Number(r.heat)||1,String(r.name||''),r.points,r.totalTime,r.result,q.phase==='gyak'?'Gyakorló':q.phase==='eles'?'Éles':String(q.phase||''),q.round,q.points == null?'':q.points,q.time == null?'':q.time,q.result == null?'':q.result,q.valid?'Igen':'Nem']));
  });
  sh.clearContents();
  sh.getRange(1,1,data.length,header.length).setValues(data);
  styleHeader_(sh,header.length);
  sh.setFrozenRows(2);
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

