// Eredmény beíró – Google Táblázat fogadó
// Másold ezt a teljes kódot a Google Apps Script szerkesztőbe.
// A korábbi fogadó ugyanazon táblázatát és SECRET beállítását használja.
// Add meg a saját Google Táblázat azonosítóját.
const SPREADSHEET_ID = 'IDE_IRD_A_TABLAZAT_AZONOSITOJAT';
const RESULTS_SHEET = 'Eredmények';
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
    const props = PropertiesService.getScriptProperties();
    const secret = props.getProperty('SECRET');
    if (!secret || p.secret !== secret) {
      return out_({ok:false, error:'Hibás vagy be nem állított közös kulcs.'});
    }
    if (!p.eventId || !Array.isArray(p.results) || !p.results.length) {
      return out_({ok:false, error:'Hiányzó versenyazonosító vagy eredmények.'});
    }
    const revisionKey = 'REV_' + p.eventId;
    const revisionMatch = String(p.revisionId || '').match(/^(\d+)-/);
    const revision = revisionMatch ? Number(revisionMatch[1]) : 0;
    const previous = Number(props.getProperty(revisionKey) || 0);
    if (previous && revision < previous) return out_({ok:true,eventId:p.eventId,stale:true,schemaVersion:2});
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sh = getSheet_(ss, RESULTS_SHEET, ['Versenyazonosító','Dátum','Versenymód','Idő használata','Mentés ideje','Versenyző','Futam','Összpont','Összidő','Végeredmény']);
    const old = sh.getDataRange().getValues();
    const headers = old[0].slice();
    // Régi változatokban eltérhet az oszlopok sorrendje.
    const columns = resultColumns_(headers);
    const keptIndexes = old.slice(1).map((r,i) => i+1).filter(i => String(old[i][columns.id]) !== String(p.eventId));
    const kept = keptIndexes.map(i => old[i]);
    const stamp = new Date();
    const added = p.results.map(r => {
      const row = Array(headers.length).fill('');
      const prior=old.slice(1).find(x=>String(x[columns.id])===String(p.eventId));
      row[columns.id]=sheetText_(p.eventId); row[columns.date]=sheetText_(p.date || '');
      row[columns.heat]=Number(r.heat)||1; row[columns.name]=sheetText_(String(r.name||'').trim());
      row[columns.points]=Number(r.points)||0; row[columns.time]=Number(r.totalTime)||0;
      row[columns.result]=Number(r.result)||0; row[columns.mode]=sheetText_(p.mode||'');
      row[columns.usedTime]=!!p.usedTime; row[columns.stamp]=stamp;
      row[columns.location] = Object.prototype.hasOwnProperty.call(p,'location') ? sheetText_(p.location || '') : (prior && prior[columns.location] || '');
      row[columns.type] = Object.prototype.hasOwnProperty.call(p,'eventType') ? (p.eventType === 'gyakorlas' ? 'Gyakorlás' : p.eventType === 'verseny' ? 'Verseny' : '') : (prior && prior[columns.type] || '');
      return row;
    });
    if (headers.length>old[0].length)
      sh.getRange(1,old[0].length+1,1,headers.length-old[0].length).setValues([headers.slice(old[0].length)]);
    const eventRows = old.slice(1).map((row,i)=>({row,index:i+1})).filter(x=>String(x.row[columns.id])===String(p.eventId));
    const managed=Object.keys(columns).map(k=>columns[k]);
    const used=new Set();
    added.forEach(row=>{
      const match=eventRows.find(x=>!used.has(x.index) && String(x.row[columns.name])===String(row[columns.name]) && Number(x.row[columns.heat])===Number(row[columns.heat]));
      const n=match ? match.index+1 : sh.getLastRow()+1;
      if (match) used.add(match.index);
      managed.forEach(i=>sh.getRange(n,i+1,1,1).setValues([[row[i]]]));
    });
    eventRows.filter(x=>!used.has(x.index)).forEach(x=>managed.forEach(i=>sh.getRange(x.index+1,i+1,1,1).clearContent()));
    const rows = kept.concat(added);
    styleHeader_(sh, headers.length);
    sh.setFrozenRows(1);
    if (rows.length) [columns.points,columns.time,columns.result].forEach(i=>sh.getRange(2,i+1,rows.length,1).setNumberFormat('0.00'));
    const ranking = updateRanking_(ss, rows, columns);
    // Régi kliens nem küld rounds mezőt: ne töröljük a korábbi részletes sorokat.
    if (p.results.some(r => Array.isArray(r.rounds))) {
      updateRounds_(ss, p, stamp);
    }
    updateEventSheet_(ss, p);
    // A kért sorrend: ranglista, majd az összesített eredményeket tartalmazó lap.
    ss.setActiveSheet(ranking);
    ss.moveActiveSheet(1);
    ss.setActiveSheet(sh);
    ss.moveActiveSheet(2);
    if (revision > previous) props.setProperty(revisionKey, String(revision));
    return out_({ok:true,eventId:p.eventId,results:added.length,schemaVersion:2});
  } catch (err) {
    return out_({ok:false,error:String(err)});
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

function updateRounds_(ss, p, stamp) {
  const headers = ['Versenyazonosító','Dátum','Helyszín','Esemény típusa','Futam','Versenyző','Szakasz','Kör','Pont','Idő','Eredmény','Érvényes','Rögzítés ideje'];
  const sh = getManagedSheet_(ss, ROUNDS_SHEET, headers);
  const values = sh.getDataRange().getValues();
  const existingHeaders = values[0].slice();
  const width = Math.max(existingHeaders.length, headers.length);
  for (let i=0; i<headers.length; i++) if (!existingHeaders[i]) existingHeaders[i]=headers[i];
  const previousRows = values.slice(1).map((r,i)=>({r,index:i+2})).filter(x=>String(x.r[0])===String(p.eventId));
  const added = [];
  p.results.forEach(r => (Array.isArray(r.rounds) ? r.rounds : []).forEach(round => {
    const row = [sheetText_(p.eventId),sheetText_(p.date||''),sheetText_(p.location||''),p.eventType === 'gyakorlas' ? 'Gyakorlás' : p.eventType === 'verseny' ? 'Verseny' : '',Number(r.heat)||1,sheetText_(String(r.name||'').trim()),round.phase === 'gyak' ? 'Gyakorló' : round.phase === 'eles' ? 'Éles' : sheetText_(round.phase||''),Number(round.round)||1,round.points == null ? '' : Number(round.points),round.time == null ? '' : Number(round.time),round.result == null ? '' : Number(round.result),round.valid ? 'Igen' : 'Nem',stamp];
    added.push(Array.from({length:width}, (_,i) => i < row.length ? row[i] : ''));
  }));
  added.forEach((row,i)=>{
    const n=i<previousRows.length?previousRows[i].index:sh.getLastRow()+1;
    sh.getRange(n,1,1,headers.length).setValues([row.slice(0,headers.length)]);
  });
  previousRows.slice(added.length).forEach(x=>sh.getRange(x.index,1,1,headers.length).clearContent());
  styleHeader_(sh,width);
  sh.setFrozenRows(1);
}

function resultColumns_(headers) {
  const definitions = {id:['Versenyazonosító'], date:['Dátum'], heat:['Futam'], name:['Versenyző'],
    points:['Pont','Összpont','Összes pont'], time:['Össz idő','Összidő','Összes idő'], result:['Végeredmény'],
    mode:['Mód','Versenymód'], usedTime:['Idő használata'], stamp:['Rögzítés ideje','Mentés ideje'],
    location:['Helyszín'], type:['Esemény típusa']};
  const normalize=v=>String(v||'').trim().toLocaleLowerCase().replace(/\s+/g,'');
  const out={};
  Object.keys(definitions).forEach(key=>{
    let i=headers.findIndex(h=>definitions[key].some(alias=>normalize(alias)===normalize(h)));
    if (i<0) {i=headers.length;headers.push(definitions[key][0]);} out[key]=i;
  });
  return out;
}
function sheetText_(value) {
  const s=String(value == null ? '' : value);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}
function updateRanking_(ss, rows, columns) {
  const heading=['Helyezés','Versenyző','Legjobb végeredmény','Pont','Össz idő','Dátum','Futam'];
  const sh = getManagedSheet_(ss, RANKING_SHEET, heading);
  const best = new Map();
  rows.forEach(r => {
    // A régi sorokban nem szerepelt típus; ezeket versenyként kezeljük.
    if (String(r[columns.type] || '').toLocaleLowerCase() === 'gyakorlás') return;
    const name = String(r[columns.name] || '').trim();
    if (!name) return;
    const key = name.toLocaleLowerCase();
    const item = {name, result:Number(r[columns.result])||0, points:Number(r[columns.points])||0, time:Number(r[columns.time])||0, date:String(r[columns.date]||''), heat:Number(r[columns.heat])||1};
    const prev = best.get(key);
    // A program pontozása szerint a nagyobb végeredmény a jobb; döntetlennél a nagyobb összidő kerül előre.
    if (!prev || item.result > prev.result || (item.result === prev.result && item.time > prev.time)) best.set(key,item);
  });
  const list = Array.from(best.values()).sort((a,b)=>b.result-a.result || b.time-a.time || a.name.localeCompare(b.name,'hu'));
  const data = [['Helyezés','Versenyző','Legjobb végeredmény','Pont','Össz idő','Dátum','Futam']];
  list.forEach((r,i)=>data.push([i+1,sheetText_(r.name),r.result,r.points,r.time,sheetText_(r.date),r.heat]));
  const oldLast=sh.getLastRow();
  sh.getRange(1,1,data.length,data[0].length).setValues(data);
  if (oldLast>data.length) sh.getRange(data.length+1,1,oldLast-data.length,data[0].length).clearContent();
  styleHeader_(sh,data[0].length);
  sh.setFrozenRows(1);
  if (list.length) sh.getRange(2,3,list.length,3).setNumberFormat('0.00');
  sh.autoResizeColumns(1,data[0].length);
  return sh;
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
    if (marker[0]==='Eredmény beíró esemény' && String(marker[1])===id) {
      // A régi kliens nem küld köröket; ne írja felül a meglévő részletes lapot.
      if (!p.results.some(r=>Array.isArray(r.rounds))) return;
      break;
    }
    sh=null;
  }
  if (!sh) throw new Error('Nem hozható létre eseménylap.');
  const keys=[];
  p.results.forEach(r=>(Array.isArray(r.rounds)?r.rounds:[]).forEach(q=>{
    const phase=String(q.phase||''),round=Number(q.round)||1,key=phase+':'+round;
    if (!keys.some(x=>x.key===key)) keys.push({key,phase,round});
  }));
  keys.sort((a,b)=>(a.phase==='gyak'?0:a.phase==='eles'?1:2)-(b.phase==='gyak'?0:b.phase==='eles'?1:2)||a.phase.localeCompare(b.phase)||a.round-b.round);
  const header=['Futam','Versenyző'].concat(keys.map(x=>sheetText_((x.phase==='gyak'?'G':x.phase==='eles'?'É':x.phase+' ')+x.round)),['Összes pont','Összes idő','Végeredmény']);
  const info=['Eredmény beíró esemény',sheetText_(id),'Dátum',sheetText_(p.date||''),'Helyszín',sheetText_(p.location||''),'Típus',p.eventType==='gyakorlas'?'Gyakorlás':p.eventType==='verseny'?'Verseny':''];
  const width=Math.max(info.length,header.length), pad=row=>Array.from({length:width},(_,i)=>i<row.length?row[i]:'');
  const data=[pad(info),pad(header)];
  p.results.forEach(r => {
    const rounds=Array.isArray(r.rounds) ? r.rounds : [];
    const cells=keys.map(k=>{
      const q=rounds.find(x=>(String(x.phase||'')+':'+(Number(x.round)||1))===k.key);
      if (!q) return '';
      return sheetText_((q.result == null ? '' : String(q.result))+'\n'+[q.points == null ? '' : q.points,q.time == null ? '' : q.time].join(' | ')+(q.valid?'':' (érvénytelen)'));
    });
    data.push(pad([Number(r.heat)||1,sheetText_(r.name||'')].concat(cells,[Number(r.points)||0,Number(r.totalTime)||0,Number(r.result)||0])));
  });
  const previousRange=sh.getDataRange().getValues();
  const oldRows=previousRange.length,oldWidth=previousRange[0].length;
  sh.getRange(1,1,data.length,width).setValues(data);
  if (oldRows>data.length) sh.getRange(data.length+1,1,oldRows-data.length,oldWidth).clearContent();
  if (oldWidth>width) sh.getRange(1,width+1,Math.min(oldRows,data.length),oldWidth-width).clearContent();
  styleHeader_(sh,width);
  sh.getRange(2,1,1,width).setFontWeight('bold').setBackground('#252540').setFontColor('#ffffff');
  if (data.length>2) sh.getRange(3,1,data.length-2,width).setWrap(true);
  sh.setFrozenRows(2);
}

function getSheet_(ss,name,headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) sh.getRange(1,1,1,headers.length).setValues([headers]);
  return sh;
}
function getManagedSheet_(ss,base,headers) {
  for (let n=0;n<1000;n++) {
    const name=base+(n?' '+n:'');
    let sh=ss.getSheetByName(name);
    if (!sh) { sh=ss.insertSheet(name); sh.getRange(1,1,1,headers.length).setValues([headers]); return sh; }
    if (!sh.getLastRow()) { sh.getRange(1,1,1,headers.length).setValues([headers]); return sh; }
    const actual=sh.getRange(1,1,1,headers.length).getValues()[0];
    if (headers.every((v,i)=>actual[i]===v)) return sh;
  }
  throw new Error('Nem hozható létre munkalap: '+base);
}
function styleHeader_(sh,cols) {
  sh.getRange(1,1,1,cols).setFontWeight('bold').setBackground('#252540').setFontColor('#ffffff');
  sh.autoResizeColumns(1,cols);
}
function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

