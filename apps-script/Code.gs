const TO = 'temp@signaturestaff.com.au';
const TOKEN = 'ssreg-2026';

function setup() {
  const p = PropertiesService.getScriptProperties();
  if (!p.getProperty('SHEET_ID')) {
    const ss = SpreadsheetApp.create('Attendance Register Log');
    ss.getSheets()[0].appendRow(['Received','Submission ID','Client','Site','Role','Date','Time','Shift','Supervisor','Worker','Status','Comment','General comments']);
    p.setProperty('SHEET_ID', ss.getId());
  }
  MailApp.getRemainingDailyQuota();
  return p.getProperty('SHEET_ID');
}

function doGet() { return out({ok:true, service:'attendance-register'}); }

function doPost(e) {
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    let d = {};
    try { d = JSON.parse(e.postData.contents); } catch (x) { d = e.parameter || {}; }
    if (d.token !== TOKEN) return out({ok:false, error:'bad token'});
    const id = String(d.submissionId || Utilities.getUuid());
    const cache = CacheService.getScriptCache();
    if (cache.get(id)) return out({ok:true, id:id, duplicate:true});
    let workers = d.workers || [];
    if (typeof workers === 'string') { try { workers = JSON.parse(workers); } catch (x) { workers = []; } }
    const c = {here:0, late:0, absent:0};
    workers.forEach(w => { const s = String(w.status||'').toLowerCase(); if (c[s] !== undefined) c[s]++; });
    const subj = 'Attendance register – ' + (d.client||'') + ' – ' + (d.role||'') + ' – ' + (d.date||'') + ' ' + (d.time||'') +
      ' – Here ' + c.here + ' Late ' + c.late + ' Absent ' + c.absent;
    const esc = s => String(s==null?'':s).replace(/[&<>]/g, m => ({'&':'&amp;','<':'&lt;','>':'&gt;'}[m]));
    const rows = workers.map((w,i) => '<tr><td>'+(i+1)+'</td><td>'+esc(w.name)+'</td><td><b>'+esc(w.status||'Not marked')+'</b></td><td>'+esc(w.comment)+'</td></tr>').join('');
    const html = '<p><b>'+esc(d.client)+'</b> · '+esc(d.site)+'<br>'+esc(d.role)+' · '+esc(d.date)+' · '+esc(d.time)+'<br>Shift '+esc(d.shift)+'</p>'+
      '<table border="1" cellpadding="6" style="border-collapse:collapse"><tr><th>#</th><th>Worker</th><th>Status</th><th>Comment</th></tr>'+rows+'</table>'+
      '<p><b>Supervisor:</b> '+esc(d.supervisor)+'<br><b>General comments:</b> '+esc(d.comments)+'</p><p style="color:#888">Submission '+id+'</p>';
    const text = workers.map((w,i) => (i+1)+'. '+w.name+' – '+(w.status||'Not marked')+(w.comment?' ('+w.comment+')':'')).join('\n') +
      '\nSupervisor: '+(d.supervisor||'')+'\nComments: '+(d.comments||'');
    MailApp.sendEmail({to:TO, subject:subj, body:text, htmlBody:html, name:'Attendance Register'});
    const sh = SpreadsheetApp.openById(setup()).getSheets()[0], now = new Date();
    (workers.length ? workers : [{}]).forEach(w => sh.appendRow([now,id,d.client,d.site,d.role,d.date,d.time,d.shift,d.supervisor,w.name||'',w.status||'',w.comment||'',d.comments||'']));
    cache.put(id, '1', 21600);
    return out({ok:true, id:id});
  } catch (err) {
    return out({ok:false, error:String(err)});
  } finally { lock.releaseLock(); }
}

function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
