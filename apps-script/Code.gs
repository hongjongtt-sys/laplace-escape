/**
 * LAPLACE 방탈출 — 기록·현황판용 앱스 스크립트
 * 시트 탭: log, classes (없으면 자동으로 만듦)
 * 배포: 배포 > 새 배포 > 웹 앱 / 실행: 나 / 액세스 권한: 모든 사용자
 */
const LOG_HEAD = ['time', 'classCode', 'team', 'teamName', 'members', 'mode', 'event', 'roomId', 'itemId', 'detail', 'remainSec', 'hintsUsed', 'gauge'];
const CLASS_HEAD = ['time', 'classCode', 'className', 'detail'];

function sheet_(name, head) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.appendRow(head); sh.setFrozenRows(1); }
  else if (sh.getLastColumn() < head.length) sh.getRange(1, 1, 1, head.length).setValues([head]);
  return sh;
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const body = JSON.parse(e.postData.contents);
    const list = body.batch || [body];
    const logRows = [], classRows = [];
    list.forEach(function (d) {
      const t = d.t ? new Date(d.t) : new Date();
      if (d.event === 'class') {
        classRows.push([t, String(d.classCode || ''), String(d.className || ''), String(d.detail || '')]);
        return;
      }
      logRows.push([t, String(d.classCode || ''), Number(d.team) || '', String(d.teamName || '').slice(0, 30),
        (d.members || []).join ? (d.members || []).join(',') : String(d.members || ''), d.mode || '',
        d.event || '', d.roomId || '', d.itemId || '', String(d.detail || ''),
        d.remainSec == null ? '' : d.remainSec, d.hintsUsed == null ? '' : d.hintsUsed, d.gauge == null ? '' : d.gauge]);
    });
    if (logRows.length) { const sh = sheet_('log', LOG_HEAD); sh.getRange(sh.getLastRow() + 1, 1, logRows.length, LOG_HEAD.length).setValues(logRows); }
    if (classRows.length) { const sh = sheet_('classes', CLASS_HEAD); sh.getRange(sh.getLastRow() + 1, 1, classRows.length, CLASS_HEAD.length).setValues(classRows); }
    return ContentService.createTextOutput('ok');
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  const code = String((e.parameter && e.parameter.classCode) || '').toUpperCase();
  const out = { classCode: code, teams: [], wrong: [] };
  if (code) {
    const sh = sheet_('log', LOG_HEAD);
    const n = sh.getLastRow();
    const rows = n > 1 ? sh.getRange(2, 1, n - 1, LOG_HEAD.length).getValues() : [];
    const teams = {}, wrong = {};
    rows.forEach(function (r) {
      if (String(r[1]).toUpperCase() !== code) return;
      const team = Number(r[2]); if (!team) return;
      const ev = r[6], time = new Date(r[0]).getTime();
      const t = teams[team] || (teams[team] = { team: team, teamName: '', mode: '', room: 'P', keyRooms: [], remainSec: null, hints: 0, last: 0, bonus: 0, status: 'playing', rank: '', gauge: null, bosses: [], finishAt: 0 });
      if (ev === 'bonus') { t.bonus += Number(r[9]) || 0; return; }
      if (r[3]) t.teamName = String(r[3]);
      if (r[5]) t.mode = r[5];
      if (time >= t.last) { t.last = time; if (r[10] !== '') t.remainSec = Number(r[10]); if (r[11] !== '') t.hints = Number(r[11]); if (r[12] !== '' && r[12] != null) t.gauge = Number(r[12]); }
      if (ev === 'start') { t.room = 'P'; t.status = 'playing'; t.keyRooms = []; t.bosses = []; t.finishAt = 0; }
      if (ev === 'boss' && t.bosses.indexOf(r[8]) < 0) t.bosses.push(r[8]);
      if (ev === 'move' && r[7]) t.room = r[7];
      if (ev === 'room_exit' && t.keyRooms.indexOf(r[7]) < 0) t.keyRooms.push(r[7]);
      if (ev === 'finish') { t.status = 'finish'; t.rank = String(r[9]); t.finishAt = time; }
      if (ev === 'timeout') t.status = 'timeout';
      if (ev === 'wrong' && r[8]) wrong[r[8]] = (wrong[r[8]] || 0) + 1;
    });
    out.teams = Object.keys(teams).map(function (k) { return teams[k]; }).filter(function (t) { return t.teamName; });
    out.wrong = Object.keys(wrong).map(function (k) { return { itemId: k, count: wrong[k] }; })
      .sort(function (a, b) { return b.count - a.count; }).slice(0, 10);
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}
