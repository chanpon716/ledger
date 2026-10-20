/***  記帳 → Google Sheets 接收端（收入 / 支出 兩張表）  ***/
/***  兩種貼法都能用：                                 ***/
/***  A. 試算表 → 擴充功能 → Apps Script（綁定式）      ***/
/***  B. script.google.com → 新增專案（獨立式，靠下面的 SHEET_ID）***/

var VERSION = '1.6.0';   // 1.6.0 = 收入自動寫入 Google 日曆；1.5.0 = 已收訂金；1.4.0 = 多日；1.3.0 = 編輯與刪除

var SHEET_ID = '10Yw9Zz7038trQhIwm9pzmQ-PdZNZI5tBDEhIAg05pIQ';

var INCOME_SHEET  = '收入';
var EXPENSE_SHEET = '支出';
var INCOME_HEADERS  = ['租用時間','服務類型','收入費用','實收費用','活動類型','鐘數','客戶/公司名稱','地點','接單來源','印相張數','備註','行家單','ID','結束日期','已收訂金','CalendarEventID'];
var EXPENSE_HEADERS = ['支出日期','Cost','類型','Notes','支出者','ID','結束日期'];

/* 日曆：留空 = 用你 Google 帳號的「主要日曆」。
   想寫進另一個日曆，填它的 ID（日曆設定 → 整合日曆 → 日曆 ID）。 */
var CALENDAR_ID = '';
var CAL_SYNC = true;     // 改成 false 就停止寫日曆（Sheet 照常）

/** 綁定式優先；獨立式就用 SHEET_ID 開啟 */
function ss_() {
  var ss = null;
  try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) {}
  if (!ss) ss = SpreadsheetApp.openById(SHEET_ID);
  return ss;
}

function getSheet_(name, headers) {
  var ss = ss_();
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.appendRow(headers);
    sh.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold').setBackground('#F2EEE4').setVerticalAlignment('middle');
    sh.setFrozenRows(1);
    sh.setColumnWidth(1, 105);
    if (name === INCOME_SHEET) {
      sh.setColumnWidth(2, 210);
      sh.setColumnWidth(7, 190);
      sh.getRange('C:D').setNumberFormat('#,##0.00');
      sh.getRange('F:F').setNumberFormat('0');
      sh.getRange('J:J').setNumberFormat('0');
    } else {
      sh.setColumnWidth(2, 110);
      sh.setColumnWidth(3, 120);
      sh.setColumnWidth(4, 240);
      sh.getRange('B:B').setNumberFormat('#,##0.00');
    }
    sh.getRange('A:A').setNumberFormat('yyyy-mm-dd');
  }

  // 舊表升級：把缺少的欄位表頭補在最後（不動原有欄位）
  var heads = name === INCOME_SHEET ? INCOME_HEADERS : EXPENSE_HEADERS;
  var need = heads.length;
  var lastCol = sh.getLastColumn();
  if (lastCol < need) {
    for (var ci = lastCol + 1; ci <= need; ci++) {
      var hName = heads[ci - 1];
      var hc = sh.getRange(1, ci);
      if (!hc.getValue()) {
        hc.setValue(hName)
          .setFontWeight('bold').setBackground('#F2EEE4').setVerticalAlignment('middle');
      }
      if (hName === '結束日期') {
        sh.getRange(1, ci, sh.getMaxRows(), 1).setNumberFormat('yyyy-mm-dd');
      } else if (hName === '已收訂金') {
        sh.getRange(1, ci, sh.getMaxRows(), 1).setNumberFormat('#,##0.00');
      }
    }
  }
  return sh;
}

/** 依 ID 找列號（1-based），找不到回 0 */
function findRow_(sh, col, id) {
  var last = sh.getLastRow();
  if (last < 2) return 0;
  var ids = sh.getRange(2, col, last - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return i + 2;
  }
  return 0;
}

function buildRow_(b, isIncome, id) {
  if (isIncome) {
    return [
      b.date ? new Date(b.date + 'T00:00:00') : new Date(),
      (b.services || []).join(', '),
      Number(b.amount) || 0,
      Number(b.received) || 0,
      b.eventType || '',
      b.hours === '' || b.hours == null ? '' : Number(b.hours),
      b.client || '',
      b.venue || '',
      b.source || '',
      b.prints === '' || b.prints == null ? '' : Number(b.prints),
      b.note || '',
      b.dealer ? 'checked' : '',
      id,
      b.endDate ? new Date(b.endDate + 'T00:00:00') : '',
      (b.deposit === '' || b.deposit == null) ? '' : (Number(b.deposit) || 0)
    ];
  }
  return [
    b.date ? new Date(b.date + 'T00:00:00') : new Date(),
    Number(b.amount) || 0,
    b.expType || '',
    b.note || '',
    b.spender || '',
    id,
    b.endDate ? new Date(b.endDate + 'T00:00:00') : ''
  ];
}

/* ---------- Google 日曆 ---------- */
var CAL_COL = 16;   // 收入表第 16 欄：CalendarEventID

function tz_() {
  try { return ss_().getSpreadsheetTimeZone() || Session.getScriptTimeZone(); }
  catch (e) { return Session.getScriptTimeZone(); }
}

/** 任何日期值 → 'yyyy-MM-dd'（用試算表時區，避免跨時區差一日） */
function ymd_(v) {
  if (!v) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, tz_(), 'yyyy-MM-dd');
  }
  var m = String(v).match(/(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})/);
  if (!m) return '';
  return m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
}

function dayDate_(ymd) {
  var p = String(ymd).split('-');
  return new Date(Number(p[0]), Number(p[1]) - 1, Number(p[2]));
}

function money_(v) {
  var n = Math.round((Number(v) || 0) * 100) / 100;
  var neg = n < 0; n = Math.abs(n);
  var parts = n.toFixed(2).split('.');
  var i = parts[0].replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  var d = parts[1].replace(/0+$/, '');
  return (neg ? '-' : '') + i + (d ? '.' + d : '');
}

function cal_() {
  if (!CAL_SYNC) return null;
  try {
    if (CALENDAR_ID) return CalendarApp.getCalendarById(CALENDAR_ID);
    return CalendarApp.getDefaultCalendar();
  } catch (e) { return null; }
}

function calTitle_(r) {
  var sv = (r.services || []).filter(function (x) { return x; });
  var t = r.client || r.eventType || sv[0] || '收入';
  if (sv.length) t += ' · ' + sv.join('、');
  return t;
}

function calDesc_(r) {
  var L = [];
  L.push('收入 ' + money_(r.amount));
  if (r.received !== '' && r.received != null && Number(r.received) !== Number(r.amount)) {
    L.push('實收 ' + money_(r.received));
  }
  if (r.deposit !== '' && r.deposit != null && Number(r.deposit)) {
    L.push('已收訂金 ' + money_(r.deposit));
  }
  if (r.eventType) L.push('活動：' + r.eventType);
  if (r.hours) L.push(r.hours + ' 小時');
  if (r.source) L.push('來源：' + r.source);
  if (r.prints) L.push('印相 ' + r.prints + ' 張');
  if (r.note) L.push('備註：' + r.note);
  L.push('— 記帳 App 自動建立');
  return L.join('\n');
}

/** 有 eventId 就更新，沒有（或事件已被刪）就新建 */
function calUpsert_(r, eventId) {
  var cal = cal_();
  if (!cal) return { ok: false, err: 'no-calendar' };

  var s = ymd_(r.date);
  if (!s) return { ok: false, err: 'no-date' };
  var e = ymd_(r.endDate) || s;
  if (e < s) e = s;

  var start = dayDate_(s);
  var end = dayDate_(e);
  end = new Date(end.getFullYear(), end.getMonth(), end.getDate() + 1); // 全天事件結束日不含當天

  var title = calTitle_(r), desc = calDesc_(r), loc = r.venue || '';
  var ev = null;
  if (eventId) { try { ev = cal.getEventById(String(eventId)); } catch (err) { ev = null; } }

  try {
    if (ev) {
      ev.setTitle(title);
      ev.setAllDayDates(start, end);
      ev.setLocation(loc);
      ev.setDescription(desc);
    } else {
      ev = cal.createAllDayEvent(title, start, end, { description: desc, location: loc });
    }
    return { ok: true, id: ev.getId() };
  } catch (err) {
    return { ok: false, err: String(err) };
  }
}

function calDelete_(eventId) {
  if (!eventId || !CAL_SYNC) return;
  var cal = cal_();
  if (!cal) return;
  try {
    var ev = cal.getEventById(String(eventId));
    if (ev) ev.deleteEvent();
  } catch (e) {}
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var b = JSON.parse(e.postData.contents);
    var isIncome = (b.type === 'income');
    var sh = isIncome ? getSheet_(INCOME_SHEET, INCOME_HEADERS)
                      : getSheet_(EXPENSE_SHEET, EXPENSE_HEADERS);
    var id = b.id || '';
    var col = isIncome ? 13 : 6;
    var mode = b.mode || '';

    // 刪除：按 ID 找到那一列直接刪掉（收入順便刪掉對應的日曆事件）
    if (mode === 'delete') {
      var delRow = id ? findRow_(sh, col, id) : 0;
      if (delRow) {
        if (isIncome) {
          try { calDelete_(sh.getRange(delRow, CAL_COL).getValue()); } catch (e2) {}
        }
        sh.deleteRow(delRow);
      }
      return out_({ ok: true, deleted: !!delRow, row: delRow });
    }

    var row = buildRow_(b, isIncome, id);

    if (id) {
      var rowNum = findRow_(sh, col, id);

      // 編輯：覆蓋原本那一列
      if (mode === 'update' && rowNum) {
        sh.getRange(rowNum, 1, 1, row.length).setValues([row]);
        var calU = { ok: true };
        if (isIncome) {
          var oldEv = sh.getRange(rowNum, CAL_COL).getValue();
          calU = calUpsert_(b, oldEv);
          if (calU.ok && !oldEv) sh.getRange(rowNum, CAL_COL).setValue(calU.id);
        }
        return out_({ ok: true, updated: true, row: rowNum, cal: calU.ok, calErr: calU.err || '' });
      }
      // 新增：ID 已存在就當重複，避免網路重試寫出兩列
      if (mode !== 'update' && rowNum) {
        return out_({ ok: true, dup: true, row: rowNum });
      }
    }

    // 沒帶 mode=update，或要更新的那列已不存在 → 直接新增
    sh.appendRow(row);
    var newRow = sh.getLastRow();
    var calA = { ok: true };
    if (isIncome) {
      calA = calUpsert_(b, '');
      if (calA.ok) sh.getRange(newRow, CAL_COL).setValue(calA.id);
    }
    return out_({ ok: true, row: newRow, appended: true, cal: calA.ok, calErr: calA.err || '' });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || '';
  if (action === 'list') return list_();
  if (action === 'calsync') return calSyncAll_();

  var name = '', ok = true, err = '';
  try { name = ss_().getName(); } catch (e2) { ok = false; err = String(e2); }
  return out_({
    ok: ok, service: 'ledger', version: VERSION, sheet: name, error: err,
    time: new Date().toISOString()
  });
}

/** ?action=calsync → 把收入表裡所有記錄補寫／更新到 Google 日曆 */
function calSyncAll_() {
  try {
    if (!CAL_SYNC) return out_({ ok: false, error: '已停用日曆同步（CAL_SYNC = false）' });
    var cal = cal_();
    if (!cal) return out_({ ok: false, error: '找不到日曆，請檢查 CALENDAR_ID 或授權' });

    var sh = getSheet_(INCOME_SHEET, INCOME_HEADERS);
    var last = sh.getLastRow();
    if (last < 2) return out_({ ok: true, created: 0, updated: 0, failed: 0, total: 0 });

    var vals = sh.getRange(2, 1, last - 1, INCOME_HEADERS.length).getValues();
    var ids = [], created = 0, updated = 0, failed = 0;

    for (var i = 0; i < vals.length; i++) {
      var r = vals[i];
      var evId = String(r[15] || '');
      if (!r[0] && !r[12]) { ids.push([evId]); continue; }

      var rec = {
        date: ymd_(r[0]),
        services: String(r[1] || '').split(',').map(function (s) { return s.trim(); }).filter(Boolean),
        amount: r[2], received: r[3], eventType: r[4], hours: r[5],
        client: r[6], venue: r[7], source: r[8], prints: r[9], note: r[10],
        endDate: ymd_(r[13]), deposit: r[14]
      };

      var res = calUpsert_(rec, evId);
      if (res.ok) {
        if (evId) updated++; else created++;
        ids.push([res.id]);
      } else {
        failed++;
        ids.push([evId]);
      }
    }

    sh.getRange(2, CAL_COL, ids.length, 1).setValues(ids);
    return out_({ ok: true, created: created, updated: updated, failed: failed, total: vals.length });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  }
}

function fmtDate_(v) {
  if (!v) return '';
  if (Object.prototype.toString.call(v) === '[object Date]') {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(v).slice(0, 10);
}
function numOrBlank_(v) {
  return (v === '' || v === null || v === undefined) ? '' : (Number(v) || 0);
}

/** ?action=list → 把 Sheet 裡的資料吐回給 App（載入歷史用） */
function list_() {
  try {
    var ss = ss_();
    var inc = [], exp = [], i, r;

    var si = ss.getSheetByName(INCOME_SHEET);
    if (si && si.getLastRow() > 1) {
      var vi = si.getRange(2, 1, si.getLastRow() - 1, INCOME_HEADERS.length).getValues();
      for (i = 0; i < vi.length; i++) {
        r = vi[i];
        if (!r[0] && !r[2]) continue;
        inc.push({
          id: String(r[12] || ('at-inc-' + (i + 1))),
          type: 'income',
          date: fmtDate_(r[0]),
          amount: Number(r[2]) || 0,
          received: Number(r[3]) || 0,
          services: String(r[1] || '').split(',').map(function (s) { return s.trim() }).filter(Boolean),
          eventType: String(r[4] || ''),
          hours: numOrBlank_(r[5]),
          prints: numOrBlank_(r[9]),
          client: String(r[6] || ''),
          venue: String(r[7] || ''),
          source: String(r[8] || ''),
          dealer: String(r[11]).toLowerCase() === 'checked' || r[11] === true,
          expType: '', spender: '',
          note: String(r[10] || ''),
          endDate: r[13] ? fmtDate_(r[13]) : '',
          deposit: (r[14] === '' || r[14] === null || r[14] === undefined) ? '' : (Number(r[14]) || 0),
          createdAt: '', synced: true
        });
      }
    }

    var se = ss.getSheetByName(EXPENSE_SHEET);
    if (se && se.getLastRow() > 1) {
      var ve = se.getRange(2, 1, se.getLastRow() - 1, EXPENSE_HEADERS.length).getValues();
      for (i = 0; i < ve.length; i++) {
        r = ve[i];
        if (!r[0] && !r[1]) continue;
        exp.push({
          id: String(r[5] || ('at-exp-' + (i + 1))),
          type: 'expense',
          date: fmtDate_(r[0]),
          amount: Number(r[1]) || 0,
          received: '',
          services: [], eventType: '', hours: '', prints: '',
          client: '', venue: '', source: '', dealer: false,
          expType: String(r[2] || ''),
          spender: String(r[4] || ''),
          note: String(r[3] || ''),
          endDate: r[6] ? fmtDate_(r[6]) : '',
          createdAt: '', synced: true
        });
      }
    }

    return out_({ ok: true, income: inc, expense: exp });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  }
}

/** 手動測試：在編輯器選這個函式按「執行」，看執行記錄 */
function testSetup() {
  var ss = ss_();
  Logger.log('試算表：' + ss.getName());
  Logger.log('收入表：' + (ss.getSheetByName(INCOME_SHEET) ? '已存在' : '尚未建立'));
  Logger.log('支出表：' + (ss.getSheetByName(EXPENSE_SHEET) ? '已存在' : '尚未建立'));
  return ss.getName();
}

function out_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
