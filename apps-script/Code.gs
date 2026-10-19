/***  記帳 → Google Sheets 接收端（收入 / 支出 兩張表）  ***/
/***  兩種貼法都能用：                                 ***/
/***  A. 試算表 → 擴充功能 → Apps Script（綁定式）      ***/
/***  B. script.google.com → 新增專案（獨立式，靠下面的 SHEET_ID）***/

var VERSION = '1.4.0';   // 1.4.0 = 支出支援多日（結束日期）；1.3.0 = 編輯與刪除

var SHEET_ID = '10Yw9Zz7038trQhIwm9pzmQ-PdZNZI5tBDEhIAg05pIQ';

var INCOME_SHEET  = '收入';
var EXPENSE_SHEET = '支出';
var INCOME_HEADERS  = ['租用時間','服務類型','收入費用','實收費用','活動類型','鐘數','客戶/公司名稱','地點','接單來源','印相張數','備註','行家單','ID'];
var EXPENSE_HEADERS = ['支出日期','Cost','類型','Notes','支出者','ID','結束日期'];

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

  // 舊表升級：補上「結束日期」欄的表頭（資料欄在最後，不動原有欄位）
  if (name === EXPENSE_SHEET) {
    var need = EXPENSE_HEADERS.length;
    var lastCol = sh.getLastColumn();
    if (lastCol < need) {
      var hc = sh.getRange(1, need);
      if (!hc.getValue()) {
        hc.setValue(EXPENSE_HEADERS[need - 1])
          .setFontWeight('bold').setBackground('#F2EEE4').setVerticalAlignment('middle');
      }
      sh.getRange(1, need, sh.getMaxRows(), 1).setNumberFormat('yyyy-mm-dd');
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
      id
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

    // 刪除：按 ID 找到那一列直接刪掉
    if (mode === 'delete') {
      var delRow = id ? findRow_(sh, col, id) : 0;
      if (delRow) sh.deleteRow(delRow);
      return out_({ ok: true, deleted: !!delRow, row: delRow });
    }

    var row = buildRow_(b, isIncome, id);

    if (id) {
      var rowNum = findRow_(sh, col, id);

      // 編輯：覆蓋原本那一列
      if (mode === 'update' && rowNum) {
        sh.getRange(rowNum, 1, 1, row.length).setValues([row]);
        return out_({ ok: true, updated: true, row: rowNum });
      }
      // 新增：ID 已存在就當重複，避免網路重試寫出兩列
      if (mode !== 'update' && rowNum) {
        return out_({ ok: true, dup: true, row: rowNum });
      }
    }

    // 沒帶 mode=update，或要更新的那列已不存在 → 直接新增
    sh.appendRow(row);
    return out_({ ok: true, row: sh.getLastRow(), appended: true });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || '';
  if (action === 'list') return list_();

  var name = '', ok = true, err = '';
  try { name = ss_().getName(); } catch (e2) { ok = false; err = String(e2); }
  return out_({
    ok: ok, service: 'ledger', version: VERSION, sheet: name, error: err,
    time: new Date().toISOString()
  });
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
