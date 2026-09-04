// ── Lab52 Inventory — Google Apps Script Backend ──────────────────────────────
// 設定步驟:
// 1. 開啟 Google Sheets → 擴充功能 → Apps Script
// 2. 貼入此檔案為 Code.gs
// 3. 貼入 inventory_app.html 的內容為 index.html
// 4. 執行 setupSheets() 一次（確保 Transits / Billing 分頁存在，Movement 加 location 欄）
// 5. 部署 → 新增部署 → 網頁應用程式（執行身分: 我 / 存取: 知道連結的所有人）
//
// Sheet 結構：
//   庫存總覽  → 產品主檔（只讀），公式自動計算庫存
//   Movement  → 出入庫紀錄（欄位: Date,Name,SKU,EAN,ASIN,exp_date,Boxes,location,note）
//   Transits  → 在途追蹤（由 setupSheets 建立）
//   Billing   → 帳單紀錄（由 setupSheets 建立）

// ── Entry Point ───────────────────────────────────────────────────────────────

function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || '';
  var cb     = (e && e.parameter && e.parameter.callback) || '';

  function jsonOut(data) {
    var str = JSON.stringify(data);
    return ContentService
      .createTextOutput(cb ? cb + '(' + str + ')' : str)
      .setMimeType(cb ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON);
  }

  if (action === 'getData') {
    return jsonOut(getAllData());
  }

  // ── 新端點: 觸發 SP-API FBA 同步並回傳結果（給獨立 HTML 面板使用）
  if (action === 'syncFba') {
    try {
      var res = syncFbaInventory();              // 打 SP-API，更新 FBA庫存 sheet
      var fba = readFbaInventory_();             // 從 sheet 讀回
      // syncFbaInventory 自己 catch 錯誤並回傳 {error}，不看回傳值會把失敗當成功
      if (res && res.error) {
        return jsonOut({ ok: false, error: res.error, fba: fba });  // fba 是舊快取
      }
      // 0 筆代表 SP-API 回傳的 ASIN 沒有一個對得上 庫存總覽，
      // 這時 writeFbaSheet_ 會保留舊資料不覆寫 —— 不能報成功，否則會誤以為已更新
      if (!res || !res.count) {
        return jsonOut({
          ok: false,
          error: '同步完成但 0 筆符合：SP-API 回傳的 ASIN 都不在「庫存總覽」中，表格仍為舊資料，請檢查 ASIN 設定',
          count: 0,
          fba: fba
        });
      }
      return jsonOut({ ok: true, count: res.count, fba: fba });
    } catch(err) {
      return jsonOut({ ok: false, error: err.message });
    }
  }

  // ── 新端點: 只回傳當前 FBA 快取（不打 SP-API）
  if (action === 'getFba') {
    return jsonOut({ ok: true, fba: readFbaInventory_() });
  }

  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Lab52 庫存系統')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── Helpers ───────────────────────────────────────────────────────────────────

function ss_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (ss) return ss;
  return SpreadsheetApp.openById('1McIxZVNBnBBmrLnJ0T5d30PLODcgcKiZRASeTobMSww');
}

function insertAtTop_(sheet, rowData) {
  sheet.insertRowAfter(1);
  sheet.getRange(2, 1, 1, rowData.length).setValues([rowData]);
}

function nowStr_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
}

function formatDate_(val) {
  if (!val) return '';
  if (val instanceof Date) return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(val).replace(/\//g, '-').trim();
}

function nextId_(sheetName) {
  var sheet = ss_().getSheetByName(sheetName);
  if (!sheet) return 1;
  var vals = sheet.getDataRange().getValues();
  var max = 0;
  for (var i = 1; i < vals.length; i++) {
    var n = parseInt(vals[i][0]);
    if (!isNaN(n) && n > max) max = n;
  }
  return max + 1;
}

function findRow_(sheet, id) {
  var vals = sheet.getDataRange().getValues();
  for (var i = 1; i < vals.length; i++) {
    if (String(vals[i][0]) === String(id)) return i + 1;
  }
  return -1;
}

// ── Main Data ─────────────────────────────────────────────────────────────────

function diagConnection() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    Logger.log('getActiveSpreadsheet: ' + (ss ? ss.getName() : 'NULL'));
    if (!ss) {
      ss = SpreadsheetApp.openById('1McIxZVNBnBBmrLnJ0T5d30PLODcgcKiZRASeTobMSww');
      Logger.log('openById: ' + (ss ? ss.getName() : 'FAILED'));
    }
    var sheets = ss.getSheets().map(function(s){ return s.getName(); });
    Logger.log('Sheets: ' + sheets.join(', '));
    var data = getAllData();
    Logger.log('products: ' + data.products.length + ' movements: ' + data.movements.length);
    return { ok: true, sheets: sheets, products: data.products.length };
  } catch(e) {
    Logger.log('diagConnection ERROR: ' + e.message);
    return { error: e.message };
  }
}

function getAllData() {
  function safe(fn, name) {
    try { return fn(); }
    catch(e) { Logger.log('[getAllData] ' + name + ' ERROR: ' + e.message); return []; }
  }
  var result = {
    products:  safe(readProducts_,        'readProducts'),
    movements: safe(readMovements_,       'readMovements'),
    taiwan:    safe(readTaiwanMovements_, 'readTaiwanMovements'),
    transits:  safe(function(){ return readSheet_('Transits'); }, 'readSheet_Transits'),
    billing:   safe(function(){ return readSheet_('Billing'); },  'readSheet_Billing'),
    fba:       safe(readFbaInventory_,    'readFbaInventory'),
    shared:    safe(readShared_,          'readShared')
  };
  return result;
}

// ── 跨裝置共用的小資料 ────────────────────────────────────────────────────────
// 手動在途與儀表板備註原本只存在瀏覽器 localStorage，換一台裝置就看不到，
// 但表格裡的備註（saveNoteInline）卻是寫進分頁的 —— 外觀一樣、行為不一樣，很容易誤會。
// 統一改用這張 key/value 分頁。不用指令碼屬性：單筆有 9KB 上限（見 TW_Backup 的教訓）。
var SHARED_SHEET_ = 'Shared';

function sharedSheet_() {
  var ss = ss_();
  var sh = ss.getSheetByName(SHARED_SHEET_);
  if (!sh) {
    sh = ss.insertSheet(SHARED_SHEET_);
    sh.appendRow(['key', 'value', 'updated_at']);
    sh.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#2D5016').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    sh.hideSheet();
  }
  return sh;
}

function readShared_() {
  var sh = sharedSheet_();
  if (sh.getLastRow() <= 1) return {};
  var out = {};
  sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(function(r) {
    var k = String(r[0] || '').trim();
    if (k) out[k] = r[1];
  });
  return out;
}

function setShared(d) {
  try {
    if (!d || !d.key) return { error: 'key 不可為空' };
    var key  = String(d.key).trim();
    var val  = d.value == null ? '' : d.value;
    var sh   = sharedSheet_();
    var last = sh.getLastRow();
    var keys = last > 1 ? sh.getRange(2, 1, last - 1, 1).getValues() : [];
    var row  = -1;
    for (var i = 0; i < keys.length; i++) {
      if (String(keys[i][0] || '').trim() === key) { row = i + 2; break; }
    }
    var now   = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
    var blank = (val === '' || val === null);
    if (row > 0) {
      if (blank) sh.deleteRow(row);   // 清空就整列刪掉，免得分頁長滿空值
      else sh.getRange(row, 2, 1, 2).setValues([[val, now]]);
    } else if (!blank) {
      sh.appendRow([key, val, now]);
    }
    return { ok: true };
  } catch(e) { return { error: e.message }; }
}

// 舊 localStorage 資料一次性搬遷用。
// 只補雲端還沒有的 key，絕不覆蓋 —— 否則會蓋掉別台裝置已經填好的值。
function setSharedBulk(pairs) {
  try {
    if (!pairs || !pairs.length) return { ok: true, added: 0 };
    var existing = readShared_();
    var sh   = sharedSheet_();
    var now  = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
    var rows = [];
    pairs.forEach(function(p) {
      var k = String((p && p.key) || '').trim();
      if (!k || existing[k] !== undefined) return;
      if (p.value === '' || p.value == null) return;
      rows.push([k, p.value, now]);
    });
    if (rows.length) sh.getRange(sh.getLastRow() + 1, 1, rows.length, 3).setValues(rows);
    return { ok: true, added: rows.length };
  } catch(e) { return { error: e.message }; }
}

function readProducts_() {
  var sheet = ss_().getSheetByName('庫存總覽');
  if (!sheet) return [];
  var vals = sheet.getDataRange().getValues();
  if (vals.length <= 1) return [];
  var hdr = vals[0].map(function(h) { return String(h).trim(); });

  // Use exact header index; fallback only when header is absent
  function colIdx(names, fallback) {
    for (var i = 0; i < names.length; i++) {
      var idx = hdr.indexOf(names[i]);
      if (idx >= 0) return idx;
    }
    return fallback;
  }
  var iName = colIdx(['産品名稱','產品名稱','名稱'], 0);
  var iSku  = colIdx(['SKU','sku'],                  1);
  var iEan  = colIdx(['EAN','ean'],                  2);
  var iAsin = colIdx(['ASIN','asin'],                3);
  var iQpc  = colIdx(['每箱pcs','每箱PCS'],          4);
  var iBox  = colIdx(['在倉總箱數'],                -1); // formula-computed total

  // Fuzzy match: any header containing '組件' or 'component' (case-insensitive)
  var iComp = -1;
  for (var hi = 0; hi < hdr.length; hi++) {
    var hl = hdr[hi].toLowerCase();
    if (hl.indexOf('組件') >= 0 || hl.indexOf('组件') >= 0 || hl.indexOf('component') >= 0) {
      iComp = hi; break;
    }
  }

  var result = [];
  for (var i = 1; i < vals.length; i++) {
    var r = vals[i];
    var name = String(r[iName] || '').trim();
    if (!name) continue;
    var eanRaw = r[iEan];
    var ean = typeof eanRaw === 'number'
      ? String(Math.round(eanRaw))
      : String(eanRaw || '').trim();
    result.push({
      name:           name,
      sku:            String(r[iSku]  || '').trim(),
      ean:            ean,
      asin:           String(r[iAsin] || '').trim(),
      qty_per_carton: parseInt(r[iQpc]) || 0,
      boxes:          iBox >= 0 ? (parseFloat(r[iBox]) || 0) : 0,
      components:     iComp >= 0 ? String(r[iComp] || '').trim() : ''
    });
  }
  return result;
}

function debugProducts() {
  var sheet = ss_().getSheetByName('庫存總覽');
  var vals  = sheet.getDataRange().getValues();
  var hdr   = vals[0].map(function(h){ return String(h).trim(); });
  var iComp = -1;
  for (var hi = 0; hi < hdr.length; hi++) {
    var hl = hdr[hi].toLowerCase();
    if (hl.indexOf('組件') >= 0 || hl.indexOf('component') >= 0) { iComp = hi; break; }
  }
  return { headers: hdr, iComp: iComp,
    products: vals.slice(1).map(function(r){
      return { name: String(r[0]||''), ean: String(r[2]||''), comp: iComp>=0 ? String(r[iComp]||'') : 'COL_NOT_FOUND' };
    }).filter(function(p){ return p.name; })
  };
}

function readMovements_() {
  var sheet = ss_().getSheetByName('Movement');
  if (!sheet) return [];
  var vals = sheet.getDataRange().getValues();
  if (vals.length <= 1) return [];
  var hdr = vals[0].map(function(h) { return String(h).trim(); });
  function mCol(names, fallback) {
    for (var i = 0; i < names.length; i++) {
      var idx = hdr.indexOf(names[i]); if (idx >= 0) return idx;
    }
    return fallback;
  }
  var iDate = mCol(['Date','date'],         0);
  var iName = mCol(['Name','name'],         1);
  var iSku  = mCol(['SKU','sku'],           2);
  var iEan  = mCol(['EAN','ean'],           3);
  var iAsin = mCol(['ASIN','asin'],         4);
  var iExp  = mCol(['exp_date','Exp_date'], 5);
  var iBox  = mCol(['Boxes','boxes'],       6);
  var iLoc  = mCol(['location','Location'], -1);
  var iNote = mCol(['note','Note'], iLoc >= 0 ? iLoc + 1 : 8);
  var result = [];
  for (var i = 1; i < vals.length; i++) {
    var r = vals[i];
    // EAN may be stored as number in Sheets — convert safely
    var eanRaw = r[iEan];
    var ean = typeof eanRaw === 'number'
      ? String(Math.round(eanRaw))
      : String(eanRaw || '').trim();
    if (!ean || ean === '0') continue;
    result.push({
      date:     formatDate_(r[iDate]),
      name:     String(r[iName] || '').trim(),
      sku:      String(r[iSku]  || '').trim(),
      ean:      ean,
      asin:     String(r[iAsin] || '').trim(),
      exp_date: formatDate_(r[iExp]),
      boxes:    parseFloat(r[iBox]) || 0,
      location: iLoc >= 0 ? (String(r[iLoc] || '').trim() || 'AMZLGS') : 'AMZLGS',
      note:     String(r[iNote] || '').trim(),
      row_idx:  i + 1
    });
  }
  return result;
}

function readSheet_(name) {
  var sheet = ss_().getSheetByName(name);
  if (!sheet) return [];
  var vals = sheet.getDataRange().getValues();
  if (vals.length <= 1) return [];
  var hdr = vals[0].map(function(h) { return String(h).trim(); });
  var tz = Session.getScriptTimeZone();
  var result = [];
  for (var i = 1; i < vals.length; i++) {
    var r = vals[i];
    if (r.every(function(c) { return c === '' || c === null; })) continue;
    var obj = {};
    hdr.forEach(function(h, j) {
      var v = r[j];
      if (v instanceof Date) obj[h] = Utilities.formatDate(v, tz, 'yyyy-MM-dd');
      else obj[h] = (v === '' || v === null || v === undefined) ? null : v;
    });
    result.push(obj);
  }
  return result;
}

// ── Movement Entry ─────────────────────────────────────────────────────────────
// Movement 欄位: Date, Name, SKU, EAN, ASIN, exp_date, Boxes, location, note

function addMovementEntry(d) {
  try {
    var sheet = ss_().getSheetByName('Movement');
    if (!sheet) return { error: 'Movement sheet not found' };
    var lastCol = sheet.getLastColumn();
    var hdr = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) { return String(h).trim(); });
    var row = hdr.map(function(h) {
      if (h === 'Date')     return d.date     || formatDate_(new Date());
      if (h === 'Name')     return d.name     || '';
      if (h === 'SKU')      return d.sku      || '';
      if (h === 'EAN')      return d.ean      || '';
      if (h === 'ASIN')     return d.asin     || '';
      if (h === 'exp_date') return d.exp_date || '';
      if (h === 'Boxes')    return parseFloat(d.boxes) || 0;
      if (h === 'location') return d.location || 'AMZLGS';
      if (h === 'note')     return d.note     || '';
      return '';
    });
    insertAtTop_(sheet, row);
    return { ok: true };
  } catch(e) { return { error: e.message }; }
}

function updateMovementNote(d) {
  try {
    var sheet = ss_().getSheetByName('Movement');
    if (!sheet) return { error: 'Movement sheet not found' };
    var rowIdx = parseInt(d.rowIdx);
    if (!rowIdx || rowIdx < 2) return { error: 'invalid row' };
    var hdr = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
      .map(function(h) { return String(h).trim(); });
    var iNote = hdr.indexOf('note');
    if (iNote < 0) iNote = hdr.indexOf('Note');
    if (iNote < 0) return { error: 'note column not found' };
    sheet.getRange(rowIdx, iNote + 1).setValue(d.note || '');
    return { ok: true };
  } catch(e) { return { error: e.message }; }
}

// ── Taiwan Warehouse (TW_Movement sheet) ──────────────────────────────────────
// 欄位: date, ean, name, sku, boxes, exp_date, note

function readTaiwanMovements_() {
  var sheet = ss_().getSheetByName('TW_Movement');
  if (!sheet) return [];
  var vals = sheet.getDataRange().getValues();
  if (vals.length <= 1) return [];
  var hdr = vals[0].map(function(h) { return String(h).trim(); });
  function col(names, fallback) {
    for (var i = 0; i < names.length; i++) {
      var idx = hdr.indexOf(names[i]); if (idx >= 0) return idx;
    }
    return fallback;
  }
  var iDate = col(['date','Date'],             0);
  var iEan  = col(['ean','EAN'],               1);
  var iName = col(['name','Name'],             2);
  var iSku  = col(['sku','SKU'],               3);
  var iBox  = col(['boxes','Boxes'],           4);
  var iQty  = col(['qty_pcs','Qty_pcs'],      -1);
  var iExp  = col(['exp_date','Exp_date'],     5);
  var iNote = col(['note','Note'],             6);
  var result = [];
  for (var i = 1; i < vals.length; i++) {
    var r = vals[i];
    var eanRaw = r[iEan];
    var ean = typeof eanRaw === 'number'
      ? String(Math.round(eanRaw))
      : String(eanRaw || '').trim();
    if (!ean || ean === '0') continue;
    result.push({
      date:     formatDate_(r[iDate]),
      ean:      ean,
      name:     String(r[iName]  || '').trim(),
      sku:      String(r[iSku]   || '').trim(),
      boxes:    parseFloat(r[iBox]) || 0,
      qty_pcs:  iQty >= 0 ? (parseInt(r[iQty]) || null) : null,
      exp_date: r[iExp] ? formatDate_(r[iExp]) : '',
      note:     String(r[iNote]  || '').trim(),
      row_idx:  i + 1
    });
  }
  return result;
}

function addTaiwanEntry(d) {
  try {
    var ss = ss_();
    var sheet = ss.getSheetByName('TW_Movement');
    if (!sheet) {
      sheet = ss.insertSheet('TW_Movement');
      sheet.appendRow(['date','ean','name','sku','boxes','qty_pcs','exp_date','note']);
      sheet.getRange(1,1,1,8).setFontWeight('bold').setBackground('#2D5016').setFontColor('#ffffff');
      sheet.setFrozenRows(1);
    }
    var hdr = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
      .map(function(h) { return String(h).trim(); });
    var row = hdr.map(function(h) {
      if (h === 'date')     return d.date     || formatDate_(new Date());
      if (h === 'ean')      return d.ean      || '';
      if (h === 'name')     return d.name     || '';
      if (h === 'sku')      return d.sku      || '';
      if (h === 'boxes')    return parseFloat(d.boxes) || 0;
      if (h === 'qty_pcs')  return d.qty_pcs != null ? parseInt(d.qty_pcs) : '';
      if (h === 'exp_date') return d.exp_date || '';
      if (h === 'note')     return d.note     || '';
      return '';
    });
    insertAtTop_(sheet, row);
    return { ok: true };
  } catch(e) { return { error: e.message }; }
}

function replaceTaiwanInventory(rows) {
  try {
    // 這是「整張覆蓋」，空清單會把台灣倉清光。
    // 前端已經擋過一次，這裡再擋一次：直接從 GAS 編輯器誤呼叫也不會清空。
    if (!rows || !rows.length) {
      return { error: '匯入清單為空，未做任何變更（要清空請直接編輯 TW_Movement 分頁）' };
    }
    var ss = ss_();
    var sheet = ss.getSheetByName('TW_Movement');
    if (!sheet) {
      sheet = ss.insertSheet('TW_Movement');
      sheet.appendRow(['date','ean','name','sku','boxes','qty_pcs','exp_date','note']);
      sheet.getRange(1,1,1,8).setFontWeight('bold').setBackground('#2D5016').setFontColor('#ffffff');
      sheet.setFrozenRows(1);
    }
    // 備份舊資料（覆蓋前先存）
    backupTaiwanSheet_(ss, sheet);
    // 清除舊資料（保留標題列）
    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.deleteRows(2, lastRow - 1);
    }
    // 批量寫入新資料
    if (rows && rows.length > 0) {
      var hdr = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
        .map(function(h) { return String(h).trim(); });
      var data = rows.map(function(d) {
        return hdr.map(function(h) {
          if (h === 'date')     return d.date     || formatDate_(new Date());
          if (h === 'ean')      return d.ean      || '';
          if (h === 'name')     return d.name     || '';
          if (h === 'sku')      return d.sku      || '';
          if (h === 'boxes')    return parseFloat(d.boxes) || 0;
          if (h === 'qty_pcs')  return d.qty_pcs  != null ? parseInt(d.qty_pcs) : '';
          if (h === 'exp_date') return d.exp_date || '';
          if (h === 'note')     return d.note     || '';
          return '';
        });
      });
      sheet.getRange(2, 1, data.length, hdr.length).setValues(data);
    }
    return { ok: true, count: rows ? rows.length : 0 };
  } catch(e) { return { error: e.message }; }
}

var TW_BACKUP_SHEET_ = 'TW_Backup';

// 備份寫進隱藏分頁而不是 ScriptProperties：單一屬性上限只有 9KB，
// 台灣倉大約 40~50 列就會塞爆，之後每次匯入都會失敗。
// 走分頁還有一個好處：日期欄位維持 Date 物件，不會被 JSON 轉成字串。
function backupTaiwanSheet_(ss, sheet) {
  var data = sheet.getDataRange().getValues();
  if (!data.length) return;
  var bk = ss.getSheetByName(TW_BACKUP_SHEET_);
  if (!bk) {
    bk = ss.insertSheet(TW_BACKUP_SHEET_);
    bk.hideSheet();
  }
  bk.clear();
  // 新分頁預設 1000 列 / 26 欄，資料超過就得先長大，否則 setValues 會 out of bounds
  if (bk.getMaxRows()    < data.length)    bk.insertRowsAfter(bk.getMaxRows(), data.length - bk.getMaxRows());
  if (bk.getMaxColumns() < data[0].length) bk.insertColumnsAfter(bk.getMaxColumns(), data[0].length - bk.getMaxColumns());
  bk.getRange(1, 1, data.length, data[0].length).setValues(data);
  PropertiesService.getScriptProperties().setProperty('TW_BACKUP_AT',
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'));
}

function restoreTaiwanBackup() {
  try {
    var ss = ss_();
    var bk = ss.getSheetByName(TW_BACKUP_SHEET_);
    if (!bk || bk.getLastRow() === 0) return { error: '沒有可還原的備份' };
    var backupData = bk.getDataRange().getValues();
    if (backupData.length < 2) return { error: '備份資料為空' };

    var sheet = ss.getSheetByName('TW_Movement');
    if (!sheet) sheet = ss.insertSheet('TW_Movement');
    sheet.clearContents();
    sheet.getRange(1, 1, backupData.length, backupData[0].length).setValues(backupData);
    return {
      ok: true,
      count: backupData.length - 1,
      at: PropertiesService.getScriptProperties().getProperty('TW_BACKUP_AT') || ''
    };
  } catch(e) { return { error: e.message }; }
}

function updateTaiwanNote(d) {
  try {
    var sheet = ss_().getSheetByName('TW_Movement');
    if (!sheet) return { error: 'TW_Movement sheet not found' };
    var rowIdx = parseInt(d.rowIdx);
    if (!rowIdx || rowIdx < 2) return { error: 'invalid row' };
    var hdr = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
      .map(function(h) { return String(h).trim(); });
    var iNote = hdr.indexOf('note');
    if (iNote < 0) iNote = hdr.indexOf('Note');
    if (iNote < 0) return { error: 'note column not found' };
    sheet.getRange(rowIdx, iNote + 1).setValue(d.note || '');
    return { ok: true };
  } catch(e) { return { error: e.message }; }
}

// ── Transits ──────────────────────────────────────────────────────────────────
// 欄位: id, ean, product_name, sku, from_location, to_location,
//       qty_cartons, exp_date, ship_date, eta_date, status,
//       arrived_date, note, created_at

// ── 在途追蹤欄位 ─────────────────────────────────────────────────────────────
// 一律加在 Transits 尾端。注意：絕對不要用 setupSheets() 來加欄位，
// 那支會 deleteSheet 再 insertSheet，整張在途紀錄會被清光。
var TRANSIT_TRACK_COLS_ = ['tracking_no', 'carrier', 'tracking_status', 'tracking_checked_at'];

function ensureTransitColumns_(sheet) {
  var lastCol = sheet.getLastColumn();
  var hdr = sheet.getRange(1, 1, 1, lastCol).getValues()[0]
    .map(function(h) { return String(h).trim(); });
  var missing = TRANSIT_TRACK_COLS_.filter(function(c) { return hdr.indexOf(c) < 0; });
  if (!missing.length) return hdr;
  var need = lastCol + missing.length;
  if (sheet.getMaxColumns() < need) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), need - sheet.getMaxColumns());
  }
  sheet.getRange(1, lastCol + 1, 1, missing.length).setValues([missing])
    .setFontWeight('bold').setBackground('#2D5016').setFontColor('#ffffff');
  return hdr.concat(missing);
}

function addTransit(d) {
  try {
    var tSheet = ss_().getSheetByName('Transits');
    if (!tSheet) return { error: 'Transits sheet not found' };
    var hdr = ensureTransitColumns_(tSheet);
    var id  = nextId_('Transits');
    var qty = Math.abs(parseFloat(d.qty_cartons) || 0);
    var fromLoc = d.from_location || 'TW';
    // 依標題名稱組列，欄位順序日後調整也不會錯位
    var vals = {
      id: id,
      ean:           d.ean          || '',
      product_name:  d.product_name || '',
      sku:           d.sku          || '',
      from_location: fromLoc,
      to_location:   d.to_location  || 'AMZLGS',
      qty_cartons:   qty,
      exp_date:      d.exp_date     || '',
      ship_date:     d.ship_date    || '',
      eta_date:      d.eta_date     || '',
      status:        'TRANSIT',
      arrived_date:  '',
      note:          d.note         || '',
      created_at:    nowStr_(),
      tracking_no:        String(d.tracking_no || '').trim(),
      carrier:            String(d.carrier     || '').trim(),
      tracking_status:    '',
      tracking_checked_at: ''
    };
    insertAtTop_(tSheet, hdr.map(function(h) {
      return vals[h] !== undefined ? vals[h] : '';
    }));
    if (fromLoc === 'TW') {
      addTaiwanEntry({
        date:     d.ship_date,
        ean:      d.ean,
        name:     d.product_name,
        sku:      d.sku,
        boxes:    -qty,
        exp_date: d.exp_date,
        note:     '在途出發 #' + id
      });
    } else {
      addMovementEntry({
        date:     d.ship_date,
        name:     d.product_name,
        sku:      d.sku,
        ean:      d.ean,
        asin:     d.asin || '',
        exp_date: d.exp_date,
        boxes:    -qty,
        location: fromLoc,
        note:     '在途出發 #' + id
      });
    }
    return { ok: true, id: id };
  } catch(e) { return { error: e.message }; }
}

// 標記為「待確認到貨」。階段一由人手動按，階段二改由每日輪詢自動設定。
// 這支不寫任何庫存 —— 入庫一律走 confirmArrival()，因為快遞的「已送達」
// 只代表送到倉庫門口，不等於海外倉已經點收入帳。
function markTransitArriving(transitId) {
  try {
    var sh = ss_().getSheetByName('Transits');
    if (!sh) return { error: 'Transits sheet not found' };
    var row = findRow_(sh, transitId);
    if (row < 0) return { error: 'Transit not found' };
    var hdr = ensureTransitColumns_(sh);
    var iStatus = hdr.indexOf('status');
    if (iStatus < 0) return { error: '找不到 status 欄位' };
    sh.getRange(row, iStatus + 1).setValue('ARRIVING');
    return { ok: true };
  } catch(e) { return { error: e.message }; }
}

// ── 自我測試：在途追蹤（階段一）───────────────────────────────────────────────
// 在 GAS 編輯器選 testTransitTracking 執行，看下方執行紀錄。
// 會建立一筆測試在途、逐項驗證，最後把自己造成的痕跡全部刪除（含台灣倉／海外倉的異動列）。
// 用不存在的英數 EAN，不會碰到任何真實產品的庫存。
var TEST_TRANSIT_EAN_ = 'ZZTEST0000001';

function testTransitTracking() {
  var out = [];
  function say(s) { out.push(s); Logger.log(s); }
  function check(label, cond, detail) {
    say((cond ? '✅ ' : '❌ ') + label + (detail ? '　（' + detail + '）' : ''));
    return cond;
  }

  var tSheet = ss_().getSheetByName('Transits');
  if (!tSheet) { say('❌ 找不到 Transits 分頁'); return out.join('\n'); }

  var id = null;
  try {
    // 1. 追蹤欄位補齊
    var hdr = ensureTransitColumns_(tSheet);
    TRANSIT_TRACK_COLS_.forEach(function(c) {
      check('欄位存在：' + c, hdr.indexOf(c) >= 0);
    });

    // 2. 建立在途（含追蹤號）
    var res = addTransit({
      ean: TEST_TRANSIT_EAN_, product_name: '【測試】請忽略', sku: 'TEST-SKU',
      from_location: 'TW', to_location: 'AMZLGS',
      qty_cartons: 3, ship_date: '2026-09-04', eta_date: '2026-09-20',
      tracking_no: 'TEST1234567890', carrier: 'dhl', note: '自動測試'
    });
    check('addTransit 成功', !!(res && res.ok), JSON.stringify(res));
    id = res && res.id;
    if (!id) throw new Error('拿不到 transit id');

    function transitRow() {
      return readSheet_('Transits').filter(function(r) { return r.id == id; })[0] || {};
    }
    function twSum() {
      return readTaiwanMovements_()
        .filter(function(m) { return m.ean === TEST_TRANSIT_EAN_; })
        .reduce(function(n, m) { return n + (parseFloat(m.boxes) || 0); }, 0);
    }
    function ovSum() {
      return readMovements_()
        .filter(function(m) { return m.ean === TEST_TRANSIT_EAN_; })
        .reduce(function(n, m) { return n + (parseFloat(m.boxes) || 0); }, 0);
    }

    var r1 = transitRow();
    check('追蹤號寫入正確', r1.tracking_no === 'TEST1234567890', '實際=' + r1.tracking_no);
    check('業者寫入正確',   r1.carrier === 'dhl',                '實際=' + r1.carrier);
    check('初始狀態 TRANSIT', r1.status === 'TRANSIT',           '實際=' + r1.status);
    check('箱數正確',        Number(r1.qty_cartons) === 3,       '實際=' + r1.qty_cartons);

    // 3. 出貨時台灣倉應自動扣掉
    check('台灣倉自動扣除 3 箱', twSum() === -3, '實際=' + twSum());
    check('此時海外倉尚未入帳',  ovSum() === 0,  '實際=' + ovSum());

    // 4. 標記待確認到貨 —— 這一步絕對不能碰庫存
    var mr = markTransitArriving(id);
    check('markTransitArriving 成功', !!(mr && mr.ok), JSON.stringify(mr));
    check('狀態轉為 ARRIVING', transitRow().status === 'ARRIVING', '實際=' + transitRow().status);
    check('ARRIVING 沒有動到台灣倉', twSum() === -3, '實際=' + twSum());
    check('ARRIVING 沒有寫進海外倉', ovSum() === 0,  '實際=' + ovSum());

    // 5. 確認到貨 → 海外倉入帳
    var cr = confirmArrival(id, { arrived_date: '2026-09-21' });
    check('confirmArrival 成功', !!(cr && cr.ok), JSON.stringify(cr));
    check('海外倉自動加入 3 箱', ovSum() === 3, '實際=' + ovSum());
    check('狀態轉為 ARRIVED', transitRow().status === 'ARRIVED', '實際=' + transitRow().status);

  } catch(e) {
    say('❌ 測試中斷：' + e.message);
  } finally {
    say('🧹 ' + cleanupTestTransit_());
  }
  return out.join('\n');
}

// 把測試 EAN 在三張分頁留下的列全部刪掉，不留痕跡
function cleanupTestTransit_() {
  var ss = ss_(), removed = 0;
  [['Transits', ['ean']], ['TW_Movement', ['ean']], ['Movement', ['EAN', 'ean']]]
    .forEach(function(pair) {
      var sh = ss.getSheetByName(pair[0]);
      if (!sh || sh.getLastRow() < 2) return;
      var vals = sh.getDataRange().getValues();
      var hdr  = vals[0].map(function(h) { return String(h).trim(); });
      var iE = -1;
      pair[1].forEach(function(name) { if (iE < 0) iE = hdr.indexOf(name); });
      if (iE < 0) return;
      for (var r = vals.length - 1; r >= 1; r--) {
        if (String(vals[r][iE] || '').trim() === TEST_TRANSIT_EAN_) {
          sh.deleteRow(r + 1); removed++;
        }
      }
    });
  return '清理完成，刪除 ' + removed + ' 列測試資料';
}

function confirmArrival(transitId, d) {
  try {
    var tSheet = ss_().getSheetByName('Transits');
    var tRow = findRow_(tSheet, transitId);
    if (tRow < 0) return { error: 'Transit not found' };
    var tVals = tSheet.getRange(tRow, 1, 1, 14).getValues()[0];
    // cols: id(0) ean(1) product_name(2) sku(3) from_location(4) to_location(5)
    //       qty_cartons(6) exp_date(7) ship_date(8) eta_date(9) status(10)
    //       arrived_date(11) note(12) created_at(13)
    var qty     = parseFloat(tVals[6]) || 0;
    var toLoc   = String(tVals[5] || 'AMZLGS');
    var expDate = d.exp_date || String(tVals[7] || '');
    if (toLoc === 'TW') {
      addTaiwanEntry({
        date:     d.arrived_date,
        ean:      String(tVals[1] || ''),
        name:     String(tVals[2] || ''),
        sku:      String(tVals[3] || ''),
        boxes:    qty,
        exp_date: expDate,
        note:     '在途到貨 #' + transitId
      });
    } else {
      addMovementEntry({
        date:     d.arrived_date,
        name:     String(tVals[2] || ''),
        sku:      String(tVals[3] || ''),
        ean:      String(tVals[1] || ''),
        exp_date: expDate,
        boxes:    qty,
        location: toLoc,
        note:     '在途到貨 #' + transitId
      });
    }
    tSheet.getRange(tRow, 11).setValue('ARRIVED');
    tSheet.getRange(tRow, 12).setValue(d.arrived_date || '');
    return { ok: true };
  } catch(e) { return { error: e.message }; }
}

function deleteTransit(id) {
  try {
    var tSheet = ss_().getSheetByName('Transits');
    var row = findRow_(tSheet, id);
    if (row < 0) return { error: 'Transit not found' };
    tSheet.deleteRow(row);
    return { ok: true };
  } catch(e) { return { error: e.message }; }
}

// ── Product Management ─────────────────────────────────────────────────────────
// 新增產品至 庫存總覽（欄 A-E: 產品名稱, SKU, EAN, ASIN, 每箱pcs）

function addProduct(d) {
  try {
    var sheet = ss_().getSheetByName('庫存總覽');
    if (!sheet) return { error: '庫存總覽 sheet not found' };
    sheet.appendRow([
      d.name || '', d.sku || '', d.ean || '', d.asin || '',
      d.qty_per_carton ? parseInt(d.qty_per_carton) : ''
    ]);
    return { ok: true };
  } catch(e) { return { error: e.message }; }
}

// ── Billing ───────────────────────────────────────────────────────────────────

function addBilling(d) {
  try {
    var id = nextId_('Billing');
    insertAtTop_(ss_().getSheetByName('Billing'), [
      id, d.date, (d.type||'').toUpperCase(), d.job_name||'',
      d.qty       ? parseFloat(d.qty)       : '',
      d.unit_fee  ? parseFloat(d.unit_fee)  : '',
      d.total_fee ? parseFloat(d.total_fee) : '',
      d.note||''
    ]);
    return { ok: true, id: id };
  } catch(e) { return { error: e.message }; }
}

function deleteBilling(id) {
  try {
    var sheet = ss_().getSheetByName('Billing');
    var row = findRow_(sheet, id);
    if (row < 0) return { error: 'Billing record not found' };
    sheet.deleteRow(row);
    return { ok: true };
  } catch(e) { return { error: e.message }; }
}

// ── Fix TW_Movement schema（舊版無 qty_pcs 欄時執行一次）────────────────────────
// 偵測 TW_Movement 是否缺少 qty_pcs 欄；若是，插入欄位並把錯位資料還原
function fixTWMovementSchema() {
  var ss = ss_();
  var sheet = ss.getSheetByName('TW_Movement');
  if (!sheet) { Logger.log('TW_Movement 不存在'); return { error: 'not found' }; }

  var lastCol = sheet.getLastColumn();
  var hdr = sheet.getRange(1, 1, 1, lastCol).getValues()[0]
    .map(function(h) { return String(h).trim(); });

  Logger.log('現有 header: ' + JSON.stringify(hdr));

  if (hdr.indexOf('qty_pcs') >= 0) {
    Logger.log('qty_pcs 已存在，無需修復');
    return { ok: true, msg: 'already ok' };
  }

  // 舊結構: date(0) ean(1) name(2) sku(3) boxes(4) exp_date(5) note(6)
  // 新結構: date(0) ean(1) name(2) sku(3) boxes(4) qty_pcs(5) exp_date(6) note(7)
  // 在 boxes 後面（col 6，1-indexed）插入 qty_pcs 欄
  var boxesIdx = hdr.indexOf('boxes');
  var insertAt = boxesIdx >= 0 ? boxesIdx + 2 : 6; // 1-indexed
  sheet.insertColumnAfter(insertAt - 1);
  sheet.getRange(1, insertAt).setValue('qty_pcs')
    .setFontWeight('bold').setBackground('#2D5016').setFontColor('#ffffff');

  // 所有資料列：qty_pcs 新欄填空字串（資料已在正確欄位，不需移動）
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, insertAt, lastRow - 1, 1).setValue('');
  }

  Logger.log('qty_pcs 欄已插入（欄 ' + insertAt + '）');
  return { ok: true, msg: 'fixed', insertedAt: insertAt };
}

// ── One-time Setup ────────────────────────────────────────────────────────────

function setupSheets() {
  var ss = ss_();

  if (!ss.getSheetByName('Billing')) {
    var b = ss.insertSheet('Billing');
    var bh = ['id','date','type','job_name','qty','unit_fee','total_fee','note'];
    b.appendRow(bh);
    b.getRange(1,1,1,bh.length).setFontWeight('bold').setBackground('#2D5016').setFontColor('#ffffff');
    b.setFrozenRows(1);
  }

  if (!ss.getSheetByName('Transits')) {
    var t = ss.insertSheet('Transits');
    var th = ['id','ean','product_name','sku','from_location','to_location',
              'qty_cartons','exp_date','ship_date','eta_date','status',
              'arrived_date','note','created_at'];
    t.appendRow(th);
    t.getRange(1,1,1,th.length).setFontWeight('bold').setBackground('#2D5016').setFontColor('#ffffff');
    t.setFrozenRows(1);
  }

  Logger.log('setupSheets 完成');
  return { ok: true };
}

// ── Migration ─────────────────────────────────────────────────────────────────
// 從舊版 schema（Products / Batches / Movements）遷移至新版
// 執行一次即可，執行後舊分頁將被刪除

function migrateToNewSchema() {
  var ss = ss_();
  var log = [];

  // ── 1. 從 Products 讀取產品資料 ──────────────────────────────────────────────
  var products = [];
  var prodSheet = ss.getSheetByName('Products');
  if (prodSheet) {
    var pv = prodSheet.getDataRange().getValues();
    var ph = pv[0].map(function(h){ return String(h).trim(); });
    var pi = {
      name: ph.indexOf('name'),
      sku:  ph.indexOf('sku'),
      ean:  ph.indexOf('ean'),
      asin: ph.indexOf('asin'),
      qpc:  ph.indexOf('qty_per_carton')
    };
    for (var i = 1; i < pv.length; i++) {
      var r = pv[i];
      var name = String(r[pi.name] || '').trim();
      if (!name) continue;
      products.push([
        name,
        String(r[pi.sku]  || '').trim(),
        String(r[pi.ean]  || '').trim(),
        String(r[pi.asin] || '').trim(),
        parseInt(r[pi.qpc]) || ''
      ]);
    }
    log.push('Products: 讀取 ' + products.length + ' 筆');
  } else {
    log.push('Products: 找不到，庫存總覽將保持現有內容');
  }

  // ── 2. 重建 庫存總覽（靜態主檔 + 公式欄）────────────────────────────────────
  // 欄位: 產品名稱(A) SKU(B) EAN(C) ASIN(D) 每箱pcs(E)
  //       在倉總箱數(F) 庫存狀態(G) 最近效期(H) 估算總pcs(I)
  // 公式指向 Movement，不再依賴 Batches
  // Movement 欄位（migration 後）:
  //   A=Date B=Name C=SKU D=EAN E=ASIN F=exp_date G=Boxes H=location I=note
  var ov = ss.getSheetByName('庫存總覽');
  if (!ov) ov = ss.insertSheet('庫存總覽');
  ov.clear();
  var maxCols = ov.getMaxColumns();
  if (maxCols > 9) ov.deleteColumns(10, maxCols - 9);
  var ovHdr = ['產品名稱','SKU','EAN','ASIN','每箱pcs','在倉總箱數','庫存狀態','最近效期','估算總pcs'];
  ov.getRange(1,1,1,9).setValues([ovHdr])
    .setFontWeight('bold').setBackground('#2D5016').setFontColor('#ffffff');
  if (products.length > 0) {
    ov.getRange(2,1,products.length,5).setValues(products);
    log.push('庫存總覽: 靜態資料寫入 ' + products.length + ' 筆');
    var fErr = [];
    for (var fi = 0; fi < products.length; fi++) {
      var row = fi + 2;
      try {
        ov.getRange(row,6).setFormula('=IFERROR(SUMIF(Movement!D:D,C'+row+',Movement!G:G),0)');
        ov.getRange(row,7).setFormula('=IF(F'+row+'<=0,"❌ 缺貨",IF(F'+row+'<10,"🟡 注意","✅ 正常"))');
        ov.getRange(row,8).setFormula('=IFERROR(TEXT(MINIFS(Movement!F:F,Movement!D:D,C'+row+',Movement!G:G,">"&0),"yyyy-mm-dd"),"")');
        ov.getRange(row,9).setFormula('=IFERROR(IF(F'+row+'=0,"",F'+row+'*E'+row+'),"")');
      } catch(fe) {
        fErr.push('row'+row+': '+fe.message);
      }
    }
    if (fErr.length > 0) {
      log.push('庫存總覽: 公式錯誤 → ' + fErr.join(' | '));
    } else {
      log.push('庫存總覽: 公式欄寫入完成');
    }
  }
  ov.setFrozenRows(1);
  log.push('庫存總覽: 重建完成');

  // ── 3. Movement：加入 location 欄 ────────────────────────────────────────────
  var mov = ss.getSheetByName('Movement');
  if (mov) {
    var lastCol = mov.getLastColumn();
    var hdr = mov.getRange(1,1,1,lastCol).getValues()[0].map(function(h){ return String(h).trim(); });
    if (hdr.indexOf('location') < 0) {
      var boxIdx = hdr.indexOf('Boxes');
      var insertCol = boxIdx >= 0 ? boxIdx + 2 : lastCol + 1;
      if (boxIdx >= 0) {
        mov.insertColumnAfter(boxIdx + 1);
      } else {
        mov.insertColumnAfter(lastCol);
      }
      mov.getRange(1, insertCol).setValue('location')
        .setFontWeight('bold').setBackground('#2D5016').setFontColor('#ffffff');
      var lastRow = mov.getLastRow();
      if (lastRow > 1) {
        var fill = Array(lastRow - 1).fill(['AMZLGS']);
        mov.getRange(2, insertCol, lastRow - 1, 1).setValues(fill);
      }
      mov.setFrozenRows(1);
      log.push('Movement: location 欄已加入（欄 ' + insertCol + '），' + (lastRow - 1) + ' 筆預設為 AMZLGS');
    } else {
      log.push('Movement: location 欄已存在，跳過');
    }
  } else {
    log.push('Movement: 找不到此分頁！');
  }

  // ── 4. 重建 Transits ──────────────────────────────────────────────────────────
  var tr = ss.getSheetByName('Transits');
  if (tr) ss.deleteSheet(tr);
  var newTr = ss.insertSheet('Transits');
  var trHdr = ['id','ean','product_name','sku','from_location','to_location',
               'qty_cartons','exp_date','ship_date','eta_date','status',
               'arrived_date','note','created_at'];
  newTr.getRange(1,1,1,trHdr.length).setValues([trHdr])
    .setFontWeight('bold').setBackground('#2D5016').setFontColor('#ffffff');
  newTr.setFrozenRows(1);
  log.push('Transits: 重建完成');

  // ── 5. 確保 Billing 存在 ──────────────────────────────────────────────────────
  if (!ss.getSheetByName('Billing')) {
    var bill = ss.insertSheet('Billing');
    var bHdr = ['id','date','type','job_name','qty','unit_fee','total_fee','note'];
    bill.getRange(1,1,1,bHdr.length).setValues([bHdr])
      .setFontWeight('bold').setBackground('#2D5016').setFontColor('#ffffff');
    bill.setFrozenRows(1);
    log.push('Billing: 新建完成');
  } else {
    log.push('Billing: 已存在，跳過');
  }

  // ── 6. 刪除舊分頁 ─────────────────────────────────────────────────────────────
  ['Products','Batches','Movements'].forEach(function(name) {
    var s = ss.getSheetByName(name);
    if (s) { ss.deleteSheet(s); log.push(name + ': 已刪除'); }
  });

  var msg = log.join('\n');
  Logger.log(msg);
  return { ok: true, log: log };
}

// ── 還原 庫存總覽（含產品資料 + 公式）────────────────────────────────────────────
// 庫存總覽遺失時執行此函數一次即可恢復

function restoreOverview() {
  var ss = ss_();
  var ov = ss.getSheetByName('庫存總覽');
  if (!ov) ov = ss.insertSheet('庫存總覽');
  ov.clear();
  var maxCols = ov.getMaxColumns();
  if (maxCols > 9) ov.deleteColumns(10, maxCols - 9);

  // 表頭
  var hdr = ['產品名稱','SKU','EAN','ASIN','每箱pcs','在倉總箱數','庫存狀態','最近效期','估算總pcs'];
  ov.getRange(1,1,1,9).setValues([hdr])
    .setFontWeight('bold').setBackground('#2D5016').setFontColor('#ffffff');

  // 產品靜態資料
  var data = [
    ['境外版兒童含鈣健齒噴霧(葡萄口味) 20mL',   'Grape_Spray_Kids',       '4714781022224','B0B5QZDBNT', 144],
    ['境外版兒童含鈣健齒噴霧(香草口味) 20mL',   'Milk_Kids_Spray',         '4714781022231','B0B5QWQ7PP', 144],
    ['境外版兒童含鈣健齒噴霧(原味口味) 20mL',   'FlavorFree_Spray_Kids',   '4714781022248','B0F6T1FCZD', 144],
    ['境外化妝品版兒童含鈣健齒噴霧(草莓口味) 20mL','XI-MN9Y-L5CY',          '4714781020015','B0FLPD32H9', 144],
    ['境外化妝品版兒童含鈣健齒噴霧(西瓜口味) 20mL','Watermelon-spray',       '4714781020022','B0FMDYP9VG', 144],
    ['境外化妝品版兒童含鈣健齒噴霧(水蜜桃口味) 20mL','Peach-spray',          '4714781020756','B0FMRSFTLF', 144],
    ['境外兩入噴噴(草莓+葡萄)',                  'D6-1MWN-Q0EN',            '4714781023580','B0DBTNJW4G',  36],
    ['境外兩入噴噴(草莓+水蜜桃)',                'TU-V4OB-07QB',            '4714781023597','B0DBT8RDPJ',  36],
    ['境外兩入噴噴(草莓+香草)',                  'M2-6V1S-09J4',            '4714781023603','B0DBTTVWB4',  36],
    ['境外兩入噴噴(草莓+西瓜)',                  'Strawberry+Watermelon',   '4714781023610','B0FMY7W7HZ',  36],
    ['境外兩入噴噴(西瓜+水蜜桃)',                'Watermelon+Peach',        '4714781023757','x',           36],
    ['境外版三入噴噴禮盒(葡萄口味)',              'F8-1THT-CLFU',            '4714781023115','B0C9T4XCJF',  60],
    ['境外化妝品三入噴噴禮盒(草莓口味)',          'StrawberrySprayBox',      '4714781020763','B0FMRSQVY6',  60],
    ['境外版兒童無氟牙膏(草莓口味) 113g',        'StrawberryNHAP',          '4714781020060','B0GQ7FHJWD',  63],
    ['境外版兒童無氟牙膏(西瓜口味) 113g',        'WatermelonNHAP',          '4714781020077','B0GQ8X3V9Q',  63],
    ['境外版成人L8020噴霧(蘋果烏龍口味) 20mL',  'AD-0001',                 '4714781020084','B0GK66X2RN', 144]
  ];
  ov.getRange(2,1,data.length,5).setValues(data);

  // 公式欄 F-I
  for (var i = 0; i < data.length; i++) {
    var r = i + 2;
    ov.getRange(r,6).setFormula('=IFERROR(SUMIF(Movement!D:D,C'+r+',Movement!G:G),0)');
    ov.getRange(r,7).setFormula('=IF(F'+r+'<=0,"❌ 缺貨",IF(F'+r+'<10,"🟡 注意","✅ 正常"))');
    ov.getRange(r,8).setFormula('=IFERROR(TEXT(MINIFS(Movement!F:F,Movement!D:D,C'+r+',Movement!G:G,">"&0),"yyyy-mm-dd"),"")');
    ov.getRange(r,9).setFormula('=IFERROR(IF(F'+r+'=0,"",F'+r+'*E'+r+'),"")');
  }
  ov.setFrozenRows(1);
  Logger.log('restoreOverview 完成：' + data.length + ' 筆產品');
  return { ok: true, count: data.length };
}

// ── 補寫 庫存總覽 公式欄（Products 已刪除後執行）────────────────────────────────
// 讀現有 庫存總覽 的列數，對 F-I 欄補寫公式
// Movement 欄位: A=Date B=Name C=SKU D=EAN E=ASIN F=exp_date G=Boxes H=location I=note

function writeOverviewFormulas() {
  var ss = ss_();
  var ov = ss.getSheetByName('庫存總覽');
  if (!ov) { Logger.log('找不到 庫存總覽'); return { error: '找不到 庫存總覽' }; }

  var lastRow = ov.getLastRow();
  if (lastRow < 2) { Logger.log('庫存總覽 沒有資料列'); return { error: '沒有資料列' }; }

  var count = lastRow - 1;
  var errors = [];

  for (var i = 0; i < count; i++) {
    var row = i + 2;
    try {
      ov.getRange(row,6).setFormula('=IFERROR(SUMIF(Movement!D:D,C'+row+',Movement!G:G),0)');
      ov.getRange(row,7).setFormula('=IF(F'+row+'<=0,"❌ 缺貨",IF(F'+row+'<10,"🟡 注意","✅ 正常"))');
      ov.getRange(row,8).setFormula('=IFERROR(TEXT(MINIFS(Movement!F:F,Movement!D:D,C'+row+',Movement!G:G,">"&0),"yyyy-mm-dd"),"")');
      ov.getRange(row,9).setFormula('=IFERROR(IF(F'+row+'=0,"",F'+row+'*E'+row+'),"")');
    } catch(e) {
      errors.push('row'+row+': '+e.message);
    }
  }

  // 確保表頭第 F-I 欄有標題
  var hdr = ov.getRange(1,1,1,9).getValues()[0];
  var hdrUpdate = false;
  var hdrNames = ['在倉總箱數','庫存狀態','最近效期','估算總pcs'];
  for (var h = 0; h < 4; h++) {
    if (!hdr[5+h]) { ov.getRange(1,6+h).setValue(hdrNames[h]); hdrUpdate = true; }
  }
  if (hdrUpdate) {
    ov.getRange(1,6,1,4).setFontWeight('bold').setBackground('#2D5016').setFontColor('#ffffff');
  }

  var msg = errors.length > 0
    ? '部分錯誤：' + errors.join(' | ')
    : '完成：' + count + ' 列公式寫入成功';
  Logger.log(msg);
  return { ok: errors.length === 0, count: count, errors: errors };
}

// ── Amazon SP-API FBA 庫存同步 ────────────────────────────────────────────────
//
// 使用步驟：
//   1. 執行 setupFbaCredentials()，把 LAB52 API.txt 的三組憑證填入後執行一次
//   2. 執行 syncFbaInventory() 測試是否成功
//   3. 執行 setupFbaTrigger() 設定每小時自動同步

function setupFbaCredentials() {
  PropertiesService.getScriptProperties().setProperties({
    'AMAZON_CLIENT_ID':     '在此貼上 Client ID',
    'AMAZON_CLIENT_SECRET': '在此貼上 Client Secret',
    'AMAZON_REFRESH_TOKEN': '在此貼上 LWA Refresh Token',
    'AMAZON_MARKETPLACE_ID': 'ATVPDKIKX0DER'
  });
  Logger.log('FBA 憑證已儲存至 PropertiesService');
}

function getSpApiToken_() {
  var p = PropertiesService.getScriptProperties();
  // trim：憑證常常是從文件複製貼進「指令碼屬性」的，很容易夾帶尾端空白或換行，
  // 那會讓 Amazon 回 invalid_client，錯誤訊息卻看不出是空白造成的
  function prop(k) { return String(p.getProperty(k) || '').trim(); }
  var resp = UrlFetchApp.fetch('https://api.amazon.com/auth/o2/token', {
    method: 'post',
    contentType: 'application/x-www-form-urlencoded',
    payload: {
      grant_type:    'refresh_token',
      refresh_token: prop('AMAZON_REFRESH_TOKEN'),
      client_id:     prop('AMAZON_CLIENT_ID'),
      client_secret: prop('AMAZON_CLIENT_SECRET')
    },
    muteHttpExceptions: true
  });
  var data = JSON.parse(resp.getContentText());
  if (!data.access_token) throw new Error('Token 取得失敗: ' + resp.getContentText());
  return data.access_token;
}

function syncFbaInventory() {
  try {
    var token = getSpApiToken_();
    var marketplaceId = PropertiesService.getScriptProperties()
      .getProperty('AMAZON_MARKETPLACE_ID') || 'ATVPDKIKX0DER';

    var url = 'https://sellingpartnerapi-na.amazon.com/fba/inventory/v1/summaries'
      + '?granularityType=Marketplace'
      + '&granularityId=' + marketplaceId
      + '&marketplaceIds=' + marketplaceId
      + '&details=true';

    var resp = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: { 'x-amz-access-token': token, 'Content-Type': 'application/json' },
      muteHttpExceptions: true
    });

    var code = resp.getResponseCode();
    if (code !== 200) throw new Error('API 錯誤 ' + code + ': ' + resp.getContentText());

    var summaries = JSON.parse(resp.getContentText()).payload.inventorySummaries || [];

    // 只保留 庫存總覽 裡有的 ASIN（不在追蹤清單的自動排除）
    var trackedAsins = readProducts_().map(function(p) { return p.asin; }).filter(Boolean);
    if (trackedAsins.length > 0) {
      summaries = summaries.filter(function(s) { return trackedAsins.indexOf(s.asin) >= 0; });
    }

    writeFbaSheet_(summaries);
    Logger.log('FBA 庫存同步完成：' + summaries.length + ' 筆');
    // 0 筆不是例外，但表格會保留舊資料不更新，一樣要讓人知道
    if (summaries.length === 0) {
      notifySyncFailure_('FBA 庫存同步',
        'SP-API 有正常回應，但回傳的 ASIN 沒有一個對得上「庫存總覽」，表格保留舊資料未更新。');
    }
    return { ok: true, count: summaries.length };
  } catch(e) {
    Logger.log('syncFbaInventory 錯誤：' + e.message);
    notifySyncFailure_('FBA 庫存同步', e.message);
    return { error: e.message };
  }
}

// 同步失敗通知。
// 原本錯誤只寫進 Logger.log 就 return 掉，從 trigger 執行時回傳值沒人接、
// Logger 沒人看，對 Google 而言每次都「執行成功」，所以既不會被停用也不會
// 有失敗摘要信 —— 這就是同步能無聲停擺 16 天的原因。改成主動寄信。
// 節流：同一種失敗 6 小時內只寄一次，免得每小時的 trigger 洗爆信箱。
var SYNC_ALERT_THROTTLE_HOURS_ = 6;

function notifySyncFailure_(what, detail) {
  try {
    var p    = PropertiesService.getScriptProperties();
    var key  = 'SYNC_ALERT_' + what;
    var last = Number(p.getProperty(key) || 0);
    var now  = new Date().getTime();
    if (last && (now - last) < SYNC_ALERT_THROTTLE_HOURS_ * 3600 * 1000) return;
    p.setProperty(key, String(now));
    GmailApp.sendEmail(RESTOCK_ALERT_EMAIL, '[Lab52 庫存] ' + what + '失敗',
      what + ' 執行失敗：\n\n' + detail
      + '\n\n請到 GAS 編輯器執行 diagFbaSync() 看是斷在哪一層。'
      + '\n（同一種失敗每 ' + SYNC_ALERT_THROTTLE_HOURS_ + ' 小時最多通知一次）');
    Logger.log('[SyncAlert] 已通知 ' + RESTOCK_ALERT_EMAIL + '：' + what);
  } catch(e) {
    // 通知自己掛掉時沒辦法再用通知回報，至少把痕跡留在指令碼屬性裡，
    // diagFbaSync() 會印出來，否則就變成「連告警壞掉都沒人知道」
    Logger.log('notifySyncFailure_ 自己也失敗了：' + e.message);
    try {
      PropertiesService.getScriptProperties().setProperty('SYNC_ALERT_LAST_ERROR',
        Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm')
        + '　' + e.message);
    } catch(_) {}
  }
}

// 前 11 欄是原本就有的，新欄位一律加在最後面 —— 保留日均銷量的邏輯是按欄位「名稱」
// 查找的，所以尾端擴充不會影響既有資料。
var FBA_HDR_ = ['SKU','ASIN','EAN','產品名稱','可出貨數量','入庫中','預留數量','最後更新','狀態','日均銷量','銷量更新日',
                '不可售','調查中','買家訂單','FC處理中','轉運'];

// unfulfillable / researching 在 API 裡是巢狀物件，但實際欄位名還沒用真實資料驗證過
// （入庫和預留都吃過「照文件推論」的虧）。所以這裡不寫死 key：
// 是數字就直接用；是物件就先找 total* 開頭的欄位，找不到才加總所有數值欄位。
function fbaTotalOf_(v) {
  if (v == null) return 0;
  if (typeof v === 'number') return v;
  if (typeof v !== 'object') return 0;
  var k;
  for (k in v) if (/^total/i.test(k) && typeof v[k] === 'number') return v[k];
  var sum = 0, found = false;
  for (k in v) if (typeof v[k] === 'number') { sum += v[k]; found = true; }
  return found ? sum : 0;
}

// SP-API 的 InboundQuantityBreakdown 有三段，但賣家後台的「入庫」欄只算前兩段：
//   inboundWorkingQuantity   已建立入庫計畫、還沒出貨      → 算入庫
//   inboundShippedQuantity   已出貨在途                    → 算入庫
//   inboundReceivingQuantity 已到倉、正在點收              → 後台不算入庫
// receiving 的貨已經在 Amazon 手上、不是在途，把它加進來會比後台多出一截
// （實例：B0DBTNJW4G working=0 shipped=0 receiving=34，後台顯示入庫 0）。
// 原本的程式只取 shipped，漏掉 working；這裡補上 working 但排除 receiving。
function fbaInboundQty_(det) {
  return (det.inboundWorkingQuantity || 0)
       + (det.inboundShippedQuantity || 0);
}

// 賣家後台的「預留」只算 買家訂單 + 運營中心處理中。
// pendingTransshipmentQuantity（運營中心轉運）後台是歸在「現貨」底下的，不算預留
// —— 後台顯示「現貨 765 = 可售 764 + 運營中心轉運 1」。
// 實例 B0DBTNJW4G：API total=60（49 + 轉運 1 + FC處理 10），後台預留顯示 59。
// 用 total 扣掉轉運而不是把兩段相加，是為了讓 Amazon 日後新增的預留類別也能被算進來。
function fbaReservedQty_(det) {
  var rq = det.reservedQuantity || {};
  return Math.max(0, (Number(rq.totalReservedQuantity)        || 0)
                   - (Number(rq.pendingTransshipmentQuantity) || 0));
}

// FBA 庫存水位判定，與前端 fbaAvailQty / 統計磚同一套口徑：
//   qty  = 可出貨 + 入庫中（不含預留，預留已被訂單佔用）
//   有日均銷量就看週數（<4 週缺貨、<10 週注意），沒有就降級用件數門檻 50
function fbaLevel_(qty, daily) {
  if (qty <= 0) return 'out';
  if (daily > 0) {
    var w = qty / (daily * 7);
    if (w < 4)  return 'out';
    if (w < 10) return 'warn';
    return 'ok';
  }
  return qty < 50 ? 'warn' : 'ok';
}

function writeFbaSheet_(summaries) {
  var ss = ss_();
  var sheet = ss.getSheetByName('FBA庫存');
  if (!sheet) {
    sheet = ss.insertSheet('FBA庫存');
    sheet.getRange(1,1,1,FBA_HDR_.length).setValues([FBA_HDR_])
      .setFontWeight('bold').setBackground('#2D5016').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }

  // 空結果就直接收工，不能往下走：下面是「先清空、後寫入」，
  // 若清完才發現沒資料可寫，整張表（含日均銷量歷史）會被洗掉
  if (summaries.length === 0) {
    Logger.log('FBA 同步：summaries 為空，保留既有資料不覆寫');
    return;
  }

  // 欄位擴充過（11 → 16），分頁欄數不夠就先長大，否則下面的 getRange 會 out of bounds
  if (sheet.getMaxColumns() < FBA_HDR_.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), FBA_HDR_.length - sheet.getMaxColumns());
  }

  // 保留現有銷量資料（按 ASIN 鍵值）
  var salesByAsin = {};
  if (sheet.getLastRow() > 1) {
    var ex = sheet.getDataRange().getValues();
    var exHdr = ex[0].map(function(h){ return String(h).trim(); });
    var eA = exHdr.indexOf('ASIN'), eS = exHdr.indexOf('日均銷量'), eD = exHdr.indexOf('銷量更新日');
    if (eA >= 0) {
      for (var ri = 1; ri < ex.length; ri++) {
        var asinKey = String(ex[ri][eA]||'').trim();
        if (asinKey) salesByAsin[asinKey] = {
          sales: eS >= 0 ? ex[ri][eS] : '',
          date:  eD >= 0 ? ex[ri][eD] : ''
        };
      }
    }
    sheet.getRange(2,1,sheet.getLastRow()-1, FBA_HDR_.length).clearContent().clearFormat();
  }

  // 確保標頭有 11 欄
  var curHdr = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0].map(function(h){ return String(h).trim(); });
  if (curHdr.length < FBA_HDR_.length) {
    sheet.getRange(1,1,1,FBA_HDR_.length).setValues([FBA_HDR_])
      .setFontWeight('bold').setBackground('#2D5016').setFontColor('#ffffff');
  }

  var asinToEan = {};
  readProducts_().forEach(function(p) { if (p.asin) asinToEan[p.asin] = p.ean || ''; });

  var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
  var rows = summaries.map(function(s) {
    var det = s.inventoryDetails || {};
    var fulfillable = det.fulfillableQuantity || 0;
    var inbound     = fbaInboundQty_(det);
    var reserved    = fbaReservedQty_(det);
    var saved       = salesByAsin[s.asin] || {};
    var level       = fbaLevel_(fulfillable + inbound, parseFloat(saved.sales) || 0);
    var status      = level === 'out' ? '❌ 缺貨' : (level === 'warn' ? '🟡 注意' : '✅ 正常');
    var ean         = asinToEan[s.asin] || '';
    var rq          = det.reservedQuantity || {};
    return [s.sellerSku||'', s.asin||'', ean, s.productName||'', fulfillable, inbound, reserved, now, status,
            saved.sales !== undefined ? saved.sales : '', saved.date || '',
            fbaTotalOf_(det.unfulfillableQuantity),        // 不可售（殘損、瑕疵…）
            fbaTotalOf_(det.researchingQuantity),          // 調查中
            Number(rq.pendingCustomerOrderQuantity) || 0,  // 預留 · 買家訂單
            Number(rq.fcProcessingQuantity)         || 0,  // 預留 · 運營中心處理中
            Number(rq.pendingTransshipmentQuantity) || 0,  // 運營中心轉運（後台歸在現貨）
            level];
  });

  sheet.getRange(2,1,rows.length, FBA_HDR_.length)
    .setValues(rows.map(function(r){ return r.slice(0, FBA_HDR_.length); }));

  for (var i = 0; i < rows.length; i++) {
    var lv = rows[i][FBA_HDR_.length];
    var bg = lv === 'out' ? '#FFCCCC' : (lv === 'warn' ? '#FFF3CC' : null);
    if (bg) sheet.getRange(i+2, 1, 1, FBA_HDR_.length).setBackground(bg);
  }
}

function readFbaInventory_() {
  var sheet = ss_().getSheetByName('FBA庫存');
  if (!sheet || sheet.getLastRow() <= 1) return [];
  var vals = sheet.getDataRange().getValues();
  var hdr  = vals[0].map(function(h) { return String(h).trim(); });
  var tz   = Session.getScriptTimeZone();
  return vals.slice(1).filter(function(r) { return r[0]; }).map(function(r) {
    var obj = {};
    hdr.forEach(function(h,j) {
      var v = r[j];
      if (v instanceof Date) obj[h] = Utilities.formatDate(v, tz, 'yyyy-MM-dd HH:mm');
      else obj[h] = (v === '' || v === null || v === undefined) ? null : v;
    });
    return obj;
  });
}

function resetFbaSheet() {
  var ss = ss_();
  var sheet = ss.getSheetByName('FBA庫存');
  if (sheet) ss.deleteSheet(sheet);
  Logger.log('FBA庫存 分頁已刪除，請執行 syncFbaInventory() 重建');
  return syncFbaInventory();
}

// ── 診斷：FBA 同步為什麼沒在跑 ────────────────────────────────────────────────
// 在 GAS 編輯器選 diagFbaSync 按執行，然後看下方「執行紀錄」。
// 逐層檢查 憑證設定 → Trigger → LWA token → SP-API → 分頁時間，指出斷在哪一層。
// 只印出設定「是否存在」與長度，不會印出憑證內容。
function diagFbaSync() {
  var lines = [];
  function say(s) { lines.push(s); Logger.log(s); }

  // 1. 憑證設定是否存在
  var p = PropertiesService.getScriptProperties();
  ['AMAZON_REFRESH_TOKEN','AMAZON_CLIENT_ID','AMAZON_CLIENT_SECRET','AMAZON_MARKETPLACE_ID']
    .forEach(function(k) {
      var v = p.getProperty(k);
      if (!v) { say('[1 設定] ' + k + ' → ★ 未設定 ★'); return; }
      var t = String(v).trim();
      say('[1 設定] ' + k + ' → 已設定（長度 ' + String(v).length + '）'
          + (t.length !== String(v).length
              ? '　★ 含前後空白／換行，實際內容長度 ' + t.length + ' ★' : ''));
      // 識別碼和密碼很容易貼反：client ID 一定是 amzn1.application-oa2-client. 開頭
      var isId = t.indexOf('amzn1.application-oa2-client.') === 0;
      if (k === 'AMAZON_CLIENT_ID' && !isId) {
        say('[1 設定] ★ CLIENT_ID 不是 amzn1.application-oa2-client. 開頭，可能貼錯欄位 ★');
      }
      if (k === 'AMAZON_CLIENT_SECRET' && isId) {
        say('[1 設定] ★ CLIENT_SECRET 填的是「用戶端識別碼」，兩個欄位貼反了 ★');
      }
    });

  // 1b. 通知管道自己有沒有壞掉（notifySyncFailure_ 寄信失敗時會留下這筆）
  var alertErr = p.getProperty('SYNC_ALERT_LAST_ERROR');
  say('[1 通知] 收件人 ' + RESTOCK_ALERT_EMAIL
      + (alertErr ? '　★ 上次寄信失敗：' + alertErr + ' ★' : '　（沒有寄信失敗紀錄）'));

  // 2. Trigger 還在不在
  var trs = ScriptApp.getProjectTriggers();
  say('[2 Trigger] 專案共 ' + trs.length + ' 個');
  trs.forEach(function(t) {
    say('[2 Trigger] ' + t.getHandlerFunction() + '（' + t.getEventType() + '）');
  });
  // sendRestockAlert 也要檢查 —— 它的 trigger 一直沒被建立過，
  // 導致補貨警示從上線以來一次都沒發出，而原本的診斷只看兩支同步、驗不出來。
  [['syncFbaInventory',     'setupFbaTrigger()'],
   ['syncFbaSalesVelocity', 'setupSalesTrigger()'],
   ['sendRestockAlert',     'setupRestockAlertTrigger()']].forEach(function(pair) {
    var found = trs.some(function(t) { return t.getHandlerFunction() === pair[0]; });
    if (!found) {
      say('[2 Trigger] ★ 找不到 ' + pair[0] + ' 的 trigger，它不會自己跑'
          + '（要啟用請執行 ' + pair[1] + '）★');
    }
  });

  // 3. LWA 換 token（Amazon 的錯誤回應只有錯誤碼，不含憑證內容）
  var token = null;
  try {
    var tResp = UrlFetchApp.fetch('https://api.amazon.com/auth/o2/token', {
      method: 'post',
      contentType: 'application/x-www-form-urlencoded',
      payload: {
        grant_type:    'refresh_token',
        refresh_token: p.getProperty('AMAZON_REFRESH_TOKEN'),
        client_id:     p.getProperty('AMAZON_CLIENT_ID'),
        client_secret: p.getProperty('AMAZON_CLIENT_SECRET')
      },
      muteHttpExceptions: true
    });
    say('[3 LWA] HTTP ' + tResp.getResponseCode());
    var tBody = tResp.getContentText();
    var tData = {};
    try { tData = JSON.parse(tBody); } catch(_) {}
    if (tData.access_token) { token = tData.access_token; say('[3 LWA] access_token 取得成功'); }
    else say('[3 LWA] ★ 失敗 ★ ' + String(tBody).slice(0, 400));
  } catch(e) {
    say('[3 LWA] ★ 例外 ★ ' + e.message);
  }

  // 4. 真的打一次 SP-API 庫存查詢
  if (!token) {
    say('[4 SP-API] 跳過（沒拿到 token，問題在第 3 層以前）');
  } else {
    var mkt = p.getProperty('AMAZON_MARKETPLACE_ID') || 'ATVPDKIKX0DER';
    try {
      var iResp = UrlFetchApp.fetch(
        'https://sellingpartnerapi-na.amazon.com/fba/inventory/v1/summaries'
          + '?granularityType=Marketplace&granularityId=' + mkt
          + '&marketplaceIds=' + mkt + '&details=true',
        { method: 'get',
          headers: { 'x-amz-access-token': token, 'Content-Type': 'application/json' },
          muteHttpExceptions: true });
      var iCode = iResp.getResponseCode();
      say('[4 SP-API] HTTP ' + iCode);
      if (iCode !== 200) {
        say('[4 SP-API] ★ 回應 ★ ' + String(iResp.getContentText()).slice(0, 500));
      } else {
        var sums = (JSON.parse(iResp.getContentText()).payload || {}).inventorySummaries || [];
        var tracked = readProducts_().map(function(x){ return x.asin; }).filter(Boolean);
        var hit = sums.filter(function(s){ return tracked.indexOf(s.asin) >= 0; });
        say('[4 SP-API] Amazon 回傳 ' + sums.length + ' 筆；庫存總覽有 ' + tracked.length
            + ' 個 ASIN，對得上 ' + hit.length + ' 筆');
        // 對不上的個別列出來：這些產品永遠不會出現在 FBA庫存，很容易被忽略
        var amzAsins = sums.map(function(s){ return s.asin; });
        var missing  = tracked.filter(function(a){ return amzAsins.indexOf(a) < 0; });
        if (missing.length) {
          say('[4 SP-API] ★ ' + missing.length + ' 個 ASIN 在「庫存總覽」但 Amazon 沒回傳，'
              + '不會出現在 FBA庫存：' + missing.join(', ') + ' ★');
        }
        if (sums.length && !hit.length) {
          say('[4 SP-API] ★ 一筆都對不上，同步會保留舊資料不覆寫 ★');
          say('[4 SP-API] Amazon 前 3 個 ASIN：' + sums.slice(0,3).map(function(s){ return s.asin; }).join(', '));
          say('[4 SP-API] 庫存總覽前 3 個 ASIN：' + tracked.slice(0,3).join(', '));
        }
        // 入庫中三段各自的數字，用來確認 fbaInboundQty_ 的修正是否吃到資料
        // API 現值 vs 分頁現值逐支對照。兩邊一致 → 分頁是新的，差異來自欄位口徑；
        // 兩邊不一致 → 分頁還沒重新同步，先跑 syncFbaInventory 再談。
        var sheetByAsin = {};
        readFbaInventory_().forEach(function(r) { if (r['ASIN']) sheetByAsin[r['ASIN']] = r; });
        hit.forEach(function(s) {
          var d   = s.inventoryDetails || {};
          var row = sheetByAsin[s.asin] || {};
          say('[4 對照] ' + s.asin
              + ' ┃ API 可出貨=' + (d.fulfillableQuantity || 0)
              + ' 入庫中=' + fbaInboundQty_(d)
              + '(w' + (d.inboundWorkingQuantity   || 0)
              + '/s' + (d.inboundShippedQuantity   || 0)
              + '/r' + (d.inboundReceivingQuantity || 0) + ')'
              + ' 預留=' + fbaReservedQty_(d)
              + ' ┃ 分頁 可出貨=' + (row['可出貨數量'] || 0)
              + ' 入庫中=' + (row['入庫中'] || 0)
              + ' 預留=' + (row['預留數量'] || 0));
        });
        // 同一個 ASIN 可能有多個 SKU，API 會回多筆。後台的商品詳情面板是看單一
        // listing 的，兩邊數字自然對不起來（面板顯示某個 SKU，我們顯示的是另一筆）。
        var byAsin = {};
        hit.forEach(function(s) { (byAsin[s.asin] = byAsin[s.asin] || []).push(s); });
        Object.keys(byAsin).filter(function(a) { return byAsin[a].length > 1; })
          .forEach(function(a) {
            var list = byAsin[a];
            say('[4 重複] ★ ' + a + ' 有 ' + list.length + ' 個 SKU：'
                + list.map(function(s) {
                    return (s.sellerSku || '?') + ' 可出貨='
                         + ((s.inventoryDetails || {}).fulfillableQuantity || 0);
                  }).join('　')
                + '　合計=' + list.reduce(function(n, s) {
                    return n + ((s.inventoryDetails || {}).fulfillableQuantity || 0);
                  }, 0) + ' ★');
          });

        // 預留的完整拆分：印原始物件，不預設欄位名，才看得出 Amazon 實際怎麼拆
        hit.filter(function(s) {
          var rq = (s.inventoryDetails || {}).reservedQuantity;
          return rq && Number(rq.totalReservedQuantity) > 0;
        }).slice(0, 5).forEach(function(s) {
          say('[4 預留明細] ' + s.asin + ' '
              + JSON.stringify(s.inventoryDetails.reservedQuantity));
        });
      }
    } catch(e) {
      say('[4 SP-API] ★ 例外 ★ ' + e.message);
    }
  }

  // 5. 分頁現在的時間戳
  var fba = ss_().getSheetByName('FBA庫存');
  if (!fba || fba.getLastRow() < 2) {
    say('[5 分頁] FBA庫存 沒有資料列');
  } else {
    var vals = fba.getDataRange().getValues();
    var hd   = vals[0].map(function(h){ return String(h).trim(); });
    var ci   = hd.indexOf('最後更新'), cs = hd.indexOf('銷量更新日');
    say('[5 分頁] 共 ' + (vals.length - 1) + ' 列，最後更新=' + (ci >= 0 ? vals[1][ci] : '?')
        + '，銷量更新日=' + (cs >= 0 ? vals[1][cs] : '?'));
  }

  return lines.join('\n');
}

function setupFbaTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'syncFbaInventory') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('syncFbaInventory').timeBased().everyHours(1).create();
  Logger.log('FBA 自動同步 Trigger 已設定（每小時）');
  return { ok: true };
}

// ── FBA 銷量同步（每日一次，取過去 7 天） ─────────────────────────────────────
// 執行步驟：
//   1. 執行 syncFbaSalesVelocity() 立即同步一次
//   2. 執行 setupSalesTrigger() 設定每天早上 7 點自動同步
// Trigger 掛的是這個名字，維持不變才不用重設 trigger。
// 這層只負責：失敗時先寄信通知，再把錯誤原樣往上丟
// （繼續讓 Google 記錄成失敗，trigger 失敗摘要才看得到）。
function syncFbaSalesVelocity() {
  try {
    var r = syncFbaSalesVelocity_();
    checkInventoryFreshness_();   // 每天順便當一次看門狗
    return r;
  } catch(e) {
    notifySyncFailure_('FBA 銷量同步', e.message);
    throw e;
  }
}

// 看門狗：庫存同步是獨立的每小時 trigger，萬一那個 trigger 被刪掉或停用，
// 它自己不會發出任何聲音（沒執行 = 沒有錯誤 = 沒有通知）。
// 2026-08-18 停擺 16 天沒人發現就是這個死角，所以由每天的銷量同步順便檢查一次。
function checkInventoryFreshness_() {
  try {
    var rows = readFbaInventory_();
    if (!rows.length) return;
    var last = rows[0]['最後更新'];
    if (!last) return;
    var t = new Date(String(last).replace(' ', 'T'));
    if (isNaN(t.getTime())) return;
    var hrs = (new Date().getTime() - t.getTime()) / 3600000;
    if (hrs >= 6) {
      notifySyncFailure_('FBA 庫存同步停擺',
        '庫存最後更新於 ' + last + '，已經 ' + Math.floor(hrs) + ' 小時沒有更新。\n'
        + '每小時的 syncFbaInventory trigger 可能被刪除或停用。\n'
        + '請執行 diagFbaSync() 檢查，第 2 段會列出目前存在的 trigger。');
    }
  } catch(e) {
    Logger.log('checkInventoryFreshness_ 失敗：' + e.message);
  }
}

// 一鍵驗證通知管道真的通。沒測過的告警等於沒有告警。
// 這支刻意「不」透過 notifySyncFailure_，也刻意不 catch —— 寄不出去就要讓例外
// 直接浮到執行紀錄。第一版走 notifySyncFailure_，結果它把自己的錯誤也吞掉，
// 寄失敗還是回傳 ok:true，等於重蹈同步無聲失敗的覆轍。
function testSyncAlert() {
  var to = RESTOCK_ALERT_EMAIL;
  Logger.log('[測試] 收件人：' + to);
  // 執行帳號要用 Session.getEffectiveUser() 才拿得到，但那需要 userinfo.email scope，
  // 不值得為此新增 scope（見 testSyncAlertTo 的說明）。看編輯器右上角的頭像即可。
  try {
    Logger.log('[測試] 今日剩餘寄信額度：' + MailApp.getRemainingDailyQuota());
  } catch(e) {
    Logger.log('[測試] 取不到寄信額度：' + e.message);
  }

  GmailApp.sendEmail(to, '[Lab52 庫存] 通知管道測試',
    '這是一封測試信。收到它，就代表 FBA 同步出問題時你會被通知。');

  Logger.log('[測試] sendEmail 呼叫完成、未拋出例外 → 信已交給 Gmail');
  Logger.log('[測試] 若仍未收到：檢查垃圾郵件，以及 ' + to + ' 端的過濾規則');
  return { ok: true, sentTo: to };
}

// 寄到指令碼屬性 TEST_ALERT_TO 指定的信箱（沒設就退回 RESTOCK_ALERT_EMAIL）。
// 用來切開兩種失敗：寄不出去（程式／授權問題）vs 寄出去了但沒送達（收件端過濾）。
//
// 這裡刻意不用 Session.getEffectiveUser() 取執行帳號 —— 那需要 userinfo.email scope，
// 而新增 oauthScopes 會讓網頁應用程式在擁有者重新授權之前失效，
// 為了一個診斷功能冒這個險不划算。要知道執行帳號直接看編輯器右上角的頭像即可。
function testSyncAlertTo() {
  var to = String(PropertiesService.getScriptProperties().getProperty('TEST_ALERT_TO') || '').trim()
           || RESTOCK_ALERT_EMAIL;
  Logger.log('[測試] 寄到：' + to);
  GmailApp.sendEmail(to, '[Lab52 庫存] 通知管道測試',
    '收到這封，就代表程式與 Gmail 授權都正常，\n'
    + 'FBA 同步出問題時的通知信也寄得出去。');
  Logger.log('[測試] sendEmail 呼叫完成、未拋出例外');
  return { ok: true, sentTo: to };
}

function syncFbaSalesVelocity_() {
  var token  = getSpApiToken_();
  var mktId  = PropertiesService.getScriptProperties().getProperty('AMAZON_MARKETPLACE_ID') || 'ATVPDKIKX0DER';
  var sheet  = ss_().getSheetByName('FBA庫存');
  if (!sheet || sheet.getLastRow() <= 1) {
    Logger.log('FBA庫存 sheet 為空，請先執行 syncFbaInventory()');
    return;
  }

  var data   = sheet.getDataRange().getValues();
  var hdr    = data[0].map(function(h){ return String(h).trim(); });
  var iAsin  = hdr.indexOf('ASIN');
  var iSales = hdr.indexOf('日均銷量');
  var iDate  = hdr.indexOf('銷量更新日');
  if (iAsin < 0) { Logger.log('找不到 ASIN 欄位'); return; }

  // 若欄位不存在則新增
  if (iSales < 0) { iSales = hdr.length;   sheet.getRange(1, iSales+1).setValue('日均銷量').setFontWeight('bold').setBackground('#2D5016').setFontColor('#ffffff'); }
  if (iDate  < 0) { iDate  = iSales === hdr.length ? hdr.length+1 : hdr.length; sheet.getRange(1, iDate+1).setValue('銷量更新日').setFontWeight('bold').setBackground('#2D5016').setFontColor('#ffffff'); }

  // 過去 7 天區間
  var now  = new Date();
  var past = new Date(now.getTime() - 7 * 24 * 3600 * 1000);
  function iso(d) { return d.toISOString().slice(0,19) + 'Z'; }
  var interval = iso(past) + '--' + iso(now);
  var today    = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd');

  var updated = 0, failed = 0;
  for (var i = 1; i < data.length; i++) {
    var asin = String(data[i][iAsin] || '').trim();
    if (!asin) continue;

    var url = 'https://sellingpartnerapi-na.amazon.com/sales/v1/orderMetrics'
      + '?marketplaceIds=' + mktId
      + '&interval=' + encodeURIComponent(interval)
      + '&granularity=total'
      + '&asin=' + asin;

    var resp = UrlFetchApp.fetch(url, {
      headers: { 'x-amz-access-token': token, 'Content-Type': 'application/json' },
      muteHttpExceptions: true
    });

    if (resp.getResponseCode() === 200) {
      var payload = JSON.parse(resp.getContentText()).payload || [];
      var units   = payload.length > 0 ? (Number(payload[0].unitCount) || 0) : 0;
      var daily   = Math.round((units / 7) * 10) / 10;
      sheet.getRange(i+1, iSales+1).setValue(daily);
      sheet.getRange(i+1, iDate+1).setValue(today);
      updated++;
    } else {
      Logger.log('ASIN ' + asin + ' 失敗: HTTP ' + resp.getResponseCode());
      failed++;
    }
    Utilities.sleep(2100); // 尊重 0.5 req/s 限制
  }
  Logger.log('銷量同步完成：' + updated + ' 筆更新，' + failed + ' 筆失敗');
  // 每支 ASIN 的失敗原本只寫 Logger，整批全掛也照樣「執行成功」
  if (failed > 0 && updated === 0) {
    notifySyncFailure_('FBA 銷量同步', '全部 ' + failed + ' 個 ASIN 都查詢失敗，日均銷量未更新。');
  }
  return { ok: true, updated: updated, failed: failed };
}

function setupSalesTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'syncFbaSalesVelocity') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('syncFbaSalesVelocity')
    .timeBased().everyDays(1).atHour(7).create();
  Logger.log('銷量自動同步 Trigger 已設定（每天早上 7 點）');
}

// ── 補貨提醒 Email ────────────────────────────────────────────────────────────
//
// 設定步驟：
//   1. 執行 setupRestockAlertTrigger() 一次，設定每天早上 9 點自動發信
//   2. 執行 sendRestockAlert() 可立即手動測試
//   3. 執行 removeRestockAlertTrigger() 可取消定時發信

var RESTOCK_ALERT_EMAIL   = 'annicewu@toothfilm.com';
// 與前端 renderDashboard 的 TARGET_DAYS 必須一致，否則信件裡的建議出貨量
// 會跟網頁上看到的對不起來（這個常數同時決定發警示的門檻與建議量的算法）
var RESTOCK_TARGET_DAYS   = 60;   // 低於此天數才發警示
var RESTOCK_URGENT_DAYS   = 28;   // 低於此天數標記為緊急

function sendRestockAlert() {
  var products = readProducts_();
  var fbaRows  = readFbaInventory_();
  var transits = readSheet_('Transits');

  var fbaByAsin = {};
  fbaRows.forEach(function(r) { if (r['ASIN']) fbaByAsin[r['ASIN']] = r; });

  // 台灣倉庫存
  var twBoxByEan = {};
  readTaiwanMovements_().forEach(function(m) {
    twBoxByEan[m.ean] = (twBoxByEan[m.ean] || 0) + (parseFloat(m.boxes) || 0);
  });

  // 海外倉庫存
  // Movement 可能混有 location=TW 的列，那些已經算在 twBoxByEan 裡，要排除避免重複計算
  var overseasBoxByEan = {};
  readMovements_().forEach(function(m) {
    if (String(m.location || 'AMZLGS').toUpperCase() === 'TW') return;
    overseasBoxByEan[m.ean] = (overseasBoxByEan[m.ean] || 0) + (parseFloat(m.boxes) || 0);
  });

  var alerts = [];
  products.forEach(function(p) {
    if (!p.asin) return;
    var fba   = fbaByAsin[p.asin]; if (!fba) return;
    var daily = Number(fba['日均銷量'] || 0); if (!daily) return;
    var qpc   = parseFloat(p.qty_per_carton) || 1;

    // 與前端補貨建議同口徑：只看 FBA 自己的量（可出貨 + 入庫中，不含預留）。
    // 原本加的在途是往「海外倉」的貨，那批到不了 FBA，不該降低 FBA 的補貨需求；
    // 真正要進 FBA 的貨 Amazon 已經算在「入庫中」，再加一次會重複。
    // 貨從哪裡出由下方的 source 欄位判斷，跟「FBA 還缺多少」是兩件事。
    var fbaQty       = parseFloat(fba['可出貨數量'] || 0);
    var effectiveQty = fbaQty + parseFloat(fba['入庫中'] || 0);
    var days = Math.round(effectiveQty / daily);
    if (days >= RESTOCK_TARGET_DAYS) return;

    var shipBoxes     = Math.ceil(Math.max(0, daily * RESTOCK_TARGET_DAYS - effectiveQty) / qpc / 5) * 5;
    var twStock       = twBoxByEan[p.ean] || 0;
    var overseasStock = overseasBoxByEan[p.ean] || 0;
    var source;
    if      (overseasStock >= shipBoxes) source = '海外倉';
    else if (overseasStock > 0)          source = '海外倉（不足）';
    else if (twStock >= shipBoxes)       source = '台灣倉';
    else if (twStock > 0)                source = '台灣倉（不足）';
    else                                 source = '缺貨';

    alerts.push({
      name: p.name, ean: p.ean, sku: p.sku, asin: p.asin,
      days: days, shipBoxes: shipBoxes, fbaQty: Math.round(fbaQty),
      daily: daily, source: source
    });
  });

  if (alerts.length === 0) {
    Logger.log('[RestockAlert] 所有產品庫存充足，不發信');
    return { ok: true, sent: false, reason: 'no alerts' };
  }

  alerts.sort(function(a, b) { return a.days - b.days; });

  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  var subject = '[Lab52 補貨提醒] ' + today + '：' + alerts.length + ' 項產品需要補貨';

  // ── 純文字版 ──
  var text = '以下產品 FBA 庫存低於 ' + RESTOCK_TARGET_DAYS + ' 天，請安排補貨：\n\n';
  alerts.forEach(function(a) {
    text += (a.days < RESTOCK_URGENT_DAYS ? '🔴 緊急' : '🟡 警示') + '  ' + a.name + '\n';
    text += '  剩餘天數：' + a.days + ' 天 ｜ 建議補：' + a.shipBoxes + ' 箱 ｜ 備貨來源：' + a.source + '\n';
    text += '  ASIN：' + a.asin + ' ｜ SKU：' + a.sku + '\n\n';
  });
  text += '\n— Lab52 庫存系統自動發送';

  // ── HTML 版 ──
  var rows = alerts.map(function(a) {
    var color = a.days < RESTOCK_URGENT_DAYS ? '#C0392B' : '#E67E22';
    var label = a.days < RESTOCK_URGENT_DAYS ? '緊急' : '警示';
    return '<tr>'
      + '<td style="padding:6px 10px"><span style="background:' + color + ';color:#fff;padding:2px 7px;border-radius:4px;font-size:12px">' + label + '</span></td>'
      + '<td style="padding:6px 10px;font-weight:600">' + a.name + '</td>'
      + '<td style="padding:6px 10px;text-align:center;color:' + color + ';font-weight:700">' + a.days + ' 天</td>'
      + '<td style="padding:6px 10px;text-align:center;font-weight:700">' + a.shipBoxes + ' 箱</td>'
      + '<td style="padding:6px 10px;color:#555">' + a.source + '</td>'
      + '<td style="padding:6px 10px;color:#888;font-size:12px">' + a.asin + '</td>'
      + '</tr>';
  }).join('');

  var html = '<div style="font-family:system-ui,sans-serif;max-width:700px">'
    + '<h2 style="color:#2D5016;margin-bottom:4px">Lab52 補貨提醒</h2>'
    + '<p style="color:#666;margin-top:0">' + today + '　共 <b>' + alerts.length + '</b> 項產品需要補貨</p>'
    + '<table style="border-collapse:collapse;width:100%;background:#fff;border:1px solid #ddd;border-radius:6px">'
    + '<thead><tr style="background:#2D5016;color:#fff">'
    + '<th style="padding:8px 10px">狀態</th><th style="padding:8px 10px">產品名稱</th>'
    + '<th style="padding:8px 10px">剩餘天數</th><th style="padding:8px 10px">建議補貨</th>'
    + '<th style="padding:8px 10px">備貨來源</th><th style="padding:8px 10px">ASIN</th>'
    + '</tr></thead><tbody>' + rows + '</tbody></table>'
    + '<p style="color:#999;font-size:12px;margin-top:16px">此信件由 Lab52 庫存系統自動發送（目標天數：' + RESTOCK_TARGET_DAYS + ' 天）</p>'
    + '</div>';

  GmailApp.sendEmail(RESTOCK_ALERT_EMAIL, subject, text, { htmlBody: html });
  Logger.log('[RestockAlert] 已發送至 ' + RESTOCK_ALERT_EMAIL + '，共 ' + alerts.length + ' 項');
  return { ok: true, sent: true, count: alerts.length };
}

function setupRestockAlertTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'sendRestockAlert') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sendRestockAlert').timeBased().everyDays(1).atHour(9).create();
  Logger.log('[RestockAlert] Trigger 已設定：每天早上 9 點');
  return { ok: true };
}

function removeRestockAlertTrigger() {
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'sendRestockAlert') ScriptApp.deleteTrigger(t);
  });
  Logger.log('[RestockAlert] Trigger 已取消');
  return { ok: true };
}

// ── 診斷：確認各分頁欄位與資料狀況 ────────────────────────────────────────────
// 執行後到 執行記錄 查看輸出，確認 EAN 是否正確讀入
function debugData() {
  // 顯示 庫存總覽 的實際欄位結構
  var ov = ss_().getSheetByName('庫存總覽');
  if (ov) {
    var ovHdr = ov.getRange(1,1,1,ov.getLastColumn()).getValues()[0];
    var ovRow2 = ov.getLastRow() > 1 ? ov.getRange(2,1,1,ov.getLastColumn()).getValues()[0] : [];
    Logger.log('庫存總覽 欄位: ' + JSON.stringify(ovHdr));
    Logger.log('庫存總覽 第2列: ' + JSON.stringify(ovRow2));
  }
  // 顯示 Movement 的實際欄位結構
  var mv = ss_().getSheetByName('Movement');
  if (mv) {
    var mvHdr = mv.getRange(1,1,1,mv.getLastColumn()).getValues()[0];
    var mvRow2 = mv.getLastRow() > 1 ? mv.getRange(2,1,1,mv.getLastColumn()).getValues()[0] : [];
    Logger.log('Movement 欄位: ' + JSON.stringify(mvHdr));
    Logger.log('Movement 第2列: ' + JSON.stringify(mvRow2));
  }
  // EAN 比對
  var products = readProducts_();
  var movements = readMovements_();
  var productEans = products.map(function(p){ return p.ean; });
  var movEans = movements.map(function(m){ return m.ean; }).filter(function(e,i,a){ return a.indexOf(e)===i; });
  var matched = productEans.filter(function(e){ return movEans.indexOf(e) >= 0; });
  Logger.log('---');
  Logger.log('產品 EAN 數: ' + productEans.length + ' | Movement EAN 種類: ' + movEans.length + ' | 對應到: ' + matched.length);
  Logger.log('產品前3 EAN: ' + JSON.stringify(productEans.slice(0,3)));
  Logger.log('Movement前3 EAN: ' + JSON.stringify(movEans.slice(0,3)));
  return { products: productEans.length, movementEans: movEans.length, matched: matched.length };
}

// ── 測試 Sales API 是否有權限 ─────────────────────────────────────────────────
// 執行此函式，在 Logs 看結果:
//   200 → 有權限，可以串銷量
//   403 → App 沒有 Selling Partner Insights 角色，需在 Developer Central 新增
//   其他 → 其他問題
function diagSalesApi() {
  var token = getSpApiToken_();
  var mktId = PropertiesService.getScriptProperties().getProperty('AMAZON_MARKETPLACE_ID') || 'ATVPDKIKX0DER';

  // 取第一個 ASIN 來測試（從 FBA庫存 sheet）
  var fbaSheet = ss_().getSheetByName('FBA庫存');
  var testAsin = '';
  if (fbaSheet && fbaSheet.getLastRow() > 1) {
    var asinCol = fbaSheet.getRange(1,1,1,fbaSheet.getLastColumn()).getValues()[0].indexOf('ASIN');
    if (asinCol >= 0) testAsin = String(fbaSheet.getRange(2, asinCol+1).getValue() || '');
  }
  if (!testAsin) { Logger.log('找不到 ASIN，請先執行 syncFbaInventory()'); return; }

  // 查詢最近 30 天銷量
  var now  = new Date();
  var past = new Date(now.getTime() - 30 * 24 * 3600 * 1000);
  function iso(d) { return d.toISOString().slice(0,19) + 'Z'; }
  var interval = iso(past) + '--' + iso(now);

  var url = 'https://sellingpartnerapi-na.amazon.com/sales/v1/orderMetrics'
    + '?marketplaceIds=' + mktId
    + '&interval=' + encodeURIComponent(interval)
    + '&granularity=total'
    + '&asin=' + testAsin;

  var resp = UrlFetchApp.fetch(url, {
    headers: { 'x-amz-access-token': token, 'Content-Type': 'application/json' },
    muteHttpExceptions: true
  });

  Logger.log('HTTP 狀態: ' + resp.getResponseCode());
  Logger.log('回應內容: ' + resp.getContentText().slice(0, 500));

  if (resp.getResponseCode() === 200) {
    Logger.log('✅ Sales API 可用！可以開始串銷量。');
  } else if (resp.getResponseCode() === 403) {
    Logger.log('❌ 403 權限不足 — 請至 Amazon Developer Central 替 App 新增「Selling Partner Insights」角色。');
  } else {
    Logger.log('⚠️ 其他錯誤，請查看回應內容。');
  }
}
