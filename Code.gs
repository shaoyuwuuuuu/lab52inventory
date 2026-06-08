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
  if (e && e.parameter && e.parameter.action === 'getData') {
    var cb = e.parameter.callback || '';
    var result = JSON.stringify(getAllData());
    return ContentService
      .createTextOutput(cb ? cb + '(' + result + ')' : result)
      .setMimeType(cb ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON);
  }
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Lab52 庫存系統')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ── Helpers ───────────────────────────────────────────────────────────────────

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }

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

function getAllData() {
  return {
    products:  readProducts_(),
    movements: readMovements_(),
    taiwan:    readTaiwanMovements_(),
    transits:  readSheet_('Transits'),
    billing:   readSheet_('Billing')
  };
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
      boxes:          iBox >= 0 ? (parseFloat(r[iBox]) || 0) : 0
    });
  }
  return result;
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

function addTransit(d) {
  try {
    var tSheet = ss_().getSheetByName('Transits');
    if (!tSheet) return { error: 'Transits sheet not found' };
    var id  = nextId_('Transits');
    var qty = Math.abs(parseFloat(d.qty_cartons) || 0);
    var fromLoc = d.from_location || 'TW';
    insertAtTop_(tSheet, [
      id,
      d.ean           || '',
      d.product_name  || '',
      d.sku           || '',
      fromLoc,
      d.to_location   || 'AMZLGS',
      qty,
      d.exp_date      || '',
      d.ship_date     || '',
      d.eta_date      || '',
      'TRANSIT',
      '',
      d.note          || '',
      nowStr_()
    ]);
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
