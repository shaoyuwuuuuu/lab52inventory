// ── Lab52 Inventory — Google Apps Script Backend ──────────────────────────────
// 設定步驟:
// 1. 開啟 Google Sheets，建好此試算表
// 2. 工具 → Apps Script → 貼入此檔案為 Code.gs
// 3. 點左側 + 新增 HTML 檔案，命名為 index，貼入 inventory_app.html 的內容
// 4. 執行 setupSheets()（執行一次即可，建立分頁與表頭）
// 5. 部署 → 新增部署 → 網頁應用程式
//    執行身分: 我 / 存取權限: 知道連結的所有人

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

function getSheetRows(name) {
  var sheet = ss_().getSheetByName(name);
  if (!sheet) return { error: 'Sheet not found: ' + name, rows: [] };
  var vals = sheet.getDataRange().getValues();
  if (vals.length <= 1) return { rows: [] };
  var hdrs = vals[0].map(function(h) { return String(h).trim(); });
  var rows = [];
  for (var i = 1; i < vals.length; i++) {
    if (vals[i].every(function(c) { return c === '' || c === null; })) continue;
    var row = {};
    hdrs.forEach(function(h, j) {
      var v = vals[i][j];
      row[h] = (v === '' || v === null || v === undefined) ? null : v;
    });
    rows.push(row);
  }
  return { rows: rows };
}

function getAllData() {
  return {
    products:  getSheetRows('Products'),
    batches:   getSheetRows('Batches'),
    movements: getSheetRows('Movements'),
    billing:   getSheetRows('Billing')
  };
}

function nextId_(name) {
  var vals = ss_().getSheetByName(name).getDataRange().getValues();
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

function nowStr_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
}

// ── One-time Setup ────────────────────────────────────────────────────────────

function setupSheets() {
  var spreadsheet = ss_();
  var defs = {
    'Products':  ['id','name','sku','ean','asin','cubic_feet_per_carton','qty_per_carton','note','created_at'],
    'Batches':   ['id','product_id','exp_date','location','current_cartons','note','created_at'],
    'Movements': ['id','batch_id','date','type','qty_cartons','note','created_at'],
    'Billing':   ['id','date','type','job_name','qty','unit_fee','total_fee','note']
  };
  Object.keys(defs).forEach(function(name) {
    var sheet = spreadsheet.getSheetByName(name);
    if (!sheet) sheet = spreadsheet.insertSheet(name);
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(defs[name]);
      sheet.getRange(1, 1, 1, defs[name].length).setFontWeight('bold').setBackground('#1a2035').setFontColor('#ffffff');
    }
  });
  setupSummarySheet_();
  Logger.log('Done! All sheets ready.');
  return { ok: true };
}

function setupSummarySheet_() {
  var spreadsheet = ss_();

  // 確保 Products 分頁存在才能建公式
  var pSheet = spreadsheet.getSheetByName('Products');
  if (!pSheet) return;

  var sheet = spreadsheet.getSheetByName('庫存總覽');
  if (!sheet) {
    sheet = spreadsheet.insertSheet('庫存總覽');
    // 移到第一個位置
    spreadsheet.setActiveSheet(sheet);
    spreadsheet.moveActiveSheet(1);
  } else {
    sheet.clearContents();
  }

  // 表頭
  var headers = ['產品名稱','SKU','ASIN','在倉總箱數','最近效期','最近效期批剩餘','每箱pcs','估算總pcs'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold').setBackground('#1a2035').setFontColor('#ffffff');

  // 公式列（最多支援 50 個產品，空的自動隱藏）
  var formulas = [];
  for (var r = 2; r <= 51; r++) {
    var pr = 'Products!A' + r; // 對應 Products 的第 r 列
    formulas.push([
      '=IFERROR(IF('+pr+'="","",VLOOKUP('+pr+',Products!A:B,2,0)),"")',
      '=IFERROR(IF('+pr+'="","",VLOOKUP('+pr+',Products!A:C,3,0)),"")',
      '=IFERROR(IF('+pr+'="","",VLOOKUP('+pr+',Products!A:E,5,0)),"")',
      '=IFERROR(IF('+pr+'="","",SUMIF(Batches!B:B,'+pr+',Batches!E:E)),"")',
      '=IFERROR(IF('+pr+'="","",TEXT(MINIFS(Batches!C:C,Batches!B:B,'+pr+',Batches!E:E,">"&0),"yyyy-mm-dd")),"")',
      '=IFERROR(IF('+pr+'="","",SUMIFS(Batches!E:E,Batches!B:B,'+pr+',Batches!C:C,MINIFS(Batches!C:C,Batches!B:B,'+pr+',Batches!E:E,">"&0))),"")',
      '=IFERROR(IF('+pr+'="","",VLOOKUP('+pr+',Products!A:G,7,0)),"")',
      '=IFERROR(IF(D'+r+'="","",D'+r+'*G'+r+'),"")',
    ]);
  }
  sheet.getRange(2, 1, formulas.length, headers.length).setFormulas(formulas);

  // 欄寬
  var widths = [200, 130, 110, 100, 100, 130, 80, 100];
  widths.forEach(function(w, i) { sheet.setColumnWidth(i + 1, w); });

  // 凍結第一列
  sheet.setFrozenRows(1);

  // 條件格式：在倉總箱數 <= 10 變紅底提醒
  var range = sheet.getRange('D2:D51');
  var rule = SpreadsheetApp.newConditionalFormatRule()
    .whenNumberLessThanOrEqualTo(10)
    .setBackground('#fce8e6')
    .setFontColor('#c5221f')
    .setRanges([range])
    .build();
  sheet.setConditionalFormatRules([rule]);
}

// ── Products ──────────────────────────────────────────────────────────────────

function addProduct(d) {
  try {
    var id = nextId_('Products');
    ss_().getSheetByName('Products').appendRow([
      id, d.name, d.sku||'', d.ean||'', d.asin||'',
      d.cubic_feet_per_carton ? parseFloat(d.cubic_feet_per_carton) : '',
      d.qty_per_carton ? parseInt(d.qty_per_carton) : '',
      d.note||'', nowStr_()
    ]);
    return { ok: true, id: id };
  } catch(e) { return { error: e.message }; }
}

function updateProduct(id, d) {
  try {
    var sheet = ss_().getSheetByName('Products');
    var row = findRow_(sheet, id);
    if (row < 0) return { error: 'Product not found' };
    sheet.getRange(row, 1, 1, 8).setValues([[
      id, d.name, d.sku||'', d.ean||'', d.asin||'',
      d.cubic_feet_per_carton ? parseFloat(d.cubic_feet_per_carton) : '',
      d.qty_per_carton ? parseInt(d.qty_per_carton) : '',
      d.note||''
    ]]);
    return { ok: true };
  } catch(e) { return { error: e.message }; }
}

function deleteProduct(id) {
  try {
    var spreadsheet = ss_();
    var pSheet = spreadsheet.getSheetByName('Products');
    var bSheet = spreadsheet.getSheetByName('Batches');
    var mSheet = spreadsheet.getSheetByName('Movements');

    // Collect batch IDs to delete
    var bVals = bSheet.getDataRange().getValues();
    var batchIds = [];
    for (var i = bVals.length - 1; i >= 1; i--) {
      if (String(bVals[i][1]) === String(id)) {
        batchIds.push(String(bVals[i][0]));
        bSheet.deleteRow(i + 1);
      }
    }
    // Delete movements for those batches
    if (batchIds.length > 0) {
      var mVals = mSheet.getDataRange().getValues();
      for (var j = mVals.length - 1; j >= 1; j--) {
        if (batchIds.indexOf(String(mVals[j][1])) >= 0) mSheet.deleteRow(j + 1);
      }
    }
    var pRow = findRow_(pSheet, id);
    if (pRow > 0) pSheet.deleteRow(pRow);
    return { ok: true };
  } catch(e) { return { error: e.message }; }
}

// ── Batches ───────────────────────────────────────────────────────────────────

function addBatch(d) {
  try {
    var id = nextId_('Batches');
    ss_().getSheetByName('Batches').appendRow([
      id, d.product_id, d.exp_date||'', d.location||'AMZLGS',
      parseFloat(d.current_cartons)||0, d.note||'', nowStr_()
    ]);
    return { ok: true, id: id };
  } catch(e) { return { error: e.message }; }
}

function deleteBatch(id) {
  try {
    var spreadsheet = ss_();
    var bSheet = spreadsheet.getSheetByName('Batches');
    var mSheet = spreadsheet.getSheetByName('Movements');
    var row = findRow_(bSheet, id);
    if (row < 0) return { error: 'Batch not found' };
    bSheet.deleteRow(row);
    var mVals = mSheet.getDataRange().getValues();
    for (var i = mVals.length - 1; i >= 1; i--) {
      if (String(mVals[i][1]) === String(id)) mSheet.deleteRow(i + 1);
    }
    return { ok: true };
  } catch(e) { return { error: e.message }; }
}

// ── Movements ─────────────────────────────────────────────────────────────────

function addMovement(d) {
  try {
    var spreadsheet = ss_();
    var mSheet = spreadsheet.getSheetByName('Movements');
    var bSheet = spreadsheet.getSheetByName('Batches');
    var id = nextId_('Movements');

    var outTypes = ['OUT_FBA','OUT_DROPSHIP','REPACK'];
    var qty = parseFloat(d.qty_cartons);
    var signed = outTypes.indexOf(d.type) >= 0 ? -Math.abs(qty) : Math.abs(qty);

    mSheet.appendRow([id, d.batch_id, d.date, d.type, signed, d.note||'', nowStr_()]);

    // Update batch current_cartons
    var bRow = findRow_(bSheet, d.batch_id);
    if (bRow > 0) {
      var cur = parseFloat(bSheet.getRange(bRow, 5).getValue()) || 0;
      bSheet.getRange(bRow, 5).setValue(Math.round((cur + signed) * 100) / 100);
    }
    return { ok: true, id: id, signed_qty: signed };
  } catch(e) { return { error: e.message }; }
}

// ── Billing ───────────────────────────────────────────────────────────────────

function addBilling(d) {
  try {
    var id = nextId_('Billing');
    ss_().getSheetByName('Billing').appendRow([
      id, d.date, (d.type||'').toUpperCase(), d.job_name||'',
      d.qty      ? parseFloat(d.qty)      : '',
      d.unit_fee ? parseFloat(d.unit_fee) : '',
      d.total_fee? parseFloat(d.total_fee): '',
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
