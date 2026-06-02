const XLSX = require('xlsx');

const wb = XLSX.utils.book_new();

// ── 實際產品資料（從 Inventory.html 解析）────────────────────────────────────
const products = [
  // [id, name, sku, ean, cubic_ft, qty_per_carton, note]
  [1,  'Lab52 Kids Oral Spray (Strawberry) 20mL', 'Stawberry_Spray_Kids', '4714781022217', 1.08, 144, ''],
  [2,  'Lab52 Kids Oral Spray (Peach) 20mL',       'Peach_Spray_Kids',    '4714781022255', 1.08, 144, ''],
  [3,  'Lab52 Kids Oral Spray (Vanilla) 20mL',     'Milk_Kids_Spray',     '4714781022231', 1.08, 144, ''],
  [4,  'Lab52 Kids Oral Spray (Grape) 20mL',       'Grape_Spray_Kids',    '4714781022224', 1.08, 144, ''],
  [5,  'Lab52 Kids Oral Spray (Watermelon) 20mL',  'Watermelon-spray',    '4714781020022', 1.08, 144, ''],
  [6,  'Lab52 Kids Oral Spray (Flavor Free) 20mL', 'FlavorFree_Spray_KIDS','4714781022248',1.08, 144, ''],
  [7,  'Lab52 Kids Oral Spray (Peach) 20mL NEW',   'PEACH-SPRAY',         '4714781020756', 1.08, 144, '新包裝'],
  [8,  'Lab52 Probiotic Oral Spray (Apple Mint)',   'AD-0001',             '4714781020084', 1.08, 144, ''],
  [9,  'Lab52 Kids Oral Spray Duo (Straw+Grape)',   'D6-1MWN-Q0EN',        '4714781023580', 1.08,  36, ''],
  [10, 'Lab52 Kids Oral Spray Duo (Straw+Vanilla)', 'M2-6V1S-09J4',        '4714781023603', 1.08,  36, ''],
  [11, 'Lab52 Kids Oral Spray Duo (Straw+Peach)',   'TU-V4OB-07QB',        '4714781023597', 1.08,  36, ''],
  [12, 'Lab52 Kids Oral Spray (Grape Trio)',        'F8-1THT-CLFU',        '4714781023115', 2.19,  60, ''],
  [13, 'Lab52 nHAP 無氟牙膏 (Strawberry)',           'StrawberryNHAP',       '4714781020060', 0.782, 63, ''],
  [14, 'Lab52 nHAP 無氟牙膏 (Watermelon)',           'WatermelonNHAP',       '4714781020077', 0.782, 63, ''],
  [15, 'PACKAGE BOX (Strawberry Trio)',             'PACKAGE BOX',         'N/A',           '',    60, '包裝材料'],
];

// ── 實際批次資料（從 Inventory.html 解析，截至 2026/05/29）──────────────────
const batches = [
  // [id, product_id, exp_date, location, current_cartons, note]
  [1,  1,  '2028-05-08', 'AMZLGS', 88,   ''],
  [2,  2,  '2028-05-07', 'AMZLGS', 14,   ''],
  [3,  3,  '2028-05-07', 'AMZLGS', 0,    '已出完'],
  [4,  3,  '2028-08-06', 'AMZLGS', 63,   ''],
  [5,  4,  '2028-09-19', 'AMZLGS', 31,   ''],
  [6,  5,  '2028-07-16', 'AMZLGS', 104,  ''],
  [7,  6,  '2028-07-23', 'AMZLGS', 38,   ''],
  [8,  7,  '2028-09-04', 'AMZLGS', 26,   '新包裝'],
  [9,  8,  '2028-07-01', 'AMZLGS', 0,    '已出完'],
  [10, 9,  '2028-06-04', 'AMZLGS', 9,    ''],
  [11, 9,  '2028-06-04', 'AMZLGS', 31,   '第二批'],
  [12, 10, '2027-05-01', 'AMZLGS', 15,   ''],
  [13, 11, '2027-11-27', 'AMZLGS', 0,    '已出完'],
  [14, 12, '2028-09-19', 'AMZLGS', 0,    '已出完'],
  [15, 13, '2029-01-08', 'AMZLGS', 48,   '新入庫'],
  [16, 14, '2029-01-07', 'AMZLGS', 47,   '新入庫'],
  [17, 15, '',           'AMZLGS', 7,    ''],
];

// ── 帳單範例（MAY 2026，從 MAY 26.html 節錄重點）────────────────────────────
const billing = [
  // [id, date, type, job_name, qty, unit_fee, total_fee, note]
  [1,  '2026-05-01', 'STORAGE',          '',                                    804.244, 0.80, 20.75,  ''],
  [2,  '2026-05-05', 'RESTICKERS',       'STRAWBERRY 10 BOX/144 PCS',           1440,    0.35, 504.00, ''],
  [3,  '2026-05-05', 'REPACK',           'STRAWBERRY GIFT SET 420PCS',          420,     1.00, 420.00, ''],
  [4,  '2026-05-05', 'POSTAGE',          'Elizabeth Vasquez',                   1,       6.27, 6.27,   ''],
  [5,  '2026-05-05', 'DROPSHIP',         'Elizabeth Vasquez',                   1,       4.00, 4.00,   'WITH BOX'],
  [6,  '2026-05-08', 'FBA',              '60 BOXES',                            60,      2.00, 120.00, ''],
  [7,  '2026-05-14', 'CHECK IN',         '',                                    1,       10.00,10.00,  ''],
  [8,  '2026-05-15', 'FBA',              'APPLE MINT *10',                      10,      2.00, 20.00,  ''],
  [9,  '2026-05-29', 'FBA',              'SWF2(32)+PSP3(30)+IND9(34)+ONT8(34)+TEB9(35)', 165, 2.00, 330.00, ''],
  [10, '2026-05-29', 'PALLET FEE',       '2 Pallets',                           2,       10.00,20.00,  ''],
  [11, '2026-05-31', 'STORAGE',          '',                                    610.924, 0.80, 15.77,  '月底結算'],
  // 月小計行（空資料，用公式算）
];

// ── 異動範例 ─────────────────────────────────────────────────────────────────
const movements = [
  // [id, batch_id, date, type, qty_cartons, note]
  [1,  1,  '2026-01-15', 'IN',          88,   '1/15 盤點入庫'],
  [2,  2,  '2026-01-15', 'IN',          16,   '1/15 盤點入庫'],
  [3,  2,  '2026-01-30', 'OUT_FBA',     -2,   'FBA 補貨'],
  [4,  5,  '2026-01-28', 'IN',          35,   '新批次入庫'],
  [5,  5,  '2026-02-05', 'OUT_FBA',     -20,  'FBA 補貨'],
  [6,  6,  '2026-01-15', 'IN',          15,   ''],
  [7,  6,  '2026-01-22', 'OUT_FBA',     -5,   ''],
  [8,  6,  '2026-01-28', 'IN',          50,   '新到貨'],
  [9,  6,  '2026-03-09', 'IN',          4,    '退貨'],
  [10, 10, '2026-01-28', 'IN',          55,   ''],
  [11, 10, '2026-02-05', 'OUT_FBA',     -30,  ''],
  [12, 10, '2026-05-08', 'OUT_DROPSHIP',-10,  ''],
  [13, 15, '2026-04-21', 'IN',          48,   '新入庫 草莓牙膏'],
  [14, 16, '2026-04-21', 'IN',          47,   '新入庫 西瓜牙膏'],
  [15, 1,  '2026-05-29', 'OUT_FBA',     -5,   ''],
  [16, 6,  '2026-05-29', 'OUT_FBA',     -104, '西瓜出庫暫記'],
];

// ── Helper ────────────────────────────────────────────────────────────────────
function makeSheet(headers, rows, colWidths) {
  const data = [headers, ...rows];
  const ws   = XLSX.utils.aoa_to_sheet(data);
  ws['!cols'] = colWidths.map(w => ({ wch: w }));
  ws['!freeze'] = { xSplit: 0, ySplit: 1 };
  return ws;
}

// ── Sheet 1: 庫存總覽（公式版）───────────────────────────────────────────────
(function() {
  const headers = [
    '產品名稱', 'SKU', 'EAN', '在倉總箱數', '估算總pcs', '估算材積(ft³)',
    '最近效期', '最近效期批剩餘', '每箱pcs', '材積(ft³/箱)', '備注', '庫存狀態'
  ];
  const rows = [];
  for (let r = 2; r <= products.length + 1; r++) {
    const pr = `Products!A${r}`;
    rows.push([
      `=IFERROR(IF(${pr}="","",VLOOKUP(${pr},Products!A:B,2,0)),"")`,
      `=IFERROR(IF(${pr}="","",VLOOKUP(${pr},Products!A:C,3,0)),"")`,
      `=IFERROR(IF(${pr}="","",VLOOKUP(${pr},Products!A:D,4,0)),"")`,
      `=IFERROR(IF(${pr}="","",SUMIF(Batches!B:B,${pr},Batches!E:E)),"")`,
      `=IFERROR(IF(D${r}="","",D${r}*I${r}),"")`,
      `=IFERROR(IF(D${r}="","",D${r}*J${r}),"")`,
      `=IFERROR(IF(${pr}="","",TEXT(MINIFS(Batches!C:C,Batches!B:B,${pr},Batches!E:E,">"&0),"yyyy-mm-dd")),"")`,
      `=IFERROR(IF(${pr}="","",SUMIFS(Batches!E:E,Batches!B:B,${pr},Batches!C:C,MINIFS(Batches!C:C,Batches!B:B,${pr},Batches!E:E,">"&0))),"")`,
      `=IFERROR(IF(${pr}="","",VLOOKUP(${pr},Products!A:G,6,0)),"")`,
      `=IFERROR(IF(${pr}="","",VLOOKUP(${pr},Products!A:F,5,0)),"")`,
      `=IFERROR(IF(${pr}="","",VLOOKUP(${pr},Products!A:H,8,0)),"")`,
      `=IF(D${r}="","",IF(D${r}<=5,"🔴 庫存偏低",IF(D${r}<=15,"🟡 注意","")))`,
    ]);
  }
  const ws = makeSheet(headers, rows,
    [32, 20, 16, 12, 12, 14, 12, 14, 10, 12, 14, 12]);
  ws['!freeze'] = { xSplit: 0, ySplit: 1 };
  XLSX.utils.book_append_sheet(wb, ws, '庫存總覽');
})();

// ── Sheet 2: Products ─────────────────────────────────────────────────────────
(function() {
  const headers = ['id','name','sku','ean','cubic_feet_per_carton','qty_per_carton','asin','note','created_at'];
  const rows = products.map(p => [
    p[0], p[1], p[2], p[3], p[4], p[5], '', p[6], '2026-06-01'
  ]);
  XLSX.utils.book_append_sheet(wb,
    makeSheet(headers, rows, [6,36,22,16,18,14,14,16,14]),
    'Products');
})();

// ── Sheet 3: Batches ──────────────────────────────────────────────────────────
(function() {
  const headers = ['id','product_id','exp_date','location','current_cartons','note','created_at'];
  const rows = batches.map(b => [
    b[0], b[1], b[2], b[3], b[4], b[5], '2026-06-01'
  ]);
  XLSX.utils.book_append_sheet(wb,
    makeSheet(headers, rows, [6,10,12,10,14,20,14]),
    'Batches');
})();

// ── Sheet 4: Movements ────────────────────────────────────────────────────────
(function() {
  const headers = ['id','batch_id','date','type','qty_cartons','note','created_at'];
  const rows = movements.map(m => [
    m[0], m[1], m[2], m[3], m[4], m[5], '2026-06-01'
  ]);
  XLSX.utils.book_append_sheet(wb,
    makeSheet(headers, rows, [6,10,12,16,12,28,14]),
    'Movements');
})();

// ── Sheet 5: Billing ──────────────────────────────────────────────────────────
(function() {
  const headers = ['id','date','type','job_name','qty','unit_fee','total_fee','note'];
  const rows = billing.map(b => [b[0],b[1],b[2],b[3],b[4],b[5],b[6],b[7]]);

  // 月小計公式列
  const sumRow = XLSX.utils.encode_row(rows.length + 1);
  rows.push(['', '', '', '📌 MAY 2026 合計', '', '',
    `=SUMIF(B2:B${rows.length + 1},"2026-05*",G2:G${rows.length + 1})`, '']);

  const ws = makeSheet(headers, rows, [6,12,16,36,10,10,12,24]);
  XLSX.utils.book_append_sheet(wb, ws, 'Billing');
})();

// ── Sheet 6: 帳單格式說明 ─────────────────────────────────────────────────────
(function() {
  const data = [
    ['欄位', '說明', '常見值'],
    ['date',      '日期 (YYYY-MM-DD)',        '2026-05-01'],
    ['type',      '費用類型',                 'STORAGE / FBA / DROPSHIP / POSTAGE / REPACK / RESTICKERS / CHECK IN / PALLET FEE'],
    ['job_name',  '工作名稱 / 客戶名稱',       '60 BOXES / Elizabeth Vasquez'],
    ['qty',       '數量（材積 ft³ 或件數）',   '804.244 / 60 / 1'],
    ['unit_fee',  '單價 ($)',                  '0.80 / 2.00 / 6.27'],
    ['total_fee', '總費用 ($)',                '20.75 / 120.00'],
    ['note',      '備注',                      'WITH BOX / 9300110990...'],
  ];
  const ws = XLSX.utils.aoa_to_sheet(data);
  ws['!cols'] = [{wch:14},{wch:28},{wch:64}];
  XLSX.utils.book_append_sheet(wb, ws, '帳單欄位說明');
})();

const outPath = 'Lab52_外倉庫存_範本.xlsx';
XLSX.writeFile(wb, outPath);
console.log('✅ 已建立：' + outPath);
console.log('   分頁：庫存總覽 / Products / Batches / Movements / Billing / 帳單欄位說明');
