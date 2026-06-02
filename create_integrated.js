const XLSX = require('xlsx');
const wb   = XLSX.utils.book_new();

// ── 產品資料（Reference）────────────────────────────────────────────────────
const productData = [
  // [SKU, 產品名稱, EAN, 材積ft³/箱, 每箱pcs]
  ['Stawberry_Spray_Kids', 'Kids Oral Spray (草莓) 20mL',         '4714781022217', 1.08,  144],
  ['Peach_Spray_Kids',     'Kids Oral Spray (水蜜桃) 20mL',        '4714781022255', 1.08,  144],
  ['Milk_Kids_Spray',      'Kids Oral Spray (香草) 20mL',          '4714781022231', 1.08,  144],
  ['Grape_Spray_Kids',     'Kids Oral Spray (葡萄) 20mL',          '4714781022224', 1.08,  144],
  ['Watermelon-spray',     'Kids Oral Spray (西瓜) 20mL',          '4714781020022', 1.08,  144],
  ['FlavorFree_Spray_KIDS','Kids Oral Spray (無口味) 20mL',        '4714781022248', 1.08,  144],
  ['PEACH-SPRAY',          'Kids Oral Spray (水蜜桃) 20mL 新包裝', '4714781020756', 1.08,  144],
  ['AD-0001',              'Probiotic Oral Spray (Apple Mint)',    '4714781020084', 1.08,  144],
  ['D6-1MWN-Q0EN',         'Kids Oral Spray Duo (草莓+葡萄)',       '4714781023580', 1.08,   36],
  ['M2-6V1S-09J4',         'Kids Oral Spray Duo (草莓+香草)',       '4714781023603', 1.08,   36],
  ['TU-V4OB-07QB',         'Kids Oral Spray Duo (草莓+水蜜桃)',     '4714781023597', 1.08,   36],
  ['F8-1THT-CLFU',         'Kids Oral Spray (葡萄 Trio)',           '4714781023115', 2.19,   60],
  ['StrawberryNHAP',       'nHAP 無氟牙膏 (草莓)',                  '4714781020060', 0.782,  63],
  ['WatermelonNHAP',       'nHAP 無氟牙膏 (西瓜)',                  '4714781020077', 0.782,  63],
  ['PACKAGE BOX',          'Package Box (Strawberry Trio)',        'N/A',           '',     60],
];

// ── 進出記錄（實際資料截至 2026/05/29）──────────────────────────────────────
// 欄位: [日期, SKU, 效期, 類型, 箱數(+入/-出), 備注]
const movements = [
  ['2026-01-15', 'Stawberry_Spray_Kids', '2028-05-08', 'IN',          88,    '1/15 盤點入庫'],
  ['2026-01-15', 'Peach_Spray_Kids',     '2028-05-07', 'IN',          16,    '1/15 盤點入庫'],
  ['2026-01-22', 'Watermelon-spray',     '2028-07-16', 'OUT_FBA',     -5,    ''],
  ['2026-01-28', 'Grape_Spray_Kids',     '2028-09-19', 'IN',          35,    '新批次入庫'],
  ['2026-01-28', 'D6-1MWN-Q0EN',        '2028-06-04', 'IN',          54,    ''],
  ['2026-01-28', 'Watermelon-spray',     '2028-07-16', 'IN',          50,    ''],
  ['2026-01-30', 'Milk_Kids_Spray',      '2028-08-06', 'IN',          40,    ''],
  ['2026-02-05', 'Grape_Spray_Kids',     '2028-09-19', 'OUT_FBA',     -20,   ''],
  ['2026-02-05', 'D6-1MWN-Q0EN',        '2028-06-04', 'OUT_FBA',     -45,   ''],
  ['2026-02-05', 'M2-6V1S-09J4',        '2027-05-01', 'IN',          55,    ''],
  ['2026-02-05', 'M2-6V1S-09J4',        '2027-05-01', 'OUT_FBA',     -30,   ''],
  ['2026-03-09', 'Grape_Spray_Kids',     '2028-09-19', 'IN',          2,     '退貨'],
  ['2026-03-09', 'Watermelon-spray',     '2028-07-16', 'IN',          4,     '退貨'],
  ['2026-03-18', 'Milk_Kids_Spray',      '2028-05-07', 'IN',          15,    '新批次入庫'],
  ['2026-03-18', 'Milk_Kids_Spray',      '2028-05-07', 'OUT_FBA',     -15,   ''],
  ['2026-03-18', 'D6-1MWN-Q0EN',        '2028-06-04', 'IN',          10,    '第二批'],
  ['2026-03-18', 'FlavorFree_Spray_KIDS','2028-07-23', 'IN',          17,    ''],
  ['2026-03-18', 'Peach_Spray_Kids',     '2028-05-07', 'OUT_FBA',     -2,    ''],
  ['2026-04-21', 'StrawberryNHAP',       '2029-01-08', 'IN',          48,    '新入庫牙膏'],
  ['2026-04-21', 'WatermelonNHAP',       '2029-01-07', 'IN',          47,    '新入庫牙膏'],
  ['2026-04-29', 'D6-1MWN-Q0EN',        '2028-06-04', 'OUT_FBA',     -10,   ''],
  ['2026-04-29', 'Milk_Kids_Spray',      '2028-08-06', 'IN',          41,    ''],
  ['2026-04-29', 'M2-6V1S-09J4',        '2027-05-01', 'IN',          34,    ''],
  ['2026-04-29', 'PEACH-SPRAY',          '2028-09-04', 'IN',          35,    '新包裝入庫'],
  ['2026-04-29', 'FlavorFree_Spray_KIDS','2028-07-23', 'OUT_FBA',     -5,    ''],
  ['2026-04-29', 'Watermelon-spray',     '2028-07-16', 'IN',          82,    ''],
  ['2026-05-08', 'Grape_Spray_Kids',     '2028-09-19', 'IN',          34,    ''],
  ['2026-05-08', 'Watermelon-spray',     '2028-07-16', 'OUT_FBA',     -60,   ''],
  ['2026-05-08', 'M2-6V1S-09J4',        '2027-05-01', 'OUT_DROPSHIP',-10,   ''],
  ['2026-05-14', 'FlavorFree_Spray_KIDS','2028-07-23', 'IN',          31,    ''],
  ['2026-05-14', 'StrawberryNHAP',       '2029-01-08', 'IN',          42,    '新到'],
  ['2026-05-14', 'WatermelonNHAP',       '2029-01-07', 'IN',          42,    '新到'],
  ['2026-05-14', 'Watermelon-spray',     '2028-07-16', 'IN',          50,    ''],
  ['2026-05-15', 'FlavorFree_Spray_KIDS','2028-07-23', 'OUT_FBA',     -5,    ''],
  ['2026-05-15', 'Stawberry_Spray_Kids', '2028-07-17', 'IN',          82,    '新批次'],
  ['2026-05-29', 'Stawberry_Spray_Kids', '2028-05-08', 'OUT_FBA',     -5,    'FBA出貨'],
  ['2026-05-29', 'Grape_Spray_Kids',     '2028-09-19', 'OUT_FBA',     -20,   'FBA出貨'],
  ['2026-05-29', 'Milk_Kids_Spray',      '2028-08-06', 'OUT_FBA',     -15,   'FBA出貨'],
  ['2026-05-29', 'D6-1MWN-Q0EN',        '2028-06-04', 'OUT_FBA',     -5,    ''],
  ['2026-05-29', 'M2-6V1S-09J4',        '2027-05-01', 'OUT_FBA',     -31,   ''],
];

// ── 帳單（MAY 2026 節錄）────────────────────────────────────────────────────
const billing = [
  ['2026-05-01', 'STORAGE',       '',                                    804.244, 0.80,  20.75,  ''],
  ['2026-05-05', 'RESTICKERS',    'STRAWBERRY 10 BOX/144 PCS',           1440,    0.35,  504.00, ''],
  ['2026-05-05', 'REPACK',        'STRAWBERRY GIFT SET 420PCS',          420,     1.00,  420.00, ''],
  ['2026-05-05', 'POSTAGE',       'Elizabeth Vasquez',                   1,       6.27,  6.27,   ''],
  ['2026-05-05', 'DROPSHIP',      'Elizabeth Vasquez',                   1,       4.00,  4.00,   'WITH BOX'],
  ['2026-05-08', 'FBA',           '60 BOXES',                            60,      2.00,  120.00, ''],
  ['2026-05-14', 'CHECK IN',      '',                                    1,       10.00, 10.00,  ''],
  ['2026-05-15', 'FBA',           'APPLE MINT *10',                      10,      2.00,  20.00,  ''],
  ['2026-05-29', 'FBA',           'SWF2+PSP3+IND9+ONT8+TEB9',           165,     2.00,  330.00, ''],
  ['2026-05-29', 'PALLET FEE',    '2 Pallets',                           2,       10.00, 20.00,  ''],
  ['2026-05-31', 'STORAGE',       '',                                    610.924, 0.80,  15.77,  '月底結算'],
];

// ── Helper ────────────────────────────────────────────────────────────────────
function addSheet(name, headers, rows, colWidths) {
  const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
  ws['!cols']   = colWidths.map(w => ({ wch: w }));
  ws['!freeze'] = { xSplit: 0, ySplit: 1, state: 'frozen' };
  XLSX.utils.book_append_sheet(wb, ws, name);
}

// ── Sheet 1: 庫存總覽 (公式全部從「進出記錄」自動拉)────────────────────────
(function() {
  const headers = [
    '產品名稱', 'SKU', 'EAN(條碼)', '現存箱數', '估算pcs', '估算材積(ft³)',
    '最近效期', '每箱pcs', '材積(ft³/箱)', '狀態'
  ];
  // B欄放 SKU，其他欄用公式從「進出記錄」和「產品資料」拉
  const skuList = productData.map(p => p[0]);
  const rows = skuList.map((sku, idx) => {
    const r = idx + 2;
    return [
      // A: 產品名稱（從產品資料 VLOOKUP）
      `=IFERROR(VLOOKUP(B${r},產品資料!A:B,2,0),"")`,
      // B: SKU（直接填）
      sku,
      // C: EAN
      `=IFERROR(VLOOKUP(B${r},產品資料!A:C,3,0),"")`,
      // D: 現存箱數（SUMIF 全部進出）
      `=IFERROR(SUMIF(進出記錄!B:B,B${r},進出記錄!E:E),0)`,
      // E: 估算 pcs
      `=IF(D${r}=0,"",D${r}*H${r})`,
      // F: 估算材積
      `=IF(D${r}=0,"",ROUND(D${r}*I${r},2))`,
      // G: 最近效期（有庫存的最早效期）
      `=IFERROR(TEXT(MINIFS(進出記錄!C:C,進出記錄!B:B,B${r},進出記錄!C:C,">"&0),"yyyy-mm-dd"),"")`,
      // H: 每箱 pcs（從產品資料）
      `=IFERROR(VLOOKUP(B${r},產品資料!A:E,5,0),"")`,
      // I: 材積
      `=IFERROR(VLOOKUP(B${r},產品資料!A:D,4,0),"")`,
      // J: 狀態
      `=IF(D${r}=0,"",IF(D${r}<=5,"🔴 庫存偏低",IF(D${r}<=15,"🟡 注意","✅ 正常")))`,
    ];
  });

  const ws = XLSX.utils.aoa_to_sheet([headers, ...rows]);
  ws['!cols']   = [32,22,16,12,12,14,12,10,12,12].map(w=>({wch:w}));
  ws['!freeze'] = { xSplit: 0, ySplit: 1, state: 'frozen' };
  XLSX.utils.book_append_sheet(wb, ws, '庫存總覽');
})();

// ── Sheet 2: 進出記錄 (你每次登入出的地方) ───────────────────────────────────
addSheet(
  '進出記錄',
  ['日期', 'SKU', '效期', '類型', '箱數(+入/-出)', '備注'],
  movements,
  [13, 22, 12, 16, 14, 30]
);

// ── Sheet 3: 產品資料 (Reference，不用常改) ──────────────────────────────────
addSheet(
  '產品資料',
  ['SKU', '產品名稱', 'EAN(條碼)', '材積(ft³/箱)', '每箱pcs', '備注'],
  productData.map(p => [...p, '']),
  [22, 32, 16, 14, 10, 20]
);

// ── Sheet 4: Billing (帳單) ───────────────────────────────────────────────────
(function() {
  const headers = ['日期', '類型', '工作名稱/客戶', '數量', '單價($)', '金額($)', '備注'];
  const lastDataRow = billing.length + 1;
  const rows = [
    ...billing,
    // 月合計列（公式）
    ['', '📌 MAY 2026 合計', '', '', '',
     `=SUMIF(A2:A${lastDataRow},"2026-05*",F2:F${lastDataRow})`, ''],
  ];
  addSheet('Billing', headers, rows, [13, 16, 36, 10, 10, 12, 24]);
})();

// ── 輸出 ──────────────────────────────────────────────────────────────────────
const outPath = 'Lab52_外倉庫存_整合版.xlsx';
XLSX.writeFile(wb, outPath);
console.log('✅ 已建立：' + outPath);
console.log('   庫存總覽公式全部從「進出記錄」自動拉，只要在進出記錄登資料即可');
