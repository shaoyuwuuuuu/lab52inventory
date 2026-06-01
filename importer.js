const cheerio = require('cheerio');
const fs = require('fs');
const db = require('./db');

function parseDate(str) {
  if (!str) return null;
  str = str.trim().replace(/\s+/g, ' ');
  let m;
  // YYYY/MM/DD or YYYY-MM-DD
  if ((m = str.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/)))
    return `${m[1]}-${m[2].padStart(2,'0')}-${m[3].padStart(2,'0')}`;
  // MM/DD/YYYY
  if ((m = str.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/)))
    return `${m[3]}-${m[1].padStart(2,'0')}-${m[2].padStart(2,'0')}`;
  // M/D/YY (short year)
  if ((m = str.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2})$/))) {
    const yr = parseInt(m[3]) < 50 ? '20'+m[3] : '19'+m[3];
    return `${yr}-${m[1].padStart(2,'0')}-${m[2].padStart(2,'0')}`;
  }
  return null;
}

function parseNum(str) {
  if (!str) return null;
  const n = parseFloat(str.replace(/[$,\s]/g, '').trim());
  return isNaN(n) ? null : n;
}

// HTML 欄位結構（含 freezebar cell）：
// ci=0: product name
// ci=1: SKU
// ci=2: EAN
// ci=3: freezebar (空白)
// ci=4: Carton
// ci=5: Cubic Feet
// ci=6: Quantity (PCS)
// ci=7: EXP date
// ci=8+: date movement columns, then Total/PCS/CUBIE FEET

function importInventory(htmlPath) {
  const html = fs.readFileSync(htmlPath, 'utf8');
  const $ = cheerio.load(html);
  const rows = $('table.waffle tbody tr');

  // 從 header row 讀取日期欄位（從 ci=8 開始）
  const headerRow = rows.eq(0);
  const dateCols = [];  // index maps to ci-8
  headerRow.find('td').each((ci, el) => {
    if (ci >= 8) {
      const label = $(el).text().trim();
      dateCols.push({ ci, date: parseDate(label), label });
    }
  });

  // 找 "Total" 欄的 headerIdx
  const totalIdx = dateCols.findIndex(c => c.label === 'Total');

  const insertProduct = db.prepare(
    `INSERT OR IGNORE INTO products (name,sku,ean,cubic_feet_per_carton,qty_per_carton) VALUES (?,?,?,?,?)`
  );
  const insertBatch = db.prepare(
    `INSERT INTO batches (product_id,exp_date,location,current_cartons,note) VALUES (?,?,?,?,?)`
  );
  const insertMove = db.prepare(
    `INSERT INTO movements (batch_id,date,type,qty_cartons,note) VALUES (?,?,?,?,?)`
  );

  let imported = { products: 0, batches: 0 };

  db.exec('BEGIN');
  try {
    rows.each((ri, row) => {
      if (ri === 0) return; // skip header
      const cells = $(row).find('td');
      if (cells.length < 4) return;

      const name = cells.eq(0).text().trim();
      const sku  = cells.eq(1).text().trim();
      const ean  = cells.eq(2).text().trim();

      // 跳過空行和分節標題
      if (!name) return;
      if (!sku && !ean) return;
      if (['需要Repack', '正貨可入FBA倉'].some(k => name.startsWith(k))) return;

      const cartonStr = cells.eq(4).text().trim();
      const cubicStr  = cells.eq(5).text().trim();
      const qtyStr    = cells.eq(6).text().trim();
      const expStr    = cells.eq(7).text().trim();

      const cartons      = parseNum(cartonStr);
      const cubic        = parseNum(cubicStr);
      const qty          = parseNum(qtyStr);
      const expDate      = parseDate(expStr);
      const qtyPerCarton = (cartons && qty && cartons > 0) ? Math.round(qty / cartons) : null;

      // 找或建產品
      let prod = db.prepare(
        'SELECT id FROM products WHERE sku=? OR (name=? AND ean=?)'
      ).get([sku||'__NONE__', name, ean||'__NONE__']);

      if (!prod) {
        const r = insertProduct.run([name, sku||null, ean||null, cubic||null, qtyPerCarton||null]);
        prod = { id: r.lastInsertRowid };
        imported.products++;
      }

      // 取 Total 欄的現存箱數
      let currentCartons = 0;
      if (totalIdx >= 0) {
        const totalCell = cells.eq(8 + totalIdx);
        currentCartons = parseNum(totalCell?.text()) || 0;
      }

      // 取各日期欄的異動
      const moveEntries = [];
      cells.each((ci, cell) => {
        if (ci < 8) return;
        const headerIdx = ci - 8;
        if (headerIdx >= dateCols.length) return;
        const colInfo = dateCols[headerIdx];
        if (!colInfo.date) return;  // 跳過非日期欄（包括 Total、PCS、CUBIE FEET）
        const val = parseNum($(cell).text());
        if (val !== null && val !== 0) {
          moveEntries.push({ date: colInfo.date, qty: val });
        }
      });

      // 建批次
      const batchRes = insertBatch.run([prod.id, expDate||null, 'AMZLGS', currentCartons, null]);
      const batchId = batchRes.lastInsertRowid;
      imported.batches++;

      // 建歷史異動
      for (const mv of moveEntries) {
        if (!mv.date) continue;
        const type = mv.qty > 0 ? 'IN_ARRIVAL' : 'OUT_SHIP';
        insertMove.run([batchId, mv.date, type, mv.qty, '匯入自Inventory.html']);
      }
    });

    db.exec('COMMIT');
  } catch (e) {
    db.exec('ROLLBACK');
    throw e;
  }

  return imported;
}

function importBilling(htmlPath) {
  const html = fs.readFileSync(htmlPath, 'utf8');
  const $ = cheerio.load(html);
  const rows = $('table.waffle tbody tr');
  let count = 0;

  const ins = db.prepare(
    `INSERT INTO billing (date,type,job_name,qty,unit_fee,total_fee,note) VALUES (?,?,?,?,?,?,?)`
  );

  db.exec('BEGIN');
  try {
    rows.each((ri, row) => {
      const cells = $(row).find('td');
      if (cells.length < 4) return;

      const dateStr  = cells.eq(0).text().trim();
      const type     = cells.eq(1).text().trim().toUpperCase();
      const jobName  = cells.eq(2).text().trim();
      const qtyStr   = cells.eq(3) ? cells.eq(3).text().trim() : '';
      const feeStr   = cells.eq(4) ? cells.eq(4).text().trim() : '';
      const totalStr = cells.eq(5) ? cells.eq(5).text().trim() : '';

      const date = parseDate(dateStr);
      if (!date || !type || type === 'DETAIL') return;

      ins.run([date, type, jobName||null, parseNum(qtyStr), parseNum(feeStr), parseNum(totalStr), null]);
      count++;
    });
    db.exec('COMMIT');
  } catch (e) {
    db.exec('ROLLBACK');
    throw e;
  }

  return count;
}

module.exports = { importInventory, importBilling };
