const express = require('express');
const path = require('path');
const methodOverride = require('method-override');
const db = require('./db');
const { importInventory, importBilling } = require('./importer');

const app = express();
const PORT = 3737;

app.set('view engine', 'ejs');
app.set('views', path.join(__dirname, 'views'));
app.use(express.static(path.join(__dirname, 'public')));
app.use(express.urlencoded({ extended: true }));
app.use(methodOverride('_method'));

// ─── Dashboard ───────────────────────────────────────────────────────────────
app.get('/', (req, res) => {
  const stats = db.prepare(`
    SELECT COUNT(DISTINCT p.id) as product_count,
           COUNT(b.id) as batch_count,
           SUM(b.current_cartons) as total_cartons
    FROM products p LEFT JOIN batches b ON b.product_id = p.id
  `).get();

  const expiringSoon = db.prepare(`
    SELECT p.name, p.sku, b.id as batch_id, b.exp_date, b.current_cartons, b.location,
           p.qty_per_carton,
           CAST(julianday(b.exp_date) - julianday('now') AS INTEGER) as days_left
    FROM batches b JOIN products p ON p.id = b.product_id
    WHERE b.exp_date IS NOT NULL AND b.exp_date <= date('now', '+180 days')
      AND b.current_cartons > 0
    ORDER BY b.exp_date ASC
    LIMIT 10
  `).all();

  const recentMoves = db.prepare(`
    SELECT m.date, m.type, m.qty_cartons, m.note,
           p.name as product_name, b.exp_date
    FROM movements m
    JOIN batches b ON b.id = m.batch_id
    JOIN products p ON p.id = b.product_id
    ORDER BY m.date DESC, m.id DESC
    LIMIT 15
  `).all();

  res.render('dashboard', { stats, expiringSoon, recentMoves });
});

// ─── Products ────────────────────────────────────────────────────────────────
app.get('/products', (req, res) => {
  const products = db.prepare(`
    SELECT p.*,
           COUNT(b.id) as batch_count,
           SUM(b.current_cartons) as total_cartons,
           MIN(CASE WHEN b.current_cartons > 0 AND b.exp_date IS NOT NULL THEN b.exp_date END) as earliest_exp,
           CAST(
             julianday(MIN(CASE WHEN b.current_cartons > 0 AND b.exp_date IS NOT NULL THEN b.exp_date END))
             - julianday('now') AS INTEGER
           ) as days_to_exp
    FROM products p LEFT JOIN batches b ON b.product_id = p.id
    GROUP BY p.id
    ORDER BY COALESCE(p.sort_order, 999), p.name
  `).all();
  const allMovements = db.prepare(`
    SELECT b.product_id, m.date, m.type, m.qty_cartons, m.note, b.exp_date
    FROM movements m JOIN batches b ON b.id = m.batch_id
    ORDER BY b.product_id, m.date DESC, m.id DESC
  `).all();
  const movementsByProduct = {};
  allMovements.forEach(m => {
    if (!movementsByProduct[m.product_id]) movementsByProduct[m.product_id] = [];
    if (movementsByProduct[m.product_id].length < 50) movementsByProduct[m.product_id].push(m);
  });

  const allBatches = db.prepare(`
    SELECT product_id, exp_date, location, current_cartons
    FROM batches WHERE current_cartons > 0
    ORDER BY product_id, exp_date ASC
  `).all();
  const batchesByProduct = {};
  allBatches.forEach(b => {
    if (!batchesByProduct[b.product_id]) batchesByProduct[b.product_id] = [];
    batchesByProduct[b.product_id].push(b);
  });

  res.render('products', { products, movementsByProduct, batchesByProduct });
});

app.get('/products/new', (req, res) => {
  res.render('product_form', { product: null, error: null });
});

app.post('/products', (req, res) => {
  const { name, sku, ean, asin, cubic_feet_per_carton, qty_per_carton, note } = req.body;
  if (!name) return res.render('product_form', { product: req.body, error: '產品名稱必填' });
  db.prepare(`INSERT INTO products (name,sku,ean,asin,cubic_feet_per_carton,qty_per_carton,note) VALUES (?,?,?,?,?,?,?)`)
    .run([name, sku||null, ean||null, asin||null, cubic_feet_per_carton||null, qty_per_carton||null, note||null]);
  res.redirect('/products');
});

app.get('/products/:id/detail.json', (req, res) => {
  const product = db.prepare('SELECT * FROM products WHERE id=?').get([req.params.id]);
  if (!product) return res.status(404).json({ error: 'not found' });
  const movements = db.prepare(`
    SELECT m.date, m.type, m.qty_cartons, m.note, b.exp_date
    FROM movements m JOIN batches b ON b.id = m.batch_id
    WHERE b.product_id=? ORDER BY m.date DESC, m.id DESC LIMIT 50
  `).all([product.id]);
  res.json({ movements });
});

app.get('/products/:id', (req, res) => {
  const product = db.prepare('SELECT * FROM products WHERE id=?').get([req.params.id]);
  if (!product) return res.redirect('/products');
  const batches = db.prepare(`
    SELECT b.*,
           (SELECT SUM(qty_cartons) FROM movements WHERE batch_id=b.id) as total_moved
    FROM batches b WHERE b.product_id=? ORDER BY b.exp_date ASC
  `).all([product.id]);
  const movements = db.prepare(`
    SELECT m.*, b.exp_date
    FROM movements m JOIN batches b ON b.id = m.batch_id
    WHERE b.product_id=? ORDER BY m.date DESC, m.id DESC LIMIT 50
  `).all([product.id]);
  res.render('product_detail', { product, batches, movements });
});


app.get('/products/:id/edit', (req, res) => {
  const product = db.prepare('SELECT * FROM products WHERE id=?').get([req.params.id]);
  if (!product) return res.redirect('/products');
  res.render('product_form', { product, error: null });
});

app.post('/products/:id/edit', (req, res) => {
  const { name, sku, ean, asin, cubic_feet_per_carton, qty_per_carton, note } = req.body;
  if (!name) {
    const product = db.prepare('SELECT * FROM products WHERE id=?').get([req.params.id]);
    return res.render('product_form', { product: { ...product, ...req.body }, error: '產品名稱必填' });
  }
  db.prepare(`UPDATE products SET name=?,sku=?,ean=?,asin=?,cubic_feet_per_carton=?,qty_per_carton=?,note=? WHERE id=?`)
    .run([name, sku||null, ean||null, asin||null, cubic_feet_per_carton||null, qty_per_carton||null, note||null, req.params.id]);
  res.redirect(`/products/${req.params.id}`);
});

app.post('/products/:id/delete', (req, res) => {
  db.prepare('DELETE FROM products WHERE id=?').run([req.params.id]);
  res.redirect('/products');
});

// ─── Batches ─────────────────────────────────────────────────────────────────
app.post('/products/:id/batches', (req, res) => {
  const { exp_date, location, current_cartons, note } = req.body;
  db.prepare(`INSERT INTO batches (product_id,exp_date,location,current_cartons,note) VALUES (?,?,?,?,?)`)
    .run([req.params.id, exp_date||null, location||'AMZLGS', parseFloat(current_cartons)||0, note||null]);
  res.redirect(`/products/${req.params.id}`);
});

app.post('/batches/:id/delete', (req, res) => {
  const batch = db.prepare('SELECT product_id FROM batches WHERE id=?').get([req.params.id]);
  db.prepare('DELETE FROM batches WHERE id=?').run([req.params.id]);
  res.redirect(batch ? `/products/${batch.product_id}` : '/products');
});

// ─── Movements ───────────────────────────────────────────────────────────────
app.get('/movements/new', (req, res) => {
  const products = db.prepare(`
    SELECT p.*, b.id as batch_id, b.exp_date, b.current_cartons, b.location
    FROM products p JOIN batches b ON b.product_id = p.id
    ORDER BY p.name, b.exp_date
  `).all();
  res.render('movement_add', { products, prefill: req.query });
});

app.post('/movements', (req, res) => {
  const { batch_id, date, type, qty_cartons, note } = req.body;
  const qty = parseFloat(qty_cartons);
  if (!batch_id || !date || !type || isNaN(qty)) return res.redirect('/movements/new');

  const outTypes = ['OUT_FBA','OUT_DROPSHIP','REPACK'];
  const signedQty = outTypes.includes(type) ? -Math.abs(qty) : Math.abs(qty);

  db.prepare(`INSERT INTO movements (batch_id,date,type,qty_cartons,note) VALUES (?,?,?,?,?)`)
    .run([batch_id, date, type, signedQty, note||null]);
  db.prepare(`UPDATE batches SET current_cartons = current_cartons + ? WHERE id=?`)
    .run([signedQty, batch_id]);

  const batch = db.prepare('SELECT product_id FROM batches WHERE id=?').get([batch_id]);
  res.redirect(batch ? `/products/${batch.product_id}` : '/');
});

// ─── Billing ─────────────────────────────────────────────────────────────────
app.get('/billing', (req, res) => {
  const month = req.query.month || '';
  const months = db.prepare(`SELECT DISTINCT substr(date,1,7) as ym FROM billing ORDER BY ym DESC`).all();
  let rows, monthTotal;

  if (month) {
    rows = db.prepare(`SELECT * FROM billing WHERE substr(date,1,7)=? ORDER BY date, id`).all([month]);
    monthTotal = db.prepare(`SELECT SUM(total_fee) as total FROM billing WHERE substr(date,1,7)=?`).get([month]);
  } else {
    rows = db.prepare(`SELECT * FROM billing ORDER BY date DESC, id DESC LIMIT 100`).all();
    monthTotal = null;
  }

  res.render('billing', { rows, months, selectedMonth: month, monthTotal });
});

app.get('/billing/new', (req, res) => {
  res.render('billing_add', { prefill: req.query });
});

app.post('/billing', (req, res) => {
  const { date, type, job_name, qty, unit_fee, total_fee, note } = req.body;
  db.prepare(`INSERT INTO billing (date,type,job_name,qty,unit_fee,total_fee,note) VALUES (?,?,?,?,?,?,?)`)
    .run([date, type?.toUpperCase(), job_name||null,
         qty?parseFloat(qty):null, unit_fee?parseFloat(unit_fee):null,
         total_fee?parseFloat(total_fee):null, note||null]);
  res.redirect('/billing');
});

app.post('/billing/:id/delete', (req, res) => {
  db.prepare('DELETE FROM billing WHERE id=?').run([req.params.id]);
  res.redirect('/billing');
});

// ─── Import ──────────────────────────────────────────────────────────────────
app.get('/import', (req, res) => {
  const dataDir = path.join(__dirname, 'AMZLGS_extracted');
  const fbaDir  = path.join(__dirname, 'FBA_extracted');
  res.render('import_page', { dataDir, fbaDir, message: null, error: null });
});

app.post('/import/inventory', (req, res) => {
  const htmlPath = path.join(__dirname, 'AMZLGS_extracted', 'Inventory.html');
  try {
    const result = importInventory(htmlPath);
    res.render('import_page', {
      dataDir: path.join(__dirname, 'AMZLGS_extracted'),
      fbaDir:  path.join(__dirname, 'FBA_extracted'),
      message: `匯入完成：新增 ${result.products} 個產品、${result.batches} 個批次`,
      error: null
    });
  } catch (e) {
    res.render('import_page', {
      dataDir: path.join(__dirname, 'AMZLGS_extracted'),
      fbaDir:  path.join(__dirname, 'FBA_extracted'),
      message: null, error: e.message
    });
  }
});

app.post('/import/billing', (req, res) => {
  const { filename } = req.body;
  const htmlPath = path.join(__dirname, 'AMZLGS_extracted', filename);
  try {
    const count = importBilling(htmlPath);
    res.render('import_page', {
      dataDir: path.join(__dirname, 'AMZLGS_extracted'),
      fbaDir:  path.join(__dirname, 'FBA_extracted'),
      message: `帳單匯入完成：${count} 筆記錄`,
      error: null
    });
  } catch (e) {
    res.render('import_page', {
      dataDir: path.join(__dirname, 'AMZLGS_extracted'),
      fbaDir:  path.join(__dirname, 'FBA_extracted'),
      message: null, error: e.message
    });
  }
});

// ─── Start ───────────────────────────────────────────────────────────────────
app.listen(PORT, () => {
  console.log(`\n✅ Lab52 庫存系統已啟動`);
  console.log(`   開啟瀏覽器：http://localhost:${PORT}\n`);
});
