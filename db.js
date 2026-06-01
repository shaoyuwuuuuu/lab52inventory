const { Database } = require('node-sqlite3-wasm');
const path = require('path');
const os = require('os');
const fs = require('fs');

// node-sqlite3-wasm on Windows requires a path without spaces
const dbDir  = path.join(os.homedir(), 'Lab52Inventory').replace(/\\/g, '/');
fs.mkdirSync(dbDir.replace(/\//g, '\\'), { recursive: true });
const dbPath = dbDir + '/inventory.db';
// Clean up stale lock directory on startup (crash recovery)
const lockDir = dbPath.replace(/\//g, '\\') + '.lock';
try { fs.rmSync(lockDir, { recursive: true, force: true }); } catch {}
const db = new Database(dbPath);
console.log(`資料庫位置：${dbPath}`);
// Ensure DB is closed on process exit to release lock
process.on('exit', () => { try { db.close(); } catch {} });
process.on('SIGINT', () => process.exit(0));
process.on('SIGTERM', () => process.exit(0));

db.exec('PRAGMA foreign_keys = ON');

db.exec(`CREATE TABLE IF NOT EXISTS products (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  name TEXT NOT NULL,
  sku TEXT,
  ean TEXT,
  cubic_feet_per_carton REAL,
  qty_per_carton INTEGER,
  note TEXT,
  created_at TEXT DEFAULT (datetime('now','localtime'))
)`);

db.exec(`CREATE TABLE IF NOT EXISTS batches (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  product_id INTEGER NOT NULL REFERENCES products(id) ON DELETE CASCADE,
  exp_date TEXT,
  location TEXT DEFAULT 'AMZLGS',
  current_cartons REAL DEFAULT 0,
  note TEXT,
  created_at TEXT DEFAULT (datetime('now','localtime'))
)`);

db.exec(`CREATE TABLE IF NOT EXISTS movements (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  batch_id INTEGER NOT NULL REFERENCES batches(id) ON DELETE CASCADE,
  date TEXT NOT NULL,
  type TEXT NOT NULL,
  qty_cartons REAL NOT NULL,
  note TEXT,
  created_at TEXT DEFAULT (datetime('now','localtime'))
)`);

db.exec(`CREATE TABLE IF NOT EXISTS billing (
  id INTEGER PRIMARY KEY AUTOINCREMENT,
  date TEXT NOT NULL,
  type TEXT NOT NULL,
  job_name TEXT,
  qty REAL,
  unit_fee REAL,
  total_fee REAL,
  note TEXT
)`);

try { db.exec('ALTER TABLE products ADD COLUMN asin TEXT'); } catch {}

module.exports = db;
