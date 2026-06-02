const XLSX = require('xlsx');

const wb = XLSX.readFile('Lab52外倉庫存表.xlsx', { cellFormula: true });
const ws = wb.Sheets['庫存總覽']; // 庫存總覽

ws['!ref'] = 'A1:J17';

for (let r = 2; r <= 17; r++) {
  // F: 在倉總箱數
  ws['F' + r] = {
    t: 'n', v: 0,
    f: 'IFERROR(SUMIF(Movement!C:C,B' + r + ',Movement!G:G),0)'
  };
  // G: 估算總pcs
  ws['G' + r] = {
    t: 'n', v: 0,
    f: 'IF(F' + r + '=0,"",F' + r + '*E' + r + ')'
  };
  // H: 最近效期（有 IN 記錄的最早 exp_date）
  ws['H' + r] = {
    t: 's', v: '',
    f: 'IFERROR(TEXT(MINIFS(Movement!F:F,Movement!C:C,B' + r + ',Movement!G:G,">"&0),"yyyy-mm-dd"),"")'
  };
  // I: 最近效期批剩餘（那個效期的淨箱數）
  ws['I' + r] = {
    t: 'n', v: 0,
    f: 'IFERROR(SUMIFS(Movement!G:G,Movement!C:C,B' + r + ',Movement!F:F,MINIFS(Movement!F:F,Movement!C:C,B' + r + ',Movement!G:G,">"&0)),"")'
  };
  // J: 庫存狀態
  ws['J' + r] = {
    t: 's', v: '',
    f: 'IF(F' + r + '=0,"",IF(F' + r + '<=5,"🔴 庫存偏低",IF(F' + r + '<=15,"🟡 注意","✅ 正常")))'
  };
}

XLSX.writeFile(wb, 'Lab52外倉庫存表.xlsx', { bookType: 'xlsx' });
console.log('完成！公式已寫入庫存總覽 F~J 欄');
