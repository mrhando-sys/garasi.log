// ===== Code.gs (Google Apps Script, terikat ke spreadsheetmu) =====
const SHEETS = { motor: 'Motor', mobil: 'Mobil' };
const FIELDS = {
  tanggal: ['tanggal', 'tgl', 'date'], merk: ['merk', 'merek', 'kendaraan', 'tipe'],
  item: ['sparepart', 'item', 'keterangan', 'barang'], harga: ['harga', 'nominal', 'biaya', 'pengeluaran'],
  jenis: ['jenis'], km: ['km', 'odometer'], bengkel: ['bengkel'], catatan: ['catatan', 'note', 'notes']
};
const TITLES = { tanggal: 'Tanggal', merk: 'Merk', item: 'Sparepart', harga: 'Harga', jenis: 'Jenis', km: 'KM', bengkel: 'Bengkel', catatan: 'Catatan' };

function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

// Cocokkan header yang SUDAH ADA; kolom yang belum ada ditambah di kanan (data lama tidak disentuh)
function mapCols(sh, create) {
  const lc = Math.max(sh.getLastColumn(), 1);
  const h = sh.getRange(1, 1, 1, lc).getValues()[0].map(v => String(v).trim().toLowerCase());
  const m = {}; let next = h.reduce((a, v, i) => v ? i + 1 : a, 0) + 1;
  for (const k in FIELDS) {
    const i = h.findIndex(x => x && FIELDS[k].some(a => x === a || x.indexOf(a) === 0));
    if (i >= 0) m[k] = i + 1;
    else if (create) { sh.getRange(1, next).setValue(TITLES[k]); m[k] = next++; }
  }
  return m;
}
function lastDataRow(sh, m) { // abaikan kolom rumus yang sudah terisi ke bawah
  const n = sh.getLastRow(); if (n < 2) return 1;
  const c = m.item || m.tanggal || 1, v = sh.getRange(1, c, n, 1).getValues();
  for (let i = n - 1; i >= 1; i--) if (v[i][0] !== '') return i + 1;
  return 1;
}
function toISO(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const d = new Date(v); return isNaN(d) ? '' : Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}
function num(v) { return typeof v === 'number' ? v : Number(String(v).replace(/[^\d]/g, '')) || 0; }

function listRows(sh) {
  const m = mapCols(sh, false), n = lastDataRow(sh, m); if (n < 2) return [];
  const vals = sh.getRange(2, 1, n - 1, sh.getLastColumn()).getValues(), r = [];
  vals.forEach((row, i) => {
    const o = { id: i + 2 };
    for (const k in FIELDS) o[k] = m[k] ? row[m[k] - 1] : '';
    if (!o.item && !o.harga) return;
    o.tanggal = toISO(o.tanggal); o.harga = num(o.harga); o.km = num(o.km); r.push(o);
  });
  return r;
}
function writeRow(sh, row, d) {
  const m = mapCols(sh, true);
  const set = (k, v) => { if (m[k]) sh.getRange(row, m[k]).setValue(v); };
  const p = String(d.tanggal).split('-');
  set('tanggal', new Date(+p[0], +p[1] - 1, +p[2])); sh.getRange(row, m.tanggal).setNumberFormat('dd-MMM-yyyy');
  set('merk', d.merk || ''); set('item', d.item);
  set('harga', Number(d.harga)); sh.getRange(row, m.harga).setNumberFormat('"Rp. "#,##0');
  set('jenis', d.jenis || ''); set('km', d.km ? Number(d.km) : ''); set('bengkel', d.bengkel || ''); set('catatan', d.catatan || '');
}

function doPost(e) {
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const b = JSON.parse(e.postData.contents);
    if (b.token !== PropertiesService.getScriptProperties().getProperty('PIN')) return out({ error: 'auth' });
    const ss = SpreadsheetApp.getActive();
    if (b.action === 'list') return out({ motor: listRows(ss.getSheetByName(SHEETS.motor)), mobil: listRows(ss.getSheetByName(SHEETS.mobil)) });
    const sh = ss.getSheetByName(SHEETS[b.kat]); if (!sh) return out({ error: 'Sheet tidak ditemukan' });
    if (b.action === 'add') writeRow(sh, lastDataRow(sh, mapCols(sh, true)) + 1, b.data);
    else if (b.action === 'edit') writeRow(sh, b.row, b.data);
    else if (b.action === 'delete') sh.deleteRow(b.row);
    return out({ ok: true });
  } catch (err) { return out({ error: String(err) }); } finally { lock.releaseLock(); }
}
// Jalankan SEKALI untuk membuat PIN: ubah teksnya dulu, lalu klik Run
function setPin() { PropertiesService.getScriptProperties().setProperty('PIN', 'GANTI-PIN-RAHASIA-KAMU'); }
