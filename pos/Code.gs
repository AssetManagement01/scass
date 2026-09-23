const LOGIN_ID = 'USER';
const LOGIN_PASS = 'USER';
const ERP_INV = 'Inventory';
const ERP_SUP = 'Supplier';
const ERP_CUS = 'Pelanggan';
const ERP_PO = 'Pembelian';
const ERP_INVCE = 'Invoice';
const ERP_KEU = 'Keuangan';
const SS_FIXED = '1wg9GiBvGzclf8_ekV-YXMqd0D0o13yBMvdnodPSU1O0';

const API_TOKEN = 'NSK-SCASS-2026';

function doGet(e) {
  e = e || { parameter: {} };
  const p = (e.parameter) || {};
  if (p.action) return apiOut_(p, dispatchApi_(p.action, p.payload, p.token));
  return HtmlService.createHtmlOutput(ERP_HTML)
    .setTitle('SCASS')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  e = e || {};
  const p = Object.assign({}, e.parameter || {});
  if (e.postData && e.postData.contents) {
    try {
      const body = JSON.parse(e.postData.contents);
      Object.keys(body).forEach(function(k){ p[k] = body[k]; });
    } catch (err) {}
  }
  const res = dispatchApi_(p.action, p.payload, p.token);
  return HtmlService.createHtmlOutput(
    '<html><body><script>try{parent.postMessage({scass:1,res:' + JSON.stringify(res) + '},"*");}catch(ex){}</script>ok</body></html>'
  ).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function apiOut_(p, res) {
  const text = JSON.stringify(res);
  if (p.callback) {
    return ContentService.createTextOutput(p.callback + '(' + text + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(text).setMimeType(ContentService.MimeType.JSON);
}

function dispatchApi_(action, payload, token) {
  try {
    if (String(token || '') !== API_TOKEN) return { ok: false, error: 'Invalid API token' };
    var data = payload;
    if (typeof payload === 'string' && payload) {
      try { data = JSON.parse(payload); } catch (e2) { data = {}; }
    }
    data = data || {};
    var fns = {
      getErpDashboard: getErpDashboard,
      getErpData: getErpData,
      simpanBarang: simpanBarang,
      ubahStok: ubahStok,
      simpanPembelian: simpanPembelian,
      simpanInvoice: simpanInvoice,
      simpanKeuangan: simpanKeuangan,
      simpanMitra: simpanMitra,
      exportErpExcel: exportErpExcel,
      importErpRows: importErpRows,
      exportInvoicePdf: exportInvoicePdf,
      ubahStatusInvoice: ubahStatusInvoice
    };
    if (!fns[action]) return { ok: false, error: 'Unknown action: ' + action };
    var out = fns[action](data);
    if (out && typeof out === 'object' && out.ok === undefined) out.ok = true;
    return out;
  } catch (err) {
    return { ok: false, error: String(err.message || err) };
  }
}

function login(payload) {
  const id = String((payload && payload.id) || '').trim();
  const pass = String((payload && payload.pass) || '');
  const idU = id.toUpperCase();
  if (idU === 'USER' && pass.toUpperCase() === 'USER') return { ok: true, role: 'USER' };
  if (idU === 'ADMIN' && pass === 'admin1234') return { ok: true, role: 'ADMIN' };
  if (idU === 'TRIAL' && pass.toUpperCase() === 'TRIAL') return { ok: true, role: 'TRIAL' };
  throw new Error('LOGIN TIDAK DI TEMUKAN');
}

function getSs_() {
  const props = PropertiesService.getScriptProperties();
  try { return SpreadsheetApp.openById(SS_FIXED); } catch (e0) {}
  const saved = props.getProperty('SS_ID');
  if (saved) {
    try { return SpreadsheetApp.openById(saved); } catch (e) {}
  }
  const active = SpreadsheetApp.getActive();
  if (active) {
    props.setProperty('SS_ID', active.getId());
    return active;
  }
  const created = SpreadsheetApp.create('SCASS ERP');
  props.setProperty('SS_ID', created.getId());
  return created;
}

function headerStyle_(sh, n) {
  sh.getRange(1, 1, 1, n).setFontWeight('bold').setBackground('#12324a').setFontColor('#ffffff');
  sh.setFrozenRows(1);
}

function ensureSheet_(name, header) {
  const ss = getSs_();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, header.length).setValues([header]);
    headerStyle_(sh, header.length);
  }
  return sh;
}

function ensureErp_() {
  ensureSheet_(ERP_INV, ['SKU', 'Nama', 'Kategori', 'Satuan', 'Stok', 'Harga Beli', 'Harga Jual', 'Min Stok', 'Status']);
  ensureSheet_(ERP_SUP, ['Kode', 'Nama', 'Telepon', 'Alamat']);
  ensureSheet_(ERP_CUS, ['Kode', 'Nama', 'Telepon', 'Alamat']);
  ensureSheet_(ERP_PO, ['Waktu', 'No PO', 'Supplier', 'SKU', 'Nama', 'Qty', 'Harga', 'Subtotal', 'Status']);
  ensureSheet_(ERP_INVCE, ['Waktu', 'No Inv', 'Pelanggan', 'SKU', 'Nama', 'Qty', 'Harga', 'Subtotal', 'Status', 'Metode', 'Bayar', 'Diskon']);
  ensureSheet_(ERP_KEU, ['Waktu', 'No Bukti', 'Jenis', 'Akun', 'Keterangan', 'Jumlah', 'Ref']);
}

function readTable_(name) {
  const sh = getSs_().getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return [];
  const values = sh.getDataRange().getValues();
  const head = values[0].map(h => String(h || '').trim());
  const out = [];
  for (let i = 1; i < values.length; i++) {
    const row = { _row: i + 1 };
    let empty = true;
    head.forEach((h, c) => {
      let v = values[i][c];
      if (v instanceof Date) v = Utilities.formatDate(v, 'Asia/Jakarta', 'dd/MM/yyyy HH:mm');
      row[h] = v;
      if (String(v || '').trim() !== '') empty = false;
    });
    if (!empty) out.push(row);
  }
  return out;
}

function nextCode_(name, col, prefix) {
  let max = 0;
  readTable_(name).forEach(r => {
    const m = String(r[col] || '').match(new RegExp('^' + prefix + '-(\\d+)$'));
    if (m) max = Math.max(max, parseInt(m[1], 10));
  });
  return prefix + '-' + String(max + 1).padStart(3, '0');
}


function seedDemo_() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('SEEDED') === '1') return;
  const ss = getSs_();
  if (!readTable_(ERP_SUP).length) {
    ss.getSheetByName(ERP_SUP).appendRow(['SUP-001', 'CV Sumber Bahan', '0812-0000-1111', 'Jakarta']);
  }
  if (!readTable_(ERP_CUS).length) {
    ss.getSheetByName(ERP_CUS).appendRow(['PLG-001', 'Pelanggan Umum', '-', '-']);
  }
  if (!readTable_(ERP_INV).length) {
    const inv = ss.getSheetByName(ERP_INV);
    inv.appendRow(['MNM-001', 'Es Teh', 'Minuman', 'porsi', 20, 1500, 4000, 5, 'AKTIF']);
    inv.appendRow(['MNM-002', 'Kopi Hitam', 'Minuman', 'porsi', 15, 2000, 5000, 5, 'AKTIF']);
    inv.appendRow(['MKN-001', 'Nasi Kucing', 'Makanan', 'porsi', 10, 3000, 7000, 5, 'AKTIF']);
    inv.appendRow(['MKN-002', 'Sate Usus', 'Makanan', 'tusuk', 30, 1000, 2500, 10, 'AKTIF']);
  }
  if (!readTable_(ERP_KEU).length) {
    const now = new Date();
    ss.getSheetByName(ERP_KEU).appendRow([now, 'K-SALDO', 'Masuk', 'Kas', 'Saldo awal', 500000, 'OPEN']);
  }
  props.setProperty('SEEDED', '1');
}

function prefixFromCategory_(kategori) {
  const clean = String(kategori || 'BRG').toUpperCase().replace(/[^A-Z0-9\s]/g, ' ').trim();
  const parts = clean.split(/\s+/).filter(Boolean);
  if (!parts.length) return 'BRG';
  if (parts.length === 1) return parts[0].slice(0, 3).padEnd(3, 'X');
  return parts.map(p => p[0]).join('').slice(0, 3).padEnd(3, 'X');
}

function pickNum_(r, keys) {
  for (var i = 0; i < keys.length; i++) {
    var v = r[keys[i]];
    if (v === null || v === undefined || v === '') continue;
    var n = Number(String(v).replace(/[^0-9.\-]/g, ''));
    if (!isNaN(n)) return n;
  }
  return 0;
}
function hasVal_(r, keys) {
  for (var i = 0; i < keys.length; i++) {
    var v = r[keys[i]];
    if (v !== null && v !== undefined && String(v).trim() !== '') return true;
  }
  return false;
}
function getErpDashboard() {
  ensureErp_();
  seedDemo_();
  const inv = readTable_(ERP_INV);
  const po = readTable_(ERP_PO).filter(r => String(r.Status || '') !== 'VOID');
  const invce = readTable_(ERP_INVCE).filter(r => String(r.Status || '') !== 'VOID');
  const keu = readTable_(ERP_KEU);
  const qtyKeys = ['Stok', 'Qty', 'QTY', 'Stock', 'Jumlah'];
  const priceKeys = ['Harga Beli', 'Harga Jual', 'Harga', 'Price', 'Cost', 'Harga Pokok'];
  const minKeys = ['Min Stok', 'Min', 'Minimum'];
  const stokNilai = inv.reduce((s, r) => s + pickNum_(r, qtyKeys) * pickNum_(r, priceKeys), 0);
  const stokRendah = inv.filter(r => hasVal_(r, qtyKeys) && pickNum_(r, qtyKeys) <= pickNum_(r, minKeys));
  const beli = po.reduce((s, r) => s + (Number(r.Subtotal) || 0), 0);
  const jual = invce.reduce((s, r) => s + (Number(r.Subtotal) || 0), 0);
  let kas = 0;
  keu.forEach(r => {
    const n = Number(r.Jumlah) || 0;
    kas += String(r.Jenis || '').toLowerCase() === 'keluar' ? -n : n;
  });
  const today = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy');
  const jualHari = invce.reduce((s, r) => {
    return String(r.Waktu || '').indexOf(today) === 0 ? s + (Number(r.Subtotal) || 0) : s;
  }, 0);
  const ss = getSs_();
  return {
    ok: true,
    sheetUrl: ss.getUrl(),
    sheetName: ss.getName(),
    kpi: {
      sku: inv.length,
      stokNilai: stokNilai,
      stokRendah: stokRendah.length,
      pembelian: beli,
      penjualan: jual,
      kas: kas,
      posHari: jualHari,
      pending: invce.filter(r => String(r.Status||'').toUpperCase()=='PENDING').length
    },
    rendah: stokRendah.slice(0, 8),
    pubHint: ss.getName()
  };
}

function getErpData() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('ERP_DATA');
  if (hit) {
    try { return JSON.parse(hit); } catch (e) {}
  }
  const ready = cache.get('ERP_READY');
  if (ready !== '1') {
    ensureErp_();
    try { seedDemo_(); } catch (e) {}
    cache.put('ERP_READY', '1', 21600);
  }
  const out = {
    inventory: readTable_(ERP_INV),
    supplier: readTable_(ERP_SUP),
    pelanggan: readTable_(ERP_CUS),
    pembelian: readTable_(ERP_PO).reverse(),
    invoice: readTable_(ERP_INVCE).reverse(),
    keuangan: readTable_(ERP_KEU).reverse()
  };
  try { cache.put('ERP_DATA', JSON.stringify(out), 45); } catch (e) {}
  return out;
}
function clearErpCache_() {
  try { CacheService.getScriptCache().removeAll(['ERP_DATA','ERP_READY']); } catch (e) {}
}

function simpanBarang(p) {
  ensureErp_();
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sh = getSs_().getSheetByName(ERP_INV);
    const nama = String(p.nama || '').trim();
    if (!nama) throw new Error('Nama barang wajib');
    const sku = String(p.sku || '').trim() || nextCode_(ERP_INV, 'SKU', prefixFromCategory_(p.kategori || 'BRG'));
    const existing = readTable_(ERP_INV).find(r => String(r.SKU) === sku);
    const row = [sku, nama, p.kategori || 'Umum', p.satuan || 'pcs', Number(p.stok) || 0, Number(p.hargaBeli) || 0, Number(p.hargaJual) || 0, Number(p.minStok) || 0, p.status || 'AKTIF'];
    if (existing) sh.getRange(existing._row, 1, 1, 9).setValues([row]);
    else sh.appendRow(row);
    return { ok: true, sku: sku };
  } finally { lock.releaseLock(); }
}

function ubahStok(p) {
  ensureErp_();
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sku = String(p.sku || '').trim();
    const item = readTable_(ERP_INV).find(r => String(r.SKU) === sku);
    if (!item) throw new Error('SKU tidak ditemukan');
    let stok;
    if (p.stok !== undefined && p.stok !== null && String(p.stok).trim() !== '') {
      stok = Math.max(0, Number(p.stok) || 0);
    } else {
      const delta = Number(p.delta) || 0;
      stok = Math.max(0, (Number(item.Stok) || 0) + delta);
    }
    getSs_().getSheetByName(ERP_INV).getRange(item._row, 5).setValue(stok);
    return { ok: true, stok: stok };
  } finally { lock.releaseLock(); }
}

function simpanMitra(p) {
  ensureErp_();
  const jenis = p.jenis === 'pelanggan' ? ERP_CUS : ERP_SUP;
  const prefix = p.jenis === 'pelanggan' ? 'PLG' : 'SUP';
  const nama = String(p.nama || '').trim();
  if (!nama) throw new Error('Nama wajib');
  const kode = String(p.kode || '').trim() || nextCode_(jenis, 'Kode', prefix);
  const existing = readTable_(jenis).find(r => String(r.Kode) === kode);
  const row = [kode, nama, p.telepon || '', p.alamat || ''];
  const sh = getSs_().getSheetByName(jenis);
  if (existing) sh.getRange(existing._row, 1, 1, 4).setValues([row]);
  else sh.appendRow(row);
  return { ok: true, kode: kode };
}

function simpanPembelian(p) {
  ensureErp_();
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const items = p.items || [];
    if (!items.length) throw new Error('Item pembelian kosong');
    const now = new Date();
    const no = String(p.noPo || '').trim() || ('PO-' + Utilities.formatDate(now, 'Asia/Jakarta', 'yyyyMMdd-HHmmss'));
    const sh = getSs_().getSheetByName(ERP_PO);
    items.forEach(it => {
      const qty = Number(it.qty) || 0;
      const harga = Number(it.harga) || 0;
      sh.appendRow([now, no, p.supplier || '', it.sku, it.nama, qty, harga, qty * harga, 'DITERIMA']);
      const inv = readTable_(ERP_INV).find(r => String(r.SKU) === String(it.sku));
      if (inv) {
        getSs_().getSheetByName(ERP_INV).getRange(inv._row, 5).setValue((Number(inv.Stok) || 0) + qty);
      }
    });
    const total = items.reduce((s, it) => s + (Number(it.qty) || 0) * (Number(it.harga) || 0), 0);
    getSs_().getSheetByName(ERP_KEU).appendRow([now, no, 'Keluar', 'Pembelian', 'PO ' + no, total, no]);
    return { ok: true, noPo: no };
  } finally { lock.releaseLock(); }
}

function simpanInvoice(p) {
  ensureErp_();
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const items = p.items || [];
    if (!items.length) throw new Error('Item invoice kosong');
    const now = new Date();
    const no = String(p.noInv || '').trim() || ('INV-' + Utilities.formatDate(now, 'Asia/Jakarta', 'yyyyMMdd-HHmmss'));
    const sh = getSs_().getSheetByName(ERP_INVCE);
    items.forEach(it => {
      const qty = Number(it.qty) || 0;
      const harga = Number(it.harga) || 0;
      const st = String(p.status || 'PENDING').toUpperCase();
      const disc = Number(p.diskon) || 0;
      const grossAll = items.reduce((s, x) => s + (Number(x.qty) || 0) * (Number(x.harga) || 0), 0);
      const share = grossAll ? (qty * harga) / grossAll : 0;
      const lineDisc = Math.round(disc * share);
      const lineNet = qty * harga - lineDisc;
      sh.appendRow([now, no, p.pelanggan || '', it.sku, it.nama, qty, harga, lineNet, st, p.metode || '', Number(p.bayar)||0, lineDisc]);
      const inv = readTable_(ERP_INV).find(r => String(r.SKU) === String(it.sku));
      if (inv) {
        const stok = Math.max(0, (Number(inv.Stok) || 0) - qty);
        getSs_().getSheetByName(ERP_INV).getRange(inv._row, 5).setValue(stok);
      }
    });
    const total = items.reduce((s, it) => s + (Number(it.qty) || 0) * (Number(it.harga) || 0), 0);
    const st = String(p.status || 'PENDING').toUpperCase();
    if (st === 'LUNAS' || st === 'PAID') {
      getSs_().getSheetByName(ERP_KEU).appendRow([now, no, 'Masuk', 'Penjualan', 'Invoice ' + no, Number(p.bayar)||total, no]);
    }
    return { ok: true, noInv: no, total: total, status: st };
  } finally { lock.releaseLock(); }
}

function simpanKeuangan(p) {
  ensureErp_();
  const jenis = String(p.jenis || 'Masuk');
  const jumlah = Number(p.jumlah) || 0;
  if (jumlah <= 0) throw new Error('Jumlah tidak valid');
  const now = new Date();
  const no = 'K-' + Utilities.formatDate(now, 'Asia/Jakarta', 'yyyyMMdd-HHmmss');
  getSs_().getSheetByName(ERP_KEU).appendRow([now, no, jenis, p.akun || 'Kas', p.keterangan || '', jumlah, '']);
  return { ok: true, noBukti: no };
}


function exportErpExcel() {
  ensureErp_();
  const ss = getSs_();
  const url = 'https://docs.google.com/spreadsheets/d/' + ss.getId() + '/export?format=xlsx';
  const res = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
  if (res.getResponseCode() >= 400) {
    throw new Error('Gagal membuat file Excel');
  }
  const bytes = res.getBlob().getBytes();
  return { ok: true, name: 'SCASS-ERP.xlsx', base64: Utilities.base64Encode(bytes) };
}

function importErpRows(p) {
  ensureErp_();
  const jenis = String((p && p.jenis) || '');
  const rows = (p && p.rows) || [];
  if (!rows.length) throw new Error('Tidak ada baris untuk diimpor');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    let n = 0;
    if (jenis === 'inventory') {
      rows.forEach(r => {
        simpanBarang({
          sku: r.SKU || r.sku,
          nama: r.Nama || r.nama,
          kategori: r.Kategori || r.kategori,
          satuan: r.Satuan || r.satuan,
          stok: r.Stok || r.stok,
          hargaBeli: r['Harga Beli'] || r.hargaBeli,
          hargaJual: r['Harga Jual'] || r.hargaJual,
          minStok: r['Min Stok'] || r.minStok,
          status: r.Status || 'AKTIF'
        });
        n++;
      });
    } else if (jenis === 'supplier' || jenis === 'pelanggan') {
      rows.forEach(r => {
        simpanMitra({
          jenis: jenis,
          kode: r.Kode || r.kode,
          nama: r.Nama || r.nama,
          telepon: r.Telepon || r.telepon,
          alamat: r.Alamat || r.alamat
        });
        n++;
      });
    } else if (jenis === 'keuangan') {
      const sh = getSs_().getSheetByName(ERP_KEU);
      rows.forEach(r => {
        sh.appendRow([
          r.Waktu || new Date(),
          r['No Bukti'] || ('K-IMP-' + n),
          r.Jenis || 'Masuk',
          r.Akun || 'Kas',
          r.Keterangan || '',
          Number(r.Jumlah) || 0,
          r.Ref || 'IMPORT'
        ]);
        n++;
      });
    } else if (jenis === 'pembelian') {
      const sh = getSs_().getSheetByName(ERP_PO);
      rows.forEach(r => {
        sh.appendRow([
          r.Waktu || new Date(),
          r['No PO'] || ('PO-IMP-' + n),
          r.Supplier || '',
          r.SKU || '',
          r.Nama || '',
          Number(r.Qty) || 0,
          Number(r.Harga) || 0,
          Number(r.Subtotal) || ((Number(r.Qty)||0)*(Number(r.Harga)||0)),
          r.Status || 'DITERIMA'
        ]);
        n++;
      });
    } else if (jenis === 'invoice') {
      const sh = getSs_().getSheetByName(ERP_INVCE);
      rows.forEach(r => {
        sh.appendRow([
          r.Waktu || new Date(),
          r['No Inv'] || ('INV-IMP-' + n),
          r.Pelanggan || '',
          r.SKU || '',
          r.Nama || '',
          Number(r.Qty) || 0,
          Number(r.Harga) || 0,
          Number(r.Subtotal) || ((Number(r.Qty)||0)*(Number(r.Harga)||0)),
          r.Status || 'LUNAS'
        ]);
        n++;
      });
    } else {
      throw new Error('Jenis import tidak dikenal');
    }
    return { ok: true, jenis: jenis, masuk: n };
  } finally { lock.releaseLock(); }
}


function exportInvoicePdf(p) {
  ensureErp_();
  const no = String((p && p.noInv) || '');
  const rows = readTable_(ERP_INVCE).filter(r => String(r['No Inv']) === no);
  if (!rows.length) throw new Error('Invoice not found');
  const ss = getSs_();
  let sh = ss.getSheetByName('PrintInv');
  if (!sh) sh = ss.insertSheet('PrintInv');
  sh.clear();
  sh.setHiddenGridlines(true);
  sh.setColumnWidths(1, 5, 110);
  sh.setColumnWidth(1, 90);
  sh.setColumnWidth(2, 200);
  sh.setColumnWidth(3, 70);
  sh.setColumnWidth(4, 110);
  sh.setColumnWidth(5, 120);

  sh.getRange('A1:E1').merge().setValue('SCASS').setFontWeight('bold').setFontSize(18).setHorizontalAlignment('center').setFontColor('#12324a');
  sh.getRange('A2:E2').merge().setValue('NewStar Kancing').setHorizontalAlignment('center').setFontSize(11).setFontColor('#5b7380');
  sh.getRange('A3:E3').merge().setValue('SALES INVOICE').setFontWeight('bold').setHorizontalAlignment('center').setBackground('#12324a').setFontColor('#ffffff').setFontSize(12);
  sh.getRange(1, 1, 3, 5).setVerticalAlignment('middle');
  sh.setRowHeight(1, 28);
  sh.setRowHeight(3, 26);

  const meta = [
    ['Invoice No.', no, '', 'Date', rows[0].Waktu || ''],
    ['Customer', rows[0].Pelanggan || '-', '', 'Status', rows[0].Status || ''],
    ['Payment', rows[0].Metode || '-', '', 'Paid', Number(rows[0].Bayar) || 0]
  ];
  sh.getRange(5, 1, 3, 5).setValues(meta);
  sh.getRange('A5:A7').setFontWeight('bold').setBackground('#eef6f4');
  sh.getRange('D5:D7').setFontWeight('bold').setBackground('#eef6f4');
  sh.getRange(5, 1, 3, 5).setBorder(true, true, true, true, true, true, '#c9dde0', SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange('E7').setNumberFormat('"Rp "#,##0');

  const start = 9;
  sh.getRange(start, 1, 1, 5).setValues([['SKU', 'Item', 'Qty', 'Price', 'Subtotal']]);
  sh.getRange(start, 1, 1, 5).setFontWeight('bold').setBackground('#1f7a6e').setFontColor('#ffffff').setHorizontalAlignment('center');
  const items = rows.map(r => [r.SKU || '', r.Nama || '', Number(r.Qty) || 0, Number(r.Harga) || 0, Number(r.Subtotal) || 0]);
  if (items.length) {
    sh.getRange(start + 1, 1, items.length, 5).setValues(items);
    sh.getRange(start + 1, 3, items.length, 1).setHorizontalAlignment('center');
    sh.getRange(start + 1, 4, items.length, 2).setNumberFormat('"Rp "#,##0');
  }
  const last = start + items.length;
  sh.getRange(start, 1, items.length + 1, 5).setBorder(true, true, true, true, true, true, '#12324a', SpreadsheetApp.BorderStyle.SOLID);

  const tot = rows.reduce((s, r) => s + (Number(r.Subtotal) || 0), 0);
  const gross = rows.reduce((s, r) => s + (Number(r.Qty) || 0) * (Number(r.Harga) || 0), 0);
  const stored = rows.reduce((s, r) => s + (Number(r.Diskon) || 0), 0);
  const disc = stored > 0 ? stored : Math.max(0, gross - tot);
  var pct = gross > 0 ? Math.round((disc / gross) * 1000) / 10 : 0;
  var pctTxt = (Math.abs(pct - Math.round(pct)) < 0.05) ? String(Math.round(pct)) : String(pct);
  sh.getRange(last + 1, 1, 1, 5).setValues([['', '', '', 'SUBTOTAL', gross]]);
  sh.getRange(last + 2, 1, 1, 5).setValues([['', '', '', 'DISCOUNT ' + pctTxt + '%', disc]]);
  sh.getRange(last + 3, 1, 1, 5).setValues([['', '', '', 'TOTAL', tot]]);
  sh.getRange(last + 1, 1, 3, 5).setFontWeight('bold');
  sh.getRange(last + 1, 5, 3, 1).setNumberFormat('"Rp "#,##0');
  sh.getRange(last + 1, 1, 3, 5).setBorder(true, true, true, true, true, true, '#0b4f37', SpreadsheetApp.BorderStyle.SOLID);
  sh.getRange(last + 3, 1, 1, 5).setBackground('#0b4f37').setFontColor('#ffffff');
  sh.getRange(last + 5, 1, 1, 5).merge().setValue('Thank you for your business').setHorizontalAlignment('center').setFontColor('#5b7380').setFontStyle('italic');
  SpreadsheetApp.flush();

  const url = 'https://docs.google.com/spreadsheets/d/' + ss.getId()
    + '/export?exportFormat=pdf&format=pdf&gid=' + sh.getSheetId()
    + '&size=A4&portrait=true&fitw=true&sheetnames=false&printtitle=false&pagenumbers=false&gridlines=false'
    + '&top_margin=0.5&bottom_margin=0.5&left_margin=0.5&right_margin=0.5';
  const res = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
  if (res.getResponseCode() >= 400) return { ok: true, url: ss.getUrl() };
  const file = DriveApp.createFile(res.getBlob().setName('SCASS-' + no + '.pdf'));
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return { ok: true, url: file.getUrl() };
}


function scassPasangOtomasi() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(tr => {
    if (tr.getHandlerFunction() === 'scassDailyJob') ScriptApp.deleteTrigger(tr);
  });
  ScriptApp.newTrigger('scassDailyJob').timeBased().atHour(21).everyDays(1).inTimezone('Asia/Jakarta').create();
  return { ok: true, jam: '21:00 WIB setiap hari', fungsi: 'scassDailyJob' };
}

function scassDailyJob() {
  ensureErp_();
  const tz = 'Asia/Jakarta';
  const now = new Date();
  const hari = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
  const label = Utilities.formatDate(now, tz, 'dd/MM/yyyy HH:mm');
  const inv = readTable_(ERP_INV);
  const po = readTable_(ERP_PO).filter(r => String(r.Status || '') !== 'VOID' && String(r.Status || '') !== 'BATAL');
  const invce = readTable_(ERP_INVCE).filter(r => String(r.Status || '').toUpperCase() !== 'BATAL' && String(r.Status || '').toUpperCase() !== 'VOID');
  const keu = readTable_(ERP_KEU);
  const rendah = inv.filter(r => (Number(r.Stok) || 0) <= (Number(r['Min Stok']) || 0));
  const pending = invce.filter(r => String(r.Status || '').toUpperCase() === 'PENDING');
  const hariIni = function(waktu) {
    return String(waktu || '').indexOf(Utilities.formatDate(now, tz, 'dd/MM/yyyy')) === 0 || String(waktu || '').indexOf(hari) === 0;
  };
  const jualHari = invce.filter(r => hariIni(r.Waktu)).reduce((s, r) => s + (Number(r.Subtotal) || 0), 0);
  const beliHari = po.filter(r => hariIni(r.Waktu)).reduce((s, r) => s + (Number(r.Subtotal) || 0), 0);
  let kas = 0;
  keu.forEach(r => {
    const n = Number(r.Jumlah) || 0;
    kas += String(r.Jenis || '').toLowerCase() === 'keluar' ? -n : n;
  });
  const sh = ensureSheet_('AutoLog', ['Waktu', 'Jenis', 'Ringkasan', 'Nilai']);
  sh.appendRow([now, 'DAILY', 'Sales today', jualHari]);
  sh.appendRow([now, 'DAILY', 'Purchases today', beliHari]);
  sh.appendRow([now, 'DAILY', 'Cash balance', kas]);
  sh.appendRow([now, 'DAILY', 'Low stock SKU', rendah.length]);
  sh.appendRow([now, 'DAILY', 'Pending invoices', pending.length]);
  rendah.forEach(r => sh.appendRow([now, 'LOW_STOCK', r.SKU + ' ' + r.Nama, Number(r.Stok) || 0]));
  pending.forEach(r => sh.appendRow([now, 'PENDING', String(r['No Inv'] || '') + ' ' + String(r.Pelanggan || ''), Number(r.Subtotal) || 0]));

  const lines = [];
  lines.push('SCASS daily job — ' + label);
  lines.push('NewStar Kancing');
  lines.push('Sales today: ' + jualHari);
  lines.push('Purchases today: ' + beliHari);
  lines.push('Cash: ' + kas);
  lines.push('Low stock: ' + rendah.length);
  rendah.slice(0, 20).forEach(r => lines.push('- ' + r.SKU + ' ' + r.Nama + ' = ' + r.Stok));
  lines.push('Pending invoices: ' + pending.length);
  pending.slice(0, 20).forEach(r => lines.push('- ' + r['No Inv'] + ' ' + r.Pelanggan + ' ' + r.Subtotal));
  const body = lines.join('\n');
  try {
    const email = Session.getEffectiveUser().getEmail();
    if (email) MailApp.sendEmail(email, 'SCASS daily — ' + hari, body);
  } catch (e) {}
  return { ok: true, jualHari: jualHari, beliHari: beliHari, kas: kas, rendah: rendah.length, pending: pending.length };
}

function scassJobStokRendah() {
  ensureErp_();
  const rendah = readTable_(ERP_INV).filter(r => (Number(r.Stok) || 0) <= (Number(r['Min Stok']) || 0));
  const sh = ensureSheet_('AutoLog', ['Waktu', 'Jenis', 'Ringkasan', 'Nilai']);
  const now = new Date();
  rendah.forEach(r => sh.appendRow([now, 'LOW_STOCK', r.SKU + ' ' + r.Nama, Number(r.Stok) || 0]));
  return { ok: true, count: rendah.length };
}

function scassJobPending() {
  ensureErp_();
  const pending = readTable_(ERP_INVCE).filter(r => String(r.Status || '').toUpperCase() === 'PENDING');
  const sh = ensureSheet_('AutoLog', ['Waktu', 'Jenis', 'Ringkasan', 'Nilai']);
  const now = new Date();
  pending.forEach(r => sh.appendRow([now, 'PENDING', String(r['No Inv'] || ''), Number(r.Subtotal) || 0]));
  return { ok: true, count: pending.length };
}

function ubahStatusInvoice(p) {
  ensureErp_();
  if (typeof clearErpCache_ === 'function') clearErpCache_();
  const no = String((p && p.noInv) || '').trim();
  let st = String((p && p.status) || '').trim().toUpperCase();
  if (!no) throw new Error('Invoice number required');
  if (st === 'PAID') st = 'LUNAS';
  if (st === 'BATAL') st = 'VOID';
  if (['PENDING','LUNAS','VOID'].indexOf(st) < 0) throw new Error('Invalid status');
  const sh = getSs_().getSheetByName(ERP_INVCE);
  const rows = readTable_(ERP_INVCE).filter(r => String(r['No Inv']) === no);
  if (!rows.length) throw new Error('Invoice not found');
  rows.forEach(r => sh.getRange(r._row, 9).setValue(st));
  if (st === 'LUNAS') {
    const keu = readTable_(ERP_KEU);
    const ada = keu.some(k => String(k.Ref) === no || String(k['No Bukti']) === no);
    if (!ada) {
      const total = rows.reduce((s, r) => s + (Number(r.Subtotal) || 0), 0);
      getSs_().getSheetByName(ERP_KEU).appendRow([new Date(), no, 'Masuk', 'Penjualan', 'Invoice ' + no, total, no]);
    }
  }
  return { ok: true, noInv: no, status: st };
}

const ERP_HTML = "<!DOCTYPE html>\n<html>\n<head>\n  <base target=\"_top\">\n  <meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">\n  <title>SCASS</title>\n  <style>\n    :root { --navy:#16324f; --teal:#1f7a6e; --card:#fff; --line:#c9dde0; --text:#16324f; --muted:#5b7380; }\n    * { box-sizing:border-box; }\n    body { margin:0; font-family:Arial,sans-serif; color:var(--text); background:linear-gradient(180deg,#e8f4ef,#e6f0f6); min-height:100vh; }\n    .erp { display:flex; min-height:100vh; }\n    .sidebar { width:248px; min-width:248px; flex-shrink:0; overflow:auto; background:linear-gradient(180deg,#0b4f37,#0d9488); color:#e8f2ef; display:flex; flex-direction:column; }\n    .side-brand { padding:18px 16px 12px; border-bottom:1px solid rgba(255,255,255,.08); }\n    .logo-wrap { background:#fff; border-radius:16px; padding:6px 10px; display:inline-block; }\n    .logo-img { display:block; width:64px; height:64px; object-fit:contain; }\n    .login-logo { padding:0; background:transparent !important; width:fit-content; align-self:center; }\n    #pageLogin .logo-wrap { background:transparent !important; display:block; width:fit-content; margin:0 auto 8px; padding:0; }\n    #pageLogin .logo-wrap img { width:132px !important; height:auto !important; display:block; margin:0 auto; }\n    \n    .login-stack { width:min(360px,88%); display:flex; flex-direction:column; align-items:stretch; }\n    .login-stack .logo-wrap, #pageLogin .logo-wrap { align-self:center; background:transparent !important; box-shadow:none !important; padding:0 !important; margin:0 auto 4px; border-radius:0; }\n    .login-stack .logo-wrap img, #pageLogin .logo-wrap img { width:108px !important; height:auto !important; background:transparent !important; mix-blend-mode:multiply; }\n    .login-stack .wordmark { text-align:center; margin:2px 0 20px; font-size:1.4rem; }\n    .login-row { display:flex; align-items:center; gap:12px; margin:0 0 10px; }\n    .login-row label { flex:0 0 78px; margin:0; font-size:13px; font-weight:600; color:#4b6570; text-align:right; }\n    .login-row input { flex:1; width:auto; margin:0; padding:9px 12px; border:1px solid #d5e4e8; border-radius:10px; font-size:14px; background:#fff; box-sizing:border-box; }\n    .login-stack .btn { width:100%; box-sizing:border-box; margin-top:8px; padding:11px 12px; }\n    .login-stack #loginInfo { text-align:center; min-height:18px; }\n\n    .login-logo .logo-img, .login-logo-img { width:180px; height:180px; margin:0 auto; object-fit:contain; display:block; }\n    .side-logo-img { width:64px; height:64px; object-fit:contain; display:block; }\n    .side-name { font-weight:800; margin-top:10px; font-size:18px; }\n    .side-sub { font-size:12px; color:#9fb6c0; margin-top:2px; }\n    .side-nav { padding:12px 10px; display:flex; flex-direction:column; gap:4px; flex:1; }\n    .nav { background:transparent; color:#cfe0e4; border:0; text-align:left; padding:10px 12px; border-radius:10px; cursor:pointer; font-size:14px; }\n    .nav:hover { background:rgba(255,255,255,.08); }\n    .nav.on { background:#14b8a6; color:#08352c; font-weight:700; }\n    .nav .hint { display:block; font-size:11px; color:#8aa3ad; font-weight:400; }\n    .side-foot { padding:12px; border-top:1px solid rgba(255,255,255,.08); }\n    .erp-main { flex:1; min-width:0; padding:16px; }\n    .crumb b { color:var(--navy); font-size:18px; }\n    .crumb { color:var(--muted); font-size:13px; margin-bottom:12px; }\n    .mod-grid { display:grid; grid-template-columns:repeat(4,1fr); gap:10px; }\n    .card { background:#fff; border:1px solid var(--line); border-radius:16px; padding:16px; box-shadow:0 8px 24px rgba(22,50,79,.08); margin-bottom:10px; }\n    .kpi .lbl { font-size:12px; color:var(--muted); text-transform:uppercase; }\n    .kpi .val { font-size:22px; font-weight:800; color:var(--navy); margin-top:4px; }\n    .btn, .btn2, .btn3, input, select { font-size:16px; padding:10px; border-radius:10px; }\n    .btn, .btn2, .btn3 { border:0; cursor:pointer; font-weight:700; }\n    .btn { background:#0d9488; color:#fff; }\n    .btn2 { background:#0f766e; color:#fff; }\n    .btn3 { background:#8a2a2a; color:#fff; }\n    .row { display:flex; flex-wrap:wrap; gap:8px; }\n    input, select { width:100%; border:1px solid var(--line); background:#f4f8fa; margin:6px 0; }\n    table { width:100%; border-collapse:collapse; font-size:13px; }\n    th { text-align:left; border-bottom:1px solid var(--line); padding:8px 4px; }\n    td { padding:8px 4px; border-bottom:1px solid #eef4f6; }\n    .muted { color:var(--muted); font-size:13px; }\n    .loginbox { max-width:420px; margin:8vh auto; text-align:center; }\n    \n    #pageLogin.hide, #pageLogin[style*=\"display: none\"], #pageLogin[style*=\"display:none\"] { display:none !important; }\n    #pageLogin { position:fixed; inset:0; display:flex; flex-direction:row; flex-wrap:nowrap !important; align-items:stretch !important; background:#eef7f5; z-index:20; padding:0; margin:0; width:100% !important; height:100% !important; }\n    .login-brand-panel { flex:7 1 70% !important; min-width:0; min-height:100% !important; background:linear-gradient(145deg,#0b4f37 0%,#0d9488 50%,#14b8a6 100%); color:#fff; display:flex; flex-direction:column; justify-content:center; padding:3rem 3.5rem; position:relative; overflow:hidden; }\n    .login-brand-panel::before { content:\"\"; position:absolute; inset:-20%; background:radial-gradient(circle at 20% 30%,rgba(255,255,255,.12) 0%,transparent 40%),radial-gradient(circle at 80% 70%,rgba(255,255,255,.08) 0%,transparent 45%); pointer-events:none; }\n    .login-brand-badge { display:inline-flex; align-items:center; gap:.4rem; background:rgba(255,255,255,.22); border-radius:999px; padding:.5rem 1.1rem; font-size:.95rem; font-weight:700; letter-spacing:.08em; margin-bottom:1.6rem; position:relative; z-index:1; width:fit-content; }\n    .login-brand-panel h1 { font-size:clamp(2.4rem,4.2vw,3.4rem); font-weight:800; line-height:1.12; margin:0 0 1.1rem; position:relative; z-index:1; letter-spacing:-.03em; }\n    .login-brand-panel p { font-size:1.18rem; opacity:.95; line-height:1.55; max-width:460px; position:relative; z-index:1; margin:0; font-weight:400; }\n    \n    #pageLogin .card.loginbox { flex:3 1 30% !important; max-width:30% !important; min-width:300px; max-width:none !important; width:auto !important; margin:0 !important; min-height:100%; display:flex; flex-direction:column; justify-content:center; align-items:center; background:transparent !important; box-shadow:none !important; border:0 !important; border-radius:0 !important; padding:2.2rem 3rem !important; text-align:center; }\n    #pageLogin .card.loginbox input { width:100%; padding:.65rem .85rem; border:1px solid #d5e4e8; border-radius:10px; font-size:.9rem; background:#fff; }\n    #pageLogin .card.loginbox input:focus { border-color:#0d9488; box-shadow:0 0 0 3px rgba(13,148,136,.18); outline:none; }\n    #pageLogin .card.loginbox .btn { width:100%; margin-top:1.25rem; padding:.8rem; background:#0d9488 !important; color:#fff !important; border:none; border-radius:10px; font-weight:700; }\n    #pageLogin .card.loginbox label { display:block; font-size:.78rem; font-weight:600; color:#5b7380; margin:.85rem 0 .3rem; text-align:left; }\n    @media (max-width:1100px){\n      #pageLogin{ flex-direction:column !important; }\n      .login-brand-panel{\n        flex:0 0 28% !important; min-height:0 !important; max-height:30vh !important;\n        padding:16px 18px !important; justify-content:flex-end !important;\n      }\n      .login-brand-badge{ margin-bottom:8px !important; font-size:.72rem !important; }\n      .login-brand-panel h1{ font-size:1.55rem !important; margin:0 0 6px !important; }\n      .login-brand-panel p{ font-size:.86rem !important; max-width:100% !important; }\n      #pageLogin .card.loginbox{\n        flex:1 1 72% !important; max-width:100% !important; min-width:0 !important;\n        width:100% !important; min-height:0 !important; padding:18px 20px 28px !important;\n        justify-content:flex-start !important;\n      }\n      .login-stack{ width:min(440px,94%) !important; }\n      .login-row{ margin:0 0 14px !important; }\n      .login-row label{ flex-basis:88px !important; font-size:15px !important; }\n      .login-row input{ padding:14px 12px !important; font-size:16px !important; min-height:48px; }\n      .login-stack .btn{ padding:14px !important; font-size:16px !important; min-height:48px; }\n      #pageLogin .logo-wrap img, .login-stack .logo-wrap img{ width:96px !important; }\n    }\n\n    \n    .item-card { position:relative; background:#f8fafb; border:1px solid #d7e8e4; border-radius:14px; padding:14px 14px 12px; margin:0 0 12px; }\n    .item-card .item-x { position:absolute; top:12px; right:12px; width:36px; height:36px; border:0; border-radius:10px; background:#fb7185; color:#fff; font-weight:800; cursor:pointer; }\n    .item-card .sku-row select, .item-card .sku-row input { width:100%; }\n    .item-grid { display:grid; grid-template-columns:1fr 1fr 1fr; gap:10px; margin-top:8px; }\n    .item-grid label, .item-card > label { display:block; font-size:12px; font-weight:700; color:#5b7380; margin:0 0 4px; text-align:left; }\n    @media (max-width:700px){ .item-grid{grid-template-columns:1fr;} }\n    .table-wrap { overflow:auto; max-height:70vh; }\n    .table-wrap table { min-width:980px; }\n    .table-wrap thead th { position:sticky; top:0; z-index:1; }\n    .modal-back { display:none; position:fixed; inset:0; background:rgba(11,79,55,.45); z-index:120; align-items:center; justify-content:center; padding:18px; }\n    .modal-back.show { display:flex; }\n    .modal-card { background:#fff; width:min(560px,100%); max-height:90vh; overflow:auto; border-radius:16px; padding:20px 20px 16px; box-shadow:0 24px 60px rgba(0,0,0,.25); }\n    .modal-card h3 { margin:0 0 12px; color:#0b4f37; }\n    .modal-grid { display:grid; grid-template-columns:1fr 1fr; gap:8px; }\n    .modal-card label { display:block; font-size:12px; font-weight:700; color:#5b7380; margin:6px 0 3px; text-align:left; }\n    .wordmark { font-weight:800; color:#0b4f37; font-size:34px; margin:10px 0 4px; letter-spacing:-.02em; }\n    .statuserr { color:#b42318; }\n    .statusok { color:#1f7a6e; }\n    .icon-pen { width:36px;height:36px;border:0;border-radius:10px;background:#0d9488;color:#fff;font-size:16px;cursor:pointer; }\n    #overlay { display:none; position:fixed; inset:0; background:rgba(22,50,79,.55); z-index:99; align-items:center; justify-content:center; color:#fff; text-align:center; }\n    #overlay.show { display:flex; }\n    .spin { width:36px; height:36px; border:4px solid #9bb; border-top-color:#2fa37a; border-radius:50%; margin:0 auto 12px; animation:r .8s linear infinite; }\n    @keyframes r { to { transform:rotate(360deg); } }\n    .sidebar img { max-width:100%; height:auto; }\n    .erp { align-items:stretch; }\n    @media (max-width:900px){\n      .mod-grid{grid-template-columns:1fr !important;}\n      .erp{flex-direction:row !important; align-items:stretch;}\n      .sidebar{width:156px !important; min-width:156px !important; min-height:100vh;}\n      .side-brand{padding:12px 10px 8px;}\n      .side-name{font-size:16px;}\n      .nav{padding:8px 10px !important; font-size:13px !important;}\n      .side-nav{flex-direction:column !important; flex-wrap:nowrap !important;}\n      .erp-main{min-width:0; padding:10px;}\n      .row{flex-direction:column !important; flex-wrap:nowrap !important;}\n      .row input, .row select, .row button, .row .btn, .row .btn2, .row .btn3{\n        width:100% !important; flex:none !important; max-width:100% !important;\n      }\n      .item-grid, .item-grid[style]{grid-template-columns:1fr !important;}\n      .card{padding:12px;}\n      input, select, .btn, .btn2, .btn3{font-size:16px !important;}\n    }\n  </style>\n</head>\n<body>\n  <div id=\"overlay\"><div><div class=\"spin\"></div><div id=\"overlayText\">Memproses...</div></div></div>\n<div id=\"invceModal\" class=\"modal-back\" onclick=\"if(event.target.id==='invceModal')closeInvce()\">\n  <div class=\"modal-card\">\n    <h3>Invoice</h3>\n    <label>No Inv</label><input id=\"ivNo\" readonly>\n    <label>Customer</label><input id=\"ivCus\" readonly>\n    <div class=\"modal-grid\">\n      <div><label>Date</label><input id=\"ivTgl\" readonly></div>\n      <div><label>Status</label>\n        <select id=\"ivStatus\">\n          <option value=\"PENDING\">PENDING</option>\n          <option value=\"LUNAS\">PAID</option>\n          <option value=\"VOID\">VOID</option>\n        </select>\n      </div>\n    </div>\n    <div class=\"muted\" id=\"ivItems\" style=\"margin:10px 0\"></div>\n    <div class=\"row\" style=\"margin-top:14px\">\n      <button class=\"btn\" onclick=\"saveInvceStatus()\">Save</button>\n      <button class=\"btn2\" type=\"button\" onclick=\"pdfInv(document.getElementById('ivNo').value)\">PDF</button>\n      <button class=\"btn3\" type=\"button\" onclick=\"closeInvce()\">Close</button>\n    </div>\n  </div>\n</div>\n<div id=\"editModal\" class=\"modal-back\" onclick=\"if(event.target.id==='editModal')closeEdit()\">\n  <div class=\"modal-card\">\n    <h3 id=\"editTitle\">Edit item</h3>\n    <label>SKU</label><input id=\"eSku\" readonly>\n    <label>Product Name</label><input id=\"eNama\">\n    <div class=\"modal-grid\">\n      <div><label>Category</label><input id=\"eKat\"></div>\n      <div><label>Uom</label><input id=\"eSat\"></div>\n      <div><label>Qty / Stock</label><input id=\"eStok\" type=\"number\"></div>\n      <div><label>Min stock</label><input id=\"eMin\" type=\"number\"></div>\n      <div><label>Cost</label><input id=\"eBeli\" type=\"number\"></div>\n      <div><label>Price</label><input id=\"eJual\" type=\"number\"></div>\n    </div>\n    <div class=\"row\" style=\"margin-top:14px\">\n      <button class=\"btn2\" onclick=\"saveEdit()\">Save</button>\n      <button class=\"btn3\" onclick=\"closeEdit()\">Cancel</button>\n    </div>\n  </div>\n</div>\n  <div id=\"pageLogin\">\n  <div class=\"login-brand-panel\">\n    <div class=\"login-brand-badge\">SCASS \u00b7 BUILD 17</div>\n    <h1>Capital with clarity.</h1>\n    <p>One workspace for inventory, purchasing, invoices, and cash \u2014 built to stay simple as the business grows.</p>\n  </div>\n  <div class=\"card loginbox\">\n    <div class=\"login-stack\">\n    <div class=\"logo-wrap login-logo\"><img src=\"data:image/jpeg;base64,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\" alt=\"SCASS\" style=\"width:120px;height:auto;display:block;margin:0 auto;background:transparent\"></div>\n    <div class=\"wordmark\">SCASS.com</div>\n    <div class=\"login-row\"><label>ID</label><input id=\"loginId\" placeholder=\"\"></div>\n    <div class=\"login-row\"><label>Password</label><input id=\"loginPass\" type=\"password\" placeholder=\"\"></div>\n    <button class=\"btn\" onclick=\"doLogin()\">Sign in</button>\n    <div id=\"loginInfo\" class=\"muted\"></div>\n    </div>\n  </div>\n</div>\n<div id=\"pageApp\" class=\"erp\" style=\"display:none\">\n    <aside class=\"sidebar\">\n      <div class=\"side-brand\">\n        <div class=\"logo-wrap\" style=\"background:#fff\"><img src=\"data:image/jpeg;base64,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\" alt=\"SCASS\" style=\"width:64px;height:auto;display:block\"></div>\n        <div class=\"side-name\">SCASS</div>\n        <div class=\"side-sub\">NewStar Kancing</div>\n      </div>\n      <nav class=\"side-nav\">\n        <button class=\"nav on\" id=\"n-dash\" onclick=\"show('dash')\">Dashboard<div class=\"hint\">Overview</div></button>\n        <button class=\"nav\" id=\"n-inv\" onclick=\"show('inv')\">Inventory<div class=\"hint\">Stock</div></button>\n        <button class=\"nav\" id=\"n-beli\" onclick=\"show('beli')\">Purchases<div class=\"hint\">Supplier receipts</div></button>\n        <button class=\"nav\" id=\"n-jual\" onclick=\"show('jual')\">Sales<div class=\"hint\">Invoices</div></button>\n        <button class=\"nav\" id=\"n-keu\" onclick=\"show('keu')\">Finance<div class=\"hint\">Cash in / out</div></button>\n        <button class=\"nav\" id=\"n-lap\" onclick=\"show('lap')\">Reports<div class=\"hint\">Internal summary</div></button>\n        <button class=\"nav\" id=\"n-master\" onclick=\"show('master')\">Master Data<div class=\"hint\">Partners</div></button>\n        <button class=\"nav\" id=\"n-set\" onclick=\"show('set')\">Settings<div class=\"hint\">System</div></button>\n      </nav>\n      <div class=\"side-foot\">\n        <div class=\"muted\" style=\"color:#9fb6c0;margin:8px 0\">USER / USER</div>\n        <button class=\"btn3\" style=\"width:100%\" onclick=\"logout()\">Sign out</button>\n      </div>\n    </aside>\n    <section class=\"erp-main\">\n      <div class=\"crumb\"><b id=\"ttl\">Dashboard</b><div id=\"sub\">Ringkasan operasional</div></div>\n      <div id=\"m-dash\">\n        <div class=\"mod-grid\">\n          <div class=\"card kpi\"><div class=\"lbl\">SKU aktif</div><div class=\"val\" id=\"kSku\">0</div></div>\n          <div class=\"card kpi\"><div class=\"lbl\">Nilai stok</div><div class=\"val\" id=\"kStok\">Rp 0</div></div>\n          <div class=\"card kpi\"><div class=\"lbl\">Stok rendah</div><div class=\"val\" id=\"kLow\">0</div></div>\n          <div class=\"card kpi\"><div class=\"lbl\">Saldo kas</div><div class=\"val\" id=\"kKas\">Rp 0</div></div>\n          <div class=\"card kpi\"><div class=\"lbl\">Pembelian</div><div class=\"val\" id=\"kBeli\">Rp 0</div></div>\n          <div class=\"card kpi\"><div class=\"lbl\">Sales</div><div class=\"val\" id=\"kJual\">Rp 0</div></div>\n          <div class=\"card kpi\"><div class=\"lbl\">Invoice hari ini</div><div class=\"val\" id=\"kPos\">Rp 0</div></div>\n          <div class=\"card kpi\"><div class=\"lbl\">Pending</div><div class=\"val\" id=\"kPub\">0</div></div>\n        </div>\n        <div class=\"card\"><b>Stok di bawah minimum</b><div id=\"dashLow\" class=\"muted\">\u2014</div></div>\n      </div>\n      <div id=\"m-inv\" style=\"display:none\">\n        <div class=\"card\">\n          <b>Tambah / ubah barang</b>\n          <div class=\"row\">\n            <input id=\"bSku\" placeholder=\"SKU (kosongkan = otomatis)\" style=\"flex:1\">\n            <input id=\"bNama\" placeholder=\"Nama barang\" style=\"flex:1\">\n          </div>\n          <div class=\"row\">\n            <input id=\"bKat\" placeholder=\"Kategori\" style=\"flex:1\">\n            <input id=\"bSat\" placeholder=\"Satuan (pcs/porsi)\" style=\"flex:1\">\n          </div>\n          <div class=\"row\">\n            <input id=\"bStok\" type=\"number\" placeholder=\"Stok\" style=\"flex:1\">\n            <input id=\"bBeli\" type=\"number\" placeholder=\"Harga beli\" style=\"flex:1\">\n            <input id=\"bJual\" type=\"number\" placeholder=\"Harga jual\" style=\"flex:1\">\n            <input id=\"bMin\" type=\"number\" placeholder=\"Min stok\" style=\"flex:1\">\n          </div>\n          <div class=\"row\"><button class=\"btn\" onclick=\"saveBarang()\">Simpan barang</button><button class=\"btn2\" onclick=\"exportCsv('inventory')\">Export CSV</button></div>\n          <div id=\"invInfo\" class=\"muted\"></div>\n        </div>\n        <div class=\"card\">\n          <b>Search</b>\n          <div class=\"item-grid\" style=\"grid-template-columns:1fr 1fr 1.4fr 1fr\">\n            <div><label>Category</label><select id=\"fKat\"></select></div>\n            <div><label>Uom</label><select id=\"fUom\"></select></div>\n            <div><label>Product Name / SKU</label><input id=\"invQ\" placeholder=\"Search SKU or name...\"></div>\n            <div><label>Status</label><select id=\"fStat\"><option value=\"\">Show All</option><option value=\"LOW\">LOW</option><option value=\"OK\">OK</option></select></div>\n          </div>\n          <div class=\"row\" style=\"margin-top:10px\">\n            <button class=\"btn2\" onclick=\"paintInv()\">Search</button>\n            <button class=\"btn\" onclick=\"exportCsv('inventory')\">Export Data</button>\n            <button class=\"btn3\" onclick=\"clearInvFilter()\">Clear Filter</button>\n          </div>\n        </div>\n        <div class=\"card\"><b>Inventory</b><div id=\"tblInv\" class=\"table-wrap\"></div></div>\n      </div>\n      <div id=\"m-beli\" style=\"display:none\">\n        <div class=\"card\">\n          <b>Penerimaan barang</b>\n          <select id=\"poSup\"></select>\n          <div id=\"poLines\"></div>\n          <button class=\"btn2\" type=\"button\" onclick=\"addPoLine()\">+ Item</button>\n          <button class=\"btn\" onclick=\"savePo()\">Save purchase</button>\n          <div id=\"poInfo\" class=\"muted\"></div>\n        </div>\n        <div class=\"card\"><b>Riwayat pembelian</b><div id=\"tblPo\"></div></div>\n      </div>\n      <div id=\"m-jual\" style=\"display:none\">\n        <div class=\"card\">\n          <b>Invoice penjualan</b>\n          <select id=\"invCus\"></select>\n          <div id=\"soLines\"></div>\n          <button class=\"btn2\" type=\"button\" onclick=\"addSoLine()\">+ Item</button>\n          <select id=\"soMetode\"><option>Cash</option><option>Transfer</option><option>QRIS</option><option>Credit</option></select>\n          <select id=\"soStatus\"><option value=\"PENDING\">PENDING</option><option value=\"LUNAS\">PAID</option><option value=\"BATAL\">VOID</option></select>\n          <input id=\"soBayar\" type=\"number\" placeholder=\"Amount paid\">\n          <div class=\"row\">\n            <input id=\"soDiskonPct\" type=\"number\" placeholder=\"Discount %\" style=\"flex:1\">\n            <input id=\"soDiskonAmt\" type=\"number\" placeholder=\"Discount amount\" style=\"flex:1\">\n            <div class=\"muted\" id=\"soTotalHint\" style=\"flex:1;padding-top:10px\"></div>\n          </div>\n          <button class=\"btn\" onclick=\"saveInv()\">Save invoice</button>\n          <div id=\"invcInfo\" class=\"muted\"></div>\n        </div>\n        <div class=\"card\"><b>Riwayat invoice</b>\n        <input id=\"invSearch\" placeholder=\"Search invoice number...\" oninput=\"renderInvce()\">\n        <div id=\"tblInvce\"></div></div>\n      </div>\n      <div id=\"m-keu\" style=\"display:none\">\n        <div class=\"card\">\n          <b>Kas masuk / keluar</b>\n          <select id=\"kJenis\"><option>Masuk</option><option>Keluar</option></select>\n          <input id=\"kAkun\" placeholder=\"Akun (Kas, Bank, Operasional)\">\n          <input id=\"kKet\" placeholder=\"Keterangan\">\n          <input id=\"kJum\" type=\"number\" placeholder=\"Jumlah\">\n          <button class=\"btn\" onclick=\"saveKeu()\">Simpan transaksi</button>\n          <div id=\"keuInfo\" class=\"muted\"></div>\n        </div>\n        <div class=\"card\"><b>Buku kas</b><div id=\"tblKeu\"></div></div>\n      </div>\n      <div id=\"m-lap\" style=\"display:none\">\n        <div class=\"card\">\n          <b>Export</b>\n          <p class=\"muted\">Unduh laporan dari sistem.</p>\n          <div class=\"row\">\n            <button class=\"btn2\" onclick=\"exportExcel()\">Export Excel semua data</button>\n            <button class=\"btn\" onclick=\"exportCsv('inventory')\">CSV Inventory</button>\n            <button class=\"btn\" onclick=\"exportCsv('supplier')\">CSV Supplier</button>\n            <button class=\"btn\" onclick=\"exportCsv('pelanggan')\">CSV Pelanggan</button>\n            <button class=\"btn\" onclick=\"exportCsv('pembelian')\">CSV Pembelian</button>\n            <button class=\"btn\" onclick=\"exportCsv('invoice')\">CSV Invoice</button>\n            <button class=\"btn\" onclick=\"exportCsv('keuangan')\">CSV Keuangan</button>\n          </div>\n          <div id=\"expInfo\" class=\"muted\"></div>\n        </div>\n        <div class=\"card\">\n          <b>Laporan ringkas</b>\n          <div class=\"muted\">Dihitung dari Inventory, Pembelian, Invoice, dan Keuangan.</div>\n          <div id=\"lapBox\"></div>\n        </div>\n      </div>\n      <div id=\"m-master\" style=\"display:none\">\n        <div class=\"card\">\n          <b>Supplier / Pelanggan</b>\n          <select id=\"mJenis\"><option value=\"supplier\">Supplier</option><option value=\"pelanggan\">Pelanggan</option></select>\n          <input id=\"mNama\" placeholder=\"Nama\">\n          <input id=\"mTel\" placeholder=\"Telepon\">\n          <input id=\"mAlamat\" placeholder=\"Alamat\">\n          <button class=\"btn\" onclick=\"saveMitra()\">Simpan mitra</button>\n          <div id=\"mitInfo\" class=\"muted\"></div>\n        </div>\n        <div class=\"card\"><b>Supplier</b><div id=\"tblSup\"></div></div>\n        <div class=\"card\"><b>Pelanggan</b><div id=\"tblCus\"></div></div>\n      </div>\n      <div id=\"m-set\" style=\"display:none\">\n        <div class=\"card\">\n          <b>Settings</b>\n          <p class=\"muted\">NewStar Kancing</p>\n          <label>Language</label>\n          <select id=\"langSel\">\n            <option value=\"en\">English</option>\n            <option value=\"id\">Bahasa Indonesia</option>\n            <option value=\"zh\">\u4e2d\u6587</option>\n            <option value=\"ar\">\u0627\u0644\u0639\u0631\u0628\u064a\u0629</option>\n          </select>\n          <button class=\"btn2\" onclick=\"applyLang()\">Apply language</button>\n          <div id=\"langInfo\" class=\"muted\"></div>\n        </div>\n        <div class=\"card\">\n          <b>Import CSV</b>\n          <p class=\"muted\">Upload file CSV (baris pertama = header). Inventory di-update berdasarkan SKU. Supplier/Pelanggan berdasarkan Kode. Transaksi ditambahkan.</p>\n          <select id=\"impJenis\">\n            <option value=\"inventory\">Inventory</option>\n            <option value=\"supplier\">Supplier</option>\n            <option value=\"pelanggan\">Pelanggan</option>\n            <option value=\"pembelian\">Pembelian</option>\n            <option value=\"invoice\">Invoice</option>\n            <option value=\"keuangan\">Keuangan</option>\n          </select>\n          <input id=\"impFile\" type=\"file\" accept=\".csv,text/csv\">\n          <button class=\"btn2\" onclick=\"importCsv()\">Import ke sistem</button>\n          <div id=\"impInfo\" class=\"muted\"></div>\n        </div>\n      </div>\n    </section>\n  </div>\n<script>\n\n  var API_URL = 'https://script.google.com/macros/s/AKfycbyxWOQQ_cVvZ_1nlvW9Zt0sABqfbTXHgc9uoZw4JR_atTPo3iVaNDZg_6A1PhDMwkoQ/exec';\n  var API_TOKEN = 'NSK-SCASS-2026';\n  function jsonpApi(action, payload){\n    return new Promise(function(resolve, reject){\n      var cb = 'scassCb' + Date.now() + Math.floor(Math.random()*1e6);\n      var done = false;\n      window[cb] = function(res){\n        done = true; try{ delete window[cb]; }catch(e){}\n        if (s && s.parentNode) s.parentNode.removeChild(s);\n        if (res && res.ok === false) reject(new Error(res.error || 'API error'));\n        else resolve(res);\n      };\n      var s = document.createElement('script');\n      s.src = API_URL + '?action=' + encodeURIComponent(action)\n        + '&token=' + encodeURIComponent(API_TOKEN)\n        + '&payload=' + encodeURIComponent(JSON.stringify(payload || {}))\n        + '&callback=' + cb;\n      s.onerror = function(){ if(!done) reject(new Error('Network error')); };\n      document.head.appendChild(s);\n      setTimeout(function(){ if(!done) reject(new Error('API timeout')); }, 28000);\n    });\n  }\n  if (!(window.google && google.script && google.script.run)) {\n    window.google = { script: { run: {\n      withSuccessHandler: function(ok){\n        return { withFailureHandler: function(fail){\n          function wrap(name){\n            return function(payload){\n              jsonpApi(name, payload).then(ok).catch(fail);\n            };\n          }\n          return {\n            getErpDashboard: wrap('getErpDashboard'),\n            getErpData: wrap('getErpData'),\n            simpanBarang: wrap('simpanBarang'),\n            ubahStok: wrap('ubahStok'),\n            simpanPembelian: wrap('simpanPembelian'),\n            simpanInvoice: wrap('simpanInvoice'),\n            simpanKeuangan: wrap('simpanKeuangan'),\n            simpanMitra: wrap('simpanMitra'),\n            exportErpExcel: wrap('exportErpExcel'),\n            importErpRows: wrap('importErpRows'),\n            exportInvoicePdf: wrap('exportInvoicePdf'), ubahStatusInvoice: wrap('ubahStatusInvoice'),\n            ubahStatusInvoice: wrap('ubahStatusInvoice')\n          };\n        }};\n      }\n    }}};\n  }\n  let data = { inventory:[], supplier:[], pelanggan:[], pembelian:[], invoice:[], keuangan:[] };\n  const META = {\n    dash:{t:'Dashboard',s:'Ringkasan operasional'},\n    inv:{t:'Inventory',s:'Stok dan harga barang'},\n    beli:{t:'Pembelian',s:'Penerimaan dari supplier'},\n    jual:{t:'Penjualan',s:'Invoice pelanggan'},\n    keu:{t:'Keuangan',s:'Kas masuk dan keluar'},\n    lap:{t:'Laporan',s:'Rekap ERP'},\n    master:{t:'Master Data',s:'Supplier dan pelanggan'},\n    set:{t:'Pengaturan',s:'Akun dan sistem'}\n  };\n  function rp(n){ return 'Rp ' + Number(n||0).toLocaleString('id-ID'); }\n  var busyTimer = null;\n  function busy(on, text){\n    document.getElementById('overlayText').innerText = text || 'Loading...';\n    document.getElementById('overlay').className = on ? 'show' : '';\n    if (busyTimer) clearTimeout(busyTimer);\n    if (on) busyTimer = setTimeout(function(){ document.getElementById('overlay').className=''; }, 12000);\n  }\n  function applyLang(){\n    var v = document.getElementById('langSel').value;\n    busy(true, 'Changing language...');\n    setTimeout(function(){\n      setLang(v);\n      busy(false);\n      var info = document.getElementById('langInfo');\n      if (info) { info.className='statusok'; info.innerText='Language updated'; }\n    }, 400);\n  }\n  function show(name){\n    Object.keys(META).forEach(function(k){\n      document.getElementById('m-'+k).style.display = k===name ? 'block' : 'none';\n      document.getElementById('n-'+k).className = 'nav' + (k===name ? ' on' : '');\n    });\n    document.getElementById('ttl').innerText = META[name].t;\n    document.getElementById('sub').innerText = META[name].s;\n    if (data && data._ready) {\n      if (name==='dash') paintDashFromCache();\n      else renderAll();\n    } else {\n      loadData(true);\n    }\n  }\n  function okLogin(role){\n    localStorage.setItem('scass_login','1');\n    localStorage.setItem('scass_role', role);\n    document.getElementById('pageLogin').classList.add('hide');\n    document.getElementById('pageLogin').style.display='none';\n    document.getElementById('pageApp').style.display='flex';\n    loadData(true);\n  }\n  function doLogin(){\n    const id = (document.getElementById('loginId').value||'').trim();\n    const pass = (document.getElementById('loginPass').value||'').trim();\n    const idU = id.toUpperCase();\n    if (idU==='USER' && pass.toUpperCase()==='USER') return okLogin('USER');\n    if (idU==='ADMIN' && pass==='admin1234') return okLogin('ADMIN');\n    if (idU==='TRIAL' && pass.toUpperCase()==='TRIAL') return okLogin('TRIAL');\n    document.getElementById('loginInfo').className='statuserr';\n    document.getElementById('loginInfo').innerText='LOGIN TIDAK DI TEMUKAN';\n  }\n  function logout(){\n    localStorage.removeItem('scass_login');\n    document.getElementById('pageApp').style.display='none';\n    document.getElementById('pageLogin').classList.remove('hide'); document.getElementById('pageLogin').style.display='flex';\n  }\n  function loadDash(){\n    busy(true, 'Memuat dashboard...');\n    google.script.run.withSuccessHandler(function(res){\n      busy(false);\n      document.getElementById('kSku').innerText = res.kpi.sku;\n      document.getElementById('kStok').innerText = rp(res.kpi.stokNilai);\n      document.getElementById('kLow').innerText = res.kpi.stokRendah;\n      document.getElementById('kKas').innerText = rp(res.kpi.kas);\n      document.getElementById('kBeli').innerText = rp(res.kpi.pembelian);\n      document.getElementById('kJual').innerText = rp(res.kpi.penjualan);\n      document.getElementById('kPos').innerText = rp(res.kpi.posHari);\n      document.getElementById('kPub').innerText = res.kpi.pending || 0;\n      if (document.getElementById('setSheet')) document.getElementById('setSheet').innerText = 'NewStar Kancing';\n      document.getElementById('dashLow').innerHTML = (res.rendah||[]).length\n        ? '<table><tr><th>SKU</th><th>Nama</th><th>Stok</th><th>Min</th></tr>' +\n          res.rendah.map(function(r){ return '<tr><td>'+r.SKU+'</td><td>'+r.Nama+'</td><td>'+r.Stok+'</td><td>'+r['Min Stok']+'</td></tr>'; }).join('') + '</table>'\n        : 'Tidak ada stok rendah.';\n    }).withFailureHandler(function(err){ busy(false); alert(err.message||err); }).getErpDashboard();\n  }\n  function paintDashFromCache(){\n    var inv=data.inventory||[], po=data.pembelian||[], invce=data.invoice||[], keu=data.keuangan||[];\n    var nilai=0, rendah=0, low=[];\n    inv.forEach(function(r){\n      var stok=Number(r.Stok||r.Qty||0), harga=Number(r['Harga Beli']||r['Harga Jual']||0), min=Number(r['Min Stok']||0);\n      nilai+=stok*harga;\n      if(String(r.Stok||r.Qty||'')!=='' && stok<=min){ rendah++; if(low.length<8) low.push(r); }\n    });\n    var beli=po.reduce(function(s,r){return s+(Number(r.Subtotal)||0);},0);\n    var jual=invce.reduce(function(s,r){return s+(Number(r.Subtotal)||0);},0);\n    var kas=0;\n    keu.forEach(function(r){ var n=Number(r.Jumlah)||0; kas += String(r.Jenis||'').toLowerCase()==='keluar'?-n:n; });\n    var pending=invce.filter(function(r){return String(r.Status||'').toUpperCase()==='PENDING';}).length;\n    if(document.getElementById('kSku')){\n      kSku.innerText=inv.length; kStok.innerText=rp(nilai); kLow.innerText=rendah; kKas.innerText=rp(kas);\n      kBeli.innerText=rp(beli); kJual.innerText=rp(jual); kPub.innerText=pending;\n    }\n    if(document.getElementById('dashLow')) document.getElementById('dashLow').innerHTML = low.length? table(low,['SKU','Nama','Stok','Min Stok']) : '\u2014';\n  }\n  function loadData(silent){\n    if (!silent) busy(true, 'Memuat data...');\n    google.script.run.withSuccessHandler(function(res){\n      busy(false);\n      data = res || {}; data._ready = true;\n      renderAll();\n      paintDashFromCache();\n    }).withFailureHandler(function(err){ busy(false); alert(err.message||err); }).getErpData();\n  }\n  function opt(list, val, label){\n    return '<option value=\"\">Pilih</option>' + list.map(function(r){\n      return '<option value=\"'+r[val]+'\">'+r[label]+' ('+r[val]+')</option>';\n    }).join('');\n  }\n  function renderAll(){\n    if (document.getElementById('soLines') && !document.getElementById('soLines').children.length) addSoLine();\n    if (document.getElementById('poLines') && !document.getElementById('poLines').children.length) addPoLine();\n\n    paintInv();\n    document.getElementById('tblPo').innerHTML = table(data.pembelian, ['Waktu','No PO','Supplier','SKU','Nama','Qty','Harga','Subtotal','Status']);\n    renderInvce();\n    document.getElementById('tblKeu').innerHTML = table(data.keuangan, ['Waktu','No Bukti','Jenis','Akun','Keterangan','Jumlah','Ref']);\n    document.getElementById('tblSup').innerHTML = table(data.supplier, ['Kode','Nama','Telepon','Alamat']);\n    document.getElementById('tblCus').innerHTML = table(data.pelanggan, ['Kode','Nama','Telepon','Alamat']);\n    document.getElementById('poSup').innerHTML = opt(data.supplier,'Nama','Nama');\n    document.getElementById('invCus').innerHTML = opt(data.pelanggan,'Nama','Nama');\n    document.getElementById('poSku').innerHTML = opt(data.inventory,'SKU','Nama');\n    document.getElementById('invSku').innerHTML = opt(data.inventory,'SKU','Nama');\n    const beli = data.pembelian.reduce(function(s,r){ return s+(Number(r.Subtotal)||0); },0);\n    const jual = data.invoice.reduce(function(s,r){ return s+(Number(r.Subtotal)||0); },0);\n    document.getElementById('lapBox').innerHTML = '<p>Total pembelian: <b>'+rp(beli)+'</b></p><p>Total invoice: <b>'+rp(jual)+'</b></p><p>SKU: <b>'+data.inventory.length+'</b></p>';\n  }\n  function table(rows, cols, extra){\n    if (!rows || !rows.length) return '<div class=\"muted\">No records yet.</div>';\n    return '<table><tr>' + cols.map(function(c){ return '<th>'+c+'</th>'; }).join('') + (extra?'<th></th>':'') + '</tr>' +\n      rows.map(function(r){\n        return '<tr>' + cols.map(function(c){\n          const v = r[c];\n          return '<td>' + (typeof v==='number' && (c.indexOf('Harga')>=0 || c==='Subtotal' || c==='Jumlah') ? rp(v) : (v==null?'':v)) + '</td>';\n        }).join('') + (extra?'<td>'+extra(r)+'</td>':'') + '</tr>';\n      }).join('') + '</table>';\n  }\n  function isiPoHarga(){\n    const sku = document.getElementById('poSku').value;\n    const it = data.inventory.find(function(r){ return String(r.SKU)===sku; });\n    if (it) document.getElementById('poHarga').value = it['Harga Beli'] || 0;\n  }\n  function isiInvHarga(){\n    const sku = document.getElementById('invSku').value;\n    const it = data.inventory.find(function(r){ return String(r.SKU)===sku; });\n    if (it) document.getElementById('invHarga').value = it['Harga Jual'] || 0;\n  }\n  function saveBarang(){\n    busy(true, 'Menyimpan barang...');\n    google.script.run.withSuccessHandler(function(res){\n      busy(false);\n      document.getElementById('invInfo').className='statusok';\n      document.getElementById('invInfo').innerText='Tersimpan SKU '+res.sku;\n      loadData();\n    }).withFailureHandler(function(err){ busy(false); alert(err.message||err); })\n      .simpanBarang({\n        sku: document.getElementById('bSku').value,\n        nama: document.getElementById('bNama').value,\n        kategori: document.getElementById('bKat').value,\n        satuan: document.getElementById('bSat').value,\n        stok: document.getElementById('bStok').value,\n        hargaBeli: document.getElementById('bBeli').value,\n        hargaJual: document.getElementById('bJual').value,\n        minStok: document.getElementById('bMin').value\n      });\n  }\n  function closeEdit(){ document.getElementById('editModal').className='modal-back'; }\n  function editStok(sku){\n    var it=(data.inventory||[]).find(function(r){ return String(r.SKU)===String(sku); });\n    if(!it) return alert('SKU not found');\n    document.getElementById('editTitle').innerText='Edit: '+(it.Nama||it.SKU);\n    eSku.value=it.SKU||'';\n    eNama.value=it.Nama||'';\n    eKat.value=it.Kategori||'';\n    eSat.value=it.Satuan||it.Uom||'PCS';\n    eStok.value=it.Stok||it.Qty||0;\n    eMin.value=it['Min Stok']||0;\n    eBeli.value=it['Harga Beli']||0;\n    eJual.value=it['Harga Jual']||it.Harga||0;\n    document.getElementById('editModal').className='modal-back show';\n  }\n  function saveEdit(){\n    busy(true,'Saving...');\n    google.script.run.withSuccessHandler(function(){\n      busy(false); closeEdit(); loadData(true);\n    }).withFailureHandler(function(err){ busy(false); alert(err.message||err); })\n      .simpanBarang({ sku:eSku.value, nama:eNama.value, kategori:eKat.value, satuan:eSat.value, stok:eStok.value, hargaBeli:eBeli.value, hargaJual:eJual.value, minStok:eMin.value });\n  }\n  function adj(sku, d){\n    busy(true, 'Mengubah stok...');\n    google.script.run.withSuccessHandler(function(){ busy(false); loadData(); })\n      .withFailureHandler(function(err){ busy(false); alert(err.message||err); })\n      .ubahStok({ sku: sku, delta: d });\n  }\n  function savePo(){\n    const items = collect('poSku');\n    if (!items.length) return alert('Add at least one item');\n    busy(true, 'Saving...');\n    google.script.run.withSuccessHandler(function(res){\n      busy(false);\n      document.getElementById('poInfo').className='statusok';\n      document.getElementById('poInfo').innerText='Saved '+res.noPo;\n      loadData();\n    }).withFailureHandler(function(err){ busy(false); alert(err.message||err); })\n      .simpanPembelian({ supplier: document.getElementById('poSup').value, items: items });\n  }\n  function skuOptions(){\n    return '<option value=\"\">SKU / item name</option>' + (data.inventory||[]).map(function(r){\n      return '<option value=\"'+r.SKU+'\">'+(r.Nama||r.SKU)+' ('+r.SKU+')</option>';\n    }).join('');\n  }\n  function itemUom(it){\n    return it.Satuan || it.Uom || it.UOM || it.Sat || it.Unit || 'PCS';\n  }\n  function itemPrice(it, jual){\n    if (jual) return it['Harga Jual'] || it.Harga || it.Price || it['Harga Beli'] || 0;\n    return it['Harga Beli'] || it.Harga || it.Cost || it.Price || 0;\n  }\n  function fillCardPrice(sel){\n    const it = (data.inventory||[]).find(function(r){ return String(r.SKU)===sel.value; });\n    const card = sel.closest('.item-card');\n    if (!card) return;\n    const price = card.querySelector('.soHarga, .poHarga');\n    const uom = card.querySelector('.soUom, .poUom');\n    if (!it) { if (uom) uom.value='PCS'; return; }\n    if (uom) uom.value = itemUom(it);\n    if (price) price.value = itemPrice(it, !!card.querySelector('.soHarga'));\n  }\n  function addSoLine(){\n    const wrap = document.getElementById('soLines');\n    const d = document.createElement('div');\n    d.className='item-card soLine';\n    d.innerHTML = '<button class=\"item-x\" type=\"button\" onclick=\"this.parentNode.remove()\">\u00d7</button>'\n      + '<label>SKU / item</label><div class=\"sku-row\"><select class=\"soSku\" onchange=\"fillCardPrice(this)\">'+skuOptions()+'</select></div>'\n      + '<div class=\"item-grid\"><div><label>Qty</label><input class=\"soQty\" type=\"number\" value=\"1\"></div>'\n      + '<div><label>Uom</label><input class=\"soUom\" placeholder=\"PCS\" value=\"PCS\"></div>'\n      + '<div><label>Price</label><input class=\"soHarga\" type=\"number\" placeholder=\"0\"></div></div>';\n    wrap.appendChild(d);\n  }\n  function addPoLine(){\n    const wrap = document.getElementById('poLines');\n    if (!wrap) return;\n    const d = document.createElement('div');\n    d.className='item-card poLine';\n    d.innerHTML = '<button class=\"item-x\" type=\"button\" onclick=\"this.parentNode.remove()\">\u00d7</button>'\n      + '<label>SKU / item</label><div class=\"sku-row\"><select class=\"poSku\" onchange=\"fillCardPrice(this)\">'+skuOptions()+'</select></div>'\n      + '<div class=\"item-grid\"><div><label>Qty</label><input class=\"poQty\" type=\"number\" value=\"1\"></div>'\n      + '<div><label>Uom</label><input class=\"poUom\" placeholder=\"PCS\" value=\"PCS\"></div>'\n      + '<div><label>Cost</label><input class=\"poHarga\" type=\"number\" placeholder=\"0\"></div></div>';\n    wrap.appendChild(d);\n  }\n  function collect(clsSku){\n    const out=[];\n    const root = clsSku==='soSku' ? '#soLines' : '#poLines';\n    document.querySelectorAll(root+' .item-card').forEach(function(card){\n      const sel=card.querySelector('select');\n      const sku=sel?sel.value:''; if(!sku) return;\n      const it=(data.inventory||[]).find(function(r){return String(r.SKU)===sku;});\n      const qty=card.querySelector('.soQty, .poQty');\n      const harga=card.querySelector('.soHarga, .poHarga');\n      out.push({sku:sku,nama:it?it.Nama:sku,qty:qty?qty.value:1,harga:harga?harga.value:0,uom:(card.querySelector('.soUom,.poUom')||{}).value||''});\n    });\n    return out;\n  }\n\n  function quoteTotal(items){\n    var gross = 0;\n    items.forEach(function(it){ gross += (Number(it.qty)||0)*(Number(it.harga)||0); });\n    var pct = Number((document.getElementById('soDiskonPct')||{}).value)||0;\n    var amt = Number((document.getElementById('soDiskonAmt')||{}).value)||0;\n    var disc = amt > 0 ? amt : gross * pct / 100;\n    if (disc > gross) disc = gross;\n    var net = gross - disc;\n    var el = document.getElementById('soTotalHint');\n    if (el) el.innerText = 'Total '+rp(net)+ (disc? ' (\u2212'+rp(disc)+')' : '');\n    return { gross:gross, disc:disc, net:net, pct:pct };\n  }\n  function saveInv(){\n    const items = collect('soSku');\n    if (!items.length) return alert('Add at least one item');\n    busy(true, 'Saving...');\n    google.script.run.withSuccessHandler(function(res){\n      busy(false);\n      document.getElementById('invcInfo').className='statusok';\n      document.getElementById('invcInfo').innerText='Saved '+res.noInv;\n      loadData();\n    }).withFailureHandler(function(err){ busy(false); alert(err.message||err); })\n      .simpanInvoice({\n        pelanggan: document.getElementById('invCus').value,\n        items: items,\n        metode: (document.getElementById('soMetode')||{}).value,\n        status: (document.getElementById('soStatus')||{}).value,\n        bayar: (document.getElementById('soBayar')||{}).value,\n        diskonPct: (document.getElementById('soDiskonPct')||{}).value,\n        diskon: quoteTotal(items).disc\n      });\n  }\n  function saveKeu(){\n    busy(true, 'Menyimpan kas...');\n    google.script.run.withSuccessHandler(function(res){\n      busy(false);\n      document.getElementById('keuInfo').className='statusok';\n      document.getElementById('keuInfo').innerText='Tersimpan '+res.noBukti;\n      loadData();\n    }).withFailureHandler(function(err){ busy(false); alert(err.message||err); })\n      .simpanKeuangan({\n        jenis: document.getElementById('kJenis').value,\n        akun: document.getElementById('kAkun').value,\n        keterangan: document.getElementById('kKet').value,\n        jumlah: document.getElementById('kJum').value\n      });\n  }\n  function saveMitra(){\n    busy(true, 'Menyimpan mitra...');\n    google.script.run.withSuccessHandler(function(res){\n      busy(false);\n      document.getElementById('mitInfo').className='statusok';\n      document.getElementById('mitInfo').innerText='Tersimpan '+res.kode;\n      loadData();\n    }).withFailureHandler(function(err){ busy(false); alert(err.message||err); })\n      .simpanMitra({\n        jenis: document.getElementById('mJenis').value,\n        nama: document.getElementById('mNama').value,\n        telepon: document.getElementById('mTel').value,\n        alamat: document.getElementById('mAlamat').value\n      });\n  }\n\n  const CSV_COLS = {\n    inventory: ['SKU','Nama','Kategori','Satuan','Stok','Harga Beli','Harga Jual','Min Stok','Status'],\n    supplier: ['Kode','Nama','Telepon','Alamat'],\n    pelanggan: ['Kode','Nama','Telepon','Alamat'],\n    pembelian: ['Waktu','No PO','Supplier','SKU','Nama','Qty','Harga','Subtotal','Status'],\n    invoice: ['Waktu','No Inv','Pelanggan','SKU','Nama','Qty','Harga','Gross','Diskon','Bersih','Status','Metode'],\n    keuangan: ['Waktu','No Bukti','Jenis','Akun','Keterangan','Jumlah','Ref']\n  };\n  function csvEscape(v){\n    const t = String(v==null?'':v);\n    if (/[\",\\n]/.test(t)) return '\"' + t.replace(/\"/g,'\"\"') + '\"';\n    return t;\n  }\n  function invoiceExportRow(r){\n    var qty = Number(r.Qty)||0;\n    var harga = Number(r.Harga)||0;\n    var gross = qty * harga;\n    var disc = Number(r.Diskon)||0;\n    var bersih = Number(r.Subtotal);\n    if (!disc && gross && bersih && gross > bersih) disc = gross - bersih;\n    if (isNaN(bersih) || bersih===0 && gross) bersih = gross - disc;\n    return Object.assign({}, r, { Gross: gross, Diskon: disc, Bersih: bersih });\n  }\n  function exportCsv(jenis){\n    const cols = CSV_COLS[jenis];\n    var rows = data[jenis] || [];\n    if (jenis === 'invoice') rows = rows.map(invoiceExportRow);\n    const lines = [cols.join(';')].concat(rows.map(function(r){\n      return cols.map(function(c){ return csvEscape(r[c]); }).join(';');\n    }));\n    const blob = new Blob(['\\ufeff'+lines.join('\\n')], {type:'application/vnd.ms-excel;charset=utf-8;'});\n    const a = document.createElement('a');\n    a.href = URL.createObjectURL(blob);\n    a.download = 'SCASS-'+jenis+'.xls';\n    a.click();\n    document.getElementById('expInfo').className='statusok';\n    document.getElementById('expInfo').innerText='File '+a.download+' diunduh.';\n  }\n  function downloadB64(name, b64, mime){\n    var bin = atob(b64);\n    var bytes = new Uint8Array(bin.length);\n    for (var i=0;i<bin.length;i++) bytes[i] = bin.charCodeAt(i);\n    var a = document.createElement('a');\n    a.href = URL.createObjectURL(new Blob([bytes], {type: mime||'application/octet-stream'}));\n    a.download = name;\n    a.click();\n  }\n  function exportExcelLocal(){\n    var parts = [];\n    Object.keys(CSV_COLS).forEach(function(jenis){\n      var cols = CSV_COLS[jenis];\n      var rows = data[jenis] || [];\n      if (jenis === 'invoice') rows = rows.map(invoiceExportRow);\n      parts.push('<h3>'+jenis+'</h3><table border=\"1\"><tr>'+cols.map(function(c){return '<th>'+c+'</th>';}).join('')+'</tr>');\n      rows.forEach(function(r){\n        parts.push('<tr>'+cols.map(function(c){return '<td>'+String(r[c]==null?'':r[c]).replace(/</g,'')+'</td>';}).join('')+'</tr>');\n      });\n      parts.push('</table>');\n    });\n    var blob = new Blob(['\\ufeff<html><body>'+parts.join('')+'</body></html>'], {type:'application/vnd.ms-excel'});\n    var a = document.createElement('a');\n    a.href = URL.createObjectURL(blob);\n    a.download = 'SCASS-ERP.xls';\n    a.click();\n  }\n  function exportExcel(){\n    busy(true, 'Menyiapkan Excel...');\n    google.script.run.withSuccessHandler(function(res){\n      busy(false);\n      if (res && res.base64) {\n        downloadB64(res.name||'SCASS-ERP.xlsx', res.base64, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');\n        document.getElementById('expInfo').className='statusok';\n        document.getElementById('expInfo').innerText = 'File '+(res.name||'SCASS-ERP.xlsx')+' diunduh.';\n        return;\n      }\n      exportExcelLocal();\n      document.getElementById('expInfo').className='statusok';\n      document.getElementById('expInfo').innerText = 'File SCASS-ERP.xls diunduh.';\n    }).withFailureHandler(function(){\n      busy(false);\n      exportExcelLocal();\n      document.getElementById('expInfo').className='statusok';\n      document.getElementById('expInfo').innerText = 'File SCASS-ERP.xls diunduh.';\n    }).exportErpExcel();\n  }\n  function parseCsvText(text){\n    const rows = [];\n    let row = [], cell = '', q = false;\n    for (let i=0;i<text.length;i++){\n      const ch = text[i], next = text[i+1];\n      if (q){\n        if (ch==='\"' && next==='\"'){ cell+='\"'; i++; }\n        else if (ch==='\"'){ q=false; }\n        else cell+=ch;\n      } else {\n        if (ch==='\"') q=true;\n        else if (ch===','){ row.push(cell); cell=''; }\n        else if (ch==='\\n'){ row.push(cell); rows.push(row); row=[]; cell=''; }\n        else if (ch!=='\\r') cell+=ch;\n      }\n    }\n    if (cell.length || row.length){ row.push(cell); rows.push(row); }\n    return rows.filter(function(r){ return r.some(function(c){ return String(c).trim()!==''; }); });\n  }\n  function importCsv(){\n    const jenis = document.getElementById('impJenis').value;\n    const file = document.getElementById('impFile').files[0];\n    if (!file) return alert('Pilih file CSV');\n    const reader = new FileReader();\n    reader.onload = function(){\n      const raw = String(reader.result||'').replace(/^\\ufeff/,'');\n      const grid = parseCsvText(raw);\n      if (grid.length < 2) return alert('CSV kosong atau tanpa data');\n      const head = grid[0].map(function(h){ return String(h||'').trim(); });\n      const rows = [];\n      for (let i=1;i<grid.length;i++){\n        const obj = {};\n        head.forEach(function(h,c){ obj[h] = grid[i][c] != null ? String(grid[i][c]).trim() : ''; });\n        rows.push(obj);\n      }\n      busy(true, 'Mengimpor '+rows.length+' baris...');\n      google.script.run.withSuccessHandler(function(res){\n        busy(false);\n        document.getElementById('impInfo').className='statusok';\n        document.getElementById('impInfo').innerText='Import selesai: '+res.masuk+' baris ke '+res.jenis+'.';\n        loadData();\n        loadDash();\n      }).withFailureHandler(function(err){ busy(false); alert(err.message||err); })\n        .importErpRows({ jenis: jenis, rows: rows });\n    };\n    reader.readAsText(file);\n  }\n\n  function uniq(list, key){\n    var m={};\n    (list||[]).forEach(function(r){ var v=String(r[key]||'').trim(); if(v) m[v]=1; });\n    return Object.keys(m).sort();\n  }\n  function fillInvFilters(){\n    var kat=document.getElementById('fKat');\n    var uom=document.getElementById('fUom');\n    if(!kat||!uom) return;\n    var kv=kat.value, uv=uom.value;\n    kat.innerHTML='<option value=\"\">- Select Category -</option>'+uniq(data.inventory,'Kategori').map(function(v){return '<option>'+v+'</option>';}).join('');\n    uom.innerHTML='<option value=\"\">- Select Uom -</option>'+uniq(data.inventory,'Satuan').map(function(v){return '<option>'+v+'</option>';}).join('');\n    kat.value=kv; uom.value=uv;\n  }\n  function clearInvFilter(){\n    var kat=document.getElementById('fKat');\n    var uom=document.getElementById('fUom');\n    var q=document.getElementById('invQ');\n    var st=document.getElementById('fStat');\n    if(kat){ kat.value=''; kat.dataset.ready=''; }\n    if(uom) uom.value='';\n    if(q) q.value='';\n    if(st) st.value='';\n    fillInvFilters();\n    paintInv();\n  }\n  function paintInv(){\n    fillInvFilters();\n    var q=String((document.getElementById('invQ')||{}).value||'').toLowerCase();\n    var kat=String((document.getElementById('fKat')||{}).value||'');\n    var uom=String((document.getElementById('fUom')||{}).value||'');\n    var st=String((document.getElementById('fStat')||{}).value||'');\n    var rows=(data.inventory||[]).filter(function(r){\n      var stok=Number(r.Stok||r.Qty||0);\n      var min=Number(r['Min Stok']||r.Min||0);\n      var status=stok<=min?'LOW':'OK';\n      if(kat && String(r.Kategori||'')!==kat) return false;\n      if(uom && String(r.Satuan||r.Uom||'')!==uom) return false;\n      if(st && status!==st) return false;\n      if(q && (String(r.SKU||'')+' '+String(r.Nama||'')).toLowerCase().indexOf(q)<0) return false;\n      return true;\n    });\n    var el=document.getElementById('tblInv');\n    if(el) el.innerHTML = renderInv(rows);\n  }\n  function renderInv(rows){\n    if(!rows||!rows.length) return '<div class=\"muted\">No items yet.</div>';\n    var head='<tr><th>#</th><th>SKU</th><th>Product Name</th><th>Category</th><th>Uom</th><th>Qty</th><th>Cost</th><th>Price</th><th>Min</th><th>Status</th><th>Aksi</th></tr>';\n    var body=rows.map(function(r,i){\n      var stok=Number(r.Stok||r.Qty||0);\n      var min=Number(r['Min Stok']||r.Min||0);\n      var low=stok<=min;\n      var sku=String(r.SKU||'').replace(/'/g,'');\n      return '<tr>'\n        +'<td>'+(i+1)+'</td>'\n        +'<td>'+ (r.SKU||'') +'</td>'\n        +'<td>'+ (r.Nama||'') +'</td>'\n        +'<td>'+ (r.Kategori||'') +'</td>'\n        +'<td>'+ (r.Satuan||r.Uom||'PCS') +'</td>'\n        +'<td>'+ stok +'</td>'\n        +'<td>'+ rp(r['Harga Beli']||r.Cost||0) +'</td>'\n        +'<td>'+ rp(r['Harga Jual']||r.Harga||r.Price||0) +'</td>'\n        +'<td>'+ min +'</td>'\n        +'<td>'+ (low?'LOW':'OK') +'</td>'\n        +'<td><button class=\"btn2\" title=\"Edit stock\" onclick=\"editStok(\\''+sku+'\\','+stok+')\">\u270e</button></td>'\n        +'</tr>';\n    }).join('');\n    return '<table><thead>'+head+'</thead><tbody>'+body+'</tbody></table>';\n  }\n\n  function renderInvce(){\n    var q = String((document.getElementById('invSearch')||{}).value||'').trim().toLowerCase();\n    var rows = (data.invoice||[]).filter(function(r){\n      if (!q) return true;\n      return String(r['No Inv']||'').toLowerCase().indexOf(q) >= 0\n        || String(r.Pelanggan||'').toLowerCase().indexOf(q) >= 0\n        || String(r.Nama||'').toLowerCase().indexOf(q) >= 0\n        || String(r.SKU||'').toLowerCase().indexOf(q) >= 0;\n    });\n    var el = document.getElementById('tblInvce');\n    if (!el) return;\n    el.innerHTML = table(rows, ['Waktu','No Inv','Pelanggan','SKU','Nama','Qty','Harga','Subtotal','Diskon','Status'], function(r){\n      return '<button class=\"icon-pen\" title=\"Edit / view\" onclick=\"openInvce(\\''+String(r['No Inv']||'').replace(/'/g,'')+'\\')\">&#9998;</button>';\n    }).replace('<th></th>','<th>Aksi</th>');\n  }\n  function openInvce(no){\n    var lines = (data.invoice||[]).filter(function(r){ return String(r['No Inv'])===String(no); });\n    if (!lines.length) return;\n    var r = lines[0];\n    var st = String(r.Status||'PENDING').toUpperCase();\n    if (st==='PAID') st='LUNAS';\n    document.getElementById('ivNo').value = no;\n    document.getElementById('ivCus').value = r.Pelanggan||'-';\n    document.getElementById('ivTgl').value = r.Waktu||'';\n    document.getElementById('ivStatus').value = (st==='VOID'||st==='BATAL')?'VOID':(st==='LUNAS'?'LUNAS':'PENDING');\n    document.getElementById('ivItems').innerText = lines.map(function(x){\n      return (x.SKU||'')+' '+ (x.Nama||'')+' x'+ (x.Qty||0);\n    }).join(' | ');\n    document.getElementById('invceModal').classList.add('show');\n  }\n  function closeInvce(){ document.getElementById('invceModal').classList.remove('show'); }\n  function saveInvceStatus(){\n    var no = document.getElementById('ivNo').value;\n    var status = document.getElementById('ivStatus').value;\n    busy(true, 'Updating status...');\n    google.script.run.withSuccessHandler(function(){\n      busy(false);\n      (data.invoice||[]).forEach(function(r){ if (String(r['No Inv'])===String(no)) r.Status = status; });\n      closeInvce();\n      renderInvce();\n    }).withFailureHandler(function(err){ busy(false); alert(err.message||err); })\n      .ubahStatusInvoice({ noInv: no, status: status });\n  }\n  function pdfInv(no){\n    var rows=(data.invoice||[]).filter(function(r){return String(r['No Inv'])===String(no);});\n    if(!rows.length) return alert('Invoice not found');\n    var r0=rows[0];\n    var gross=0, disc=0, tot=0;\n    rows.forEach(function(r){\n      var g=(Number(r.Qty)||0)*(Number(r.Harga)||0);\n      var net=Number(r.Subtotal)||0;\n      var d=Number(r.Diskon)||0;\n      if(!d && g>net) d=g-net;\n      gross+=g; disc+=d; tot+=net|| (g-d);\n    });\n    var pct=gross>0? Math.round(disc/gross*1000)/10 : 0;\n    if(Math.abs(pct-Math.round(pct))<0.05) pct=Math.round(pct);\n    var st=String(r0.Status||'PENDING').toUpperCase();\n    if(st==='LUNAS') st='PAID';\n    function rp2(n){ return 'Rp '+Number(n||0).toLocaleString('id-ID'); }\n    var lines=rows.map(function(r){\n      return '<tr><td>'+(r.SKU||'')+'</td><td>'+(r.Nama||'')+'</td><td style=\"text-align:center\">'+(r.Qty||0)+'</td><td style=\"text-align:right\">'+rp2(r.Harga)+'</td><td style=\"text-align:right\">'+rp2((Number(r.Qty)||0)*(Number(r.Harga)||0))+'</td></tr>';\n    }).join('');\n    var html='<!DOCTYPE html><html><head><meta charset=\"utf-8\"><title>SCASS '+no+'</title>'\n      +'<style>body{font-family:Arial,sans-serif;color:#16324f;padding:28px;}'\n      +'h1{margin:0;text-align:center;font-size:22px} .sub{text-align:center;color:#5b7380;margin:4px 0 14px}'\n      +'.bar{background:#12324a;color:#fff;text-align:center;padding:8px;font-weight:700;letter-spacing:.04em}'\n      +'table{width:100%;border-collapse:collapse;margin-top:12px;font-size:13px}'\n      +'th,td{border:1px solid #c9dde0;padding:7px 8px}'\n      +'th{background:#1f7a6e;color:#fff}'\n      +'.tot td{font-weight:700}'\n      +'.tot .g{background:#0b4f37;color:#fff}'\n      +'.meta td{border:1px solid #c9dde0}'\n      +'.lbl{background:#eef6f4;font-weight:700;width:18%}'\n      +'@media print{button{display:none} body{padding:0}}</style></head><body>'\n      +'<button onclick=\"window.print()\" style=\"margin-bottom:12px;padding:8px 14px;background:#0d9488;color:#fff;border:0;border-radius:8px;font-weight:700;cursor:pointer\">Print / Save PDF</button>'\n      +'<h1>SCASS</h1><div class=\"sub\">NewStar Kancing</div>'\n      +'<div class=\"bar\">SALES INVOICE</div>'\n      +'<table class=\"meta\" style=\"margin-top:10px\"><tr><td class=\"lbl\">Invoice No.</td><td>'+no+'</td><td class=\"lbl\">Date</td><td>'+(r0.Waktu||'')+'</td></tr>'\n      +'<tr><td class=\"lbl\">Customer</td><td>'+(r0.Pelanggan||'-')+'</td><td class=\"lbl\">Status</td><td>'+st+'</td></tr>'\n      +'<tr><td class=\"lbl\">Payment</td><td>'+(r0.Metode||'-')+'</td><td class=\"lbl\">Paid</td><td>'+rp2(r0.Bayar)+'</td></tr></table>'\n      +'<table><thead><tr><th>SKU</th><th>Item</th><th>Qty</th><th>Price</th><th>Subtotal</th></tr></thead><tbody>'\n      +lines\n      +'<tr class=\"tot\"><td colspan=\"3\"></td><td>SUBTOTAL</td><td style=\"text-align:right\">'+rp2(gross)+'</td></tr>'\n      +'<tr class=\"tot\"><td colspan=\"3\"></td><td>DISCOUNT '+pct+'%</td><td style=\"text-align:right\">'+rp2(disc)+'</td></tr>'\n      +'<tr class=\"tot g\"><td colspan=\"3\"></td><td>TOTAL</td><td style=\"text-align:right\">'+rp2(tot)+'</td></tr>'\n      +'</tbody></table>'\n      +'<p style=\"text-align:center;color:#5b7380;margin-top:28px\">Thank you for your business</p>'\n      +'</body></html>';\n    var w=window.open('','_blank');\n    if(!w){ alert('Izinkan pop-up untuk melihat invoice'); return; }\n    w.document.write(html);\n    w.document.close();\n  }\n  function xlsInv(no){\n    const rows=(data.invoice||[]).filter(function(r){return String(r['No Inv'])===String(no);});\n    const cols=['Waktu','No Inv','Pelanggan','SKU','Nama','Qty','Harga','Subtotal','Status'];\n    const lines=[cols.join(';')].concat(rows.map(function(r){return cols.map(function(c){return String(r[c]==null?'':r[c]);}).join(';');}));\n    const blob=new Blob(['\\ufeff'+lines.join('\\r\\n')],{type:'application/vnd.ms-excel;charset=utf-8;'});\n    const a=document.createElement('a'); a.href=URL.createObjectURL(blob); a.download='SCASS-INV-'+no+'.xls'; a.click();\n  }\n\n  document.addEventListener('input', function(e){\n    if (!e.target) return;\n    if (/soQty|soHarga|soDiskon|invQty|invHarga/.test(e.target.id||e.target.className||'')) {\n      try { quoteTotal(collect('soSku')); } catch(ex) {}\n    }\n  });\n  document.addEventListener('keydown', function(e){\n    if (e.key==='Enter' && document.getElementById('pageApp').style.display==='none') doLogin();\n  });\n  if (localStorage.getItem('scass_login')==='1') {\n    document.getElementById('pageLogin').classList.add('hide'); document.getElementById('pageLogin').style.display='none';\n    document.getElementById('pageApp').style.display='flex';\n    setLang(lang); loadData(true);\n  }\n</script>\n</body>\n</html>\n";
