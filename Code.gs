// Web app dijalankan sebagai pengguna yang membuka (appsscript.json). Tiap akun punya Spreadsheet
// sendiri di Drive-nya, dibuat otomatis oleh onboard(), jadi data antar akun terpisah.
const HEAD = {
  Motor: ['nama', 'km', 'pajak_jatuh_tempo'],
  Servis: ['id', 'motor', 'tanggal', 'km', 'tipe', 'bengkel', 'total', 'catatan', 'link_nota'],
  Item: ['servis_id', 'komponen', 'aksi', 'nama', 'biaya'],
  Jadwal: ['komponen', 'interval_km', 'interval_bulan'],
  Config: ['key', 'value'],
};
// Angka contoh, sesuaikan dengan buku servis. Kosongkan sel untuk menonaktifkan salah satu interval.
const JADWAL = [
  ['Oli mesin', 2000, 2], ['CVT', 8000, 6], ['Injeksi', 10000, 12], ['Gardan', 8000, 6],
  ['Rantai', 15000, 12], ['Kampas rem', 12000, 12], ['Busi', 10000, 12], ['Aki', '', 18],
];
const DAY = 864e5;

const UP = () => PropertiesService.getUserProperties();
let SS_ = null;
const ss_ = () => SS_ || (SS_ = SpreadsheetApp.openById(UP().getProperty('ssid')));
const sheet_ = n => ss_().getSheetByName(n);
const dstr_ = x => x instanceof Date ? Utilities.formatDate(x, Session.getScriptTimeZone(), 'yyyy-MM-dd') : (x || '');

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index').setTitle('Rekap Bengkel')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function rows_(n) {
  const v = sheet_(n).getDataRange().getValues();
  const h = v.shift();
  return v.filter(r => r[0] !== '').map(r => Object.fromEntries(h.map((k, i) => [k, r[i]])));
}

function cekMotor_(nama, km) {
  nama = String(nama || '').trim();
  km = Number(km);
  if (!nama || !isFinite(km) || km < 0) throw new Error('Nama motor dan km wajib diisi.');
  return [nama, km];
}

// Kunjungan pertama: buat Sheet pribadi + motor pertama + trigger email harian.
function onboard(nama, km) {
  [nama, km] = cekMotor_(nama, km);
  const lock = LockService.getUserLock();
  lock.waitLock(20000);
  let warn = '';
  try {
    if (UP().getProperty('ssid')) return getData();
    const ss = SpreadsheetApp.create('Rekap Bengkel');
    Object.keys(HEAD).forEach(n => ss.insertSheet(n).appendRow(HEAD[n]));
    ss.deleteSheet(ss.getSheets()[0]);
    ss.getSheetByName('Jadwal').getRange(2, 1, JADWAL.length, 3).setValues(JADWAL);
    ss.getSheetByName('Motor').appendRow([nama, km, '']);
    const email = Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail();
    ss.getSheetByName('Config').appendRow(['email', email]);
    UP().setProperty('ssid', ss.getId());
    SS_ = ss;
    try {
      ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'dailyReminder').forEach(ScriptApp.deleteTrigger);
      ScriptApp.newTrigger('dailyReminder').timeBased().everyDays(1).atHour(8).create();
    } catch (e) { warn = 'Pengingat email gagal diaktifkan: ' + e.message; }
  } finally { lock.releaseLock(); }
  const d = getData();
  d.warn = warn;
  return d;
}

function addMotor(nama, km) {
  [nama, km] = cekMotor_(nama, km);
  if (rows_('Motor').some(m => m.nama === nama)) throw new Error('Nama motor sudah ada.');
  sheet_('Motor').appendRow([nama, km, '']);
  return getData();
}

// last: {km, date} komponen terakhir dikerjakan, atau null.
function hitung_(j, last, km, today) {
  if (!last) return { state: 'kosong' };
  const sisaKm = j.interval_km ? last.km + j.interval_km - km : null;
  let sisaHari = null;
  if (j.interval_bulan) {
    const due = new Date(last.date);
    due.setMonth(due.getMonth() + j.interval_bulan);
    sisaHari = Math.ceil((due - today) / DAY);
  }
  const lewat = (sisaKm !== null && sisaKm < 0) || (sisaHari !== null && sisaHari < 0);
  // shortcut: ambang "dekat" tetap 10% interval km / 14 hari, jadikan kolom Config jika perlu diatur.
  const dekat = (sisaKm !== null && sisaKm <= j.interval_km * 0.1) || (sisaHari !== null && sisaHari <= 14);
  return { state: lewat ? 'lewat' : dekat ? 'dekat' : 'aman', sisaKm, sisaHari };
}

function status_() {
  const motors = rows_('Motor'), jadwal = rows_('Jadwal'), servis = rows_('Servis'), items = rows_('Item');
  const sById = Object.fromEntries(servis.map(s => [s.id, s]));
  const today = new Date();
  const out = [];
  motors.forEach(m => {
    jadwal.forEach(j => {
      let last = null;
      items.forEach(it => {
        const s = sById[it.servis_id];
        if (it.komponen !== j.komponen || !s || s.motor !== m.nama) return;
        if (!last || new Date(s.tanggal) > last.date || (+new Date(s.tanggal) === +last.date && s.km > last.km)) last = { km: Number(s.km), date: new Date(s.tanggal) };
      });
      out.push(Object.assign({ motor: m.nama, komponen: j.komponen }, hitung_(j, last, Number(m.km), today)));
    });
    if (m.pajak_jatuh_tempo) {
      const sisaHari = Math.ceil((new Date(m.pajak_jatuh_tempo) - today) / DAY);
      out.push({ motor: m.nama, komponen: 'Pajak STNK', sisaKm: null, sisaHari, state: sisaHari < 0 ? 'lewat' : sisaHari <= 30 ? 'dekat' : 'aman' });
    }
  });
  return out;
}

function getData() {
  if (!UP().getProperty('ssid')) return { onboarding: true };
  const items = rows_('Item');
  const servis = rows_('Servis').map(s => Object.assign(s, {
    tanggal: dstr_(s.tanggal),
    items: items.filter(i => i.servis_id === s.id),
  })).sort((a, b) => b.tanggal.localeCompare(a.tanggal));
  return {
    email: (rows_('Config').find(c => c.key === 'email') || {}).value || '',
    motors: rows_('Motor').map(m => Object.assign(m, { pajak_jatuh_tempo: dstr_(m.pajak_jatuh_tempo) })),
    jadwal: rows_('Jadwal'),
    servis,
    status: status_(),
  };
}

function saveServis(p) {
  const km = Number(p.km);
  if (!p.motor || !p.tanggal || !isFinite(km) || !p.items || !p.items.length) throw new Error('Motor, tanggal, km, dan minimal 1 item wajib diisi.');
  const lock = LockService.getUserLock();
  lock.waitLock(20000);
  try {
    const id = Utilities.getUuid().slice(0, 8);
    const items = p.items.map(i => [id, i.komponen, i.aksi, i.nama || '', Number(i.biaya) || 0]);
    const total = items.reduce((a, r) => a + r[4], 0);
    let link = '';
    if (p.nota) {
      const it = DriveApp.getFoldersByName('Nota Bengkel');
      const folder = it.hasNext() ? it.next() : DriveApp.createFolder('Nota Bengkel');
      link = folder.createFile(Utilities.newBlob(Utilities.base64Decode(p.nota.b64), p.nota.mime, p.nota.name)).getUrl();
    }
    sheet_('Servis').appendRow([id, p.motor, new Date(p.tanggal), km, p.tipe, p.bengkel || '', total, p.catatan || '', link]);
    const si = sheet_('Item');
    si.getRange(si.getLastRow() + 1, 1, items.length, 5).setValues(items);
    updateKm(p.motor, km, true);
  } finally { lock.releaseLock(); }
  return getData();
}

// onlyUp: dari form servis, km tidak boleh menurunkan odometer (mis. input servis lama).
function updateKm(motor, km, onlyUp) {
  km = Number(km);
  if (!isFinite(km)) throw new Error('Km tidak valid.');
  const s = sheet_('Motor'), v = s.getDataRange().getValues();
  for (let r = 1; r < v.length; r++) {
    if (v[r][0] === motor) { if (!onlyUp || km > v[r][1]) s.getRange(r + 1, 2).setValue(km); return getData(); }
  }
  throw new Error('Motor tidak ditemukan: ' + motor);
}

function setPajak(motor, tanggal) {
  const s = sheet_('Motor'), v = s.getDataRange().getValues();
  for (let r = 1; r < v.length; r++) {
    if (v[r][0] === motor) { s.getRange(r + 1, 3).setValue(tanggal ? new Date(tanggal) : ''); return getData(); }
  }
  throw new Error('Motor tidak ditemukan: ' + motor);
}

function deleteServis(id) {
  const lock = LockService.getUserLock();
  lock.waitLock(20000);
  try {
    ['Servis', 'Item'].forEach(n => {
      const s = sheet_(n), v = s.getDataRange().getValues();
      for (let r = v.length - 1; r >= 1; r--) if (v[r][0] === id) s.deleteRow(r + 1);
    });
  } finally { lock.releaseLock(); }
  return getData();
}

// Trigger harian (milik tiap pengguna): satu email berisi semua yang dekat/lewat.
function dailyReminder() {
  if (!UP().getProperty('ssid')) return;
  const due = status_().filter(s => s.state === 'dekat' || s.state === 'lewat');
  if (!due.length) return;
  const to = (rows_('Config').find(c => c.key === 'email') || {}).value;
  if (!to) return;
  const line = s => {
    const p = [];
    if (s.sisaKm !== null) p.push(s.sisaKm < 0 ? 'lewat ' + -s.sisaKm + ' km' : 'sisa ' + s.sisaKm + ' km');
    if (s.sisaHari !== null) p.push(s.sisaHari < 0 ? 'lewat ' + -s.sisaHari + ' hari' : 'sisa ' + s.sisaHari + ' hari');
    return `- [${s.motor}] ${s.komponen}: ${p.join(', ')}${s.state === 'lewat' ? ' (LEWAT)' : ''}`;
  };
  MailApp.sendEmail(to, 'Pengingat servis motor', due.map(line).join('\n'));
}

function selfTest() {
  const today = new Date('2026-06-01');
  const j = { interval_km: 2000, interval_bulan: 2 };
  const a = (c, m) => { if (!c) throw new Error('gagal: ' + m); };
  a(hitung_(j, null, 0, today).state === 'kosong', 'kosong');
  a(hitung_(j, { km: 1000, date: new Date('2026-05-20') }, 1500, today).state === 'aman', 'aman');
  a(hitung_(j, { km: 1000, date: new Date('2026-05-20') }, 2900, today).state === 'dekat', 'dekat km');
  a(hitung_(j, { km: 1000, date: new Date('2026-05-20') }, 3100, today).state === 'lewat', 'lewat km');
  a(hitung_(j, { km: 1000, date: new Date('2026-03-01') }, 1100, today).state === 'lewat', 'lewat hari');
  a(hitung_({ interval_km: '', interval_bulan: 18 }, { km: 0, date: new Date('2026-05-01') }, 99999, today).state === 'aman', 'aki hanya bulan');
  console.log('selfTest OK');
}
