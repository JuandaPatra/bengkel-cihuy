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
  if (!last) return { state: 'kosong', lastKm: null, lastDate: '', targetKm: null, dueDate: '', intervalKm: j.interval_km || null, intervalBulan: j.interval_bulan || null };
  const targetKm = j.interval_km ? last.km + j.interval_km : null;
  const sisaKm = targetKm !== null ? targetKm - km : null;
  let sisaHari = null;
  let dueDate = null;
  if (j.interval_bulan) {
    const due = new Date(last.date);
    due.setMonth(due.getMonth() + j.interval_bulan);
    dueDate = due;
    const tMid = new Date(today.getFullYear(), today.getMonth(), today.getDate());
    const dMid = new Date(due.getFullYear(), due.getMonth(), due.getDate());
    sisaHari = Math.round((dMid - tMid) / DAY);
  }
  const lewat = (sisaKm !== null && sisaKm < 0) || (sisaHari !== null && sisaHari < 0);
  // shortcut: ambang "dekat" tetap 10% interval km / 14 hari, jadikan kolom Config jika perlu diatur.
  const dekat = (sisaKm !== null && sisaKm <= j.interval_km * 0.1) || (sisaHari !== null && sisaHari <= 14);
  return {
    state: lewat ? 'lewat' : dekat ? 'dekat' : 'aman',
    sisaKm,
    sisaHari,
    lastKm: last.km,
    lastDate: dstr_(last.date),
    targetKm,
    dueDate: dueDate ? dstr_(dueDate) : '',
    intervalKm: j.interval_km || null,
    intervalBulan: j.interval_bulan || null,
  };
}

function status_() {
  const motors = rows_('Motor'), jadwal = rows_('Jadwal'), servis = rows_('Servis'), items = rows_('Item');
  const sById = Object.fromEntries(servis.map(s => [s.id, s]));
  const today = new Date();
  const tMidnight = new Date(today.getFullYear(), today.getMonth(), today.getDate());
  const out = [];
  motors.forEach(m => {
    const motorKm = Number(m.km) || 0;
    jadwal.forEach(j => {
      let last = null;
      items.forEach(it => {
        const s = sById[it.servis_id];
        if (it.komponen !== j.komponen || !s || s.motor !== m.nama) return;
        if (!last || new Date(s.tanggal) > last.date || (+new Date(s.tanggal) === +last.date && s.km > last.km)) last = { km: Number(s.km), date: new Date(s.tanggal) };
      });
      out.push(Object.assign({ motor: m.nama, motorKm, komponen: j.komponen }, hitung_(j, last, motorKm, today)));
    });
    if (m.pajak_jatuh_tempo) {
      const pDate = new Date(m.pajak_jatuh_tempo);
      const pMidnight = new Date(pDate.getFullYear(), pDate.getMonth(), pDate.getDate());
      const sisaHari = Math.round((pMidnight - tMidnight) / DAY);
      out.push({
        motor: m.nama,
        motorKm,
        komponen: 'Pajak STNK',
        sisaKm: null,
        sisaHari,
        state: sisaHari < 0 ? 'lewat' : sisaHari <= 30 ? 'dekat' : 'aman',
        dueDate: dstr_(pDate),
        lastKm: null,
        lastDate: '',
        targetKm: null,
        intervalKm: null,
        intervalBulan: null,
      });
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

// ==========================================
// PENGINGAT EMAIL INFORMATIF & TEMPLATE HTML
// ==========================================

const TIPS_KOMPONEN = {
  'Pajak STNK': 'Siapkan STNK asli & KTP pemilik untuk pembayaran di Samsat terdekat atau bayar online melalui aplikasi SIGNAL / Samsat Digital.',
  'Oli mesin': 'Ganti oli tepat waktu untuk menjaga pelumasan mesin dan mencegah keausan komponen silinder.',
  'CVT': 'Servis dan bersihkan area CVT (roller, kampas ganda, mangkok) dan periksa kelayakan v-belt.',
  'Gardan': 'Ganti oli transmisi/gardan berkala agar roda gigi rasio tetap terlumasi maksimal dan tidak bersuara kasar.',
  'Kampas rem': 'Periksa ketebalan kampas rem depan/belakang serta volume minyak rem demi keselamatan berkendara.',
  'Injeksi': 'Lakukan pembersihan throttle body dan injector untuk menjaga performa tarikan dan kehematan BBM.',
  'Rantai': 'Bersihkan kotoran pada rantai, lumasi dengan chain lube, dan setel ketegangannya agar tidak kendur.',
  'Busi': 'Ganti busi berkala agar pembakaran bahan bakar tetap optimal dan mesin mudah dinyalakan.',
  'Aki': 'Periksa voltase aki dan bersihkan kutub aki agar sistem kelistrikan dan starter tetap prima.',
};

function formatTglIndo_(d, withDay) {
  if (!d) return '-';
  let date;
  if (d instanceof Date) {
    date = d;
  } else if (typeof d === 'string' && /^\d{4}-\d{2}-\d{2}/.test(d)) {
    const parts = d.slice(0, 10).split('-');
    date = new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
  } else {
    date = new Date(d);
  }
  if (isNaN(date.getTime())) return String(d);
  const months = [
    'Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni',
    'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'
  ];
  const days = ['Minggu', 'Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'];
  const formatted = date.getDate() + ' ' + months[date.getMonth()] + ' ' + date.getFullYear();
  return withDay ? days[date.getDay()] + ', ' + formatted : formatted;
}

function escHtml_(s) {
  return String(s ?? '').replace(/[&<>"']/g, c => ({
    '&': '&amp;',
    '<': '&lt;',
    '>': '&gt;',
    '"': '&quot;',
    "'": '&#39;'
  }[c]));
}

function getItemUrgency_(s) {
  if (s.state === 'lewat') {
    return {
      type: 'lewat',
      label: 'LEWAT JATUH TEMPO',
      bgBadge: '#fee2e2',
      colorBadge: '#b91c1c',
      borderBadge: '#fca5a5',
      borderColor: '#ef4444',
      icon: '🚨'
    };
  }
  if (s.sisaHari === 0 || s.sisaKm === 0) {
    return {
      type: 'today',
      label: 'JATUH TEMPO HARI INI',
      bgBadge: '#ffedd5',
      colorBadge: '#c2410c',
      borderBadge: '#fdba74',
      borderColor: '#f97316',
      icon: '🔔'
    };
  }
  return {
    type: 'dekat',
    label: 'MENDEKATI JADWAL',
    bgBadge: '#fef3c7',
    colorBadge: '#b45309',
    borderBadge: '#fcd34d',
    borderColor: '#eab308',
    icon: '⏳'
  };
}

function buildSisaKeterangan_(s) {
  const parts = [];
  if (s.sisaKm !== null && s.sisaKm !== undefined) {
    if (s.sisaKm < 0) {
      parts.push(`Melebihi batas ${(-s.sisaKm).toLocaleString('id-ID')} km`);
    } else if (s.sisaKm === 0) {
      parts.push('Target km tercapai hari ini');
    } else {
      parts.push(`Sisa ${s.sisaKm.toLocaleString('id-ID')} km lagi`);
    }
  }
  if (s.sisaHari !== null && s.sisaHari !== undefined) {
    if (s.sisaHari < 0) {
      parts.push(`Terlambat ${-s.sisaHari} hari`);
    } else if (s.sisaHari === 0) {
      parts.push('Jatuh tempo HARI INI');
    } else {
      parts.push(`Sisa ${s.sisaHari} hari lagi`);
    }
  }
  return parts.join(' • ');
}

function buildReminderEmailPayload_(due, appUrl, sheetUrl) {
  const today = new Date();
  const tanggalHariIni = formatTglIndo_(today, true);

  const hasLewat = due.some(s => s.state === 'lewat');
  const hasHariIni = due.some(s => s.sisaHari === 0 || s.sisaKm === 0);

  let subject = '';
  if (due.length === 1) {
    const s = due[0];
    if (s.state === 'lewat') {
      subject = `⚠️ [Lewat Batas] ${s.komponen} (${s.motor}) Perlu Segera Ditangani - Rekap Bengkel`;
    } else if (s.sisaHari === 0 || s.sisaKm === 0) {
      subject = `🔔 [Hari Ini] ${s.komponen} (${s.motor}) Jatuh Tempo Hari Ini - Rekap Bengkel`;
    } else {
      subject = `🔔 [Pengingat] ${s.komponen} (${s.motor}) Mendekati Batas Servis - Rekap Bengkel`;
    }
  } else {
    if (hasLewat) {
      subject = `⚠️ [Penting] ${due.length} Jadwal Servis & Pajak Lewat Batas - Rekap Bengkel`;
    } else if (hasHariIni) {
      subject = `🔔 [Hari Ini] ${due.length} Jadwal Servis & Pajak Jatuh Tempo Hari Ini - Rekap Bengkel`;
    } else {
      subject = `🔔 [Pengingat] ${due.length} Jadwal Servis & Pajak Mendekati Waktu - Rekap Bengkel`;
    }
  }

  // Kelompokkan per motor
  const byMotor = {};
  due.forEach(s => {
    if (!byMotor[s.motor]) byMotor[s.motor] = [];
    byMotor[s.motor].push(s);
  });

  // Plain Text Version
  let plain = `[REKAP BENGKEL] PENGINGAT SERVIS & PAJAK MOTOR\n`;
  plain += `${tanggalHariIni}\n\n`;
  plain += `Halo,\nBerikut adalah jadwal servis berkala dan pajak kendaraan Anda yang memerlukan perhatian:\n\n`;

  Object.entries(byMotor).forEach(([motor, items]) => {
    const motorKm = items[0].motorKm != null ? `${Number(items[0].motorKm).toLocaleString('id-ID')} km` : '-';
    plain += `========================================\n`;
    plain += `🏍️ MOTOR: ${motor.toUpperCase()} (Km saat ini: ${motorKm})\n`;
    plain += `========================================\n`;

    items.forEach((item, idx) => {
      const urg = getItemUrgency_(item);
      const sisaDesc = buildSisaKeterangan_(item);
      const tip = TIPS_KOMPONEN[item.komponen] || 'Lakukan pengecekan berkala demi performa dan kenyamanan berkendara.';

      plain += `\n${idx + 1}. ${item.komponen}\n`;
      plain += `   Status       : [${urg.label}]\n`;
      plain += `   Keterangan   : ${sisaDesc}\n`;

      if (item.komponen === 'Pajak STNK') {
        if (item.dueDate) plain += `   Jatuh Tempo  : ${formatTglIndo_(item.dueDate)}\n`;
      } else {
        if (item.lastKm != null) {
          plain += `   Servis Lalu  : ${item.lastKm.toLocaleString('id-ID')} km ${item.lastDate ? '(' + formatTglIndo_(item.lastDate) + ')' : ''}\n`;
        }
        if (item.targetKm != null) {
          plain += `   Target Servis: ${item.targetKm.toLocaleString('id-ID')} km\n`;
        }
        if (item.intervalKm || item.intervalBulan) {
          const ivl = [];
          if (item.intervalKm) ivl.push(`${item.intervalKm.toLocaleString('id-ID')} km`);
          if (item.intervalBulan) ivl.push(`${item.intervalBulan} bulan`);
          plain += `   Interval     : Setiap ${ivl.join(' / ')}\n`;
        }
      }
      plain += `   💡 Saran     : ${tip}\n`;
    });
    plain += `\n`;
  });

  plain += `----------------------------------------\n`;
  if (appUrl) {
    plain += `Buka Aplikasi Rekap Bengkel:\n${appUrl}\n\n`;
  }
  if (sheetUrl) {
    plain += `Buka Spreadsheet Catatan Servis:\n${sheetUrl}\n\n`;
  }
  plain += `Email ini dikirim otomatis oleh trigger harian Rekap Bengkel.\n`;

  // HTML Version
  let itemsHtml = '';
  Object.entries(byMotor).forEach(([motor, items]) => {
    const motorKm = items[0].motorKm != null ? `${Number(items[0].motorKm).toLocaleString('id-ID')} km` : '-';

    itemsHtml += `
      <!-- Motor Section -->
      <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="margin-top:20px;border-collapse:collapse;border:1px solid #e2e8f0;border-radius:8px;overflow:hidden;background-color:#ffffff;">
        <tr>
          <td style="padding:14px 18px;background-color:#0f172a;color:#ffffff;">
            <table role="presentation" width="100%" cellpadding="0" cellspacing="0">
              <tr>
                <td style="font-size:16px;font-weight:700;letter-spacing:0.5px;color:#ffffff;">
                  <span style="font-size:18px;margin-right:6px;">🏍️</span> ${escHtml_(motor)}
                </td>
                <td style="text-align:right;font-size:13px;color:#94a3b8;">
                  Km saat ini: <strong style="color:#f8fafc;">${escHtml_(motorKm)}</strong>
                </td>
              </tr>
            </table>
          </td>
        </tr>
        <tr>
          <td style="padding:16px 18px;background-color:#f8fafc;">
    `;

    items.forEach((item) => {
      const urg = getItemUrgency_(item);
      const sisaDesc = buildSisaKeterangan_(item);
      const tip = TIPS_KOMPONEN[item.komponen] || 'Lakukan pengecekan berkala demi performa dan keselamatan berkendara.';

      let detailRows = '';
      if (item.komponen === 'Pajak STNK') {
        detailRows = `
          <tr>
            <td style="padding:4px 0;font-size:13px;color:#64748b;width:120px;">Jatuh Tempo:</td>
            <td style="padding:4px 0;font-size:13px;color:#1e293b;font-weight:600;">${escHtml_(formatTglIndo_(item.dueDate))}</td>
          </tr>
        `;
      } else {
        if (item.lastKm != null) {
          detailRows += `
            <tr>
              <td style="padding:4px 0;font-size:13px;color:#64748b;width:120px;">Servis Terakhir:</td>
              <td style="padding:4px 0;font-size:13px;color:#1e293b;font-weight:600;">${item.lastKm.toLocaleString('id-ID')} km ${item.lastDate ? '<span style="font-weight:400;color:#64748b;">(' + escHtml_(formatTglIndo_(item.lastDate)) + ')</span>' : ''}</td>
            </tr>
          `;
        }
        if (item.targetKm != null) {
          detailRows += `
            <tr>
              <td style="padding:4px 0;font-size:13px;color:#64748b;width:120px;">Target Servis:</td>
              <td style="padding:4px 0;font-size:13px;color:#1e293b;font-weight:600;">${item.targetKm.toLocaleString('id-ID')} km</td>
            </tr>
          `;
        }
        if (item.intervalKm || item.intervalBulan) {
          const ivl = [];
          if (item.intervalKm) ivl.push(`${item.intervalKm.toLocaleString('id-ID')} km`);
          if (item.intervalBulan) ivl.push(`${item.intervalBulan} bulan`);
          detailRows += `
            <tr>
              <td style="padding:4px 0;font-size:13px;color:#64748b;width:120px;">Interval Jadwal:</td>
              <td style="padding:4px 0;font-size:13px;color:#1e293b;">Setiap ${escHtml_(ivl.join(' / '))}</td>
            </tr>
          `;
        }
      }

      itemsHtml += `
        <!-- Item Card -->
        <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="margin-bottom:12px;border:1px solid #e2e8f0;border-left:5px solid ${urg.borderColor};border-radius:6px;background-color:#ffffff;box-shadow:0 1px 2px rgba(0,0,0,0.04);">
          <tr>
            <td style="padding:14px 16px;">
              <table role="presentation" width="100%" cellpadding="0" cellspacing="0">
                <tr>
                  <td style="font-size:16px;font-weight:700;color:#0f172a;">
                    ${urg.icon} ${escHtml_(item.komponen)}
                  </td>
                  <td style="text-align:right;">
                    <span style="display:inline-block;padding:3px 8px;font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:0.5px;color:${urg.colorBadge};background-color:${urg.bgBadge};border:1px solid ${urg.borderBadge};border-radius:4px;">
                      ${escHtml_(urg.label)}
                    </span>
                  </td>
                </tr>
              </table>

              <div style="margin:8px 0 10px;font-size:14px;font-weight:600;color:${urg.borderColor};">
                ${escHtml_(sisaDesc)}
              </div>

              <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="margin-bottom:10px;border-top:1px dashed #e2e8f0;padding-top:8px;">
                ${detailRows}
              </table>

              <div style="background-color:#f1f5f9;border-radius:4px;padding:8px 12px;font-size:12px;color:#475569;line-height:1.4;">
                <strong style="color:#334155;">💡 Saran:</strong> ${escHtml_(tip)}
              </div>
            </td>
          </tr>
        </table>
      `;
    });

    itemsHtml += `
          </td>
        </tr>
      </table>
    `;
  });

  const bannerColor = hasLewat ? '#ef4444' : hasHariIni ? '#f97316' : '#3b82f6';
  const bannerBg = hasLewat ? '#fef2f2' : hasHariIni ? '#fff7ed' : '#eff6ff';
  const bannerBorder = hasLewat ? '#fecaca' : hasHariIni ? '#ffedd5' : '#dbeafe';
  const bannerTitle = hasLewat
    ? 'Perhatian: Ada jadwal servis atau pajak yang telah lewat batas!'
    : hasHariIni
    ? 'Pemberitahuan: Ada jadwal servis atau pajak yang jatuh tempo hari ini!'
    : 'Pengingat: Ada jadwal perawatan motor yang mendekati batas waktu/km.';

  const html = `
<!DOCTYPE html>
<html lang="id">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>${escHtml_(subject)}</title>
</head>
<body style="margin:0;padding:0;background-color:#f1f5f9;font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,Helvetica,Arial,sans-serif;color:#1e293b;line-height:1.5;">
  <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background-color:#f1f5f9;padding:24px 12px;">
    <tr>
      <td align="center">
        <!-- Main Card Container -->
        <table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="max-width:600px;background-color:#ffffff;border-radius:10px;overflow:hidden;border:1px solid #e2e8f0;box-shadow:0 4px 6px -1px rgba(0,0,0,0.05);">
          
          <!-- Header Banner -->
          <tr>
            <td style="background-color:#0b0b0c;border-top:5px solid #e10600;padding:24px 24px 20px;text-align:center;">
              <div style="font-size:24px;font-weight:900;letter-spacing:1px;text-transform:uppercase;color:#ffffff;font-family:'Arial Black',Impact,sans-serif;">
                REKAP<span style="color:#e10600;">BENGKEL</span>
              </div>
              <div style="font-size:13px;color:#94a3b8;margin-top:4px;letter-spacing:0.5px;">
                Pengingat Servis &amp; Pajak Motor
              </div>
            </td>
          </tr>

          <!-- Subheader / Date info -->
          <tr>
            <td style="padding:16px 24px 8px;background-color:#ffffff;">
              <table role="presentation" width="100%" cellpadding="0" cellspacing="0">
                <tr>
                  <td style="font-size:13px;color:#64748b;">
                    📅 <strong>${escHtml_(tanggalHariIni)}</strong>
                  </td>
                  <td style="text-align:right;font-size:12px;color:#64748b;">
                    Total: <strong style="color:#0f172a;">${due.length} item</strong>
                  </td>
                </tr>
              </table>
            </td>
          </tr>

          <!-- Alert Banner -->
          <tr>
            <td style="padding:8px 24px 12px;">
              <div style="background-color:${bannerBg};border:1px solid ${bannerBorder};border-left:4px solid ${bannerColor};padding:12px 14px;border-radius:6px;font-size:13px;color:#1e293b;">
                <strong>${escHtml_(bannerTitle)}</strong>
                <div style="margin-top:2px;font-size:12px;color:#64748b;">Periksa rincian di bawah ini dan jadwalkan perawatan motor Anda.</div>
              </div>
            </td>
          </tr>

          <!-- Content List -->
          <tr>
            <td style="padding:0 24px 20px;">
              ${itemsHtml}
            </td>
          </tr>

          <!-- Call To Action -->
          <tr>
            <td style="padding:16px 24px 24px;text-align:center;background-color:#f8fafc;border-top:1px solid #e2e8f0;">
              ${appUrl ? `
                <a href="${escHtml_(appUrl)}" target="_blank" style="display:inline-block;background-color:#e10600;color:#ffffff;text-decoration:none;padding:12px 28px;border-radius:6px;font-size:14px;font-weight:700;letter-spacing:0.5px;text-transform:uppercase;box-shadow:0 2px 4px rgba(225,6,0,0.25);">
                  Buka Rekap Bengkel
                </a>
              ` : ''}
              ${sheetUrl ? `
                <div style="margin-top:12px;">
                  <a href="${escHtml_(sheetUrl)}" target="_blank" style="font-size:12px;color:#64748b;text-decoration:underline;">
                    Lihat Catatan di Google Spreadsheet
                  </a>
                </div>
              ` : ''}
            </td>
          </tr>

          <!-- Footer -->
          <tr>
            <td style="padding:16px 24px;background-color:#0b0b0c;text-align:center;font-size:11px;color:#64748b;line-height:1.6;">
              Email ini dikirim otomatis oleh pengingat harian sistem <strong>Rekap Bengkel</strong>.<br>
              Data tersimpan aman di Google Drive pribadi Anda.
            </td>
          </tr>

        </table>
      </td>
    </tr>
  </table>
</body>
</html>
  `;

  return { subject, plain, html };
}

function sendReminderEmail_(to, due) {
  let appUrl = '';
  try {
    appUrl = ScriptApp.getService().getUrl();
  } catch (e) {}
  if (!appUrl || appUrl.indexOf('http') !== 0) {
    appUrl = 'https://script.google.com/macros/s/AKfycbz5gfC5REH8r8Xhid8fzP_hP0ts6sfcv96XUQgtqDzybltdQ-DidS94tK4TwTM5GVD8qw/exec';
  }

  let sheetUrl = '';
  try {
    sheetUrl = ss_().getUrl();
  } catch (e) {}

  const payload = buildReminderEmailPayload_(due, appUrl, sheetUrl);

  MailApp.sendEmail({
    to: to,
    subject: payload.subject,
    body: payload.plain,
    htmlBody: payload.html,
    name: 'Rekap Bengkel',
  });
}

// Trigger harian (milik tiap pengguna): satu email berisi semua yang dekat/lewat.
function dailyReminder() {
  if (!UP().getProperty('ssid')) return;
  const due = status_().filter(s => s.state === 'dekat' || s.state === 'lewat');
  if (!due.length) return;
  const to = (rows_('Config').find(c => c.key === 'email') || {}).value;
  if (!to) return;
  sendReminderEmail_(to, due);
}

// Fungsi pengujian: bisa dipanggil dari UI atau langsung dijalankan dari Apps Script editor
function sendTestReminderEmail() {
  if (!UP().getProperty('ssid')) throw new Error('Aplikasi belum di-onboard.');
  const to = (rows_('Config').find(c => c.key === 'email') || {}).value || Session.getActiveUser().getEmail();
  if (!to) throw new Error('Email penerima tidak ditemukan di pengaturan Config.');

  let due = status_().filter(s => s.state === 'dekat' || s.state === 'lewat');
  if (!due.length) {
    const motors = rows_('Motor');
    const motorName = motors.length ? motors[0].nama : 'Motor Anda';
    const motorKm = motors.length ? Number(motors[0].km) : 25000;
    due = [
      {
        motor: motorName,
        motorKm: motorKm,
        komponen: 'Pajak STNK',
        sisaKm: null,
        sisaHari: 0,
        state: 'dekat',
        dueDate: new Date(),
        lastKm: null,
        lastDate: null,
        targetKm: null,
        intervalKm: null,
        intervalBulan: null,
      },
      {
        motor: motorName,
        motorKm: motorKm,
        komponen: 'Oli mesin',
        sisaKm: -120,
        sisaHari: null,
        state: 'lewat',
        dueDate: null,
        lastKm: motorKm - 2120,
        lastDate: new Date(Date.now() - 60 * DAY),
        targetKm: motorKm - 120,
        intervalKm: 2000,
        intervalBulan: 2,
      }
    ];
  }

  sendReminderEmail_(to, due);
  return 'Email pengingat berhasil dikirim ke ' + to;
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

  const p = buildReminderEmailPayload_([
    {
      motor: 'CB150R',
      motorKm: 24500,
      komponen: 'Pajak STNK',
      sisaKm: null,
      sisaHari: 0,
      state: 'dekat',
      dueDate: today,
    }
  ], 'https://example.com/app', 'https://example.com/sheet');
  a(p.subject.includes('CB150R'), 'payload subject has motor');
  a(p.plain.includes('CB150R') && p.plain.includes('Pajak STNK'), 'plain body has item');
  a(p.html.includes('REKAP') && p.html.includes('Pajak STNK'), 'html body has content');

  console.log('selfTest OK');
}
