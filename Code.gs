const SHEET_NAME = 'Absensi';
const TIMEZONE = 'Asia/Jakarta';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Absensi Siswa')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow(['Waktu Kirim', 'Tanggal', 'Nomor Urut', 'Nama', 'Kelas', 'Kehadiran']);
    sh.setFrozenRows(1);
  }
  sh.getRange(1, 1, 1, 6).setFontWeight('bold').setBackground('#123b35').setFontColor('#ffffff');
  sh.autoResizeColumns(1, 6);
  return 'Sheet absensi siap digunakan.';
}

function submitAttendance(data) {
  if (!data || !Number.isInteger(Number(data.nomor)) || Number(data.nomor) < 1 || Number(data.nomor) > 34) {
    return {ok:false, message:'Nomor urut harus antara 1 sampai 34.'};
  }
  const nama = String(data.nama || '').trim();
  const kelas = String(data.kelas || '').trim();
  const status = String(data.status || '').trim();
  if (!nama || !kelas || !['Hadir','Izin','Sakit'].includes(status)) {
    return {ok:false, message:'Data belum lengkap atau status tidak valid.'};
  }

  // Lock mencegah dua pengiriman bersamaan membuat nomor ganda.
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sh = ss.getSheetByName(SHEET_NAME);
    if (!sh) { setup(); sh = ss.getSheetByName(SHEET_NAME); }
    const today = Utilities.formatDate(new Date(), TIMEZONE, 'yyyy-MM-dd');
    const last = sh.getLastRow();
    if (last >= 2) {
      const existing = sh.getRange(2, 2, last - 1, 2).getValues();
      const duplicate = existing.some(row =>
        String(row[0]) === today && Number(row[1]) === Number(data.nomor)
      );
      if (duplicate) return {ok:false, message:'Nomor urut ini sudah mengirim absensi hari ini. Hubungi guru jika perlu koreksi.'};
    }
    sh.appendRow([new Date(), today, Number(data.nomor), nama, kelas, status]);
    return {ok:true, message:'Absensi berhasil dikirim. Terima kasih!'};
  } finally {
    lock.releaseLock();
  }
}
