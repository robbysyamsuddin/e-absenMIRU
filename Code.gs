// ============================================================
// code.gs — E-Absen Madrasah
// Update: tambah CORS header agar bisa diakses dari GitHub Pages
// ============================================================

function doGet(e) {
  // Jika ada parameter 'nama', proses absensi
  if (e && e.parameter && e.parameter.nama) {
    const hasil = submitAbsensi(e.parameter.nama);
    return ContentService.createTextOutput(hasil).setMimeType(
      ContentService.MimeType.TEXT,
    );
  }

  // Jika tidak ada parameter, tampilkan halaman HTML (opsional, tidak wajib)
  return ContentService.createTextOutput(
    "E-Absen Madrasah API aktif.",
  ).setMimeType(ContentService.MimeType.TEXT);
}

// Untuk mendukung preflight CORS dari browser modern
function doOptions(e) {
  return ContentService.createTextOutput("").setMimeType(
    ContentService.MimeType.TEXT,
  );
}

function submitAbsensi(nama) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Absensi");
  const timezone = "GMT+7";
  const date = new Date();

  const tglSaja = Utilities.formatDate(date, timezone, "yyyy-MM-dd");
  const bulan = Utilities.formatDate(date, timezone, "MMMM yyyy");

  // Validasi: cek nama terdaftar di Data_Guru
  const guruSheet = ss.getSheetByName("Data_Guru");
  const daftarGuru = guruSheet
    .getRange(2, 1, guruSheet.getLastRow() - 1, 1)
    .getValues()
    .flat()
    .filter(String); // hapus baris kosong

  if (!daftarGuru.includes(nama)) {
    return "❌ Error: Nama tidak terdaftar di sistem!";
  }

  // Validasi: cek sudah absen hari ini
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const tglLog =
      data[i][0] instanceof Date
        ? Utilities.formatDate(data[i][0], timezone, "yyyy-MM-dd")
        : "";
    if (data[i][1] === nama && tglLog === tglSaja) {
      return "⚠️ " + nama + ", Anda sudah absen hari ini.";
    }
  }

  // Simpan data absensi
  sheet.appendRow([date, nama, "Hadir", bulan]);
  return "✅ Berhasil! Terima kasih " + nama + ", selamat bertugas.";
}
