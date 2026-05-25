// Menampilkan halaman HTML saat webapp diakses
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Aplikasi Audit Internal')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Fungsi untuk memproses data dari frontend
function processAuditForm(formData) {
  try {
    // 1. ID Folder Google Drive Anda (Tempat menyimpan gambar TTD)
    const folderId = "1xjVi1fxtIcZm28BoPKSyFoKJA_8-Elei"; 
    const folder = DriveApp.getFolderById(folderId);
    
    // Proses Tanda Tangan Auditor
    let auditorSignUrl = "";
    if (formData.auditorSignature) {
      const auditorSignData = formData.auditorSignature.split(',')[1];
      const auditorBlob = Utilities.newBlob(Utilities.base64Decode(auditorSignData), 'image/png', formData.auditorName + '_auditor_sign.png');
      const auditorFile = folder.createFile(auditorBlob);
      auditorSignUrl = auditorFile.getUrl(); // Mengambil link gambar TTD di Drive
    }

    // Proses Tanda Tangan Auditee
    let auditeeSignUrl = "";
    if (formData.auditeeSignature) {
      const auditeeSignData = formData.auditeeSignature.split(',')[1];
      const auditeeBlob = Utilities.newBlob(Utilities.base64Decode(auditeeSignData), 'image/png', formData.auditeeName + '_auditee_sign.png');
      const auditeeFile = folder.createFile(auditeeBlob);
      auditeeSignUrl = auditeeFile.getUrl(); // Mengambil link gambar TTD di Drive
    }

    // 2. ID Google Sheet Anda (Sebagai Database)
    const sheetId = "1Cmccso6hkwHPJNX1jZy5arK5OAOLnhMBCp89JwJBG50";
    const ss = SpreadsheetApp.openById(sheetId);
    const sheet = ss.getSheets()[0]; // Otomatis menggunakan sheet pertama (paling kiri)
    
    sheet.appendRow([
      new Date(),                     // Kolom A: Timestamp
      formData.auditorName,           // Kolom B: Nama Auditor
      formData.auditeeName,           // Kolom C: Nama Auditee
      formData.checklistData,         // Kolom D: Hasil Checklist
      formData.pengamatan,            // Kolom E: Form Pengamatan
      formData.lks,                   // Kolom F: Form LKS
      auditorSignUrl,                 // Kolom G: Link TTD Auditor
      auditeeSignUrl                  // Kolom H: Link TTD Auditee
    ]);
    
    return { success: true, message: "Data Audit Berhasil Disimpan ke Google Sheet!" };
  } catch (error) {
    return { success: false, message: "Terjadi kesalahan: " + error.message };
  }
}