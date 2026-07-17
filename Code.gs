// Menampilkan halaman HTML saat webapp diakses
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Aplikasi Audit Internal Matrika')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// Konfigurasi ID (PENTING: Sesuaikan dengan ID Google Drive & Sheet Anda)
const FOLDER_ID = "1xjVi1fxtIcZm28BoPKSyFoKJA_8-Elei";
const SHEET_ID = "1Cmccso6hkwHPJNX1jZy5arK5OAOLnhMBCp89JwJBG50";

// --- Fungsi Utilitas ---
function getSheet(sheetName) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(sheetName);
  
  // Buat sheet jika belum ada (untuk kemudahan setup pertama kali)
  if (!sheet) {
    if (sheetName === 'checklist_template') {
       sheet = ss.insertSheet(sheetName);
       sheet.appendRow(['Kategori', 'No', 'Pertanyaan']);
    } else if (sheetName === 'data_pengamatan') {
       sheet = ss.insertSheet(sheetName);
       sheet.appendRow(['ID', 'Timestamp', 'Auditor', 'Auditee', 'Kategori Menu', 'Data Checklist (JSON)', 'Status', 'Tanda Tangan Auditor', 'Tanda Tangan Auditee', 'Hash_Integritas']);
    } else if (sheetName === 'signatures_db') {
       sheet = ss.insertSheet(sheetName);
       sheet.appendRow(['NIK', 'TandaTanganBase64', 'PIN_Hash', 'Created_At']);
    } else if (sheetName === 'audit_trail') {
       sheet = ss.insertSheet(sheetName);
       sheet.appendRow(['Timestamp', 'Action', 'NIK', 'IP_Address', 'Target_ID', 'Data_Hash']);
    } else if (sheetName === 'list_departemen') {
       sheet = ss.insertSheet(sheetName);
       sheet.appendRow(['ID', 'Kategori', 'Departemen']);
    }
  }
  return sheet;
}

function saveImageToDrive(base64Data, filename) {
  if (!base64Data) return "";
  try {
    const folder = DriveApp.getFolderById(FOLDER_ID);
    const data = base64Data.split(',')[1];
    const blob = Utilities.newBlob(Utilities.base64Decode(data), 'image/png', filename);
    const file = folder.createFile(blob);
    
    // Set permission agar image bisa diload di tag HTML <img>
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return "https://drive.google.com/thumbnail?id=" + file.getId() + "&sz=w1000";
  } catch(e) {
    return "Error Upload: " + e.message;
  }
}

// --- Fungsi Kriptografi & Audit (UU ITE) ---
function hashSHA256(text) {
  const rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text);
  let txtHash = '';
  for (let i = 0; i < rawHash.length; i++) {
    let hashVal = rawHash[i];
    if (hashVal < 0) {
      hashVal += 256;
    }
    if (hashVal.toString(16).length == 1) {
      txtHash += '0';
    }
    txtHash += hashVal.toString(16);
  }
  return txtHash;
}

function logAuditTrail(action, nik, ipAddress, targetId, dataHash) {
  const sheet = getSheet("audit_trail");
  sheet.appendRow([new Date(), action, nik, ipAddress || "Unknown", targetId, dataHash || ""]);
}

function handleRegisterSignature(nik, signatureData, pin) {
  try {
    const sheet = getSheet("signatures_db");
    const data = sheet.getDataRange().getValues();
    const pinHash = hashSHA256(pin);
    
    // Konversi Base64 ke URL Google Drive
    let finalSignatureUrl = signatureData;
    if (signatureData && signatureData.startsWith('data:image')) {
      finalSignatureUrl = saveImageToDrive(signatureData, `TTD_${nik}_${new Date().getTime()}.png`);
    }
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(nik).trim()) {
        sheet.getRange(i + 1, 2).setValue(finalSignatureUrl);
        sheet.getRange(i + 1, 3).setValue(pinHash);
        sheet.getRange(i + 1, 4).setValue(new Date());
        return { success: true, message: "Tanda tangan berhasil diperbarui!" };
      }
    }
    sheet.appendRow([nik, finalSignatureUrl, pinHash, new Date()]);
    return { success: true, message: "Tanda tangan berhasil didaftarkan!" };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function handleVerifySignature(nik, pin) {
  try {
    const sheet = getSheet("signatures_db");
    const data = sheet.getDataRange().getValues();
    const pinHash = hashSHA256(pin);
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(nik).trim()) {
        if (String(data[i][2]).trim() === pinHash) {
          return { success: true, signatureData: data[i][1], message: "Verifikasi berhasil!" };
        } else {
          return { success: false, message: "PIN Salah!" };
        }
      }
    }
    return { success: false, message: "Tanda tangan tidak ditemukan." };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// --- Handler API ---

function handleLogin(nik, password) {
  try {
    const userSheet = getSheet("user");
    if (!userSheet) return { success: false, message: "Sheet 'user' tidak ditemukan." };
    
    const data = userSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(nik).trim() && String(data[i][1]).trim() === String(password).trim()) {
        return { success: true, nama: String(data[i][2]).trim(), message: "Login Berhasil!" };
      }
    }
    return { success: false, message: "NIK atau Password salah!" };
  } catch (error) {
    return { success: false, message: "Error sistem: " + error.message };
  }
}

function handleGetChecklist(kategori) {
  try {
    const sheet = getSheet("checklist_template");
    const data = sheet.getDataRange().getValues();
    let questions = [];
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim().toUpperCase() === kategori.toUpperCase()) {
        questions.push({
          no: data[i][1],
          type: data[i][2],
          kriteria: data[i][3],
          pertanyaan: data[i][4]
        });
      }
    }
    
    // Fallback Mock Data jika sheet kosong (untuk demo)
    if (questions.length === 0) {
      questions = [
        { no: "1.", type: 1, kriteria: "", pertanyaan: "KOMITMEN DAN TANGGUNG JAWAB" },
        { no: "1.1.", type: 2, kriteria: "", pertanyaan: "KEBIJAKAN HALAL" },
        { no: "1.1.c", type: 4, kriteria: "1", pertanyaan: "Ada kegiatan sosialisasi Kebijakan Halal; sebutkan bentuk sosialisasinya" }
      ];
    }
    
    return { success: true, data: questions };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function handleGetDepartemen(kategori) {
  try {
    const sheet = getSheet("list_departemen");
    const data = sheet.getDataRange().getValues();
    let depts = [];
    
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]).trim().toUpperCase() === kategori.toUpperCase()) {
        depts.push({
          id: data[i][0],
          nama: data[i][2]
        });
      }
    }
    
    return { success: true, data: depts };
  } catch(e) {
    return { success: false, message: e.message };
  }
}


function handleSaveData(formData, status) {
  try {
    const sheet = getSheet("data_pengamatan");
    let id = formData.id;
    let isUpdate = false;
    let rowIndex = -1;
    
    if (id) {
      isUpdate = true;
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === id) {
          rowIndex = i + 1;
          break;
        }
      }
    } else {
      id = Utilities.getUuid();
    }
    
    let urlAuditor = formData.auditorSignature || "";
    let urlAuditee = formData.auditeeSignature || "";
    let hashIntegritas = "";

    if (status === 'Submitted') {
       if (urlAuditor.startsWith('data:image')) {
         urlAuditor = saveImageToDrive(urlAuditor, id + "_auditor.png");
       }
       if (urlAuditee.startsWith('data:image')) {
         urlAuditee = saveImageToDrive(urlAuditee, id + "_auditee.png");
       }
       
       // Generate Hash Integritas Dokumen (Pasal 11 UU ITE)
       const documentDataString = JSON.stringify({
         id: id,
         auditor: formData.auditorName,
         auditee: formData.auditeeName,
         kategori: formData.kategori,
         checklist: formData.checklistData
       });
       hashIntegritas = hashSHA256(documentDataString);
       
       // Record Audit Trail
       logAuditTrail("SUBMIT_REPORT_SIGNED", formData.nik || "Unknown", formData.ipAddress, id, hashIntegritas);
    } else {
       logAuditTrail("SAVE_DRAFT", formData.nik || "Unknown", formData.ipAddress, id, "");
    }

    const rowData = [
      id,
      new Date(),
      formData.auditorName,
      formData.auditeeName,
      formData.kategori,
      JSON.stringify(formData.checklistData),
      status,
      urlAuditor,
      urlAuditee,
      hashIntegritas
    ];

    if (isUpdate && rowIndex > -1) {
      sheet.getRange(rowIndex, 1, 1, rowData.length).setValues([rowData]);
    } else {
      sheet.appendRow(rowData);
    }
    
    return { success: true, hash: hashIntegritas, message: `Data berhasil disimpan sebagai ${status}!` };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function handleGetList(auditorName) {
  try {
    const sheet = getSheet("data_pengamatan");
    const data = sheet.getDataRange().getValues();
    let list = [];
    
    for (let i = 1; i < data.length; i++) {
      list.push({
        id: data[i][0],
        tanggal: data[i][1],
        auditor: data[i][2],
        auditee: data[i][3],
        kategori: data[i][4],
        checklistData: data[i][5], // Kirim string JSON agar bisa diproses Export di client
        status: data[i][6]
      });
    }
    list.reverse();
    return { success: true, data: list };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function handleDeleteData(id) {
  try {
    const sheet = getSheet("data_pengamatan");
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === id) {
        sheet.deleteRow(i + 1);
        return { success: true, message: "Data berhasil dihapus!" };
      }
    }
    return { success: false, message: "Data tidak ditemukan." };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function handleGetDetail(id) {
  try {
    const sheet = getSheet("data_pengamatan");
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === id) {
        return { 
          success: true, 
          data: {
            id: data[i][0],
            tanggal: data[i][1],
            auditorName: data[i][2],
            auditeeName: data[i][3],
            kategori: data[i][4],
            checklistData: JSON.parse(data[i][5]),
            status: data[i][6],
            auditorSignature: data[i][7],
            auditeeSignature: data[i][8],
            hashIntegritas: data[i][9]
          }
        };
      }
    }
    return { success: false, message: "Data tidak ditemukan." };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

// --- POST Entry Point ---
function doPost(e) {
  try {
    const postData = JSON.parse(e.postData.contents);
    let result = {};
    
    switch (postData.action) {
      case "login":
        result = handleLogin(postData.nik, postData.password);
        break;
      case "registerSignature":
        result = handleRegisterSignature(postData.nik, postData.signatureData, postData.pin);
        break;
      case "verifySignature":
        result = handleVerifySignature(postData.nik, postData.pin);
        break;
      case "getChecklist":
        result = handleGetChecklist(postData.kategori);
        break;
      case "getDepartemen":
        result = handleGetDepartemen(postData.kategori);
        break;
      case "saveDraft":
        result = handleSaveData(postData.formData, "Draft");
        break;
      case "submitFinal":
        result = handleSaveData(postData.formData, "Submitted");
        break;
      case "getList":
        result = handleGetList(postData.auditorName);
        break;
      case "getDetail":
        result = handleGetDetail(postData.id);
        break;
      case "deleteData":
        result = handleDeleteData(postData.id);
        break;
      default:
        result = { success: false, message: "Action tidak dikenal." };
    }
    
    return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: error.message })).setMimeType(ContentService.MimeType.JSON);
  }
}

// --- FUNGSI IMPORT DATA (JALANKAN SEKALI DARI EDITOR) ---
// Pilih fungsi importWarehouseData di dropdown (atas) lalu klik "Run" / "Jalankan"
function importWarehouseData() {
  const sheet = getSheet("checklist_template");
  
  const questions = [
    { no: "1.1.c", pertanyaan: "Ada kegiatan sosialisasi Kebijakan Halal; sebutkan bentuk sosialisasinya" },
    { no: "2.f", pertanyaan: "Semua bahan baku (ingredient) memiliki kode label yang sesuai standard. Lakukan telusur kesesuaian secara fisik yang meliputi: nama supplier, principle, asal negara principle, COA" },
    { no: "3.3.G.a", pertanyaan: "Penyimpanan di gudang dilakukan sesuai prosedur kehalalan dan terhindar dari kontaminasi bahan najis/haram:\n- Bahan baku, bahan tambahan, bahan kemasan dan pelengkap\n* Bahan baku pembuat kemasan dan pelengkap\n* Produk jadi kemasan dan pelengkap\n- Bahan penolong\n- Bahan pelumas, bahan pencucian" },
    { no: "3.3.G.b", pertanyaan: "Semua bahan dan produk yang disimpan di gudang sudah jelas status kehalalannya" },
    { no: "3.3.G.c", pertanyaan: "Gudang hanya digunakan untuk penyimpanan bahan dan produk yang tidak terkontaminasi bahan najis/haram" },
    { no: "3.3.G.d", pertanyaan: "Dokumentasi penerimaan, penyimpanan dan pengeluaran bahan dan produk di gudang dilakukan dengan baik" },
    { no: "3.3.G.e", pertanyaan: "Tangki-tangki penyimpanan (minyak goreng, air) dalam kondisi tertutup dan dapat mencegah masuknya bahan najis/haram atau benda asing" },
    { no: "3.3.G.f", pertanyaan: "Alat untuk menangani atau memindahkan bahan dan produk (forklift, lori, reach truck, dll) dalam kondisi bersih dan tidak terkontaminasi bahan najis/haram" },
    { no: "3.3.G.h", pertanyaan: "Bahan dan produk di gudang yang diketahui tidak memenuhi persyaratan kehalalan telah dipisahkan dan ditindaklanjuti sesuai prosedur" },
    { no: "3.3.H.a", pertanyaan: "Dilakukan pemeriksaan kehalalan fisik alat transportasi pada saat dilakukan pembongkaran/unloading bahan atau produk" },
    { no: "3.3.H.b", pertanyaan: "Dilakukan pemeriksaan kehalalan fisik alat transportasi pada saat dilakukan pemuatan/loading bahan atau produk" },
    { no: "3.3.H.d", pertanyaan: "Alat transportasi yang diketahui tidak memenuhi persyaratan kehalalan telah dipisahkan dan ditindaklanjuti sesuai prosedur" },
    { no: "4.2.b", pertanyaan: "Dapat dilakukan penelusuran produk dengan baik terhadap bahan yang digunakan dan proses yang dilakukan untuk produk tersebut" }
  ];
  
  // Menggunakan kategori 'Halal' agar sinkron dengan tombol di aplikasi
  const kategori = "Halal"; 

  questions.forEach((q) => {
    sheet.appendRow([kategori, q.no, q.pertanyaan]);
  });
  
  Logger.log("Data dari Excel berhasil diimport ke sheet 'checklist_template'!");
}