/**
 * BACKEND GOOGLE APPS SCRIPT (GAS) - STANDALONE VERSION (LUAR PULAU)
 * --------------------------------------------------------------------------
 * Dibuat langsung dari Standalone Apps Script (script.google.com) milik Akun B.
 * Google Sheet milik Akun A tetap BERSIH murni (tanpa Extension Apps Script).
 */

// KONFIGURASI UTAMA SPREADSHEET & MULTI-REGION MAP
const SECRET_TOKEN = "VISIT_SECRET_KEY_2026"; // Secret Token Frontend
const SHEET_KUNJUNGAN = "Kunjungan"; // Nama Sheet Kunjungan Utama
const SHEET_MASTER_TOKO = "Master_Toko"; // Nama Sheet Master Toko
const SHEET_DETAIL_AUDIT = "Detail Audit"; // Nama Sheet Detail Visit (Tab 2)
const SHEET_MASTER_PRODUK = "Master_Produk"; // Nama Sheet Master Produk
const SPREADSHEET_MASTER_PRODUK_ID = "1qWupc78oFaCpqc1oki9BLsn2BPWAGR31743LHL8qMSQ"; // ID Spreadsheet Master Produk Akun B

// MAP REGION SPREADSHEETS (PUSAT DI 1 FILE GAS)
const REGION_SPREADSHEETS = {
  LUAR_PULAU: {
    spreadsheetId: "1JpG5MhUmz1qBtXp58b8FgV-_avCjVZKjNzuc824SUcE",
    folderFotoId: "1vFpmiVzhu3qSZCxH938ORtEtaUa3cjqo",
    folderFotoName: "Foto_Kunjungan_App_Luar_Pulau"
  },
  DALAM_PULAU: {
    spreadsheetId: "1ZF8bAChTAYKMbmfeAamE9R0ge6hzS5AeddptOwJauW0",
    folderFotoId: "1npp75ufziJSsO1oHGcn9uJKTVEq2SBfZ",
    folderFotoName: "Foto_Kunjungan_App_Dalam_Pulau"
  }
};

/**
 * Helper Membuka Spreadsheet Berdasarkan Region (Standalone Script)
 */
function getTargetSpreadsheet(region = "LUAR_PULAU") {
  const regConfig = REGION_SPREADSHEETS[region] || REGION_SPREADSHEETS.LUAR_PULAU;
  return SpreadsheetApp.openById(regConfig.spreadsheetId);
}

/**
 * Helper Mendapatkan Folder Google Drive Berdasarkan Region
 */
function getTargetFolder(region = "LUAR_PULAU") {
  const regConfig = REGION_SPREADSHEETS[region] || REGION_SPREADSHEETS.LUAR_PULAU;
  if (regConfig.folderFotoId && regConfig.folderFotoId.trim() !== "") {
    try {
      return DriveApp.getFolderById(regConfig.folderFotoId.trim());
    } catch (e) {}
  }
  
  const folders = DriveApp.getFoldersByName(regConfig.folderFotoName);
  if (folders.hasNext()) {
    return folders.next();
  }
  return DriveApp.createFolder(regConfig.folderFotoName);
}

/**
 * Helper Ambil Data Master Produk (Dynamic Mapping dari Spreadsheet Akun B)
 */
function fetchMasterProdukData() {
  let ss;
  try {
    ss = SpreadsheetApp.openById(SPREADSHEET_MASTER_PRODUK_ID);
  } catch (e) {
    ss = getTargetSpreadsheet("LUAR_PULAU");
  }

  let sheet = ss.getSheetByName(SHEET_MASTER_PRODUK) || ss.getSheets()[0];
  if (!sheet) {
    return createJsonResponse({ status: "success", count: 0, data: {} });
  }
  
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) {
    return createJsonResponse({ status: "success", count: 0, data: {} });
  }
  
  const result = {};
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const barcode = String(row[2] || "").trim();
    if (!barcode) continue;
    
    result[barcode] = {
      brand: String(row[0] || "").trim() || "CIMORY GROUP",
      katagori: String(row[1] || "").trim() || "YOGURT",
      nama_barang: String(row[3] || "").trim(),
      packsize: String(row[4] || "").trim()
    };
  }
  
  return createJsonResponse({ status: "success", count: Object.keys(result).length, data: result });
}

/**
 * Helper Ambil Data Master Toko Berdasarkan Region
 */
function fetchMasterTokoData(region = "LUAR_PULAU") {
  const ss = getTargetSpreadsheet(region);
  const sheet = ss.getSheetByName(SHEET_MASTER_TOKO);
  if (!sheet) {
    return createJsonResponse({ status: "error", message: `Sheet '${SHEET_MASTER_TOKO}' tidak ditemukan di Spreadsheet ${region}.` });
  }
  
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) {
    return createJsonResponse({ status: "success", data: [] });
  }
  
  const result = [];
  
  // Map baris master toko (Mulai baris index 1)
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[1] && !row[2]) continue; // Skip jika kode/nama toko kosong
    
    result.push({
      account: String(row[0] || ""),        // Kolom A (ACCOUNT)
      kodeToko: String(row[1] || ""),       // Kolom B (KODE TOKO)
      namaToko: String(row[2] || ""),       // Kolom C (NAMA TOKO)
      kodeCrew: String(row[3] || ""),       // Kolom D (KODE CREW)
      namaCrew: String(row[4] || ""),       // Kolom E (NAMA CREW)
      rute: String(row[5] || ""),           // Kolom F (RUTE)
      tipeToko: String(row[6] || "")        // Kolom G (TIPE TOKO)
    });
  }
  
  return createJsonResponse({ status: "success", count: result.length, data: result });
}

/**
 * Helper Label Kolom Foto Sesuai Format Native AppSheet
 */
function getAppSheetPhotoLabel(colIndex) {
  const colNum = colIndex + 1;
  if (colNum === 15) return "FOTO SELFIE (DEPAN TOKO)";
  if (colNum === 18) return "FOTO TG SEWA";
  if (colNum === 24) return "FOTO ASSET 1";
  if (colNum === 25) return "FOTO ASSET 2";
  if (colNum === 27) return "FOTO UPDATE PLANOGRAM";

  // Promo Khusus
  if (colNum === 134) return "FOTO PSM 1";
  if (colNum === 135) return "FOTO PSM 2";
  if (colNum === 153) return "FOTO MURAH 1";
  if (colNum === 154) return "FOTO MURAH 2";
  if (colNum === 172) return "FOTO PWP_KHUSUS 1";
  if (colNum === 173) return "FOTO PWP_KHUSUS 2";
  if (colNum === 191) return "FOTO JSM 1";
  if (colNum === 192) return "FOTO JSM 2";
  if (colNum === 210) return "FOTO GANTUNG 1";
  if (colNum === 211) return "FOTO GANTUNG 2";

  // Before / After (Col 213 - 224)
  if (colNum >= 213 && colNum <= 218) return `FOTO BEFORE ${colNum - 212}`;
  if (colNum >= 219 && colNum <= 224) return `FOTO AFTER ${colNum - 218}`;

  // HYSU
  if (colNum === 225) return "FOTO HYSU";

  // Rak Reguler (Col 226 - 240)
  if (colNum >= 226 && colNum <= 240) return `FOTO RAK REG ${colNum - 225}`;

  // Extra Display (Col 241 - 244)
  if (colNum >= 241 && colNum <= 244) return `FOTO EXTRA ${colNum - 240}`;

  return `FOTO COL ${colNum}`;
}

/**
 * Listener HTTP GET - Mengambil Data Master Toko & Health Check
 */
function doGet(e) {
  try {
    const action = e && e.parameter ? e.parameter.action : "";
    
    // Action 1: Ambil Master Toko dari Sheet Master_Toko Sesuai Region
    if (action === "getMasterToko") {
      const region = (e && e.parameter && e.parameter.region) ? e.parameter.region : "LUAR_PULAU";
      return fetchMasterTokoData(region);
    }

    // Action 2: Ambil Master Produk dari Sheet Master_Produk
    if (action === "getMasterProduk") {
      return fetchMasterProdukData();
    }
    
    return createJsonResponse({ status: "success", message: "Standalone Web App Script Active & Connected!" });
  } catch (error) {
    return createJsonResponse({ status: "error", message: error.toString() });
  }
}

/**
 * Listener HTTP POST - Menerima & Menyimpan Payload Data 244 Kolom + Foto Base64
 */
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    
    // Parse JSON Payload
    const requestData = JSON.parse(e.postData.contents || "{}");
    
    // Check if request is action getMasterToko / getMasterProduk via POST
    if (requestData.action === "getMasterToko") {
      const region = requestData.region || "LUAR_PULAU";
      return fetchMasterTokoData(region);
    }
    if (requestData.action === "getMasterProduk") {
      return fetchMasterProdukData();
    }

    const token = requestData.secret_token;
    const rowData = requestData.row_data;
    
    // 1. Validasi Token Keamanan
    if (token !== SECRET_TOKEN) {
      return createJsonResponse({ status: "error", message: "Akses Ditolak: Secret Token tidak cocok!" });
    }
    
    if (!rowData || !Array.isArray(rowData)) {
      return createJsonResponse({ status: "error", message: "Payload row_data tidak valid!" });
    }
    
    const region = requestData.region || "LUAR_PULAU";
    
    // Buka Sheet Target Berdasarkan Region (LUAR_PULAU / DALAM_PULAU)
    const ss = getTargetSpreadsheet(region);
    let sheet = ss.getSheetByName(SHEET_KUNJUNGAN);
    if (!sheet) {
      sheet = ss.getSheets()[0]; // Fallback ke sheet pertama jika sheet "Kunjungan" belum di-rename
    }
    
    // 2. Generate ID_VISIT Unik (8 Karakter Hexadecimal Lowercase)
    const idVisit = generateUniqueIdVisit(sheet).toLowerCase();
    rowData[0] = idVisit; // Set Kolom Index 1 (ID_VISIT)
    
    // 3. Set Tanggal & Waktu (Gunakan input custom check-in dari frontend jika ada, atau fallback ke server time)
    const now = new Date();
    const timeStr = Utilities.formatDate(now, ss.getSpreadsheetTimeZone(), "HH:mm:ss");
    const dateStr = Utilities.formatDate(now, ss.getSpreadsheetTimeZone(), "dd/MM/yyyy");
    
    if (!rowData[2] || rowData[2] === "") rowData[2] = timeStr; // Index 3: WAKTU
    if (!rowData[3] || rowData[3] === "") rowData[3] = dateStr; // Index 4: TANGGAL
    
    const timeHHMMSS = rowData[2].toString().replace(/:/g, ""); // "233158"
    
    // 4. Dapatkan / Buat Folder Google Drive Penampung Foto & Ambil Header Baris 1 secara Dinamis
    const driveFolder = getTargetFolder(region);
    let sheetHeaders = [];
    if (sheet.getLastRow() >= 1) {
      sheetHeaders = sheet.getRange(1, 1, 1, Math.max(244, sheet.getLastColumn())).getValues()[0];
    }
    
    // 5. Proses Upload Gambar Base64 ke Google Drive (Format: {idVisit}.{NAMA_HEADER_BARIS_1}.{HHMMSS})
    for (let i = 0; i < rowData.length; i++) {
      const val = rowData[i];
      if (typeof val === "string" && val.startsWith("data:image/")) {
        const headerName = (sheetHeaders[i] && sheetHeaders[i].toString().trim() !== "") 
          ? sheetHeaders[i].toString().trim() 
          : getAppSheetPhotoLabel(i);
        
        // Format Nama File Presisi 100% AppSheet: {idVisit}.{NAMA_HEADER_BARIS_1}.{HHMMSS}
        const appSheetFileName = `${idVisit}.${headerName}.${timeHHMMSS}`;
        
        const photoUrl = saveBase64ToDrive(val, driveFolder, appSheetFileName);
        rowData[i] = photoUrl; // Ganti string Base64 raksasa dengan URL Google Drive Direct Link
      }
    }
    
    // 6. Pastikan Panjang Array Tepat 244 Kolom
    while (rowData.length < 244) {
      rowData.push("");
    }
    
    // 7. Simpan Baris Baru ke Google Sheets Luar Pulau
    sheet.appendRow(rowData);

    // 7.5 Simpan Detail Visit ke Sheet Detail Audit (Multi-row support per product)
    const rowDetailList = requestData.row_detail_list || (requestData.row_detail_data ? [requestData.row_detail_data] : []);
    
    if (Array.isArray(rowDetailList) && rowDetailList.length > 0) {
      let detailSheet = ss.getSheetByName(SHEET_DETAIL_AUDIT) || ss.getSheetByName("Detail Audit") || ss.getSheetByName("Detail_Kunjungan");
      if (!detailSheet) {
        detailSheet = ss.insertSheet(SHEET_DETAIL_AUDIT);
      }
      
      rowDetailList.forEach(rowDetailData => {
        if (!Array.isArray(rowDetailData) || rowDetailData.every(v => v === '')) return;

        const detailId = generateUniqueIdVisit(detailSheet).toLowerCase();
        rowDetailData[0] = detailId; // DETAIL_VISIT ID
        rowDetailData[1] = idVisit;   // Foreign Key: ID_VISIT (Link ke main visit)
        rowDetailData[2] = rowData[1]; // MAP (from main visit)
        rowDetailData[3] = rowData[2]; // WAKTU (from main visit)
        rowDetailData[4] = rowData[3]; // TANGGAL (from main visit)
        
        // Upload Base64 photos in Detail row
        for (let j = 0; j < rowDetailData.length; j++) {
          const valD = rowDetailData[j];
          if (typeof valD === "string" && valD.startsWith("data:image/")) {
            const fileNameD = `${idVisit}.DETAIL_COL${j+1}.${timeHHMMSS}`;
            rowDetailData[j] = saveBase64ToDrive(valD, driveFolder, fileNameD);
          }
        }
        
        while (rowDetailData.length < 71) {
          rowDetailData.push("");
        }
        
        detailSheet.appendRow(rowDetailData);
      });
    }

    // 8. Auto-Sort Seluruh Baris Berdasarkan Tanggal (Kolom D / Col 4) & Waktu (Kolom C / Col 3)
    const lastRow = sheet.getLastRow();
    if (lastRow > 2) {
      sheet.getRange(2, 1, lastRow - 1, 244).sort([
        { column: 4, ascending: true }, // Tanggal (Ascending)
        { column: 3, ascending: true }  // Waktu (Ascending)
      ]);
    }
    
    return createJsonResponse({
      status: "success",
      id_visit: idVisit,
      message: "Data kunjungan toko berhasil disimpan di Google Sheets Luar Pulau!",
      timestamp: dateStr + " " + timeStr
    });
    
  } catch (err) {
    return createJsonResponse({ status: "error", message: "Server Error: " + err.toString() });
  } finally {
    lock.releaseLock();
  }
}

/**
 * Membuat ID_VISIT 8-karakter hex unik dengan loop collision check di kolom A
 */
function generateUniqueIdVisit(sheet) {
  const existingIds = new Set();
  const lastRow = sheet.getLastRow();
  
  if (lastRow > 1) {
    const ids = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (ids[i][0]) existingIds.add(ids[i][0].toString().toLowerCase());
    }
  }
  
  let newId = "";
  let isUnique = false;
  
  while (!isUnique) {
    newId = Math.floor(Math.random() * 0xFFFFFFFF).toString(16).toLowerCase().padStart(8, '0');
    if (!existingIds.has(newId)) {
      isUnique = true;
    }
  }
  
  return newId;
}

/**
 * Menyimpan String Base64 Image ke File Google Drive & Mengembalikan View Link
 */
function saveBase64ToDrive(base64Data, folder, fileName) {
  const splitData = base64Data.split(",");
  const contentType = splitData[0].split(";")[0].replace("data:", "");
  const byteData = Utilities.base64Decode(splitData[1]);
  
  const blob = Utilities.newBlob(byteData, contentType, `${fileName}.jpg`);
  const file = folder.createFile(blob);
  
  // Set Izin view Anyone with Link agar AppSheet bisa langsung nampilin foto
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  
  // Return URL Direct Viewer
  return "https://lh3.googleusercontent.com/d/" + file.getId();
}

/**
 * Mendapatkan Folder Google Drive (Bisa via ID Folder Akun A atau Auto-Create di Drive Akun B)
 */
function getTargetFolder() {
  if (FOLDER_FOTO_ID && FOLDER_FOTO_ID.trim() !== "") {
    try {
      return DriveApp.getFolderById(FOLDER_FOTO_ID.trim());
    } catch (e) {}
  }
  
  const folders = DriveApp.getFoldersByName(FOLDER_FOTO_NAME);
  if (folders.hasNext()) {
    return folders.next();
  }
  return DriveApp.createFolder(FOLDER_FOTO_NAME);
}

/**
 * Helper Output Response JSON
 */
function createJsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
