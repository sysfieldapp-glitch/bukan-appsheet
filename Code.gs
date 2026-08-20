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
const SPREADSHEET_MASTER_PRODUK_ID =
  "1qWupc78oFaCpqc1oki9BLsn2BPWAGR31743LHL8qMSQ"; // ID Spreadsheet Master Produk Akun B

// MAP REGION SPREADSHEETS (PUSAT DI 1 FILE GAS)
const REGION_SPREADSHEETS = {
  LUAR_PULAU: {
    spreadsheetId: "1kUWJIQxtSkjebZMualR2bIV6HyGxp-baDVz1s-7KspU", // LP4_Data Internal (AppSheet 05. RO MDS LP 4)
    folderFotoId: "1vFpmiVzhu3qSZCxH938ORtEtaUa3cjqo",
    folderFotoName: "Foto_Kunjungan_App_Luar_Pulau",
    photoPrefix: "MDS LP4_",
  },
  LUAR_PULAU_BARU: {
    spreadsheetId: "1S__W_tKymV2xwqx_-vthpPt5jn5u7t3ePlgPqM1opMM", // LP3_Data Internal (AppSheet 05. RO MDS LP 3)
    folderFotoId: "1vFpmiVzhu3qSZCxH938ORtEtaUa3cjqo",
    folderFotoName: "Foto_Kunjungan_App_Luar_Pulau_Baru",
    photoPrefix: "MDS LP3_",
  },
  DALAM_PULAU: {
    spreadsheetId: "1asDdjDm0kUfFmICLtkhJ5VBhmKUels2c-H8Cd22qYvk",
    folderFotoId: "1npp75ufziJSsO1oHGcn9uJKTVEq2SBfZ",
    folderFotoName: "Foto_Kunjungan_App_Dalam_Pulau",
    photoPrefix: "MDS DK1_",
  },
};

/**
 * Helper Membuka Spreadsheet Berdasarkan Region (Standalone Script)
 */
function getTargetSpreadsheet(region = "LUAR_PULAU") {
  const regConfig =
    REGION_SPREADSHEETS[region] || REGION_SPREADSHEETS.LUAR_PULAU;
  return SpreadsheetApp.openById(regConfig.spreadsheetId);
}

/**
 * Helper Mendapatkan Folder Google Drive Berdasarkan Region
 */
function getTargetFolder(region = "LUAR_PULAU") {
  const regConfig =
    REGION_SPREADSHEETS[region] || REGION_SPREADSHEETS.LUAR_PULAU;

  // Opsi 1: Coba ambil folder by ID (paling presisi)
  if (regConfig.folderFotoId && regConfig.folderFotoId.trim() !== "") {
    try {
      return DriveApp.getFolderById(regConfig.folderFotoId.trim());
    } catch (e) {}
  }

  // Opsi 2: Cari by nama, atau buat baru
  try {
    const folders = DriveApp.getFoldersByName(regConfig.folderFotoName);
    if (folders.hasNext()) {
      return folders.next();
    }
    return DriveApp.createFolder(regConfig.folderFotoName);
  } catch (err) {}

  // Opsi 3: Fallback root — kalau ini juga gagal, return null (JANGAN throw)
  try {
    return DriveApp.getRootFolder();
  } catch (e) {}

  // Semua DriveApp access gagal (permission error Akun B) → return null safely
  return null;
}

/**
 * Helper Ambil Data Master Produk (Dynamic Mapping dari Spreadsheet Akun B)
 */
function fetchMasterProdukData(callback = "") {
  let ss;
  try {
    ss = SpreadsheetApp.openById(SPREADSHEET_MASTER_PRODUK_ID);
  } catch (e) {
    ss = getTargetSpreadsheet("LUAR_PULAU");
  }

  let sheet = ss.getSheetByName(SHEET_MASTER_PRODUK) || ss.getSheets()[0];
  if (!sheet) {
    return createJsonResponse(
      { status: "success", count: 0, data: {} },
      callback,
    );
  }

  const data = sheet.getDataRange().getValues();
  const result = {};
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const barcode = String(row[2] || "").trim();
    if (!barcode) continue;

    result[barcode] = {
      brand: String(row[0] || "").trim() || "CIMORY GROUP",
      katagori: String(row[1] || "").trim() || "YOGURT",
      nama_barang: String(row[3] || "").trim(),
      packsize: String(row[4] || "").trim(),
    };
  }

  return createJsonResponse(
    { status: "success", count: Object.keys(result).length, data: result },
    callback,
  );
}

/**
 * Helper Ambil Data Master Toko Berdasarkan Region
 */
function fetchMasterTokoData(region = "LUAR_PULAU", callback = "") {
  const ss = getTargetSpreadsheet(region);
  const sheet = ss.getSheetByName(SHEET_MASTER_TOKO);
  if (!sheet) {
    return ContentService.createTextOutput(
      JSON.stringify({
        status: "error",
        message: `Sheet '${SHEET_MASTER_TOKO}' tidak ditemukan di Spreadsheet ${region}.`,
      }),
    ).setMimeType(ContentService.MimeType.JSON);
  }

  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) {
    return ContentService.createTextOutput(
      JSON.stringify({ status: "success", data: [] }),
    ).setMimeType(ContentService.MimeType.JSON);
  }

  const result = [];

  // Map baris master toko (Mulai baris index 1)
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (!row[1] && !row[2]) continue; // Skip jika kode/nama toko kosong

    result.push({
      account: String(row[0] || ""), // Kolom A (ACCOUNT)
      kodeToko: String(row[1] || ""), // Kolom B (KODE TOKO)
      namaToko: String(row[2] || ""), // Kolom C (NAMA TOKO)
      kodeCrew: String(row[3] || ""), // Kolom D (KODE CREW)
      namaCrew: String(row[4] || ""), // Kolom E (NAMA CREW)
      rute: String(row[5] || ""), // Kolom F (RUTE)
      tipeToko: String(row[6] || ""), // Kolom G (TIPE TOKO)
    });
  }

  return ContentService.createTextOutput(
    JSON.stringify({ status: "success", count: result.length, data: result }),
  ).setMimeType(ContentService.MimeType.JSON);
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
  if (colNum === 141) return "FOTO PSM 1";
  if (colNum === 142) return "FOTO PSM 2";
  if (colNum === 161) return "FOTO MURAH 1";
  if (colNum === 162) return "FOTO MURAH 2";
  if (colNum === 181) return "FOTO PWP_KHUSUS 1";
  if (colNum === 182) return "FOTO PWP_KHUSUS 2";
  if (colNum === 201) return "FOTO JSM 1";
  if (colNum === 202) return "FOTO JSM 2";
  if (colNum === 221) return "FOTO GANTUNG 1";
  if (colNum === 222) return "FOTO GANTUNG 2";

  // Before / After (Col 224 - 235)
  if (colNum >= 224 && colNum <= 229) return `FOTO BEFORE ${colNum - 223}`;
  if (colNum >= 230 && colNum <= 235) return `FOTO AFTER ${colNum - 229}`;

  // HYSU
  if (colNum === 236) return "FOTO HYSU";

  // Rak Reguler (Col 237 - 251)
  if (colNum >= 237 && colNum <= 251) return `FOTO RAK REG ${colNum - 236}`;

  // Extra Display (Col 252 - 255)
  if (colNum >= 252 && colNum <= 255) return `FOTO EXTRA ${colNum - 251}`;

  return `FOTO COL ${colNum}`;
}

/**
 * Listener HTTP GET - Mengambil Data Master Toko & Master Produk (Sesuai Pola gas.js)
 */
function doGet(e) {
  try {
    const action = e && e.parameter ? e.parameter.action : "";

    // Action 1: Ambil Master Toko dari Sheet Master_Toko Sesuai Region
    if (action === "getMasterToko") {
      const region =
        e && e.parameter && e.parameter.region
          ? e.parameter.region
          : "LUAR_PULAU";
      return fetchMasterTokoData(region);
    }

    // Action 2: Ambil Master Produk dari Sheet Master_Produk
    if (action === "getMasterProduk") {
      return fetchMasterProdukData();
    }

    // Action 3: getVisitedToday via GET (100% Bebas CORS Blocked Error di Browser)
    if (action === "getVisitedToday") {
      const token = e && e.parameter ? e.parameter.secret_token : "";
      if (token !== SECRET_TOKEN) {
        return createJsonResponse({
          status: "error",
          message: "Akses Ditolak: Token tidak cocok!",
        });
      }
      const idCrewClean = ((e && e.parameter && e.parameter.id_crew) || "")
        .toString()
        .trim()
        .toLowerCase();
      const visitedCodes = [];
      const regionsToScan = ["LUAR_PULAU", "LUAR_PULAU_BARU", "DALAM_PULAU"];

      regionsToScan.forEach((reg) => {
        try {
          const ss = getTargetSpreadsheet(reg);
          const sheet = ss.getSheetByName(SHEET_KUNJUNGAN) || ss.getSheets()[0];
          const lastRow = sheet.getLastRow();
          if (lastRow >= 2) {
            const tz = ss.getSpreadsheetTimeZone();
            const todayDateStr = Utilities.formatDate(
              new Date(),
              tz,
              "dd/MM/yyyy",
            );

            const range = sheet.getRange(2, 1, lastRow - 1, 14);
            const values = range.getValues();

            values.forEach((row) => {
              const rawDate = row[3];
              if (isDateToday(rawDate, tz)) {
                const rowIdCrew = row[7] ? row[7].toString().trim().toLowerCase() : "";
                const rowNamaCrew = row[8] ? row[8].toString().trim().toLowerCase() : "";
                let kodeToko = row[11] ? row[11].toString().trim().toUpperCase() : "";

                if (!kodeToko && row[12]) {
                  const m = row[12].toString().match(/^[A-Z0-9]+/i);
                  if (m) kodeToko = m[0].toUpperCase();
                }

                const crewNumMatch = idCrewClean.match(/\d+/);
                const crewNum = crewNumMatch ? crewNumMatch[0] : "";

                const isCrewMatch = !idCrewClean ||
                                    rowIdCrew === idCrewClean ||
                                    rowIdCrew.includes(idCrewClean) ||
                                    rowNamaCrew.includes(idCrewClean) ||
                                    (crewNum && crewNum.length >= 2 && (rowIdCrew.includes(crewNum) || rowNamaCrew.includes(crewNum))) ||
                                    rowNamaCrew.includes("yohandi");

                if (isCrewMatch && kodeToko && !visitedCodes.includes(kodeToko)) {
                  visitedCodes.push(kodeToko);
                }
              }
            });
          }
        } catch (errReg) {}
      });

      return createJsonResponse({
        status: "success",
        count: visitedCodes.length,
        visited_stores: visitedCodes,
        visited: visitedCodes,
      });
    }

    return ContentService.createTextOutput(
      JSON.stringify({
        status: "success",
        message: "Backend Google Apps Script Active!",
      }),
    ).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(
      JSON.stringify({ status: "error", message: error.toString() }),
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Helper Memeriksa Apakah Tanggal Adalah Hari Ini (Super Fleksibel untuk M/d/yyyy, dd/MM/yyyy, yyyy-MM-dd, Date Object)
 */
function isDateToday(rawDate, tz) {
  if (!rawDate) return false;
  const now = new Date();
  const todayD = now.getDate();
  const todayM = now.getMonth() + 1;
  const todayY = now.getFullYear();

  if (Object.prototype.toString.call(rawDate) === "[object Date]") {
    try {
      const formatted = Utilities.formatDate(rawDate, tz || "GMT+8", "yyyy-MM-dd");
      const p = formatted.split("-");
      return parseInt(p[0], 10) === todayY && parseInt(p[1], 10) === todayM && parseInt(p[2], 10) === todayD;
    } catch (eDate) {
      return rawDate.getDate() === todayD && (rawDate.getMonth() + 1) === todayM && rawDate.getFullYear() === todayY;
    }
  }

  const s = rawDate.toString().trim();
  if (!s) return false;

  if (s.includes("/")) {
    const parts = s.split("/");
    if (parts.length === 3) {
      const p1 = parseInt(parts[0], 10);
      const p2 = parseInt(parts[1], 10);
      const p3 = parseInt(parts[2], 10);
      if ((p1 === todayM && p2 === todayD && p3 === todayY) || (p1 === todayD && p2 === todayM && p3 === todayY)) {
        return true;
      }
    }
  } else if (s.includes("-")) {
    const parts = s.split("T")[0].split("-");
    if (parts.length === 3) {
      const y = parseInt(parts[0], 10);
      const m = parseInt(parts[1], 10);
      const d = parseInt(parts[2], 10);
      if (y === todayY && m === todayM && d === todayD) return true;
    }
  }
  return false;
}

/**
 * Listener HTTP POST - Menerima & Menyimpan Payload Data 244 Kolom + Foto Base64
 */
function doPost(e) {
  try {
    // Parse JSON Payload
    const requestData = JSON.parse(e.postData.contents || "{}");

    // Check if request is action getMasterToko / getMasterProduk via POST (Read-only — No Lock Needed)
    if (requestData.action === "getMasterToko") {
      const region = requestData.region || "LUAR_PULAU";
      return fetchMasterTokoData(region, requestData.callback || "");
    }
    if (requestData.action === "getMasterProduk") {
      return fetchMasterProdukData(requestData.callback || "");
    }

    // Action: getVisitedToday (Read-only — No Lock Needed)
    if (requestData.action === "getVisitedToday") {
      const token = requestData.secret_token;
      if (token !== SECRET_TOKEN) {
        return createJsonResponse({
          status: "error",
          message: "Akses Ditolak: Token tidak cocok!",
        });
      }
      const idCrewClean = (requestData.id_crew || "")
        .toString()
        .trim()
        .toLowerCase();
      const visitedCodes = [];
      const regionsToScan = ["LUAR_PULAU", "LUAR_PULAU_BARU", "DALAM_PULAU"];

      regionsToScan.forEach((reg) => {
        try {
          const ss = getTargetSpreadsheet(reg);
          const sheet = ss.getSheetByName(SHEET_KUNJUNGAN) || ss.getSheets()[0];
          const lastRow = sheet.getLastRow();
          if (lastRow >= 2) {
            const tz = ss.getSpreadsheetTimeZone();
            const range = sheet.getRange(2, 1, lastRow - 1, 14); // kolom 1-14
            const values = range.getValues();

            values.forEach((row) => {
              const rawDate = row[3];
              if (isDateToday(rawDate, tz)) {
                const rowIdCrew = row[7] ? row[7].toString().trim().toLowerCase() : "";
                const rowNamaCrew = row[8] ? row[8].toString().trim().toLowerCase() : "";
                let kodeToko = row[11] ? row[11].toString().trim().toUpperCase() : "";

                if (!kodeToko && row[12]) {
                  const m = row[12].toString().match(/^[A-Z0-9]+/i);
                  if (m) kodeToko = m[0].toUpperCase();
                }

                const isCrewMatch = !idCrewClean ||
                                    rowIdCrew === idCrewClean ||
                                    rowIdCrew.includes(idCrewClean) ||
                                    rowNamaCrew.includes(idCrewClean) ||
                                    (idCrewClean.includes("ro036") && rowNamaCrew.includes("yohandi"));

                if (isCrewMatch && kodeToko && !visitedCodes.includes(kodeToko)) {
                  visitedCodes.push(kodeToko);
                }
              }
            });
          }
        } catch (errReg) {}
      });

      return createJsonResponse({
        status: "success",
        count: visitedCodes.length,
        visited_stores: visitedCodes,
        visited: visitedCodes
      });
    }

    // 🔒 PROSES WRITE DATA KUNJUNGAN (Memerlukan Script Lock)
    const lock = LockService.getScriptLock();
    try {
      lock.waitLock(30000);
    } catch (lockErr) {
      return createJsonResponse({
        status: "error",
        message:
          "Sistem sedang sibuk memproses pengiriman data lain. Silakan coba lagi.",
      });
    }

    // 1. Verifikasi Secret Token
    const secretToken = requestData.secret_token;
    if (secretToken !== SECRET_TOKEN) {
      return createJsonResponse({
        status: "error",
        message: "Akses Ditolak: Secret Token tidak cocok!",
      });
    }

    const rowData = requestData.row_data;
    if (!rowData || !Array.isArray(rowData)) {
      return createJsonResponse({
        status: "error",
        message: "Payload row_data tidak valid!",
      });
    }

    const region = requestData.region || "LUAR_PULAU";

    // Buka Sheet Target Berdasarkan Region (LUAR_PULAU / LUAR_PULAU_BARU / DALAM_PULAU)
    const ss = getTargetSpreadsheet(region);
    let sheet = ss.getSheetByName(SHEET_KUNJUNGAN);
    if (!sheet) {
      sheet = ss.getSheets()[0];
    }

    // 2. Gunakan ID_VISIT Konsisten dari Frontend (Mencegah Duplikasi Row/Foto Saat Retry)
    let idVisit = (requestData.id_visit || requestData.idVisit || "")
      .toString()
      .trim()
      .toLowerCase();
    if (!idVisit) {
      idVisit = generateUniqueIdVisit(sheet).toLowerCase();
    }
    rowData[0] = idVisit; // Set Kolom Index 1 (ID_VISIT)

    // 3. Set Tanggal & Waktu Presisi Native AppSheet (Wajib Ada Detik HH:mm:ss)
    const now = new Date();
    const tz = ss.getSpreadsheetTimeZone();

    // Format Waktu: Jika dari frontend cuma "17:20", tambahkan detik random (01-59)
    let finalTimeStr = rowData[2] ? rowData[2].toString().trim() : "";
    if (!finalTimeStr) {
      finalTimeStr = Utilities.formatDate(now, tz, "HH:mm:ss");
    } else if (/^\d{1,2}:\d{2}$/.test(finalTimeStr)) {
      // Format "HH:mm" atau "H:mm" -> Tambah detik random "HH:mm:ss"
      const randomSec = Math.floor(Math.random() * 58 + 1)
        .toString()
        .padStart(2, "0");
      finalTimeStr = `${finalTimeStr}:${randomSec}`;
    }
    rowData[2] = finalTimeStr; // Index 3: WAKTU (Kolom C) -> "17:20:43"

    // Format Tanggal: Presisi M/d/yyyy (misal "8/14/2026") TANPA pergeseran jam/zona waktu
    let dateFormattedMdy = "";
    let finalDateStrRaw = rowData[3] ? rowData[3].toString().trim() : "";
    if (finalDateStrRaw && finalDateStrRaw.includes("/")) {
      const parts = finalDateStrRaw.split("/");
      if (parts.length === 3) {
        let m = parseInt(parts[0]);
        let d = parseInt(parts[1]);
        let y = parseInt(parts[2]);
        if (m > 12) {
          const tmp = m;
          m = d;
          d = tmp;
        }
        dateFormattedMdy = `${m}/${d}/${y}`;
      }
    }
    if (!dateFormattedMdy) {
      dateFormattedMdy = Utilities.formatDate(now, tz, "M/d/yyyy");
    }
    rowData[3] = dateFormattedMdy;

    // 4. Dapatkan Folder Google Drive & Header Baris 1
    const regConfig =
      REGION_SPREADSHEETS[region] || REGION_SPREADSHEETS["LUAR_PULAU"];
    const photoPrefix = regConfig.photoPrefix || "";
    const driveFolder = getTargetFolder(region);
    const driveFolderAvailable = driveFolder !== null;
    let sheetHeaders = [];
    if (sheet.getLastRow() >= 1) {
      sheetHeaders = sheet
        .getRange(1, 1, 1, Math.max(244, sheet.getLastColumn()))
        .getValues()[0];
    }

    // 5. Pre-calculate Path Foto & Ganti Base64 dengan Text Path AppSheet
    const pendingPhotos = [];
    let photosSuccess = 0;
    let photosTotal = 0;
    const photoDetails = [];

    const timePartsStr = rowData[2].toString().split(":");
    let photoHour = parseInt(timePartsStr[0]) || 12;
    let photoMin = parseInt(timePartsStr[1]) || 0;
    let photoSec = parseInt(timePartsStr[2]) || 0;

    const photoDateObj = new Date();
    photoDateObj.setHours(photoHour, photoMin, photoSec, 0);
    let photoIndex = 0;

    for (let i = 0; i < rowData.length; i++) {
      const val = rowData[i];
      if (typeof val === "string" && val.startsWith("data:image/")) {
        photosTotal++;
        if (photoIndex > 0) {
          const secAdd =
            Math.random() < 0.45 ? 0 : Math.floor(Math.random() * 3 + 1);
          photoDateObj.setSeconds(photoDateObj.getSeconds() + secAdd);
        }
        photoIndex++;

        const currentPhotoTimeHHMMSS = Utilities.formatDate(
          photoDateObj,
          tz,
          "HHmmss",
        );
        const headerName =
          sheetHeaders[i] && sheetHeaders[i].toString().trim() !== ""
            ? sheetHeaders[i].toString().trim()
            : getAppSheetPhotoLabel(i);

        const appSheetFileName = `${photoPrefix}${idVisit}.${headerName}.${currentPhotoTimeHHMMSS}`;
        rowData[i] = `Kunjungan_Images/${appSheetFileName}.jpg`;

        if (driveFolderAvailable) {
          pendingPhotos.push({
            base64Data: val,
            fileName: appSheetFileName,
            headerName: headerName,
          });
        } else {
          photoDetails.push({
            name: headerName,
            uploaded: false,
            reason: "Drive tidak accessible",
          });
        }
      }
    }

    // Cek Anti-Duplikat pada Sheet
    let isVisitAlreadyInserted = false;
    if (sheet.getLastRow() > 1) {
      const existingVisitIds = sheet
        .getRange(2, 1, sheet.getLastRow() - 1, 1)
        .getValues();
      isVisitAlreadyInserted = existingVisitIds.some(
        (r) =>
          r[0] &&
          r[0].toString().trim().toLowerCase() === idVisit.toLowerCase(),
      );
    }

    if (isVisitAlreadyInserted) {
      const respDateStr = Utilities.formatDate(now, tz, "dd/MM/yyyy");
      const respTimeStr = Utilities.formatDate(now, tz, "HH:mm:ss");
      return createJsonResponse({
        status: "success",
        id_visit: idVisit,
        message:
          "Data kunjungan terkonfirmasi sudah aman tersimpan di Google Sheets!",
        timestamp: respDateStr + " " + respTimeStr,
      });
    }

    // 6. SIMPAN BARIS UTAMA KE SHEET KUNJUNGAN
    const maxSheetCols = Math.max(1, sheet.getMaxColumns());
    let cleanRowData = rowData;
    if (cleanRowData.length > maxSheetCols) {
      cleanRowData = cleanRowData.slice(0, maxSheetCols);
    }
    for (let k = 0; k < cleanRowData.length; k++) {
      if (cleanRowData[k] !== null && cleanRowData[k] !== undefined) {
        cleanRowData[k] = cleanRowData[k].toString();
      } else {
        cleanRowData[k] = "";
      }
    }

    sheet.appendRow(cleanRowData);
    try {
      const lastR = sheet.getLastRow();
      // Paksa SELURUH sel dalam baris menjadi Plain Text '@' (Rata Kiri, persis AppSheet Native)
      sheet.getRange(lastR, 1, 1, cleanRowData.length).setNumberFormat("@");
      // Khusus Tanggal (Kolom D / Index 4), set format M/d/yyyy
      sheet.getRange(lastR, 4).setNumberFormat("M/d/yyyy");
    } catch (eFmt) {}

    // 7. SIMPAN BARIS DETAIL KE SHEET DETAIL AUDIT
    let detailSheet = null;
    const rowDetailList =
      requestData.row_detail_list ||
      (requestData.row_detail_data ? [requestData.row_detail_data] : []);

    if (Array.isArray(rowDetailList) && rowDetailList.length > 0) {
      detailSheet =
        ss.getSheetByName(SHEET_DETAIL_AUDIT) ||
        ss.getSheetByName("Detail Audit") ||
        ss.getSheetByName("Detail_Kunjungan");
      if (!detailSheet) {
        detailSheet = ss.insertSheet(SHEET_DETAIL_AUDIT);
      }

      const validDetailRows = rowDetailList.filter(
        (r) => Array.isArray(r) && !r.every((v) => v === ""),
      );
      const itemCount = validDetailRows.length;

      const timeParts = rowData[2].toString().split(":");
      const mainHour = parseInt(timeParts[0]) || 12;
      const mainMin = parseInt(timeParts[1]) || 0;
      const mainSec = parseInt(timeParts[2]) || 0;

      const mainDateObj = new Date();
      mainDateObj.setHours(mainHour, mainMin, mainSec, 0);

      let currentScanTimeMs = mainDateObj.getTime() - itemCount * (35 * 1000);

      let detailSheetHeaders = [];
      if (detailSheet.getLastRow() >= 1) {
        detailSheetHeaders = detailSheet
          .getRange(1, 1, 1, Math.max(71, detailSheet.getLastColumn()))
          .getValues()[0];
      }

      const existingDetailIds = new Set();
      if (detailSheet && detailSheet.getLastRow() > 1) {
        const detailValues = detailSheet
          .getRange(2, 1, detailSheet.getLastRow() - 1, 1)
          .getValues();
        detailValues.forEach((r) => {
          if (r[0]) existingDetailIds.add(r[0].toString().trim().toLowerCase());
        });
      }

      validDetailRows.forEach((rowDetailData, idx) => {
        let detailId = "";
        do {
          detailId = Math.floor(Math.random() * 0xffffffff)
            .toString(16)
            .padStart(8, "0");
        } while (existingDetailIds.has(detailId));
        existingDetailIds.add(detailId);

        rowDetailData[0] = detailId;
        rowDetailData[1] = idVisit;
        rowDetailData[2] = rowData[1];

        const jitterSec = Math.floor(Math.random() * 30 + 25);
        currentScanTimeMs += jitterSec * 1000;

        if (currentScanTimeMs > mainDateObj.getTime()) {
          currentScanTimeMs =
            mainDateObj.getTime() - (itemCount - idx) * 15 * 1000;
        }

        const scanDateObj = new Date(currentScanTimeMs);
        const scanTimeStr = Utilities.formatDate(
          scanDateObj,
          ss.getSpreadsheetTimeZone(),
          "HH:mm:ss",
        );

        rowDetailData[3] = scanTimeStr;
        rowDetailData[4] = rowData[3];

        for (let j = 0; j < rowDetailData.length; j++) {
          const valD = rowDetailData[j];
          if (typeof valD === "string" && valD.startsWith("data:image/")) {
            const headerNameD =
              detailSheetHeaders[j] &&
              detailSheetHeaders[j].toString().trim() !== ""
                ? detailSheetHeaders[j].toString().trim()
                : `DETAIL_COL${j + 1}`;
            const fileNameD = `${photoPrefix}${detailId}.${headerNameD}.${scanTimeStr.replace(/:/g, "")}`;
            rowDetailData[j] = `Kunjungan_Images/${fileNameD}.jpg`;
          }
        }

        let cleanRow = Array.isArray(rowDetailData) ? rowDetailData : [];
        const maxDetailCols = Math.max(1, detailSheet.getMaxColumns());
        if (cleanRow.length > maxDetailCols) {
          cleanRow = cleanRow.slice(0, maxDetailCols);
        }

        for (let kd = 0; kd < cleanRow.length; kd++) {
          if (cleanRow[kd] !== null && cleanRow[kd] !== undefined) {
            cleanRow[kd] = cleanRow[kd].toString();
          } else {
            cleanRow[kd] = "";
          }
        }

        detailSheet.appendRow(cleanRow);
        try {
          const lastDR = detailSheet.getLastRow();
          detailSheet.getRange(lastDR, 1, 1, cleanRow.length).setNumberFormat("@");
          detailSheet.getRange(lastDR, 4).setNumberFormat("HH:mm:ss");
          detailSheet.getRange(lastDR, 5).setNumberFormat("M/d/yyyy");
        } catch (eFmtD) {}
      });
    }

    // 8. PROSES FOTO DRIVE SETELAH DATA SHEET SUKSES 100% TERULIS
    if (driveFolderAvailable && pendingPhotos.length > 0) {
      pendingPhotos.forEach((item) => {
        const saveRes = saveBase64ToDrive(
          item.base64Data,
          driveFolder,
          item.fileName,
        );
        if (saveRes.uploaded) photosSuccess++;
        photoDetails.push({
          name: item.headerName,
          uploaded: saveRes.uploaded,
        });
      });
    }

    // 9. Auto-Sort Seluruh Baris Berdasarkan Tanggal & Waktu
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow > 2 && lastCol > 0) {
      try {
        sheet.getRange(2, 1, lastRow - 1, lastCol).sort([
          { column: 4, ascending: true },
          { column: 3, ascending: true },
        ]);
      } catch (eSort) {}
    }

    if (detailSheet) {
      const dLastRow = detailSheet.getLastRow();
      const dLastCol = detailSheet.getLastColumn();
      if (dLastRow > 2 && dLastCol > 0) {
        try {
          detailSheet.getRange(2, 1, dLastRow - 1, dLastCol).sort([
            { column: 5, ascending: true },
            { column: 4, ascending: true },
          ]);
        } catch (eSortDetail) {}
      }
    }

    const respDateStr = Utilities.formatDate(now, tz, "dd/MM/yyyy");
    const respTimeStr = Utilities.formatDate(now, tz, "HH:mm:ss");

    return createJsonResponse({
      status: "success",
      id_visit: idVisit,
      message: `Data kunjungan berhasil disimpan di Google Sheets! (${photosSuccess}/${photosTotal} foto fisik ter-upload ke Drive, seluruh path terisi rapi)`,
      photo_stats: {
        total: photosTotal,
        uploaded: photosSuccess,
        details: photoDetails,
      },
      timestamp: respDateStr + " " + respTimeStr,
    });
  } catch (err) {
    return createJsonResponse({
      status: "error",
      message: "Server Error: " + err.toString(),
    });
  } finally {
    lock.releaseLock();
  }
}

/**
 * Membuat ID_VISIT 8-karakter hex unik dengan loop collision check di kolom A
 */
function generateUniqueIdVisit(sheet) {
  let unique = false;
  let id = "";
  const existingIds = new Set();

  if (sheet && sheet.getLastRow() > 1) {
    const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    values.forEach((r) => {
      if (r[0]) existingIds.add(r[0].toString().trim().toLowerCase());
    });
  }

  while (!unique) {
    id = Math.floor(Math.random() * 0xffffffff)
      .toString(16)
      .padStart(8, "0");
    if (!existingIds.has(id)) {
      unique = true;
    }
  }
  return id;
}

/**
 * Menyimpan String Gambar Base64 ke Google Drive
 */
function saveBase64ToDrive(base64Data, folder, fileName) {
  const parts = base64Data.split(",");
  const contentType = parts[0].split(";")[0].replace("data:", "");
  const decodedData = Utilities.base64Decode(parts[1]);

  let extension = "png";
  if (contentType.includes("jpeg") || contentType.includes("jpg")) {
    extension = "jpg";
  } else if (contentType.includes("webp")) {
    extension = "webp";
  }

  const fullFileName = `${fileName}.${extension}`;
  const blob = Utilities.newBlob(decodedData, contentType, fullFileName);
  let isUploaded = false;

  // Direct Upload ke Drive (Super Kencang ~0.3s per foto)
  try {
    folder.createFile(blob);
    isUploaded = true;
  } catch (e) {
    isUploaded = false;
  }

  return {
    path: `Kunjungan_Images/${fullFileName}`,
    uploaded: isUploaded,
  };
}

/**
 * Helper Output Response JSON (Mendukung JSONP Callback jika ada)
 */
function createJsonResponse(obj, callback = "") {
  const jsonString = JSON.stringify(obj);
  if (callback && String(callback).trim() !== "") {
    return ContentService.createTextOutput(
      `${callback}(${jsonString})`,
    ).setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(jsonString).setMimeType(
    ContentService.MimeType.JSON,
  );
}
