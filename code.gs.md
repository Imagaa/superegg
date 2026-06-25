// GERBANG UTAMA API (Menerima Request dari Vercel)
function doPost(e) {
  try {
    const requestData = JSON.parse(e.postData.contents);
    const action = requestData.action;
    const payload = requestData.data;
    
    let result;
    if (action === 'getSalesList') result = getSalesListInternal();
    else if (action === 'getStoresBySales') result = getStoresBySalesInternal(payload.salesName);
    else if (action === 'submitLaporan') result = submitLaporanInternal(payload);
    else if (action === 'submitTitikBaru') result = submitTitikBaruInternal(payload);
    else if (action === 'submitEndCustomer') result = submitEndCustomerInternal(payload);
    else if (action === 'getMapData') result = getMapDataInternal();
    else if (action === 'getRoutingData') result = getRoutingDataInternal(payload);
    else return respondJSON({ success: false, message: "Action API tidak dikenal." });
    
    return respondJSON(result);
  } catch (error) {
    return respondJSON({ success: false, message: "Gagal memproses API: " + error.toString() });
  }
}

function respondJSON(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheetByGid(ss, gid) {
  const sheets = ss.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    if (sheets[i].getSheetId() == gid) return sheets[i];
  }
  return null;
}

// ==========================================
// FUNGSI INJEKSI DATA PRESISI
// ==========================================

function appendPengirimanAman(sheet, rowData, mode = "Laporan") {
  const colData = sheet.getRange("B:B").getValues(); 
  let lastRow = 0;
  for (let i = colData.length - 1; i >= 0; i--) {
    if (colData[i][0] !== "") {
      lastRow = i + 1;
      break;
    }
  }
  const targetRow = lastRow + 1;
  
  // Injeksi Kolom A-K (Indeks 0-10)
  sheet.getRange(targetRow, 1, 1, 11).setValues([rowData.slice(0, 11)]);
  
  // Injeksi Kolom N-Q (Indeks 13-16) melewati L & M -> panjangnya 4 kolom
  sheet.getRange(targetRow, 14, 1, 4).setValues([rowData.slice(13, 17)]);
  
  // PERBAIKAN FINAL: Menggunakan .setValue() agar titik koma (;) diterima dengan sempurna
  sheet.getRange(targetRow, 4).setValue('=IF(C' + targetRow + '=""; ""; LEFT(C' + targetRow + '; 6))');
  sheet.getRange(targetRow, 7).setValue('=IFERROR(LEFT(E' + targetRow + '; SEARCH("-"; E' + targetRow + ')-1); "")');
  
  // Injeksi Rumus Kolom L dan M berdasarkan Mode (juga menggunakan .setValue)
  if (mode === "Laporan") {
    sheet.getRange(targetRow, 12).setValue("=(K" + targetRow + "-H" + targetRow + ")*3500");
    sheet.getRange(targetRow, 13).setValue("=I" + targetRow + "-2720*(K" + targetRow + "-H" + targetRow + ")");
  } else if (mode === "Baru") {
    sheet.getRange(targetRow, 12).setValue("=J" + targetRow + "*3500");
    sheet.getRange(targetRow, 13).setValue("=I" + targetRow + "-(2720*J" + targetRow + ")");
  }
}

// Modifikasi agar bisa menerima parameter warna latar (bgColor)
function appendMasterAman(sheet, rowData, bgColor = null) {
  const colData = sheet.getRange("A:A").getValues(); 
  let lastRow = 0;
  for (let i = colData.length - 1; i >= 0; i--) {
    if (colData[i][0] !== "") {
      lastRow = i + 1;
      break;
    }
  }
  const targetRow = lastRow + 1;
  sheet.getRange(targetRow, 1, 1, 8).setValues([rowData]); 
  
  // Jika warna dikirim, blok dari Kolom A sampai H (8 kolom)
  if (bgColor) {
    sheet.getRange(targetRow, 1, 1, 8).setBackground(bgColor);
  }
}

// ==========================================
// LOGIKA INTERNAL API
// ==========================================

function getSalesListInternal() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getSheetByGid(ss, 1772680131); 
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  let uniqueSales = [];
  for (let i = 1; i < data.length; i++) {
    let sales = String(data[i][7] || "").trim(); 
    if (sales !== "" && sales.toLowerCase() !== "pic sales") uniqueSales.push(sales);
  }
  return [...new Set(uniqueSales)].sort();
}

function getStoresBySalesInternal(salesName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pengirimanSheet = getSheetByGid(ss, 1790905023); 
  if (!pengirimanSheet) return ["Error: Gagal terhubung ke Tab Pengiriman"];
  
  const targetSales = String(salesName).trim().toLowerCase();
  const pengirimanData = pengirimanSheet.getDataRange().getValues();
  const validStores = new Map();
  
  for (let i = 1; i < pengirimanData.length; i++) {
    let tokoPengiriman = String(pengirimanData[i][2] || "").trim(); 
    let picPengiriman = String(pengirimanData[i][5] || "").trim().toLowerCase();  
    let statusBayar = String(pengirimanData[i][13] || "").trim().toLowerCase(); 
    let stokLalu = pengirimanData[i][10]; 
    
    if (statusBayar === 'belum' && picPengiriman === targetSales && tokoPengiriman !== "") {
      validStores.set(tokoPengiriman, { nama: tokoPengiriman, stokLalu: stokLalu }); 
    }
  }
  if (validStores.size === 0) return ["Error: Tidak ada tagihan 'Belum' untuk Sales ini"];
  return Array.from(validStores.values()).sort((a, b) => a.nama.localeCompare(b.nama));
}

function getMapDataInternal() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getSheetByGid(ss, 1772680131); 
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues();
  let mapData = [];
  
  for (let i = 1; i < data.length; i++) {
    let namaWarung = String(data[i][1] || "").trim(); 
    let lat = parseFloat(data[i][5]);                 
    let lng = parseFloat(data[i][6]);                 
    let pic = String(data[i][7] || "").trim();        

    if (namaWarung !== "" && !isNaN(lat) && !isNaN(lng)) {
      mapData.push({ nama: namaWarung, lat: lat, lng: lng, pic: pic });
    }
  }
  return mapData;
}

function submitLaporanInternal(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pengirimanSheet = getSheetByGid(ss, 1790905023); 
  if (!pengirimanSheet) return { success: false, message: "Tab Pengiriman tidak ditemukan." };

  const folder = DriveApp.getFolderById("1mTW30DulUaMJOXgUzS-RUFJ-vIut71du");
  const contentType = data.fotoBase64.substring(5, data.fotoBase64.indexOf(';'));
  const bytes = Utilities.base64Decode(data.fotoBase64.split(',')[1]);
  const safeTimestamp = data.timestamp.replace(/[\/\:]/g, '-'); 
  const fileName = `LAPORAN_${data.sales}_${safeTimestamp}_${data.toko}.jpg`;
  
  const file = folder.createFile(Utilities.newBlob(bytes, contentType, fileName));
  const fileUrl = file.getUrl(); 

  const pengirimanData = pengirimanSheet.getDataRange().getValues();
  let targetRowIndex = -1;
  let lastInvoiceNum = 0; 

  for (let i = 1; i < pengirimanData.length; i++) {
    let tokoPengiriman = String(pengirimanData[i][2] || "").trim().toLowerCase();
    let picPengiriman = String(pengirimanData[i][5] || "").trim().toLowerCase();
    let statusBayar = String(pengirimanData[i][13] || "").trim().toLowerCase();
    
    if (statusBayar === 'belum' && picPengiriman === data.sales.toLowerCase() && tokoPengiriman === data.toko.toLowerCase()) {
      targetRowIndex = i + 1; 
    }
    let invStr = String(pengirimanData[i][1] || "").trim();
    if (invStr.startsWith("INV")) {
      let numPart = parseInt(invStr.replace("INV", ""), 10);
      if (!isNaN(numPart) && numPart > lastInvoiceNum) lastInvoiceNum = numPart;
    }
  }

  if (targetRowIndex === -1) return { success: false, message: "Data toko tidak ditemukan." };

  let today = new Date();
  // UPDATE BARIS LAMA
  pengirimanSheet.getRange(targetRowIndex, 1).setValue(today); // Timpa tanggal menjadi hari ini
  pengirimanSheet.getRange(targetRowIndex, 14).setValue('Perlu Direview'); 
  pengirimanSheet.getRange(targetRowIndex, 15).setValue(data.keterangan);  
  pengirimanSheet.getRange(targetRowIndex, 16).setValue(fileUrl); 
  pengirimanSheet.getRange(targetRowIndex, 17).setValue('Sudah dikunjungi'); // Eksekusi Kolom Q
  pengirimanSheet.getRange(targetRowIndex, 8).setValue(data.sisaTelur); 
  pengirimanSheet.getRange(targetRowIndex, 9).setValue(data.setoranClean); 
  
  if (data.status === 'Restock') {
    pengirimanSheet.getRange(targetRowIndex, 10).setValue(data.jumlahRestock); 
    
    let nextWeek = new Date(today.getTime() + 7 * 24 * 60 * 60 * 1000); // Kalkulasi H+7
    let newInvoice = "INV" + String(lastInvoiceNum + 1).padStart(3, '0');
    let barangStatis = 'TAO001 - Paket Telur Asin Omega 10 Butir (Agen)';
    
    let newRow = [today, newInvoice, data.toko, "", barangStatis, data.sales, "", "", "", "", data.jumlahRestock, "", "", "Belum", "", "", nextWeek];
    appendPengirimanAman(pengirimanSheet, newRow, "Laporan");
  } else {
    pengirimanSheet.getRange(targetRowIndex, 10).setValue("Stop"); 
  }
  return { success: true, message: "Laporan berhasil masuk ke spreadsheet!" };
}

function submitTitikBaruInternal(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = getSheetByGid(ss, 1772680131);
  const pengirimanSheet = getSheetByGid(ss, 1790905023); 
  
  if (!masterSheet || !pengirimanSheet) return { success: false, message: "Tab tidak ditemukan." };

  const folder = DriveApp.getFolderById("1mTW30DulUaMJOXgUzS-RUFJ-vIut71du");
  const contentType = data.fotoBase64.substring(5, data.fotoBase64.indexOf(';'));
  const bytes = Utilities.base64Decode(data.fotoBase64.split(',')[1]);
  const safeTimestamp = data.timestamp.replace(/[\/\:]/g, '-'); 
  const fileName = `BUKA-TITIK_${data.pic}_${safeTimestamp}_${data.namaWarung}.jpg`;
  
  const file = folder.createFile(Utilities.newBlob(bytes, contentType, fileName));
  const fileUrl = file.getUrl(); 

  const masterData = masterSheet.getDataRange().getValues();
  let lastWrgNum = 0;
  for (let i = 1; i < masterData.length; i++) {
    let idStr = String(masterData[i][0] || "").trim();
    if (idStr.startsWith("WRG")) {
      let numPart = parseInt(idStr.replace("WRG", ""), 10);
      if (!isNaN(numPart) && numPart > lastWrgNum) lastWrgNum = numPart;
    }
  }
  let newWrgId = "WRG" + String(lastWrgNum + 1).padStart(3, '0');
  
  // Ambil data akurat untuk baris rumus
  const colDataM = masterSheet.getRange("A:A").getValues();
  let lastRowM = 0;
  for (let i = colDataM.length - 1; i >= 0; i--) { if (colDataM[i][0] !== "") { lastRowM = i + 1; break; } }
  let newRowMasterNum = lastRowM + 1;
  
  let displayFormula = `=A${newRowMasterNum}&" - "&B${newRowMasterNum}`;
  let displayTeksResult = `${newWrgId} - ${data.namaWarung}`;
  let masterRow = [newWrgId, data.namaWarung, data.pemilik, data.hp, displayFormula, data.lat, data.long, data.pic];
  appendMasterAman(masterSheet, masterRow);

  const pengirimanData = pengirimanSheet.getDataRange().getValues();
  let lastInvoiceNum = 0;
  for (let i = 1; i < pengirimanData.length; i++) {
    let invStr = String(pengirimanData[i][1] || "").trim();
    if (invStr.startsWith("INV")) {
      let numPart = parseInt(invStr.replace("INV", ""), 10);
      if (!isNaN(numPart) && numPart > lastInvoiceNum) lastInvoiceNum = numPart;
    }
  }
  
  let today = new Date();
  let nextWeek = new Date(today.getTime() + 7 * 24 * 60 * 60 * 1000); // Kalkulasi H+7
  let newInvoiceId = "INV" + String(lastInvoiceNum + 1).padStart(3, '0');
  let keteranganTeks = data.keterangan ? "TITIK BARU - " + data.keterangan : "TITIK BARU";
  let barangStatis = 'TAO001 - Paket Telur Asin Omega 10 Butir (Agen)';
  
  let pengirimanRow = [today, newInvoiceId, displayTeksResult, "", barangStatis, data.pic, "", 0, 0, data.jumlahAwal, 0, "", "", "Belum", keteranganTeks, fileUrl, nextWeek];
  appendPengirimanAman(pengirimanSheet, pengirimanRow, "Baru");

  return { success: true, message: "Titik Baru berhasil diregistrasi!" };
}

function submitEndCustomerInternal(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterSheet = getSheetByGid(ss, 1772680131);
  const pengirimanSheet = getSheetByGid(ss, 1790905023); 
  
  if (!masterSheet || !pengirimanSheet) return { success: false, message: "Tab tidak ditemukan." };

  const folder = DriveApp.getFolderById("1mTW30DulUaMJOXgUzS-RUFJ-vIut71du");
  const contentType = data.fotoBase64.substring(5, data.fotoBase64.indexOf(';'));
  const bytes = Utilities.base64Decode(data.fotoBase64.split(',')[1]);
  const safeTimestamp = data.timestamp.replace(/[\/\:]/g, '-'); 
  const fileName = `EndCustomer_${data.pic}_${safeTimestamp}_${data.namaUser}.jpg`;
  
  const file = folder.createFile(Utilities.newBlob(bytes, contentType, fileName));
  const fileUrl = file.getUrl(); 

  const masterData = masterSheet.getDataRange().getValues();
  let lastCstNum = 0;
  for (let i = 1; i < masterData.length; i++) {
    let idStr = String(masterData[i][0] || "").trim();
    if (idStr.startsWith("CST")) {
      let numPart = parseInt(idStr.replace("CST", ""), 10);
      if (!isNaN(numPart) && numPart > lastCstNum) lastCstNum = numPart;
    }
  }
  let newCstId = "CST" + String(lastCstNum + 1).padStart(3, '0');
  
  // Ambil data akurat untuk baris rumus
  const colDataM = masterSheet.getRange("A:A").getValues();
  let lastRowM = 0;
  for (let i = colDataM.length - 1; i >= 0; i--) { if (colDataM[i][0] !== "") { lastRowM = i + 1; break; } }
  let newRowMasterNum = lastRowM + 1;
  
  let formatNamaWarung = "Nama User " + data.namaUser; 
  let displayFormula = `=A${newRowMasterNum}&" - "&B${newRowMasterNum}`;
  let displayTeksResult = `${newCstId} - ${formatNamaWarung}`;
  
  let masterRow = [newCstId, formatNamaWarung, data.namaUser, data.noWa, displayFormula, "-", "-", data.pic];
  // Eksekusi fungsi dengan parameter warna kuning standar ("#FFFF00")
  appendMasterAman(masterSheet, masterRow, "#FFFF00");

  const pengirimanData = pengirimanSheet.getDataRange().getValues();
  let lastInvoiceNum = 0;
  for (let i = 1; i < pengirimanData.length; i++) {
    let invStr = String(pengirimanData[i][1] || "").trim();
    if (invStr.startsWith("INV")) {
      let numPart = parseInt(invStr.replace("INV", ""), 10);
      if (!isNaN(numPart) && numPart > lastInvoiceNum) lastInvoiceNum = numPart;
    }
  }
  
  let newInvoiceId = "INV" + String(lastInvoiceNum + 1).padStart(3, '0');
  let keteranganTeks = data.keterangan ? "END USER - " + data.keterangan : "END USER";
  let barangStatis = 'TAO001 - Paket Telur Asin Omega 10 Butir (Agen)';
  
  let pengirimanRow = [new Date(), newInvoiceId, displayTeksResult, "", barangStatis, data.pic, "", 0, data.totalSetoran, data.jumlahBeli, 0, "", "", "Perlu Direview", keteranganTeks, fileUrl, "-"];
  appendPengirimanAman(pengirimanSheet, pengirimanRow, "Baru");

  return { success: true, message: "Transaksi Pembeli End User berhasil direkam!" };
}

// ==========================================
// ALAT SINKRONISASI DATA LAMA 
// ==========================================

function sinkronisasiFotoLama() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = getSheetByGid(ss, 1790905023); 
  const folder = DriveApp.getFolderById("1mTW30DulUaMJOXgUzS-RUFJ-vIut71du");
  const files = folder.getFiles();
  const data = sheet.getDataRange().getValues();

  let countBerhasil = 0;

  while (files.hasNext()) {
    const file = files.next();
    const fileName = file.getName();
    const fileUrl = file.getUrl();

    let parts = fileName.replace('.jpg', '').replace('.jpeg', '').replace('.png', '').split('_');
    
    if (parts.length >= 4) {
      let salesName = parts[1].toLowerCase();
      let timeString = parts[2]; 
      let tokoName = parts.slice(3).join('_').toLowerCase(); 

      let fileD = -1, fileM = -1, fileY = -1;
      let fileTimeMatches = timeString.match(/(\d{1,2})-(\d{1,2})-(\d{4})/);
      
      if (fileTimeMatches) {
         fileD = parseInt(fileTimeMatches[1], 10);
         fileM = parseInt(fileTimeMatches[2], 10);
         fileY = parseInt(fileTimeMatches[3], 10);
      }

      for (let i = 1; i < data.length; i++) {
        let rowDate = new Date(data[i][0]); 
        let rowToko = String(data[i][2]).toLowerCase();
        let rowSales = String(data[i][5]).toLowerCase();
        let rowLink = String(data[i][15] || ""); 

        if (isNaN(rowDate.getTime())) continue; 
        
        let sheetD = rowDate.getDate();
        let sheetM = rowDate.getMonth() + 1; 
        let sheetY = rowDate.getFullYear();

        let isDateMatch = (fileD === sheetD && fileM === sheetM && fileY === sheetY);
        
        if (isDateMatch && rowToko.includes(tokoName) && rowSales === salesName && rowLink === "") {
          sheet.getRange(i + 1, 16).setValue(fileUrl); 
          data[i][15] = fileUrl; 
          countBerhasil++;
          break; 
        }
      }
    }
  }
  Logger.log("Sinkronisasi selesai! Berhasil menyambungkan " + countBerhasil + " foto lama.");
}
function getRoutingDataInternal(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const pengirimanSheet = getSheetByGid(ss, 1790905023);
  const masterSheet = getSheetByGid(ss, 1772680131);
  
  if (!pengirimanSheet || !masterSheet) return { success: false, message: "Tab tidak ditemukan" };
  
  const targetSales = String(payload.salesName).trim().toLowerCase();
  const startDate = payload.startDate ? new Date(payload.startDate) : null;
  const endDate = payload.endDate ? new Date(payload.endDate) : null;
  if (endDate) endDate.setHours(23, 59, 59, 999); // Kunci hingga jam 23:59 batas akhir
  
  const pengirimanData = pengirimanSheet.getDataRange().getValues();
  const masterData = masterSheet.getDataRange().getValues();
  
  // MAPPING DATA KOORDINAT MASTER
  const masterCoords = new Map();
  for (let i = 1; i < masterData.length; i++) {
    let displayToko = String(masterData[i][4] || "").trim(); // Kolom E (Display)
    let lat = parseFloat(masterData[i][5]); // Kolom F
    let lng = parseFloat(masterData[i][6]); // Kolom G
    if (displayToko !== "") {
      masterCoords.set(displayToko, { lat: lat, lng: lng });
    }
  }
  
  // FILTERING & JOIN PENGIRIMAN
  const validStores = new Map();
  for (let i = 1; i < pengirimanData.length; i++) {
    let tglTransaksi = new Date(pengirimanData[i][0]); // Kolom A
    let tokoPengiriman = String(pengirimanData[i][2] || "").trim(); // Kolom C
    let picPengiriman = String(pengirimanData[i][5] || "").trim().toLowerCase();  
    let statusBayar = String(pengirimanData[i][13] || "").trim().toLowerCase(); 
    let stokLalu = pengirimanData[i][10]; // Kolom K
    
    let isDateValid = true;
    if (startDate && endDate && !isNaN(tglTransaksi.getTime())) {
      if (tglTransaksi < startDate || tglTransaksi > endDate) isDateValid = false;
    }
    
    if (statusBayar === 'belum' && picPengiriman === targetSales && tokoPengiriman !== "" && isDateValid) {
      let coords = masterCoords.get(tokoPengiriman);
      if (coords && !isNaN(coords.lat) && !isNaN(coords.lng)) {
        validStores.set(tokoPengiriman, { 
          nama: tokoPengiriman, 
          lat: coords.lat, 
          lng: coords.lng, 
          stokLalu: stokLalu 
        });
      }
    }
  }
  
  let resultArr = Array.from(validStores.values()).sort((a, b) => a.nama.localeCompare(b.nama));
  if (resultArr.length === 0) return { success: false, message: "Tidak ada data tagihan/rute untuk rentang kriteria ini." };
  
  return { success: true, data: resultArr };
}