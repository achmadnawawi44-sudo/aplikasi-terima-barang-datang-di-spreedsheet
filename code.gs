// ID Spreadsheet untuk menyimpan data
var SPREADSHEET_ID = '13RIUQQjKVV2Mh8-UncGBh0Lh7xj0t-uc0HHu3G_msqU'; // Isi dengan ID Spreadsheet Anda
var SHEET_NAME = 'PenerimaanBarang';
var FOLDER_ID = ''; // Isi dengan ID Folder Drive Anda

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Sistem Penerimaan Barang')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Fungsi untuk memproses surat jalan (format CSV) - DIPERBAIKI
function prosesSuratJalan(fileContent, fileName) {
  console.log('Memulai proses surat jalan...');
  console.log('Nama file:', fileName);
  try {
    // Validasi input
    if (!fileContent || !fileName) {
      throw new Error('File content atau nama file tidak valid');
    }
    
    // Simpan file ke Drive
    console.log('Menyimpan file ke Drive...');
    const folder = getFolder();
    const blob = createBlob(fileContent, fileName);
    console.log('Blob berhasil dibuat');
    const file = folder.createFile(blob);
    console.log('File berhasil dibuat di Drive:', file.getId());
    
    // Proses file CSV - DIPERBAIKI
    console.log('Memproses konten CSV...');
    const base64Data = fileContent.indexOf(',') > -1 ? fileContent.split(',')[1] : fileContent;
    const csvContent = Utilities.base64Decode(base64Data);
    console.log('Base64 berhasil didecode');
    const csvBlob = Utilities.newBlob(csvContent, 'text/csv', fileName);
    const csvString = csvBlob.getDataAsString();
    console.log('CSV string:', csvString.substring(0, 100) + '...'); // Log sebagian konten
    
    // Parse data CSV
    console.log('Parsing CSV...');
    const dataSuratJalan = parseCsvSuratJalan(csvString);
    console.log('Data surat jalan:', JSON.stringify(dataSuratJalan.slice(0, 2)));
    
    // Ekstrak nomor surat jalan dari nama file
    const nomorSuratJalan = extractNomorSuratJalan(fileName);
    
    return {
      success: true,
      message: `Surat jalan ${nomorSuratJalan} berhasil diupload`,
      data: dataSuratJalan,
      nomorSuratJalan: nomorSuratJalan,
      fileUrl: file.getUrl()
    };
  } catch (e) {
    console.error('Error dalam prosesSuratJalan:', e);
    return {
      success: false,
      message: 'Gagal memproses surat jalan: ' + e.message,
      data: null
    };
  }
}

// Fungsi untuk memproses data barang (format CSV/Excel) - TETAP SAMA
function prosesDataBarang(fileContent, fileName) {
  try {
    // Simpan file ke Drive
    const folder = getFolder();
    const blob = createBlob(fileContent, fileName);
    const file = folder.createFile(blob);
    
    // Proses file berdasarkan ekstensi
    let dataBarang = [];
    
    if (fileName.endsWith('.csv')) {
      // Proses CSV
      const csvContent = Utilities.base64Decode(fileContent.split(',')[1]);
      const csvBlob = Utilities.newBlob(csvContent, 'text/csv', fileName);
      const csvString = csvBlob.getDataAsString();
      dataBarang = parseCsvDataBarang(csvString);
    } else if (fileName.endsWith('.xlsx') || fileName.endsWith('.xls')) {
      // Proses Excel (menggunakan Drive API)
      dataBarang = parseExcelData(file.getId());
    }
    
    return {
      success: true,
      message: 'Data barang berhasil diproses',
      data: dataBarang,
      fileUrl: file.getUrl()
    };
  } catch (e) {
    return {
      success: false,
      message: 'Gagal memproses data barang: ' + e.message,
      data: null
    };
  }
}

// Fungsi untuk membandingkan barang dengan surat jalan - TETAP SAMA
function bandingkanBarang(barangDiterima, suratJalanData) {
  try {
    // Normalisasi data (case insensitive dan trim whitespace)
    const normalizedBarangDiterima = barangDiterima.map(item => ({
      ...item,
      barcode: item.barcode.trim().toLowerCase()
    }));
    
    const normalizedSuratJalan = suratJalanData.map(item => ({
      ...item,
      barcode: item.barcode.trim().toLowerCase()
    }));
    
    // Buat salinan array untuk hasil
    const hasil = JSON.parse(JSON.stringify(normalizedBarangDiterima));
    
    // Bandingkan setiap barang yang diterima
    hasil.forEach(item => {
      const itemSuratJalan = normalizedSuratJalan.find(sj => sj.barcode === item.barcode);
      
      if (itemSuratJalan) {
        item.qtySuratJalan = itemSuratJalan.qty;
        
        if (item.qtyDiterima > itemSuratJalan.qty) {
          item.status = 'Barang Lebih';
        } else if (item.qtyDiterima < itemSuratJalan.qty) {
          item.status = 'Barang Kurang';
        } else {
          item.status = 'Sesuai';
        }
      } else {
        item.qtySuratJalan = 0;
        item.status = 'Tidak ada di surat jalan';
      }
    });
    
    // Tambahkan barang dari surat jalan yang belum diterima
    normalizedSuratJalan.forEach(sjItem => {
      const sudahDiterima = hasil.some(item => item.barcode === sjItem.barcode);
      
      if (!sudahDiterima) {
        hasil.push({
          barcode: sjItem.barcode,
          nama: sjItem.nama,
          qtyDiterima: 0,
          qtySuratJalan: sjItem.qty,
          status: 'Belum diterima'
        });
      }
    });
    
    return {
      success: true,
      message: 'Perbandingan berhasil',
      barangDiterima: hasil
    };
  } catch (e) {
    return {
      success: false,
      message: 'Gagal membandingkan: ' + e.message,
      barangDiterima: barangDiterima
    };
  }
}

// Fungsi untuk menyimpan data ke spreadsheet - TETAP SAMA
function simpanKeSpreadsheet(nomorSuratJalan, dataBarang) {
  try {
    // Buka atau buat spreadsheet
    let spreadsheet;
    try {
      spreadsheet = SpreadsheetApp.openById(SPREADSHEET_ID);
    } catch (e) {
      spreadsheet = SpreadsheetApp.create('Data Penerimaan Barang');
      SPREADSHEET_ID = spreadsheet.getId();
    }
    
    // Dapatkan atau buat sheet
    let sheet = spreadsheet.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = spreadsheet.insertSheet(SHEET_NAME);
      // Buat header
      sheet.getRange(1, 1, 1, 7).setValues([[
        'Tanggal', 'Nomor Surat Jalan', 'Barcode', 'Nama Barang', 
        'Qty Diterima', 'Qty Surat Jalan', 'Selisih', 'Status'
      ]]);
    }
    
    // Siapkan data untuk disimpan
    const newData = [];
    const timestamp = new Date();
    
    dataBarang.forEach(item => {
      newData.push([
        timestamp,
        nomorSuratJalan,
        item.barcode,
        item.nama,
        item.qtyDiterima,
        item.qtySuratJalan,
        item.selisih,
        item.status
      ]);
    });
    
    // Simpan data
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1, 1, newData.length, 8).setValues(newData);
    
    // Format tanggal
    sheet.getRange(lastRow + 1, 1, newData.length, 1)
      .setNumberFormat('dd/MM/yyyy HH:mm:ss');
    
    return {
      success: true,
      message: 'Data berhasil disimpan',
      spreadsheetUrl: spreadsheet.getUrl()
    };
  } catch (e) {
    return {
      success: false,
      message: 'Gagal menyimpan data: ' + e.message
    };
  }
}

// ========== FUNGSI BANTUAN ========== //

// Fungsi untuk parsing CSV surat jalan - DIPERBAIKI
function parseCsvSuratJalan(csvString) {
  console.log('Memulai parsing CSV surat jalan...');
  const lines = csvString.split('\n').filter(line => line.trim() !== '');
  
  if (lines.length < 2) {
    throw new Error('File CSV kosong atau hanya berisi header');
  }
  
  const headers = lines[0].split(',').map(h => h.trim().toLowerCase());
  console.log('Header ditemukan:', headers);
  
  // Validasi kolom wajib
  const requiredColumns = ['barcode', 'nama', 'qty'];
  const missingColumns = requiredColumns.filter(col => !headers.includes(col));
  
  if (missingColumns.length > 0) {
    throw new Error(`Format CSV tidak valid. Kolom wajib tidak ditemukan: ${missingColumns.join(', ')}. Kolom yang ada: ${headers.join(', ')}`);
  }
  
  const barcodeIndex = headers.indexOf('barcode');
  const namaIndex = headers.indexOf('nama');
  const qtyIndex = headers.indexOf('qty');
  
  return lines.slice(1).map(line => {
    const values = line.split(',');
    return {
      barcode: values[barcodeIndex].trim(),
      nama: values[namaIndex].trim(),
      qty: parseInt(values[qtyIndex].trim()) || 0
    };
  }).filter(item => item.barcode && item.nama); // Hapus baris kosong
}

// Fungsi untuk parsing CSV data barang - TETAP SAMA
function parseCsvDataBarang(csvString) {
  const lines = csvString.split('\n').filter(line => line.trim() !== '');
  const headers = lines[0].split(',').map(h => h.trim().toLowerCase());
  
  // Validasi kolom wajib
  const requiredColumns = ['barcode', 'nama'];
  const missingColumns = requiredColumns.filter(col => !headers.includes(col));
  
  if (missingColumns.length > 0) {
    throw new Error(`Format CSV tidak valid. Kolom wajib: ${requiredColumns.join(', ')}`);
  }
  
  const barcodeIndex = headers.indexOf('barcode');
  const namaIndex = headers.indexOf('nama');
  
  return lines.slice(1).map(line => {
    const values = line.split(',');
    return {
      barcode: values[barcodeIndex].trim(),
      nama: values[namaIndex].trim()
    };
  }).filter(item => item.barcode && item.nama); // Hapus baris kosong
}

// Fungsi untuk parsing Excel data barang (menggunakan Drive API) - TETAP SAMA
function parseExcelData(fileId) {
  // Konversi file Excel ke Google Spreadsheet
  const file = Drive.Files.get(fileId);
  const spreadsheet = Drive.Files.copy(
    {title: 'Temp Import Data - ' + new Date().getTime(), mimeType: MimeType.GOOGLE_SHEETS},
    fileId
  );
  
  // Baca data dari sheet
  const sheet = SpreadsheetApp.openById(spreadsheet.id).getSheets()[0];
  const data = sheet.getDataRange().getValues();
  
  // Hapus file temporary
  Drive.Files.remove(spreadsheet.id);
  
  // Asumsikan baris pertama adalah header
  const headers = data[0].map(h => h.toString().trim().toLowerCase());
  
  // Validasi kolom wajib
  const requiredColumns = ['barcode', 'nama'];
  const missingColumns = requiredColumns.filter(col => !headers.includes(col));
  
  if (missingColumns.length > 0) {
    throw new Error(`Format Excel tidak valid. Kolom wajib: ${requiredColumns.join(', ')}`);
  }
  
  const barcodeIndex = headers.indexOf('barcode');
  const namaIndex = headers.indexOf('nama');
  
  return data.slice(1).map(row => {
    return {
      barcode: row[barcodeIndex].toString().trim(),
      nama: row[namaIndex].toString().trim()
    };
  }).filter(item => item.barcode && item.nama); // Hapus baris kosong
}

// Fungsi untuk membuat blob dari file content - DIPERBAIKI
function createBlob(fileContent, fileName) {
  try {
    const base64Data = fileContent.indexOf(',') > -1 ? fileContent.split(',')[1] : fileContent;
    const decoded = Utilities.base64Decode(base64Data);
    const mimeType = getMimeType(fileName);
    return Utilities.newBlob(decoded, mimeType, fileName);
  } catch (e) {
    throw new Error('Format file tidak valid: ' + e.message);
  }
}

// Fungsi untuk mendapatkan folder - TETAP SAMA
function getFolder() {
  try {
    if (FOLDER_ID) {
      return DriveApp.getFolderById(FOLDER_ID);
    }
    return DriveApp.getRootFolder();
  } catch (e) {
    throw new Error('Folder tidak ditemukan. Pastikan FOLDER_ID benar.');
  }
}

// Fungsi untuk menentukan mime type - TETAP SAMA
function getMimeType(fileName) {
  const extension = fileName.split('.').pop().toLowerCase();
  switch (extension) {
    case 'csv': return 'text/csv';
    case 'xlsx': return 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    case 'xls': return 'application/vnd.ms-excel';
    case 'pdf': return 'application/pdf';
    case 'jpg': case 'jpeg': return 'image/jpeg';
    case 'png': return 'image/png';
    default: return 'application/octet-stream';
  }
}

// Fungsi untuk ekstrak nomor surat jalan dari nama file - TETAP SAMA
function extractNomorSuratJalan(fileName) {
  // Contoh: "SJ-2023-001.csv" → "2023-001"
  const match = fileName.match(/(?:SJ|SURATJALAN|SURAT_JALAN)[-_]?([\w-]+)/i);
  return match ? match[1].replace(/\.[^/.]+$/, "") : // Hapus ekstensi file
    fileName.replace(/\.[^/.]+$/, ""); // Gunakan nama file tanpa ekstensi
}
