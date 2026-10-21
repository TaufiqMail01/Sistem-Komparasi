import * as XLSX from 'xlsx';
import mammoth from 'mammoth';

// Fungsi untuk membaca dan memparsing file Excel / CSV
export const readFileData = (file) => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    const isCsv = file.name.toLowerCase().endsWith('.csv');

    reader.onload = (e) => {
      try {
        const data = e.target.result;
        let workbook;

        if (isCsv) {
          workbook = XLSX.read(data, { type: 'string', cellText: true, cellDates: true });
        } else {
          workbook = XLSX.read(data, { type: 'binary', cellText: true, cellDates: true });
        }

        const sheetNames = workbook.SheetNames;
        const sheetsData = {};

        sheetNames.forEach(sheetName => {
          const worksheet = workbook.Sheets[sheetName];
          // Mengubah sheet ke JSON dengan baris kosong diisi string kosong
          const json = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
          sheetsData[sheetName] = json;
        });

        resolve({ sheetNames, sheetsData });
      } catch (error) {
        reject(error);
      }
    };

    reader.onerror = (error) => reject(error);

    if (isCsv) {
      reader.readAsText(file);
    } else {
      reader.readAsBinaryString(file);
    }
  });
};

// Fungsi untuk membaca teks dari dokumen Word (.docx)
export const readWordData = async (file) => {
  try {
    const arrayBuffer = await file.arrayBuffer();
    const result = await mammoth.extractRawText({ arrayBuffer });
    return result.value;
  } catch (error) {
    console.error("Gagal membaca dokumen Word:", error);
    throw error;
  }
};

// Fungsi untuk memformat nilai sel tabel agar aman dari salah deteksi angka panjang
export const formatCellValue = (headerName, value) => {
  if (value === null || value === undefined) return "";
  
  const lowerHeader = String(headerName || "").toLowerCase();
  
  // Deteksi kolom nomor identitas atau nomor penting agar TIDAK diformat jadi Rupiah
  const isIdOrNumberField = 
    lowerHeader.includes('bpjs') || 
    lowerHeader.includes('nik') || 
    lowerHeader.includes('nrk') || 
    lowerHeader.includes('nip') || 
    lowerHeader.includes('id') || 
    lowerHeader.includes('no') ||
    lowerHeader.includes('kode') ||
    lowerHeader.includes('telp') ||
    lowerHeader.includes('phone');

  if (isIdOrNumberField) {
    return String(value).trim();
  }

  // Jika berupa angka finansial murni / numerik biasa
  const num = Number(value);
  if (!isNaN(num) && typeof value !== 'boolean' && String(value).trim() !== "") {
    return num.toLocaleString('id-ID');
  }

  return String(value);
};

// Fungsi bantu parsing angka format Indonesia ke angka sistem
export const parseIndonesianNumber = (val) => {
  if (val === null || val === undefined) return NaN;
  if (typeof val === 'number') return val;
  
  let s = String(val).trim();
  // Hilangkan awalan mata uang jika ada
  s = s.replace(/Rp\.?/gi, '').trim();

  // Jika format menggunakan titik sebagai pemisah ribuan dan koma sebagai desimal (contoh: 1.500.000,50)
  if (s.includes(',') && s.includes('.')) {
    s = s.replace(/\./g, '').replace(',', '.');
  } else if (s.includes(',')) {
    // Jika hanya ada koma (kemungkinan desimal atau pemisah ribuan lokal)
    s = s.replace(',', '.');
  }

  const parsed = parseFloat(s);
  return isNaN(parsed) ? NaN : parsed;
};