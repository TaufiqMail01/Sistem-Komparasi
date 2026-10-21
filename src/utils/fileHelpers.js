import * as XLSX from 'xlsx';
import mammoth from 'mammoth';

// Helper untuk memparsing angka dengan format desimal Indonesia
export const parseIndonesianNumber = (val) => {
  if (val === null || val === undefined || val === "") return NaN;
  if (typeof val === 'number') return val;
  let s = String(val).trim();
  s = s.replace(/Rp\.?/gi, '').trim(); 
  if (s.includes(',') && s.includes('.')) {
    s = s.replace(/\./g, '').replace(',', '.');
  } else if (s.includes(',')) {
    s = s.replace(',', '.');
  }
  const num = Number(s);
  return isNaN(num) ? NaN : num;
};

// Parser khusus Excel / CSV
export const readFileData = (file) => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array', cellDates: true });
        const sheetNames = workbook.SheetNames;
        
        const sheetsData = {};
        sheetNames.forEach(name => {
          const worksheet = workbook.Sheets[name];
          const rawRows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" });
          
          let headerRowIndex = 0;
          const keywords = ['NO', 'ID', 'KODE', 'NRK', 'NIP', 'NAMA', 'KARYAWAN', 'NIK', 'EMPLOYEE', 'NAMA LENGKAP'];

          for (let i = 0; i < rawRows.length; i++) {
            const rowValues = rawRows[i].map(cell => String(cell).trim().toUpperCase());
            const matchCount = keywords.filter(keyword => rowValues.includes(keyword)).length;
            if (matchCount > 0) {
              headerRowIndex = i;
              break;
            }
          }

          const rawJson = XLSX.utils.sheet_to_json(worksheet, { range: headerRowIndex, defval: "" });
          const json = Array.isArray(rawJson) ? rawJson.filter(row => {
            return row && Object.values(row).some(val => val !== "" && val !== null && val !== undefined);
          }) : [];

          sheetsData[name] = json;
        });

        resolve({ workbook, sheetNames, sheetsData });
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = (error) => reject(error);
    reader.readAsArrayBuffer(file);
  });
};

// Parser Dokumen Word (.docx)
export const readWordData = async (file) => {
  const arrayBuffer = await file.arrayBuffer();
  const result = await mammoth.extractRawText({ arrayBuffer });
  return result.value;
};

// Format Tampilan Sel Tabel
export const formatCellValue = (headerName, val) => {
  if (val === null || val === undefined || val === "") return "-";
  
  const lowerHeader = headerName.toLowerCase();
  const excludeCurrencyKeywords = ['no', 'nomor', 'nik', 'nrk', 'kode', 'id', 'telp', 'phone', 'npwp', 'rekening'];
  const isExcludedFromCurrency = excludeCurrencyKeywords.some(keyword => {
    const regex = new RegExp(`(^|[^a-z])${keyword}([^a-z]|$)`, 'i');
    return regex.test(lowerHeader);
  });

  if (isExcludedFromCurrency) {
    return String(val);
  }

  const currencyKeywords = [
    'gaji', 'tunjangan', 'potongan', 'bonus', 'insentif', 'lembur', 'thr', 
    'bruto', 'thp', 'total', 'bayar', 'umk', 'penalty', 'kompensasi', 
    'pajak', 'pph', 'nominal', 'jumlah', 'harga', 'biaya', 'rp', 'saltab', 'hasil_formula', 'iuran',
    'overtime', 'jht', 'jkatan', 'jp', 'kesehatan', 'jkk', 'jkm', 'jshk', 'disnaker', 'fix allowance', 'ter max', 'ter min'
  ];
  const isCurrency = currencyKeywords.some(keyword => lowerHeader.includes(keyword));

  const numVal = parseIndonesianNumber(val);
  if (isCurrency && !isNaN(numVal) && typeof val !== 'boolean') {
    return new Intl.NumberFormat('id-ID', { style: 'currency', currency: 'IDR', maximumFractionDigits: 0 }).format(numVal);
  }

  const dateKeywords = ['tanggal', 'date', 'tgl', 'start', 'end', 'join', 'expired', 'lahir'];
  const isDateHeader = dateKeywords.some(keyword => lowerHeader.includes(keyword));

  if (isDateHeader) {
    if (val instanceof Date) {
      const d = String(val.getDate()).padStart(2, '0');
      const m = String(val.getMonth() + 1).padStart(2, '0');
      const y = val.getFullYear();
      return `${d}-${m}-${y}`;
    }
    if (typeof val === 'string' && /^\d{4}-\d{2}-\d{2}/.test(val)) {
      const parts = val.split('T')[0].split('-');
      if (parts.length === 3) return `${parts[2]}-${parts[1]}-${parts[0]}`;
    }
    if (typeof val === 'number' && val > 30000 && val < 60000) {
      const utc_days = Math.floor(val - 25569);
      const utc_value = utc_days * 86400;
      const dateInfo = new Date(utc_value * 1000);
      const d = String(dateInfo.getUTCDate()).padStart(2, '0');
      const m = String(dateInfo.getUTCMonth() + 1).padStart(2, '0');
      const y = dateInfo.getUTCFullYear();
      return `${d}-${m}-${y}`;
    }
  }

  if (typeof val === 'object') return JSON.stringify(val);
  return String(val);
};