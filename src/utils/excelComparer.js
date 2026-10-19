import * as XLSX from 'xlsx';

/**
 * Membaca file Excel/CSV dan mengembalikan daftar sheet & data JSON
 */
export const readExcelFile = (file) => {
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
          const json = XLSX.utils.sheet_to_json(worksheet, { defval: "" });
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

/**
 * Engine Utama Komparasi 2 Dataset
 */
export const compareDatasets = (data1, data2, options = {}) => {
  const {
    primaryKey = "",
    ignoreCase = true,
    trimWhitespace = true,
    numericTolerance = 0,
    selectedColumns = []
  } = options;

  const normalizeVal = (val) => {
    if (val === null || val === undefined) return "";
    let s = String(val);
    if (trimWhitespace) s = s.trim();
    if (ignoreCase) s = s.toLowerCase();
    return s;
  };

  const areValuesEqual = (val1, val2) => {
    const n1 = Number(val1);
    const n2 = Number(val2);

    // Cek jika kedua nilai numerik
    if (!isNaN(n1) && !isNaN(n2) && val1 !== "" && val2 !== "") {
      return Math.abs(n1 - n2) <= numericTolerance;
    }

    return normalizeVal(val1) === normalizeVal(val2);
  };

  // Kumpulkan seluruh kolom unik dari kedua file
  const cols1 = data1.length > 0 ? Object.keys(data1[0]) : [];
  const cols2 = data2.length > 0 ? Object.keys(data2[0]) : [];
  const allHeaders = Array.from(new Set([...cols1, ...cols2]));

  const activeHeaders = selectedColumns.length > 0 
    ? allHeaders.filter(h => selectedColumns.includes(h))
    : allHeaders;

  const results = [];
  let counts = { total: 0, matched: 0, modified: 0, added: 0, deleted: 0 };

  if (primaryKey) {
    // === KOMPARASI BERDASARKAN PRIMARY KEY ===
    const map2 = new Map();
    data2.forEach(row => {
      const keyVal = normalizeVal(row[primaryKey]);
      if (keyVal) map2.set(keyVal, row);
    });

    const processedKeysIn2 = new Set();

    data1.forEach((row1, idx) => {
      const keyVal = normalizeVal(row1[primaryKey]);
      const row2 = map2.get(keyVal);

      if (!row2) {
        // Data Ada di File 1 tapi Tidak Ada di File 2
        results.push({
          id: `row-del-${idx}`,
          key: row1[primaryKey] || `Baris ${idx + 1}`,
          status: 'DELETED',
          diffCount: activeHeaders.length,
          row1,
          row2: null,
          diffs: {}
        });
        counts.deleted++;
      } else {
        processedKeysIn2.add(keyVal);
        let diffCount = 0;
        const diffs = {};

        activeHeaders.forEach(col => {
          const v1 = row1[col] ?? "";
          const v2 = row2[col] ?? "";
          const isEqual = areValuesEqual(v1, v2);

          diffs[col] = { v1, v2, isEqual };
          if (!isEqual) diffCount++;
        });

        const status = diffCount === 0 ? 'MATCHED' : 'MODIFIED';
        if (status === 'MATCHED') counts.matched++;
        else counts.modified++;

        results.push({
          id: `row-mod-${idx}`,
          key: row1[primaryKey],
          status,
          diffCount,
          row1,
          row2,
          diffs
        });
      }
    });

    // Cari data baru di File 2 (ADDED)
    data2.forEach((row2, idx) => {
      const keyVal = normalizeVal(row2[primaryKey]);
      if (!processedKeysIn2.has(keyVal)) {
        results.push({
          id: `row-add-${idx}`,
          key: row2[primaryKey] || `Baris ${idx + 1}`,
          status: 'ADDED',
          diffCount: activeHeaders.length,
          row1: null,
          row2,
          diffs: {}
        });
        counts.added++;
      }
    });

  } else {
    // === KOMPARASI BERDASARKAN URUTAN BARIS (ROW BY ROW) ===
    const maxRows = Math.max(data1.length, data2.length);

    for (let i = 0; i < maxRows; i++) {
      const row1 = data1[i] || null;
      const row2 = data2[i] || null;

      if (row1 && !row2) {
        results.push({
          id: `row-${i}`,
          key: `Baris ${i + 1}`,
          status: 'DELETED',
          diffCount: activeHeaders.length,
          row1, row2: null, diffs: {}
        });
        counts.deleted++;
      } else if (!row1 && row2) {
        results.push({
          id: `row-${i}`,
          key: `Baris ${i + 1}`,
          status: 'ADDED',
          diffCount: activeHeaders.length,
          row1: null, row2, diffs: {}
        });
        counts.added++;
      } else {
        let diffCount = 0;
        const diffs = {};

        activeHeaders.forEach(col => {
          const v1 = row1[col] ?? "";
          const v2 = row2[col] ?? "";
          const isEqual = areValuesEqual(v1, v2);

          diffs[col] = { v1, v2, isEqual };
          if (!isEqual) diffCount++;
        });

        const status = diffCount === 0 ? 'MATCHED' : 'MODIFIED';
        if (status === 'MATCHED') counts.matched++;
        else counts.modified++;

        results.push({
          id: `row-${i}`,
          key: `Baris ${i + 1}`,
          status,
          diffCount,
          row1, row2, diffs
        });
      }
    }
  }

  counts.total = results.length;
  return { results, counts, headers: activeHeaders, allHeaders };
};

/**
 * Export hasil rekonsiliasi ke Excel
 */
export const exportToExcel = (results, headers, filename = "Hasil_Rekonsiliasi.xlsx") => {
  const exportData = results.map(item => {
    const row = {
      'Status': item.status,
      'Kunci / Baris': item.key,
      'Jumlah Beda Kolom': item.diffCount
    };

    headers.forEach(h => {
      if (item.status === 'MATCHED') {
        row[`${h}`] = item.row1 ? item.row1[h] : "";
      } else {
        row[`${h} (File 1)`] = item.row1 ? item.row1[h] : "";
        row[`${h} (File 2)`] = item.row2 ? item.row2[h] : "";
      }
    });

    return row;
  });

  const worksheet = XLSX.utils.json_to_sheet(exportData);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Hasil Komparasi");
  XLSX.writeFile(workbook, filename);
};