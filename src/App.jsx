import React, { useState, useMemo } from 'react';
import * as XLSX from 'xlsx';
import { 
  Upload, FileSpreadsheet, ArrowRightLeft, Download, 
  Settings, Key, Sliders, Layers, CheckCircle, AlertTriangle, 
  PlusCircle, MinusCircle, Eye, CheckCircle2, AlertCircle, FilePlus, FileMinus, X, Calculator, Plus, RotateCcw, Database, Type, Binary
} from 'lucide-react';

export default function App() {
  const [file1, setFile1] = useState(null);
  const [file2, setFile2] = useState(null);

  const [sheet1, setSheet1] = useState('');
  const [sheet2, setSheet2] = useState('');

  const [config, setConfig] = useState({
    primaryKey: '',
    ignoreCase: true,
    trimWhitespace: true,
    numericTolerance: 0.001,
    selectedColumns: []
  });

  const [enableFormula, setEnableFormula] = useState(false);
  const [formulaExpr, setFormulaExpr] = useState('');

  const [filter, setFilter] = useState('ALL');
  const [selectedRowDetail, setSelectedRowDetail] = useState(null);

  const readFileData = (file) => {
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

  const handleUploadFile1 = async (file) => {
    try {
      const res = await readFileData(file);
      setFile1({ fileName: file.name, ...res });
      setSheet1(res.sheetNames[0]);
    } catch (err) {
      alert("Gagal membaca File 1. Pastikan file berformat Excel/CSV valid.");
    }
  };

  const handleUploadFile2 = async (file) => {
    try {
      const res = await readFileData(file);
      setFile2({ fileName: file.name, ...res });
      setSheet2(res.sheetNames[0]);
    } catch (err) {
      alert("Gagal membaca File 2. Pastikan file berformat Excel/CSV valid.");
    }
  };

  const resetAllData = () => {
    setFile1(null);
    setFile2(null);
    setSheet1('');
    setSheet2('');
    setFormulaExpr('');
    setEnableFormula(false);
  };

  const commonHeaders = useMemo(() => {
    if (!file1 || !file2) return [];
    const data1 = file1.sheetsData[sheet1] || [];
    const data2 = file2.sheetsData[sheet2] || [];
    const h1 = data1.length > 0 ? Object.keys(data1[0]) : [];
    const h2 = data2.length > 0 ? Object.keys(data2[0]) : [];
    return Array.from(new Set([...h1, ...h2]));
  }, [file1, file2, sheet1, sheet2]);

  const insertColumnTagToFormula = (col) => {
    setFormulaExpr(prev => prev + `[${col}]`);
  };

  const evaluateFormula = (row, expr, headers) => {
    if (!row) return 0;
    try {
      let evaluatedExpr = expr;
      headers.forEach(h => {
        const val = Number(row[h]) || 0;
        const regex = new RegExp(`\\[${h}\\]`, 'g');
        evaluatedExpr = evaluatedExpr.replace(regex, val);
      });
      const result = Function(`'use strict'; return (${evaluatedExpr})`)();
      return isNaN(result) ? 0 : Number(result.toFixed(2));
    } catch (e) {
      return "Error Rumus";
    }
  };

  // Format universal untuk Rupiah, Tanggal, dan Teks
  const formatCellValue = (headerName, val) => {
    if (val === null || val === undefined || val === "") return "-";
    
    const lowerHeader = headerName.toLowerCase();
    
    // Deteksi komprehensif untuk data keuangan / nominal / gaji
    const currencyKeywords = [
      'gaji', 'tunjangan', 'potongan', 'bonus', 'insentif', 'lembur', 'thr', 
      'bruto', 'thp', 'total', 'bayar', 'umk', 'penalty', 'bpjs', 'kompensasi', 
      'pajak', 'pph', 'nominal', 'jumlah', 'harga', 'biaya', 'rp', 'saltab'
    ];
    const isCurrency = currencyKeywords.some(keyword => lowerHeader.includes(keyword));

    if (isCurrency && !isNaN(Number(val)) && typeof val !== 'boolean') {
      return new Intl.NumberFormat('id-ID', { style: 'currency', currency: 'IDR', maximumFractionDigits: 0 }).format(Number(val));
    }

    // Deteksi format tanggal (dd-mm-yyyy)
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
    }

    if (typeof val === 'object') return JSON.stringify(val);
    return String(val);
  };

  const comparisonResult = useMemo(() => {
    if (!file1 || !file2 || !sheet1 || !sheet2) return null;

    const data1 = file1.sheetsData[sheet1] || [];
    const data2 = file2.sheetsData[sheet2] || [];
    const { primaryKey, ignoreCase, trimWhitespace, numericTolerance, selectedColumns } = config;

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
      if (!isNaN(n1) && !isNaN(n2) && val1 !== "" && val2 !== "") {
        return Math.abs(n1 - n2) <= numericTolerance;
      }
      return normalizeVal(val1) === normalizeVal(val2);
    };

    let allHeaders = commonHeaders;
    if (enableFormula && formulaExpr.trim()) {
      allHeaders = [...allHeaders, "Hasil_Formula"];
    }

    const activeHeaders = selectedColumns.length > 0 
      ? allHeaders.filter(h => selectedColumns.includes(h) || h === "Hasil_Formula")
      : allHeaders;

    const results = [];
    let counts = { total: 0, matched: 0, modified: 0, added: 0, deleted: 0 };

    const enrichRow = (row) => {
      if (!row) return null;
      const newRow = { ...row };
      if (enableFormula && formulaExpr.trim()) {
        const originalCols = Object.keys(row);
        newRow["Hasil_Formula"] = evaluateFormula(row, formulaExpr, originalCols);
      }
      return newRow;
    };

    const enrichedData1 = data1.map(enrichRow);
    const enrichedData2 = data2.map(enrichRow);

    if (primaryKey && primaryKey !== '__INDEX__') {
      const map2 = new Map();
      enrichedData2.forEach(row => {
        const keyVal = normalizeVal(row[primaryKey]);
        if (keyVal) map2.set(keyVal, row);
      });

      const processedKeysIn2 = new Set();

      enrichedData1.forEach((row1, idx) => {
        const keyVal = normalizeVal(row1[primaryKey]);
        const row2 = map2.get(keyVal);

        if (!row2) {
          results.push({
            id: `row-del-${idx}`,
            key: row1[primaryKey] !== undefined ? String(row1[primaryKey]) : `Baris ${idx + 1}`,
            status: 'DELETED',
            diffCount: activeHeaders.length,
            row1, row2: null, diffs: {}
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
            key: row1[primaryKey] !== undefined ? String(row1[primaryKey]) : `Baris ${idx + 1}`,
            status,
            diffCount,
            row1, row2, diffs
          });
        }
      });

      enrichedData2.forEach((row2, idx) => {
        const keyVal = normalizeVal(row2[primaryKey]);
        if (!processedKeysIn2.has(keyVal)) {
          results.push({
            id: `row-add-${idx}`,
            key: row2[primaryKey] !== undefined ? String(row2[primaryKey]) : `Baris ${idx + 1}`,
            status: 'ADDED',
            diffCount: activeHeaders.length,
            row1: null, row2, diffs: {}
          });
          counts.added++;
        }
      });

    } else {
      const maxRows = Math.max(enrichedData1.length, enrichedData2.length);
      for (let i = 0; i < maxRows; i++) {
        const row1 = enrichedData1[i] || null;
        const row2 = enrichedData2[i] || null;

        if (row1 && !row2) {
          results.push({
            id: `row-${i}`, key: `Baris ${i + 1}`, status: 'DELETED',
            diffCount: activeHeaders.length, row1, row2: null, diffs: {}
          });
          counts.deleted++;
        } else if (!row1 && row2) {
          results.push({
            id: `row-${i}`, key: `Baris ${i + 1}`, status: 'ADDED',
            diffCount: activeHeaders.length, row1: null, row2, diffs: {}
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
            id: `row-${i}`, key: `Baris ${i + 1}`, status,
            diffCount, row1, row2, diffs
          });
        }
      }
    }

    counts.total = results.length;
    return { results, counts, headers: activeHeaders };
  }, [file1, file2, sheet1, sheet2, config, enableFormula, formulaExpr, commonHeaders]);

  const filteredResults = useMemo(() => {
    if (!comparisonResult) return [];
    if (filter === 'ALL') return comparisonResult.results;
    return comparisonResult.results.filter(r => r.status === filter);
  }, [comparisonResult, filter]);

  const exportToExcel = () => {
    if (!comparisonResult) return;
    const exportData = comparisonResult.results.map(item => {
      const row = {
        'Status': item.status,
        'Kunci / Identifier': item.key,
        'Jumlah Beda Kolom': item.diffCount
      };
      comparisonResult.headers.forEach(h => {
        if (item.status === 'MATCHED') {
          row[h] = item.row1 ? item.row1[h] : "";
        } else {
          row[`${h} (File 1)`] = item.row1 ? item.row1[h] : "";
          row[`${h} (File 2)`] = item.row2 ? item.row2[h] : "";
        }
      });
      return row;
    });

    const worksheet = XLSX.utils.json_to_sheet(exportData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Hasil Rekonsiliasi");
    XLSX.writeFile(workbook, "Laporan_Hasil_Rekonsiliasi.xlsx");
  };

  return (
    <div className="bg-slate-50 text-slate-800 font-sans antialiased min-h-screen flex flex-col">
      <header className="bg-gradient-to-r from-slate-900 via-slate-800 to-indigo-950 text-white border-b border-slate-800 sticky top-0 z-30 shadow-md">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-3.5">
          <div className="flex flex-col md:flex-row md:items-center md:justify-between gap-3">
            <div className="flex items-center gap-3">
              <div className="p-2.5 bg-blue-600/30 rounded-xl border border-blue-400/30 text-blue-400">
                <Layers className="w-6 h-6" />
              </div>
              <div>
                <div className="flex items-center gap-2">
                  <span className="px-2 py-0.5 text-[10px] font-bold bg-blue-500/20 text-blue-300 border border-blue-500/30 rounded-full uppercase tracking-wider">
                    Universal Excel Reconciliation
                  </span>
                  <span className="text-xs text-slate-400 hidden sm:inline">| Skripsi Program Science (Kode 156)</span>
                </div>
                <h1 className="text-lg md:text-xl font-bold tracking-tight text-white">
                  Sistem Komparasi & Rekonsiliasi Data Excel Universal
                </h1>
              </div>
            </div>

            <div className="flex items-center gap-2 text-xs">
              <button onClick={exportToExcel} disabled={!comparisonResult} className="flex items-center gap-1.5 px-3.5 py-2 bg-emerald-700 hover:bg-emerald-600 disabled:opacity-40 disabled:cursor-not-allowed text-white font-medium rounded-xl shadow-sm transition-all border border-emerald-500/30">
                <Download className="w-4 h-4" />
                <span>Ekspor Laporan (.xlsx)</span>
              </button>
            </div>
          </div>
        </div>
      </header>

      <main className="flex-1 max-w-7xl w-full mx-auto px-4 sm:px-6 lg:px-8 py-6 space-y-6">
        <section className="bg-white rounded-2xl border border-slate-200 shadow-sm p-5 space-y-5">
          <div className="flex flex-col md:flex-row md:items-center justify-between gap-4 pb-4 border-b border-slate-100">
            <div>
              <h2 className="text-base font-semibold text-slate-900 flex items-center gap-2">
                <FileSpreadsheet className="w-5 h-5 text-blue-600" />
                1. Unggah File & Pilih Sheet
              </h2>
              <p className="text-xs text-slate-500 mt-0.5">
                Unggah file Excel / CSV apa pun (File 1 = Data Lama / Acuan, File 2 = Data Baru / Pembanding).
              </p>
            </div>
            <button onClick={resetAllData} className="text-xs font-medium text-slate-500 hover:text-rose-600 flex items-center gap-1 self-start md:self-auto px-3 py-1.5 rounded-lg border border-slate-200 hover:bg-slate-50 transition-colors">
              <RotateCcw className="w-3.5 h-3.5" />
              Reset File & Data
            </button>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
            <div className="border-2 border-dashed border-slate-200 hover:border-amber-400 bg-slate-50/50 rounded-2xl p-4 transition-all relative group">
              <input type="file" accept=".xlsx, .xls, .csv" onChange={(e) => e.target.files[0] && handleUploadFile1(e.target.files[0])} className="absolute inset-0 w-full h-full opacity-0 cursor-pointer z-10" />
              <div className="flex items-start gap-3">
                <div className="p-3 bg-amber-100 text-amber-700 rounded-xl flex-shrink-0">
                  <FileSpreadsheet className="w-6 h-6" />
                </div>
                <div className="flex-1 min-w-0">
                  <span className="text-[10px] font-bold text-amber-700 uppercase tracking-wider block">FILE 1 (DATA ACUAN / LAMA)</span>
                  <p className="text-sm font-semibold text-slate-800 truncate mt-0.5">{file1 ? file1.fileName : "Pilih atau drag File 1 ke sini..."}</p>
                  <p className="text-[11px] text-slate-400">Format .xlsx, .xls, .csv</p>
                  
                  {file1 && (
                    <div className="mt-2.5 relative z-20">
                      <label className="text-[11px] font-medium text-slate-600 block mb-1">Pilih Sheet File 1:</label>
                      <select value={sheet1} onChange={(e) => setSheet1(e.target.value)} className="w-full bg-white text-xs border border-slate-300 rounded-lg px-2.5 py-1.5">
                        {file1.sheetNames.map(s => <option key={s} value={s}>{s}</option>)}
                      </select>
                    </div>
                  )}
                </div>
              </div>
            </div>

            <div className="border-2 border-dashed border-slate-200 hover:border-blue-400 bg-slate-50/50 rounded-2xl p-4 transition-all relative group">
              <input type="file" accept=".xlsx, .xls, .csv" onChange={(e) => e.target.files[0] && handleUploadFile2(e.target.files[0])} className="absolute inset-0 w-full h-full opacity-0 cursor-pointer z-10" />
              <div className="flex items-start gap-3">
                <div className="p-3 bg-blue-100 text-blue-700 rounded-xl flex-shrink-0">
                  <FileSpreadsheet className="w-6 h-6" />
                </div>
                <div className="flex-1 min-w-0">
                  <span className="text-[10px] font-bold text-blue-700 uppercase tracking-wider block">FILE 2 (DATA PEMBANDING / BARU)</span>
                  <p className="text-sm font-semibold text-slate-800 truncate mt-0.5">{file2 ? file2.fileName : "Pilih atau drag File 2 ke sini..."}</p>
                  <p className="text-[11px] text-slate-400">Format .xlsx, .xls, .csv</p>
                  
                  {file2 && (
                    <div className="mt-2.5 relative z-20">
                      <label className="text-[11px] font-medium text-slate-600 block mb-1">Pilih Sheet File 2:</label>
                      <select value={sheet2} onChange={(e) => setSheet2(e.target.value)} className="w-full bg-white text-xs border border-slate-300 rounded-lg px-2.5 py-1.5">
                        {file2.sheetNames.map(s => <option key={s} value={s}>{s}</option>)}
                      </select>
                    </div>
                  )}
                </div>
              </div>
            </div>
          </div>
        </section>

        <section className="bg-white rounded-2xl border border-slate-200 shadow-sm p-5 space-y-5">
          <div className="flex items-center justify-between pb-3 border-b border-slate-100">
            <h2 className="text-base font-semibold text-slate-900 flex items-center gap-2">
              <Sliders className="w-5 h-5 text-indigo-600" />
              2. Pengaturan Parameter Komparasi & Formula Engine
            </h2>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-3 gap-5">
            <div className="space-y-1.5">
              <label className="text-xs font-semibold text-slate-700 flex items-center gap-1.5">
                <Key className="w-3.5 h-3.5 text-amber-500" />
                Kolom Kunci Utama (Primary Key):
              </label>
              <select 
                value={config.primaryKey} 
                onChange={(e) => setConfig({ ...config, primaryKey: e.target.value })} 
                disabled={!file1 || !file2}
                className="w-full bg-slate-50 text-xs font-medium text-slate-800 rounded-xl border border-slate-300 px-3 py-2 disabled:opacity-50"
              >
                <option value="__INDEX__">Pencocokan Baris (Nomor Urut / Index)</option>
                {commonHeaders.map(col => <option key={col} value={col}>{col}</option>)}
              </select>
              <p className="text-[11px] text-slate-400">Pilih kolom ID Unik (NRK, NIK, ID, Kode) agar data tidak tertukar.</p>
            </div>

            <div className="space-y-2">
              <label className="text-xs font-semibold text-slate-700 flex items-center gap-1.5">
                <Type className="w-3.5 h-3.5 text-indigo-500" />
                Aturan Pemrosesan Teks:
              </label>
              <div className="space-y-2 pt-0.5">
                <label className="flex items-center gap-2 text-xs text-slate-700 cursor-pointer">
                  <input type="checkbox" checked={config.ignoreCase} onChange={(e) => setConfig({ ...config, ignoreCase: e.target.checked })} className="w-4 h-4 text-indigo-600 rounded border-slate-300" />
                  <span>Abaikan Kapitalisasi (Case-Insensitive)</span>
                </label>
                <label className="flex items-center gap-2 text-xs text-slate-700 cursor-pointer">
                  <input type="checkbox" checked={config.trimWhitespace} onChange={(e) => setConfig({ ...config, trimWhitespace: e.target.checked })} className="w-4 h-4 text-indigo-600 rounded border-slate-300" />
                  <span>Abaikan Spasi Awal & Akhir</span>
                </label>
              </div>
            </div>

            <div className="space-y-1.5">
              <label className="text-xs font-semibold text-slate-700 flex items-center gap-1.5">
                <Calculator className="w-3.5 h-3.5 text-emerald-500" />
                Toleransi Selisih Angka:
              </label>
              <input type="number" value={config.numericTolerance} step="0.001" min="0" onChange={(e) => setConfig({ ...config, numericTolerance: parseFloat(e.target.value) || 0 })} className="w-full bg-slate-50 text-xs text-slate-800 rounded-xl border border-slate-300 px-3 py-2" />
              <p className="text-[11px] text-slate-400">Selisih nominal di bawah nilai ini diabaikan.</p>
            </div>
          </div>

          <div className="pt-4 border-t border-slate-100 space-y-3 bg-indigo-50/40 p-4 rounded-xl border border-indigo-100">
            <div className="flex items-center justify-between">
              <label className="text-xs font-bold text-indigo-900 flex items-center gap-1.5">
                <Binary className="w-4 h-4 text-indigo-600" />
                Modul Evaluasi Formula Kustom (Dynamic Custom Formula Engine):
              </label>
              <label className="flex items-center gap-2 text-xs text-indigo-800 font-medium cursor-pointer">
                <input type="checkbox" checked={enableFormula} onChange={(e) => setEnableFormula(e.target.checked)} className="w-4 h-4 text-indigo-600 rounded border-indigo-300" />
                <span>Aktifkan Pengolahan Formula</span>
              </label>
            </div>

            {enableFormula && (
              <div className="space-y-2.5 animate-fade-in">
                <input type="text" placeholder="Contoh: [Gaji] + [Tunjangan] - [Potongan]" value={formulaExpr} onChange={(e) => setFormulaExpr(e.target.value)} className="w-full bg-white text-xs font-mono text-slate-800 rounded-xl border border-indigo-200 px-3.5 py-2.5 shadow-sm" />
                <div>
                  <span className="text-[11px] font-semibold text-slate-600 block mb-1">Klik nama kolom untuk menyisipkan ke rumus:</span>
                  <div className="flex flex-wrap gap-1.5 max-h-24 overflow-y-auto">
                    {commonHeaders.map(col => (
                      <button key={col} onClick={() => insertColumnTagToFormula(col)} className="px-2 py-1 bg-white border border-slate-300 text-xs rounded hover:border-indigo-600 font-mono">
                        +{col}
                      </button>
                    ))}
                  </div>
                </div>
              </div>
            )}
          </div>
        </section>

        {comparisonResult && (
          <section className="space-y-4">
            <div className="grid grid-cols-2 md:grid-cols-5 gap-3">
              <div className="bg-white p-4 rounded-2xl border border-slate-200 shadow-sm flex flex-col justify-between">
                <div className="flex items-center justify-between text-slate-500">
                  <span className="text-[11px] font-bold uppercase tracking-wider">Total Record</span>
                  <Database className="w-4 h-4 text-slate-400" />
                </div>
                <div className="mt-2">
                  <div className="text-2xl font-bold text-slate-900">{comparisonResult.counts.total}</div>
                </div>
              </div>

              <div className="bg-white p-4 rounded-2xl border border-slate-200 shadow-sm flex flex-col justify-between">
                <div className="flex items-center justify-between text-emerald-600">
                  <span className="text-[11px] font-bold uppercase tracking-wider">Sama (Identik)</span>
                  <CheckCircle className="w-4 h-4" />
                </div>
                <div className="mt-2">
                  <div className="text-2xl font-bold text-emerald-600">{comparisonResult.counts.matched}</div>
                </div>
              </div>

              <div className="bg-white p-4 rounded-2xl border border-slate-200 shadow-sm flex flex-col justify-between">
                <div className="flex items-center justify-between text-amber-600">
                  <span className="text-[11px] font-bold uppercase tracking-wider">Ada Perbedaan</span>
                  <AlertTriangle className="w-4 h-4" />
                </div>
                <div className="mt-2">
                  <div className="text-2xl font-bold text-amber-600">{comparisonResult.counts.modified}</div>
                </div>
              </div>

              <div className="bg-white p-4 rounded-2xl border border-slate-200 shadow-sm flex flex-col justify-between">
                <div className="flex items-center justify-between text-blue-600">
                  <span className="text-[11px] font-bold uppercase tracking-wider">Baru (Added)</span>
                  <PlusCircle className="w-4 h-4" />
                </div>
                <div className="mt-2">
                  <div className="text-2xl font-bold text-blue-600">{comparisonResult.counts.added}</div>
                </div>
              </div>

              <div className="bg-white p-4 rounded-2xl border border-slate-200 shadow-sm col-span-2 md:col-span-1 flex flex-col justify-between">
                <div className="flex items-center justify-between text-rose-600">
                  <span className="text-[11px] font-bold uppercase tracking-wider">Dihapus</span>
                  <MinusCircle className="w-4 h-4" />
                </div>
                <div className="mt-2">
                  <div className="text-2xl font-bold text-rose-600">{comparisonResult.counts.deleted}</div>
                </div>
              </div>
            </div>

            <div className="flex items-center gap-2 bg-white p-2 rounded-xl border border-slate-200 shadow-sm overflow-x-auto">
              {[
                { key: 'ALL', label: 'Semua Status' },
                { key: 'MATCHED', label: 'Identik' },
                { key: 'MODIFIED', label: 'Beda' },
                { key: 'ADDED', label: 'Baru' },
                { key: 'DELETED', label: 'Dihapus' },
              ].map(tab => (
                <button
                  key={tab.key}
                  onClick={() => setFilter(tab.key)}
                  className={`px-3 py-1.5 rounded-lg text-xs font-semibold transition-colors whitespace-nowrap ${filter === tab.key ? 'bg-indigo-600 text-white' : 'text-slate-600 hover:bg-slate-100'}`}
                >
                  {tab.label}
                </button>
              ))}
            </div>

            <div className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden">
              <div className="overflow-x-auto max-h-[500px]">
                <table className="w-full text-left text-xs">
                  <thead className="bg-slate-100 text-slate-700 uppercase font-semibold sticky top-0 z-10 border-b border-slate-200">
                    <tr>
                      <th className="p-3 w-12 text-center">Aksi</th>
                      <th className="p-3 w-28">Status</th>
                      <th className="p-3 w-36">Kunci Baris</th>
                      {comparisonResult.headers.slice(0, 3).map(h => (
                        <th key={h} className="p-3 min-w-[140px]">{h}</th>
                      ))}
                      {comparisonResult.headers.length > 3 && (
                        <th className="p-3 text-slate-400 italic">+ {comparisonResult.headers.length - 3} Kolom Lainnya (Cek Detail)</th>
                      )}
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {filteredResults.length === 0 ? (
                      <tr>
                        <td colSpan={comparisonResult.headers.length + 3} className="p-8 text-center text-slate-400">
                          Tidak ada data yang cocok dengan filter.
                        </td>
                      </tr>
                    ) : (
                      filteredResults.map(row => (
                        <tr key={row.id} className="hover:bg-slate-50 transition-colors">
                          <td className="p-3 text-center">
                            <button
                              onClick={() => setSelectedRowDetail(row)}
                              className="p-1.5 text-indigo-600 hover:bg-indigo-50 rounded-lg transition-colors shadow-2xs"
                              title="Lihat Detail Side-by-Side"
                            >
                              <Eye className="w-4 h-4" />
                            </button>
                          </td>
                          <td className="p-3">
                            {row.status === 'MATCHED' && <span className="inline-flex items-center gap-1 px-2.5 py-1 rounded-full bg-emerald-100 text-emerald-800"><CheckCircle2 className="w-3.5 h-3.5" /> Sesuai</span>}
                            {row.status === 'MODIFIED' && <span className="inline-flex items-center gap-1 px-2.5 py-1 rounded-full bg-amber-100 text-amber-800"><AlertCircle className="w-3.5 h-3.5" /> Beda</span>}
                            {row.status === 'ADDED' && <span className="inline-flex items-center gap-1 px-2.5 py-1 rounded-full bg-blue-100 text-blue-800"><FilePlus className="w-3.5 h-3.5" /> Baru</span>}
                            {row.status === 'DELETED' && <span className="inline-flex items-center gap-1 px-2.5 py-1 rounded-full bg-rose-100 text-rose-800"><FileMinus className="w-3.5 h-3.5" /> Hapus</span>}
                          </td>
                          <td className="p-3 font-medium text-slate-800">{row.key}</td>
                          {comparisonResult.headers.slice(0, 3).map(h => {
                            const diff = row.diffs[h];
                            const isDiff = diff && !diff.isEqual;
                            return (
                              <td key={h} className={`p-3 ${isDiff ? 'bg-amber-50 font-semibold text-amber-900 border-l-2 border-amber-400' : 'text-slate-600'}`}>
                                {row.status === 'MATCHED' ? formatCellValue(h, row.row1 ? row.row1[h] : "") : (
                                  <div className="flex flex-col gap-0.5">
                                    <span>{formatCellValue(h, row.row1 ? row.row1[h] : "")}</span>
                                    {isDiff && <span className="text-rose-600 font-bold text-[11px] pt-0.5 border-t border-dashed border-amber-200">➔ {formatCellValue(h, row.row2 ? row.row2[h] : "")}</span>}
                                  </div>
                                )}
                              </td>
                            );
                          })}
                          {comparisonResult.headers.length > 3 && (
                            <td className="p-3 text-slate-400 italic">Klik ikon mata (👁) untuk lihat selengkapnya</td>
                          )}
                        </tr>
                      ))
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          </section>
        )}
      </main>

      {/* Modal Detail Audit dengan Desain Lebih Rapi & Profesional */}
      {selectedRowDetail && comparisonResult && (
        <div className="fixed inset-0 z-50 bg-slate-900/60 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-white rounded-3xl shadow-2xl max-w-4xl w-full max-h-[90vh] flex flex-col overflow-hidden animate-fade-in border border-slate-100">
            
            {/* Header Modal */}
            <div className="p-5 border-b border-slate-100 flex items-center justify-between bg-slate-900 text-white">
              <div className="flex items-center gap-3">
                <div className="p-2 bg-indigo-600 rounded-xl text-white">
                  <Database className="w-5 h-5" />
                </div>
                <div>
                  <h3 className="font-bold text-base">Detail Audit Baris / Record</h3>
                  <p className="text-xs text-slate-300">Identifier (Primary Key): <span className="font-mono text-amber-400 font-bold">{selectedRowDetail.key}</span></p>
                </div>
              </div>
              <div className="flex items-center gap-3">
                <span className={`px-3 py-1 rounded-full text-xs font-bold uppercase tracking-wider ${
                  selectedRowDetail.status === 'MATCHED' ? 'bg-emerald-500/20 text-emerald-300 border border-emerald-500/30' :
                  selectedRowDetail.status === 'MODIFIED' ? 'bg-amber-500/20 text-amber-300 border border-amber-500/30' :
                  selectedRowDetail.status === 'ADDED' ? 'bg-blue-500/20 text-blue-300 border border-blue-500/30' :
                  'bg-rose-50/20 text-rose-300 border border-rose-500/30'
                }`}>
                  {selectedRowDetail.status}
                </span>
                <button onClick={() => setSelectedRowDetail(null)} className="p-2 bg-slate-800 hover:bg-slate-700 rounded-xl text-slate-300 hover:text-white transition-colors">
                  <X className="w-5 h-5" />
                </button>
              </div>
            </div>
            
            {/* Body Content */}
            <div className="p-6 overflow-y-auto space-y-4 flex-1 bg-slate-50/50">
              <div className="bg-amber-50/80 border border-amber-200/80 p-3.5 rounded-2xl text-xs text-amber-900 flex items-center justify-between shadow-xs">
                <span className="font-medium">Perbandingan atribut dan perhitungan nilai secara berdampingan. Atribut yang berbeda disorot kuning.</span>
                <span className="font-bold bg-amber-200/70 px-2.5 py-1 rounded-lg text-amber-950">{selectedRowDetail.diffCount} Kolom Berbeda</span>
              </div>

              <div className="bg-white rounded-2xl border border-slate-200/80 shadow-xs overflow-hidden">
                <table className="w-full text-xs">
                  <thead className="bg-slate-100/80 text-slate-700 font-semibold border-b border-slate-200">
                    <tr>
                      <th className="p-3.5 text-left w-1/3">Nama Kolom / Atribut</th>
                      <th className="p-3.5 text-left w-1/3 text-slate-600">File 1 (Data Lama / Acuan)</th>
                      <th className="p-3.5 text-left w-1/3 text-slate-800">File 2 (Data Baru / Pembanding)</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {comparisonResult.headers.map(h => {
                      const diff = selectedRowDetail.diffs[h];
                      const isEqual = diff ? diff.isEqual : true;
                      return (
                        <tr key={h} className={!isEqual ? 'bg-amber-50/60 font-semibold' : 'hover:bg-slate-50/50 transition-colors'}>
                          <td className="p-3.5 font-medium text-slate-800 border-r border-slate-100 flex items-center gap-2">
                            <span className={`w-2 h-2 rounded-full ${!isEqual ? 'bg-amber-500' : 'bg-slate-300'}`}></span>
                            {h}
                          </td>
                          <td className={`p-3.5 border-r border-slate-100 font-mono ${!isEqual ? 'text-rose-700 bg-rose-50/40' : 'text-slate-600'}`}>
                            {formatCellValue(h, selectedRowDetail.row1 ? selectedRowDetail.row1[h] : "")}
                          </td>
                          <td className={`p-3.5 font-mono ${!isEqual ? 'text-emerald-700 bg-emerald-50/40 font-bold' : 'text-slate-600'}`}>
                            {formatCellValue(h, selectedRowDetail.row2 ? selectedRowDetail.row2[h] : "")}
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </div>

            {/* Footer Modal */}
            <div className="p-4 border-t border-slate-100 bg-white flex justify-end">
              <button onClick={() => setSelectedRowDetail(null)} className="px-5 py-2.5 bg-slate-900 hover:bg-slate-800 text-white text-xs font-semibold rounded-xl transition-colors shadow-sm">
                Tutup Jendela Detail
              </button>
            </div>

          </div>
        </div>
      )}
    </div>
  );
}