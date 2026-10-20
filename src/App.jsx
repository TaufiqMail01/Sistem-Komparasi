import React, { useState, useMemo } from 'react';
import * as XLSX from 'xlsx';
import { 
  Upload, FileSpreadsheet, ArrowRightLeft, Download, 
  Settings, Key, Sliders, Layers, CheckCircle, AlertTriangle, 
  PlusCircle, MinusCircle, Eye, CheckCircle2, AlertCircle, FilePlus, FileMinus, X, Calculator, Plus, RotateCcw, Database, Type, Binary, CheckSquare, Square, Search, Filter, Target
} from 'lucide-react';

export default function App() {
  const [file1, setFile1] = useState(null);
  const [file2, setFile2] = useState(null);

  const [sheet1, setSheet1] = useState('');
  const [sheet2, setSheet2] = useState('');

  // Mode Operasi: 'COMPARE' (Mencari Perbedaan) atau 'CALCULATE' (Hitung Rumus Saja)
  const [operationMode, setOperationMode] = useState('COMPARE');

  const [config, setConfig] = useState({
    primaryKeys: [],
    ignoreCase: true,
    trimWhitespace: true,
    numericTolerance: 0.001,
    selectedColumns: []
  });

  const [pkSearchQuery, setPkSearchQuery] = useState('');
  
  // State Formula (Hanya aktif di Mode Kalkulasi)
  const [enableFormula, setEnableFormula] = useState(false);
  const [formulaTargetColumn, setFormulaTargetColumn] = useState('');
  const [customTargetName, setCustomTargetName] = useState('');
  const [formulaExpr, setFormulaExpr] = useState('');

  const [filter, setFilter] = useState('ALL');
  const [nameSearchFilter, setNameSearchFilter] = useState('');
  const [selectedRowDetail, setSelectedRowDetail] = useState(null);
  const [selectedRowIds, setSelectedRowIds] = useState([]);

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
            const json = XLSX.utils.sheet_to_json(worksheet, { defval: "" }).filter(row => {
              return Object.values(row).some(val => val !== "" && val !== null && val !== undefined);
            });
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
    setOperationMode('COMPARE');
    setFormulaExpr('');
    setEnableFormula(false);
    setFormulaTargetColumn('');
    setCustomTargetName('');
    setConfig(prev => ({ ...prev, primaryKeys: [] }));
    setPkSearchQuery('');
    setNameSearchFilter('');
    setSelectedRowIds([]);
  };

  const commonHeaders = useMemo(() => {
    if (!file1 && !file2) return [];
    const h1 = file1 && sheet1 ? (file1.sheetsData[sheet1]?.[0] ? Object.keys(file1.sheetsData[sheet1][0]) : []) : [];
    const h2 = file2 && sheet2 ? (file2.sheetsData[sheet2]?.[0] ? Object.keys(file2.sheetsData[sheet2][0]) : []) : [];
    const merged = Array.from(new Set([...h1, ...h2]));
    return merged.filter(col => !col.startsWith('__EMPTY'));
  }, [file1, file2, sheet1, sheet2]);

  const filteredPkHeaders = useMemo(() => {
    if (!pkSearchQuery.trim()) return commonHeaders;
    return commonHeaders.filter(col => col.toLowerCase().includes(pkSearchQuery.toLowerCase()));
  }, [commonHeaders, pkSearchQuery]);

  const togglePrimaryKey = (col) => {
    setConfig(prev => {
      const exists = prev.primaryKeys.includes(col);
      if (exists) {
        return { ...prev, primaryKeys: prev.primaryKeys.filter(item => item !== col) };
      } else {
        return { ...prev, primaryKeys: [...prev.primaryKeys, col] };
      }
    });
  };

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

  const formatCellValue = (headerName, val) => {
    if (val === null || val === undefined || val === "") return "-";
    
    const lowerHeader = headerName.toLowerCase();
    
    const currencyKeywords = [
      'gaji', 'tunjangan', 'potongan', 'bonus', 'insentif', 'lembur', 'thr', 
      'bruto', 'thp', 'total', 'bayar', 'umk', 'penalty', 'bpjs', 'kompensasi', 
      'pajak', 'pph', 'nominal', 'jumlah', 'harga', 'biaya', 'rp', 'saltab', 'hasil_formula'
    ];
    const isCurrency = currencyKeywords.some(keyword => lowerHeader.includes(keyword));

    if (isCurrency && !isNaN(Number(val)) && typeof val !== 'boolean') {
      return new Intl.NumberFormat('id-ID', { style: 'currency', currency: 'IDR', maximumFractionDigits: 0 }).format(Number(val));
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
    }

    if (typeof val === 'object') return JSON.stringify(val);
    return String(val);
  };

  const comparisonResult = useMemo(() => {
    if (operationMode === 'CALCULATE') {
      const targetFile = file2 || file1;
      const targetSheet = file2 ? sheet2 : sheet1;
      if (!targetFile || !targetSheet) return null;

      const data = targetFile.sheetsData[targetSheet] || [];
      const { primaryKeys, ignoreCase, trimWhitespace } = config;

      const normalizeVal = (val) => {
        if (val === null || val === undefined) return "";
        let s = String(val);
        if (trimWhitespace) s = s.trim();
        if (ignoreCase) s = s.toLowerCase();
        return s;
      };

      const getRowKeyString = (row) => {
        if (!primaryKeys || primaryKeys.length === 0) return null;
        const keys = primaryKeys.map(k => normalizeVal(row[k]));
        if (keys.some(k => k === "")) return null;
        return keys.join(' | ');
      };

      let targetColName = "";
      if (enableFormula && formulaExpr.trim()) {
        targetColName = formulaTargetColumn === "CUSTOM" ? (customTargetName.trim() || "Kolom_Hasil_Formula") : (formulaTargetColumn || "Hasil_Formula");
      }

      let allHeaders = commonHeaders;
      if (targetColName && !allHeaders.includes(targetColName)) {
        allHeaders = [...allHeaders, targetColName];
      }

      const results = [];
      data.forEach((row, idx) => {
        const cleanRow = {};
        Object.keys(row).forEach(k => {
          if (!k.startsWith('__EMPTY')) cleanRow[k] = row[k];
        });

        if (targetColName && formulaExpr.trim()) {
          const originalCols = Object.keys(cleanRow);
          cleanRow[targetColName] = evaluateFormula(cleanRow, formulaExpr, originalCols);
        }

        const compositeKey = getRowKeyString(cleanRow) || `Baris ${idx + 1}`;
        results.push({
          id: `calc-${idx}`,
          key: compositeKey,
          status: 'CALCULATED',
          row2: cleanRow,
          diffs: {}
        });
      });

      return { results, counts: { total: results.length }, headers: allHeaders, targetColName };
    }

    if (!file1 || !file2 || !sheet1 || !sheet2) return null;

    const data1 = file1.sheetsData[sheet1] || [];
    const data2 = file2.sheetsData[sheet2] || [];
    const { primaryKeys, ignoreCase, trimWhitespace, numericTolerance, selectedColumns } = config;

    const normalizeVal = (val) => {
      if (val === null || val === undefined) return "";
      let s = String(val);
      if (trimWhitespace) s = s.trim();
      if (ignoreCase) s = s.toLowerCase();
      return s;
    };

    const getRowKeyString = (row) => {
      if (!primaryKeys || primaryKeys.length === 0) return null;
      const keys = primaryKeys.map(k => normalizeVal(row[k]));
      if (keys.some(k => k === "")) return null;
      return keys.join(' | ');
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
    const activeHeaders = selectedColumns.length > 0 
      ? allHeaders.filter(h => selectedColumns.includes(h))
      : allHeaders;

    const results = [];
    let counts = { total: 0, matched: 0, modified: 0, added: 0, deleted: 0 };

    const enrichRow = (row) => {
      if (!row) return null;
      const newRow = {};
      Object.keys(row).forEach(k => {
        if (!k.startsWith('__EMPTY')) newRow[k] = row[k];
      });
      return newRow;
    };

    const enrichedData1 = data1.map(enrichRow).filter(row => primaryKeys.length === 0 || getRowKeyString(row) !== null);
    const enrichedData2 = data2.map(enrichRow).filter(row => primaryKeys.length === 0 || getRowKeyString(row) !== null);

    if (primaryKeys && primaryKeys.length > 0) {
      const map2 = new Map();
      enrichedData2.forEach(row => {
        const compositeKey = getRowKeyString(row);
        if (compositeKey) map2.set(compositeKey, row);
      });

      const processedKeysIn2 = new Set();

      enrichedData1.forEach((row1, idx) => {
        const compositeKey = getRowKeyString(row1);
        if (!compositeKey) return;
        const row2 = map2.get(compositeKey);

        if (!row2) {
          results.push({
            id: `row-del-${idx}`,
            key: compositeKey,
            status: 'DELETED',
            diffCount: activeHeaders.length,
            row1, row2: null, diffs: {}
          });
          counts.deleted++;
        } else {
          processedKeysIn2.add(compositeKey);
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
            key: compositeKey,
            status,
            diffCount,
            row1, row2, diffs
          });
        }
      });

      enrichedData2.forEach((row2, idx) => {
        const compositeKey = getRowKeyString(row2);
        if (compositeKey && !processedKeysIn2.has(compositeKey)) {
          results.push({
            id: `row-add-${idx}`,
            key: compositeKey,
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
    return { results, counts, headers: activeHeaders, targetColName: "" };
  }, [file1, file2, sheet1, sheet2, operationMode, config, enableFormula, formulaExpr, formulaTargetColumn, customTargetName, commonHeaders]);

  const filteredResults = useMemo(() => {
    if (!comparisonResult) return [];
    let res = comparisonResult.results;

    if (operationMode === 'COMPARE' && filter !== 'ALL') {
      res = res.filter(r => r.status === filter);
    }

    if (nameSearchFilter.trim()) {
      const q = nameSearchFilter.toLowerCase();
      res = res.filter(r => {
        if (r.key.toLowerCase().includes(q)) return true;
        const matchRow1 = r.row1 && Object.values(r.row1).some(val => String(val).toLowerCase().includes(q));
        const matchRow2 = r.row2 && Object.values(r.row2).some(val => String(val).toLowerCase().includes(q));
        return matchRow1 || matchRow2;
      });
    }

    return res;
  }, [comparisonResult, operationMode, filter, nameSearchFilter]);

  const toggleRowSelection = (id) => {
    setSelectedRowIds(prev => 
      prev.includes(id) ? prev.filter(item => item !== id) : [...prev, id]
    );
  };

  const toggleSelectAll = () => {
    const allFilteredIds = filteredResults.map(r => r.id);
    const isAllSelected = allFilteredIds.every(id => selectedRowIds.includes(id));
    
    if (isAllSelected) {
      setSelectedRowIds(prev => prev.filter(id => !allFilteredIds.includes(id)));
    } else {
      const combined = Array.from(new Set([...selectedRowIds, ...allFilteredIds]));
      setSelectedRowIds(combined);
    }
  };

  const exportToExcel = () => {
    if (!comparisonResult) return;
    
    const targetData = selectedRowIds.length > 0 
      ? comparisonResult.results.filter(r => selectedRowIds.includes(r.id))
      : filteredResults;

    if (targetData.length === 0) {
      alert("Tidak ada data yang dipilih untuk diekspor!");
      return;
    }

    const exportData = targetData.map(item => {
      if (operationMode === 'CALCULATE') {
        return {
          'Composite Primary Key': item.key,
          ...(item.row2 || {})
        };
      }

      const row = {
        'Status Komparasi': item.status === 'MATCHED' ? 'Identik (Sama)' : 
                            item.status === 'MODIFIED' ? 'Berbeda (Modified)' : 
                            item.status === 'ADDED' ? 'Baru (Added)' : 'Dihapus (Deleted)',
        'Composite Primary Key': item.key,
        'Jumlah Kolom Beda': item.diffCount
      };

      comparisonResult.headers.forEach(h => {
        if (item.status === 'MATCHED') {
          row[`${h} (Data Sama)`] = item.row1 ? item.row1[h] : "";
        } else {
          row[`${h} [Lama / File 1]`] = item.row1 ? item.row1[h] : "-";
          row[`${h} [Baru / File 2]`] = item.row2 ? item.row2[h] : "-";
        }
      });
      return row;
    });

    const worksheet = XLSX.utils.json_to_sheet(exportData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, operationMode === 'CALCULATE' ? 'Hasil_Kalkulasi' : 'Rekonsiliasi');
    const suffix = operationMode === 'CALCULATE' ? '_Kalkulasi_Formula' : `_${filter}`;
    XLSX.writeFile(workbook, `Laporan${suffix}.xlsx`);
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
                <h1 className="text-lg md:text-xl font-bold tracking-tight text-white">
                  Sistem Komparasi & Kalkulasi Data Excel
                </h1>
                <p className="text-xs text-slate-300 mt-0.5">
                  Dikembangkan untuk Penelitian Skripsi Program Studi Teknik Informatika &mdash; Universitas Bina Sarana Informatika.
                </p>
              </div>
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
                1. Pilih Mode Kerja & Unggah File
              </h2>
              <p className="text-xs text-slate-500 mt-0.5">
                Pilih apakah Anda ingin murni **Mencari Perbedaan** antar file atau melakukan **Kalkulasi/Hitung Rumus**.
              </p>
            </div>
            <button onClick={resetAllData} className="text-xs font-medium text-slate-500 hover:text-rose-600 flex items-center gap-1 self-start md:self-auto px-3 py-1.5 rounded-lg border border-slate-200 hover:bg-slate-50 transition-colors">
              <RotateCcw className="w-3.5 h-3.5" />
              Reset File & Data
            </button>
          </div>

          <div className="flex justify-center my-2">
            <div className="bg-slate-100 p-1.5 rounded-2xl flex items-center gap-1.5 max-w-lg w-full border border-slate-200/80 shadow-inner">
              <button
                onClick={() => { setOperationMode('COMPARE'); setEnableFormula(false); }}
                className={`flex-1 flex items-center justify-center gap-2.5 py-3 px-4 rounded-xl text-xs font-bold transition-all duration-300 ${
                  operationMode === 'COMPARE' 
                    ? 'bg-white text-indigo-950 shadow-md border border-slate-200/60 scale-[1.02]' 
                    : 'text-slate-500 hover:text-slate-800 hover:bg-slate-200/60'
                }`}
              >
                <ArrowRightLeft className={`w-4 h-4 ${operationMode === 'COMPARE' ? 'text-indigo-600' : 'text-slate-400'}`} />
                <span>Mode Komparasi File</span>
              </button>

              <button
                onClick={() => setOperationMode('CALCULATE')}
                className={`flex-1 flex items-center justify-center gap-2.5 py-3 px-4 rounded-xl text-xs font-bold transition-all duration-300 ${
                  operationMode === 'CALCULATE' 
                    ? 'bg-white text-emerald-950 shadow-md border border-slate-200/60 scale-[1.02]' 
                    : 'text-slate-500 hover:text-slate-800 hover:bg-slate-200/60'
                }`}
              >
                <Calculator className={`w-4 h-4 ${operationMode === 'CALCULATE' ? 'text-emerald-600' : 'text-slate-400'}`} />
                <span>Mode Kalkulasi Rumus</span>
              </button>
            </div>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 gap-4 pt-2">
            <div className="border-2 border-dashed border-slate-200 hover:border-amber-400 bg-slate-50/50 rounded-2xl p-4 transition-all relative group">
              <input type="file" accept=".xlsx, .xls, .csv" onChange={(e) => e.target.files[0] && handleUploadFile1(e.target.files[0])} className="absolute inset-0 w-full h-full opacity-0 cursor-pointer z-10" />
              <div className="flex items-start gap-3">
                <div className="p-3 bg-amber-100 text-amber-700 rounded-xl flex-shrink-0">
                  <FileSpreadsheet className="w-6 h-6" />
                </div>
                <div className="flex-1 min-w-0">
                  <span className="text-[10px] font-bold text-amber-700 uppercase tracking-wider block">
                    {operationMode === 'CALCULATE' ? 'FILE 1 (OPSIONAL / ACUAN)' : 'FILE 1 (DATA REFERENSI / LAMA)'}
                  </span>
                  <p className="text-sm font-semibold text-slate-800 truncate mt-0.5">{file1 ? file1.fileName : "Pilih atau drag File 1 ke sini..."}</p>
                  
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
                  <span className="text-[10px] font-bold text-blue-700 uppercase tracking-wider block">
                    {operationMode === 'CALCULATE' ? 'FILE 2 (DATA UTAMA YANG DIHITUNG)' : 'FILE 2 (DATA PEMBANDING / BARU)'}
                  </span>
                  <p className="text-sm font-semibold text-slate-800 truncate mt-0.5">{file2 ? file2.fileName : "Pilih atau drag File 2 ke sini..."}</p>
                  
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
              2. Pengaturan Parameter & Kolom Kunci Acuan
            </h2>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-3 gap-5">
            <div className="space-y-2 md:col-span-1 flex flex-col">
              <label className="text-xs font-semibold text-slate-700 flex items-center gap-1.5">
                <Key className="w-3.5 h-3.5 text-amber-500" />
                Pilih Kolom Kunci Utama / Primary Key:
              </label>

              <div className="relative">
                <Search className="w-3.5 h-3.5 text-slate-400 absolute left-3 top-2.5" />
                <input 
                  type="text" 
                  placeholder="Cari nama kolom..." 
                  value={pkSearchQuery}
                  onChange={(e) => setPkSearchQuery(e.target.value)}
                  disabled={!file1 && !file2}
                  className="w-full bg-slate-50 text-xs text-slate-800 rounded-xl border border-slate-300 pl-8 pr-3 py-1.5 disabled:opacity-50"
                />
              </div>

              <div className="bg-slate-50 border border-slate-200 rounded-xl p-3 h-36 overflow-y-auto space-y-1.5">
                {(!file1 && !file2) ? (
                  <p className="text-[11px] text-slate-400 italic">Unggah file terlebih dahulu.</p>
                ) : filteredPkHeaders.length === 0 ? (
                  <p className="text-[11px] text-slate-400 italic">Kolom tidak ditemukan.</p>
                ) : (
                  filteredPkHeaders.map(col => {
                    const isSelected = config.primaryKeys.includes(col);
                    return (
                      <div 
                        key={col} 
                        onClick={() => togglePrimaryKey(col)}
                        className={`flex items-center justify-between p-2 rounded-lg cursor-pointer text-xs transition-colors ${isSelected ? 'bg-indigo-50 border border-indigo-200 text-indigo-900 font-semibold' : 'hover:bg-slate-100 text-slate-700'}`}
                      >
                        <span className="truncate">{col}</span>
                        {isSelected ? <CheckSquare className="w-4 h-4 text-indigo-600 flex-shrink-0 ml-2" /> : <Square className="w-4 h-4 text-slate-300 flex-shrink-0 ml-2" />}
                      </div>
                    );
                  })
                )}
              </div>
            </div>

            <div className="space-y-4 md:col-span-1">
              <div className="space-y-2">
                <label className="text-xs font-semibold text-slate-700 flex items-center gap-1.5">
                  <Type className="w-3.5 h-3.5 text-indigo-500" />
                  Aturan Pemrosesan Teks & Toleransi:
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
            </div>

            <div className="space-y-1.5 md:col-span-1 bg-indigo-50/50 border border-indigo-200/60 p-4 rounded-xl flex flex-col justify-between">
              <div>
                <span className="text-xs font-bold text-indigo-900 block mb-1">Status Mode Aktif:</span>
                <p className="text-[11px] text-indigo-800 leading-relaxed">
                  {operationMode === 'COMPARE' 
                    ? 'Mode Komparasi File: Fokus membandingkan selisih data Sama, Beda, Baru, atau Dihapus secara presisi.' 
                    : 'Mode Kalkulasi Rumus: Memproses rumus per baris dengan Primary Key sebagai identitas acuan.'}
                </p>
              </div>
            </div>
          </div>

          {operationMode === 'CALCULATE' && (
            <div className="pt-4 border-t border-slate-100 space-y-3 bg-emerald-50/40 p-4 rounded-xl border border-emerald-200">
              <div className="flex items-center justify-between">
                <label className="text-xs font-bold text-emerald-950 flex items-center gap-1.5">
                  <Target className="w-4 h-4 text-emerald-600" />
                  Modul Kalkulasi Rumus & Pengisian Kolom Target:
                </label>
                <label className="flex items-center gap-2 text-xs text-emerald-900 font-medium cursor-pointer">
                  <input type="checkbox" checked={enableFormula} onChange={(e) => setEnableFormula(e.target.checked)} className="w-4 h-4 text-emerald-600 rounded border-emerald-300" />
                  <span>Aktifkan Rumus Kalkulasi</span>
                </label>
              </div>

              {enableFormula && (
                <div className="space-y-3 animate-fade-in pt-1">
                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                    <div>
                      <label className="text-xs font-semibold text-slate-700 block mb-1">Pilih Kolom Tujuan (Target Kolom yang Ingin Diisi):</label>
                      <select 
                        value={formulaTargetColumn} 
                        onChange={(e) => setFormulaTargetColumn(e.target.value)} 
                        className="w-full bg-white text-xs border border-emerald-300 rounded-xl px-3 py-2"
                      >
                        <option value="">-- Buat Kolom Baru Otomatis (Hasil_Formula) --</option>
                        {commonHeaders.map(col => (
                          <option key={col} value={col}>Timpa / Isi Kolom: {col}</option>
                        ))}
                        <option value="CUSTOM">+ Buat Nama Kolom Baru Lainnya...</option>
                      </select>
                    </div>

                    {formulaTargetColumn === 'CUSTOM' && (
                      <div>
                        <label className="text-xs font-semibold text-slate-700 block mb-1">Ketik Nama Kolom Baru:</label>
                        <input 
                          type="text" 
                          placeholder="Contoh: Total_Take_Home_Pay" 
                          value={customTargetName} 
                          onChange={(e) => setCustomTargetName(e.target.value)} 
                          className="w-full bg-white text-xs border border-emerald-300 rounded-xl px-3 py-2"
                        />
                      </div>
                    )}
                  </div>

                  <div>
                    <label className="text-xs font-semibold text-slate-700 block mb-1">Rumus Kalkulasi (Gunakan Kurung Siku `[...]`):</label>
                    <input type="text" placeholder="Contoh: [Gaji_Pokok] + [Tunjangan] - [Potongan]" value={formulaExpr} onChange={(e) => setFormulaExpr(e.target.value)} className="w-full bg-white text-xs font-mono text-slate-800 rounded-xl border border-emerald-300 px-3.5 py-2.5 shadow-sm" />
                  </div>

                  <div>
                    <span className="text-[11px] font-semibold text-slate-600 block mb-1">Klik nama kolom untuk menyisipkan ke rumus:</span>
                    <div className="flex flex-wrap gap-1.5 max-h-24 overflow-y-auto">
                      {commonHeaders.map(col => (
                        <button key={col} onClick={() => insertColumnTagToFormula(col)} className="px-2 py-1 bg-white border border-slate-300 text-xs rounded hover:border-emerald-600 font-mono">
                          +{col}
                        </button>
                      ))}
                    </div>
                  </div>
                </div>
              )}
            </div>
          )}
        </section>

        {comparisonResult && (
          <section className="space-y-4">
            {operationMode === 'COMPARE' ? (
              <div className="grid grid-cols-2 md:grid-cols-5 gap-3">
                <div className="bg-white p-4 rounded-2xl border border-slate-200 shadow-sm flex flex-col justify-between">
                  <div className="flex items-center justify-between text-slate-500">
                    <span className="text-[11px] font-bold uppercase tracking-wider">Total Record</span>
                    <Database className="w-4 h-4 text-slate-400" />
                  </div>
                  <div className="mt-2"><div className="text-2xl font-bold text-slate-900">{comparisonResult.counts.total}</div></div>
                </div>

                <div className="bg-white p-4 rounded-2xl border border-slate-200 shadow-sm flex flex-col justify-between">
                  <div className="flex items-center justify-between text-emerald-600">
                    <span className="text-[11px] font-bold uppercase tracking-wider">Sama (Identik)</span>
                    <CheckCircle className="w-4 h-4" />
                  </div>
                  <div className="mt-2"><div className="text-2xl font-bold text-emerald-600">{comparisonResult.counts.matched}</div></div>
                </div>

                <div className="bg-white p-4 rounded-2xl border border-slate-200 shadow-sm flex flex-col justify-between">
                  <div className="flex items-center justify-between text-amber-600">
                    <span className="text-[11px] font-bold uppercase tracking-wider">Perbedaan</span>
                    <AlertTriangle className="w-4 h-4" />
                  </div>
                  <div className="mt-2"><div className="text-2xl font-bold text-amber-600">{comparisonResult.counts.modified}</div></div>
                </div>

                <div className="bg-white p-4 rounded-2xl border border-slate-200 shadow-sm flex flex-col justify-between">
                  <div className="flex items-center justify-between text-blue-600">
                    <span className="text-[11px] font-bold uppercase tracking-wider">Baru</span>
                    <PlusCircle className="w-4 h-4" />
                  </div>
                  <div className="mt-2"><div className="text-2xl font-bold text-blue-600">{comparisonResult.counts.added}</div></div>
                </div>

                <div className="bg-white p-4 rounded-2xl border border-slate-200 shadow-sm col-span-2 md:col-span-1 flex flex-col justify-between">
                  <div className="flex items-center justify-between text-rose-600">
                    <span className="text-[11px] font-bold uppercase tracking-wider">Dihapus</span>
                    <MinusCircle className="w-4 h-4" />
                  </div>
                  <div className="mt-2"><div className="text-2xl font-bold text-rose-600">{comparisonResult.counts.deleted}</div></div>
                </div>
              </div>
            ) : (
              <div className="bg-emerald-50 border border-emerald-200 p-4 rounded-2xl flex items-center justify-between">
                <div>
                  <span className="text-xs font-bold text-emerald-900 uppercase tracking-wider block">Mode Kalkulasi Selesai</span>
                  <p className="text-xs text-emerald-700 mt-0.5">Berhasil memproses perhitungan untuk total **{comparisonResult.counts.total} baris data** dengan Primary Key sebagai acuan.</p>
                </div>
                <div className="text-2xl font-bold text-emerald-700 font-mono">{comparisonResult.counts.total} Baris</div>
              </div>
            )}

            <div className="flex flex-col lg:flex-row items-center justify-between gap-3 bg-white p-3 rounded-2xl border border-slate-200 shadow-sm">
              {operationMode === 'COMPARE' ? (
                <div className="flex items-center gap-1.5 overflow-x-auto w-full lg:w-auto">
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
                      className={`px-3.5 py-1.5 rounded-xl text-xs font-semibold transition-colors whitespace-nowrap ${filter === tab.key ? 'bg-indigo-600 text-white shadow-xs' : 'text-slate-600 hover:bg-slate-100'}`}
                    >
                      {tab.label}
                    </button>
                  ))}
                </div>
              ) : (
                <div className="text-xs font-bold text-slate-700 px-2">Hasil Perhitungan Kalkulasi Data</div>
              )}

              <div className="flex flex-col sm:flex-row items-center gap-2 w-full lg:w-auto">
                <div className="relative w-full sm:w-64">
                  <Filter className="w-3.5 h-3.5 text-slate-400 absolute left-3 top-2.5" />
                  <input 
                    type="text" 
                    placeholder="Cari nama / ID tertentu..." 
                    value={nameSearchFilter}
                    onChange={(e) => setNameSearchFilter(e.target.value)}
                    className="w-full bg-slate-50 text-xs text-slate-800 rounded-xl border border-slate-300 pl-8 pr-3 py-1.5"
                  />
                  {nameSearchFilter && (
                    <button onClick={() => setNameSearchFilter('')} className="absolute right-2.5 top-2 text-slate-400 hover:text-slate-600 text-xs font-bold">×</button>
                  )}
                </div>

                <button 
                  onClick={exportToExcel} 
                  disabled={!comparisonResult || filteredResults.length === 0} 
                  className="flex items-center gap-1.5 px-3.5 py-1.5 bg-emerald-700 hover:bg-emerald-600 disabled:opacity-40 disabled:cursor-not-allowed text-white font-medium text-xs rounded-xl shadow-sm transition-all border border-emerald-500/30 whitespace-nowrap w-full sm:w-auto justify-center"
                >
                  <Download className="w-3.5 h-3.5" />
                  <span>Ekspor ({selectedRowIds.length > 0 ? `${selectedRowIds.length} Dipilih` : `Semua Data`})</span>
                </button>
              </div>
            </div>

            <div className="bg-white rounded-2xl border border-slate-200 shadow-sm overflow-hidden">
              <div className="overflow-x-auto max-h-[500px]">
                <table className="w-full text-left text-xs">
                  <thead className="bg-slate-100 text-slate-700 uppercase font-semibold sticky top-0 z-10 border-b border-slate-200">
                    <tr>
                      <th className="p-3 w-10 text-center">
                        <input 
                          type="checkbox" 
                          onChange={toggleSelectAll}
                          checked={filteredResults.length > 0 && filteredResults.every(r => selectedRowIds.includes(r.id))}
                          className="w-4 h-4 text-indigo-600 rounded border-slate-300 cursor-pointer"
                        />
                      </th>
                      {operationMode === 'COMPARE' && <th className="p-3 w-12 text-center">Aksi</th>}
                      {operationMode === 'COMPARE' && <th className="p-3 w-28">Status</th>}
                      <th className="p-3 w-48">Composite Key (Acuan)</th>
                      {comparisonResult.headers.slice(0, 3).map(h => (
                        <th key={h} className="p-3 min-w-[140px]">{h}</th>
                      ))}
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {filteredResults.length === 0 ? (
                      <tr>
                        <td colSpan={comparisonResult.headers.length + 4} className="p-8 text-center text-slate-400">
                          Tidak ada data yang cocok.
                        </td>
                      </tr>
                    ) : (
                      filteredResults.map(row => {
                        const isChecked = selectedRowIds.includes(row.id);
                        return (
                          <tr key={row.id} className={`transition-colors ${isChecked ? 'bg-indigo-50/50' : 'hover:bg-slate-50'}`}>
                            <td className="p-3 text-center">
                              <input 
                                type="checkbox" 
                                checked={isChecked}
                                onChange={() => toggleRowSelection(row.id)}
                                className="w-4 h-4 text-indigo-600 rounded border-slate-300 cursor-pointer"
                              />
                            </td>
                            {operationMode === 'COMPARE' && (
                              <td className="p-3 text-center">
                                <button
                                  onClick={() => setSelectedRowDetail(row)}
                                  className="p-1.5 text-indigo-600 hover:bg-indigo-50 rounded-lg transition-colors shadow-2xs"
                                  title="Lihat Detail"
                                >
                                  <Eye className="w-4 h-4" />
                                </button>
                              </td>
                            )}
                            {operationMode === 'COMPARE' && (
                              <td className="p-3">
                                {row.status === 'MATCHED' && <span className="inline-flex items-center gap-1 px-2.5 py-1 rounded-full bg-emerald-100 text-emerald-800"><CheckCircle2 className="w-3.5 h-3.5" /> Sesuai</span>}
                                {row.status === 'MODIFIED' && <span className="inline-flex items-center gap-1 px-2.5 py-1 rounded-full bg-amber-100 text-amber-800"><AlertCircle className="w-3.5 h-3.5" /> Beda</span>}
                                {row.status === 'ADDED' && <span className="inline-flex items-center gap-1 px-2.5 py-1 rounded-full bg-blue-100 text-blue-800"><FilePlus className="w-3.5 h-3.5" /> Baru</span>}
                                {row.status === 'DELETED' && <span className="inline-flex items-center gap-1 px-2.5 py-1 rounded-full bg-rose-100 text-rose-800"><FileMinus className="w-3.5 h-3.5" /> Hapus</span>}
                              </td>
                            )}
                            <td className="p-3 font-mono font-bold text-indigo-900 bg-slate-50">{row.key}</td>
                            {comparisonResult.headers.slice(0, 3).map(h => {
                              const val = operationMode === 'CALCULATE' 
                                ? (row.row2 ? row.row2[h] : "") 
                                : (row.status === 'MATCHED' ? (row.row1 ? row.row1[h] : "") : (row.row2 ? row.row2[h] : ""));
                              const isTargetCol = h === comparisonResult.targetColName;
                              return (
                                <td key={h} className={`p-3 font-mono ${isTargetCol ? 'bg-emerald-50 text-emerald-900 font-bold border-l-2 border-emerald-400' : 'text-slate-600'}`}>
                                  {formatCellValue(h, val)}
                                </td>
                              );
                            })}
                          </tr>
                        );
                      })
                    )}
                  </tbody>
                </table>
              </div>
            </div>
          </section>
        )}
      </main>

      {selectedRowDetail && comparisonResult && operationMode === 'COMPARE' && (
        <div className="fixed inset-0 z-50 bg-slate-900/60 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-white rounded-3xl shadow-2xl max-w-4xl w-full max-h-[90vh] flex flex-col overflow-hidden animate-fade-in border border-slate-100">
            <div className="p-5 border-b border-slate-100 flex items-center justify-between bg-slate-900 text-white">
              <div className="flex items-center gap-3">
                <div className="p-2 bg-indigo-600 rounded-xl text-white">
                  <Database className="w-5 h-5" />
                </div>
                <div>
                  <h3 className="font-bold text-base">Detail Audit Baris / Record</h3>
                  <p className="text-xs text-slate-300">Composite Primary Key: <span className="font-mono text-amber-400 font-bold">{selectedRowDetail.key}</span></p>
                </div>
              </div>
              <button onClick={() => setSelectedRowDetail(null)} className="p-2 bg-slate-800 hover:bg-slate-700 rounded-xl text-slate-300 hover:text-white transition-colors">
                <X className="w-5 h-5" />
              </button>
            </div>
            
            <div className="p-6 overflow-y-auto space-y-4 flex-1 bg-slate-50/50">
              <div className="bg-amber-50/80 border border-amber-200/80 p-3.5 rounded-2xl text-xs text-amber-900 flex items-center justify-between shadow-xs">
                <span className="font-medium">Perbandingan atribut data secara presisi berdasarkan kombinasi Composite Primary Key.</span>
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

            <div className="p-4 border-t border-slate-100 bg-white flex justify-end">
              <button onClick={() => setSelectedRowDetail(null)} className="px-5 py-2.5 bg-slate-900 hover:bg-slate-800 text-white text-xs font-semibold rounded-xl transition-colors shadow-sm">
                Tutup Jendela Detail
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Footer Bawah Profesional Skripsi (Tanpa Garis Pemisah Atas) */}
      <footer className="bg-slate-900 text-slate-400 border-t border-slate-800 mt-auto py-6">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 text-center space-y-1">
          <p className="font-medium text-slate-300 text-xs">
            &copy; {new Date().getFullYear()} <span className="text-indigo-400 font-bold">Taufiq Ismail</span>. All rights reserved.
          </p>
          <p className="text-[10px] text-slate-400">
            Dibangun dengan React.js, Vite, Tailwind CSS, & SheetJS (xlsx).
          </p>
        </div>
      </footer>
    </div>
  );
}