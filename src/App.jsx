import React, { useState, useMemo, useEffect } from 'react';
import * as XLSX from 'xlsx';
import * as Diff from 'diff';
import { 
  FileSpreadsheet, ArrowRightLeft, Download, 
  Key, Sliders, CheckCircle, AlertTriangle, 
  PlusCircle, MinusCircle, Eye, CheckCircle2, AlertCircle, FilePlus, FileMinus, Calculator, RotateCcw, Database, Type, CheckSquare, Square, Search, Filter, Target, Sparkles
} from 'lucide-react';

import Header from './components/Header';
import WordModule from './components/WordModule';
import DetailModal from './components/DetailModal';
import { readFileData, readWordData, formatCellValue, parseIndonesianNumber } from './utils/fileHelpers';

export default function App() {
  const [darkMode, setDarkMode] = useState(true);
  const [appMode, setAppMode] = useState('EXCEL');

  const [file1, setFile1] = useState(null);
  const [file2, setFile2] = useState(null);
  const [sheet1, setSheet1] = useState('');
  const [sheet2, setSheet2] = useState('');

  const [operationMode, setOperationMode] = useState('COMPARE');
  const [config, setConfig] = useState({
    primaryKeys: [],
    ignoreCase: true,
    trimWhitespace: true,
    numericTolerance: 0.001,
    selectedColumns: []
  });

  const [pkSearchQuery, setPkSearchQuery] = useState('');
  const [enableFormula, setEnableFormula] = useState(false);
  const [formulaTargetColumn, setFormulaTargetColumn] = useState('');
  const [customTargetName, setCustomTargetName] = useState('');
  const [formulaExpr, setFormulaExpr] = useState('');
  const [formulaColumnSearch, setFormulaColumnSearch] = useState('');

  const [filter, setFilter] = useState('ALL');
  const [nameSearchFilter, setNameSearchFilter] = useState('');
  const [selectedRowDetail, setSelectedRowDetail] = useState(null);
  const [selectedRowIds, setSelectedRowIds] = useState([]);

  const [wordFile1, setWordFile1] = useState(null);
  const [wordFile2, setWordFile2] = useState(null);
  const [wordText1, setWordText1] = useState('');
  const [wordText2, setWordText2] = useState('');

  const handleUploadFile1 = async (file) => {
    const fileName = file.name.toLowerCase();
    if (!fileName.endsWith('.xlsx') && !fileName.endsWith('.xls') && !fileName.endsWith('.csv')) {
      alert("⚠️ Peringatan: Format file tidak sesuai! Mohon unggah file berformat Excel (.xlsx, .xls) atau CSV.");
      return;
    }
    try {
      const res = await readFileData(file);
      setFile1({ fileName: file.name, ...res });
      setSheet1(res.sheetNames[0]);
    } catch (err) {
      alert("Gagal membaca File 1. Pastikan file Excel valid.");
    }
  };

  const handleUploadFile2 = async (file) => {
    const fileName = file.name.toLowerCase();
    if (!fileName.endsWith('.xlsx') && !fileName.endsWith('.xls') && !fileName.endsWith('.csv')) {
      alert("⚠️ Peringatan: Format file tidak sesuai! Mohon unggah file berformat Excel (.xlsx, .xls) atau CSV.");
      return;
    }
    try {
      const res = await readFileData(file);
      setFile2({ fileName: file.name, ...res });
      setSheet2(res.sheetNames[0]);
    } catch (err) {
      alert("Gagal membaca File 2. Pastikan file Excel valid.");
    }
  };

  const handleUploadWord1 = async (file) => {
    const fileName = file.name.toLowerCase();
    if (!fileName.endsWith('.docx')) {
      alert("⚠️ Peringatan: Format file tidak sesuai! Mohon unggah dokumen Word berformat .docx.");
      return;
    }
    try {
      const text = await readWordData(file);
      setWordFile1({ fileName: file.name });
      setWordText1(text);
    } catch (err) {
      alert("Gagal membaca Dokumen Word 1.");
    }
  };

  const handleUploadWord2 = async (file) => {
    const fileName = file.name.toLowerCase();
    if (!fileName.endsWith('.docx')) {
      alert("⚠️ Peringatan: Format file tidak sesuai! Mohon unggah dokumen Word berformat .docx.");
      return;
    }
    try {
      const text = await readWordData(file);
      setWordFile2({ fileName: file.name });
      setWordText2(text);
    } catch (err) {
      alert("Gagal membaca Dokumen Word 2.");
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
    setFormulaColumnSearch('');
    setConfig(prev => ({ ...prev, primaryKeys: [] }));
    setPkSearchQuery('');
    setNameSearchFilter('');
    setSelectedRowIds([]);
    setWordFile1(null);
    setWordFile2(null);
    setWordText1('');
    setWordText2('');
  };

  const commonHeaders = useMemo(() => {
    if (!file1 && !file2) return [];
    const h1 = file1 && sheet1 ? (file1.sheetsData[sheet1]?.[0] ? Object.keys(file1.sheetsData[sheet1][0]) : []) : [];
    const h2 = file2 && sheet2 ? (file2.sheetsData[sheet2]?.[0] ? Object.keys(file2.sheetsData[sheet2][0]) : []) : [];
    const merged = Array.from(new Set([...h1, ...h2]));
    return merged.filter(col => !col.startsWith('__EMPTY'));
  }, [file1, file2, sheet1, sheet2]);

  useEffect(() => {
    if (commonHeaders.length > 0 && config.primaryKeys.length === 0) {
      const pkKeywords = ['nik', 'nip', 'nrk', 'id', 'kode', 'no', 'employee', 'nomor'];
      const foundKey = commonHeaders.find(h => 
        pkKeywords.some(kw => h.toLowerCase() === kw || h.toLowerCase().includes(kw))
      );
      if (foundKey) {
        setConfig(prev => ({ ...prev, primaryKeys: [foundKey] }));
      } else if (commonHeaders.length > 0) {
        setConfig(prev => ({ ...prev, primaryKeys: [commonHeaders[0]] }));
      }
    }
  }, [commonHeaders]);

  const filteredPkHeaders = useMemo(() => {
    if (!pkSearchQuery.trim()) return commonHeaders;
    return commonHeaders.filter(col => col.toLowerCase().includes(pkSearchQuery.toLowerCase()));
  }, [commonHeaders, pkSearchQuery]);

  const filteredFormulaHeaders = useMemo(() => {
    if (!formulaColumnSearch.trim()) return commonHeaders;
    return commonHeaders.filter(col => col.toLowerCase().includes(formulaColumnSearch.toLowerCase()));
  }, [commonHeaders, formulaColumnSearch]);

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
        const val = parseIndonesianNumber(row[h]) || 0;
        const regex = new RegExp(`\\[${h}\\]`, 'g');
        evaluatedExpr = evaluatedExpr.replace(regex, val);
      });

      if (!/^[\d\+\-\*\/\.\(\)\s]+$/.test(evaluatedExpr)) {
        return "Format Rumus Tidak Valid";
      }

      const result = Function(`'use strict'; return (${evaluatedExpr})`)();
      return isNaN(result) ? 0 : Number(result.toFixed(2));
    } catch (e) {
      return "Error Rumus";
    }
  };

  const comparisonResult = useMemo(() => {
    const normalizeVal = (val) => {
      if (val === null || val === undefined) return "";
      let s = String(val);
      if (config.trimWhitespace) {
        s = s.trim().replace(/\s+/g, ' ');
      }
      if (config.ignoreCase) s = s.toLowerCase();
      return s;
    };

    if (operationMode === 'CALCULATE') {
      const targetFile = file2 || file1;
      const targetSheet = file2 ? sheet2 : sheet1;
      if (!targetFile || !targetSheet) return null;

      const data = targetFile.sheetsData[targetSheet] || [];
      const { primaryKeys } = config;

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
    const { primaryKeys, numericTolerance, selectedColumns } = config;

    const getRowKeyString = (row) => {
      if (!primaryKeys || primaryKeys.length === 0) return null;
      const keys = primaryKeys.map(k => normalizeVal(row[k]));
      if (keys.some(k => k === "")) return null;
      return keys.join(' | ');
    };

    const areValuesEqual = (val1, val2) => {
      const s1 = val1 === null || val1 === undefined ? "" : String(val1).trim();
      const s2 = val2 === null || val2 === undefined ? "" : String(val2).trim();

      const n1 = parseIndonesianNumber(val1);
      const n2 = parseIndonesianNumber(val2);
      if (!isNaN(n1) && !isNaN(n2) && s1 !== "" && s2 !== "") {
        return Math.abs(n1 - n2) <= numericTolerance;
      }

      const finalV1 = config.ignoreCase ? s1.toLowerCase() : s1;
      const finalV2 = config.ignoreCase ? s2.toLowerCase() : s2;

      return finalV1 === finalV2;
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

  const wordDiffResult = useMemo(() => {
    if (!wordText1 && !wordText2) return null;
    return Diff.diffWordsWithSpace(wordText1, wordText2);
  }, [wordText1, wordText2]);

  return (
    <div className={`min-h-screen flex flex-col font-sans antialiased transition-all duration-300 ${darkMode ? 'bg-slate-950 text-slate-100 selection:bg-indigo-500 selection:text-white' : 'bg-slate-50/80 text-slate-800 selection:bg-indigo-600 selection:text-white'}`}>
      
      <Header 
        appMode={appMode} 
        setAppMode={setAppMode} 
        darkMode={darkMode} 
        setDarkMode={setDarkMode} 
      />

      <main className="flex-1 max-w-7xl w-full mx-auto px-4 sm:px-6 lg:px-8 py-6 space-y-6">
        
        {appMode === 'EXCEL' && (
          <>
            <section className={`rounded-3xl border shadow-sm p-6 space-y-6 transition-all ${darkMode ? 'bg-slate-900/90 border-slate-800 shadow-xl shadow-black/10' : 'bg-white border-slate-200/90 shadow-sm'}`}>
              <div className={`flex flex-col md:flex-row md:items-center justify-between gap-4 pb-4 border-b ${darkMode ? 'border-slate-800' : 'border-slate-100'}`}>
                <div>
                  <h2 className="text-base font-bold flex items-center gap-2">
                    <div className="p-1.5 bg-blue-500/10 text-blue-500 rounded-xl">
                      <FileSpreadsheet className="w-4 h-4" />
                    </div>
                    Pilih Mode Operasi & Unggah Dokumen Target
                  </h2>
                  <p className={`text-xs mt-0.5 ${darkMode ? 'text-slate-400' : 'text-slate-500'}`}>
                    Unggah dokumen laporan Anda (Excel atau CSV) untuk dikomparasi atau dikalkulasi otomatis.
                  </p>
                </div>
                <button 
                  onClick={resetAllData} 
                  className={`text-xs font-semibold flex items-center gap-1.5 px-3.5 py-2 rounded-xl border transition-all ${
                    darkMode 
                      ? 'border-slate-700 hover:bg-slate-800 text-slate-300' 
                      : 'border-slate-200 hover:bg-rose-50 hover:border-rose-200 text-slate-600 hover:text-rose-600'
                  }`}
                >
                  <RotateCcw className="w-3.5 h-3.5" />
                  Reset Seluruh Data
                </button>
              </div>

              <div className="flex justify-center my-2">
                <div className={`p-1.5 rounded-2xl flex items-center gap-1.5 max-w-md w-full border shadow-inner ${darkMode ? 'bg-slate-950 border-slate-800' : 'bg-slate-100 border-slate-200/80'}`}>
                  <button
                    onClick={() => { setOperationMode('COMPARE'); setEnableFormula(false); }}
                    className={`flex-1 flex items-center justify-center gap-2 py-3 px-4 rounded-xl text-xs font-bold transition-all duration-300 ${
                      operationMode === 'COMPARE' 
                        ? (darkMode ? 'bg-slate-800 text-white shadow-md border border-slate-700' : 'bg-white text-indigo-950 shadow-md border border-slate-200/60 scale-[1.02]') 
                        : (darkMode ? 'text-slate-400 hover:text-white' : 'text-slate-500 hover:text-slate-800')
                    }`}
                  >
                    <ArrowRightLeft className={`w-4 h-4 ${operationMode === 'COMPARE' ? 'text-indigo-500' : 'text-slate-400'}`} />
                    <span>Komparasi File</span>
                  </button>

                  <button
                    onClick={() => setOperationMode('CALCULATE')}
                    className={`flex-1 flex items-center justify-center gap-2 py-3 px-4 rounded-xl text-xs font-bold transition-all duration-300 ${
                      operationMode === 'CALCULATE' 
                        ? (darkMode ? 'bg-slate-800 text-white shadow-md border border-slate-700' : 'bg-white text-emerald-950 shadow-md border border-slate-200/60 scale-[1.02]') 
                        : (darkMode ? 'text-slate-400 hover:text-white' : 'text-slate-500 hover:text-slate-800')
                    }`}
                  >
                    <Calculator className={`w-4 h-4 ${operationMode === 'CALCULATE' ? 'text-emerald-500' : 'text-slate-400'}`} />
                    <span>Kalkulasi Rumus</span>
                  </button>
                </div>
              </div>

              <div className="grid grid-cols-1 md:grid-cols-2 gap-5 pt-2">
                <div className={`border-2 border-dashed rounded-3xl p-5 transition-all relative group overflow-hidden ${darkMode ? 'border-slate-800 bg-slate-950/40 hover:border-amber-500/50' : 'border-slate-200 bg-slate-50/50 hover:border-amber-400 hover:bg-amber-50/20'}`}>
                  <input type="file" accept=".xlsx, .xls, .csv" onChange={(e) => e.target.files[0] && handleUploadFile1(e.target.files[0])} className="absolute inset-0 w-full h-full opacity-0 cursor-pointer z-10" />
                  <div className="flex items-start gap-4">
                    <div className="p-3.5 bg-amber-500/10 text-amber-500 rounded-2xl flex-shrink-0 group-hover:scale-110 transition-transform">
                      <FileSpreadsheet className="w-6 h-6" />
                    </div>
                    <div className="flex-1 min-w-0">
                      <span className="text-[10px] font-bold text-amber-500 uppercase tracking-widest block mb-1">
                        {operationMode === 'CALCULATE' ? 'FILE 1 (OPSIONAL / ACUAN)' : 'FILE 1 (DATA REFERENSI / LAMA)'}
                      </span>
                      <p className={`text-sm font-semibold truncate ${darkMode ? 'text-slate-200' : 'text-slate-800'}`}>{file1 ? file1.fileName : "Klik atau seret file Excel/CSV..."}</p>
                      
                      {file1 && (
                        <div className="mt-3 relative z-20 animate-fade-in">
                          <label className={`text-[11px] font-medium block mb-1 ${darkMode ? 'text-slate-400' : 'text-slate-600'}`}>Pilih Lembar Sheet:</label>
                          <select value={sheet1} onChange={(e) => setSheet1(e.target.value)} className={`w-full text-xs border rounded-xl px-3 py-2 font-medium ${darkMode ? 'bg-slate-900 border-slate-700 text-slate-100' : 'bg-white border-slate-300'}`}>
                            {file1.sheetNames.map(s => <option key={s} value={s}>{s}</option>)}
                          </select>
                        </div>
                      )}
                    </div>
                  </div>
                </div>

                <div className={`border-2 border-dashed rounded-3xl p-5 transition-all relative group overflow-hidden ${darkMode ? 'border-slate-800 bg-slate-950/40 hover:border-blue-500/50' : 'border-slate-200 bg-slate-50/50 hover:border-blue-400 hover:bg-blue-50/20'}`}>
                  <input type="file" accept=".xlsx, .xls, .csv" onChange={(e) => e.target.files[0] && handleUploadFile2(e.target.files[0])} className="absolute inset-0 w-full h-full opacity-0 cursor-pointer z-10" />
                  <div className="flex items-start gap-4">
                    <div className="p-3.5 bg-blue-500/10 text-blue-500 rounded-2xl flex-shrink-0 group-hover:scale-110 transition-transform">
                      <FileSpreadsheet className="w-6 h-6" />
                    </div>
                    <div className="flex-1 min-w-0">
                      <span className="text-[10px] font-bold text-blue-500 uppercase tracking-widest block mb-1">
                        {operationMode === 'CALCULATE' ? 'FILE 2 (DATA UTAMA YANG DIHITUNG)' : 'FILE 2 (DATA PEMBANDING / BARU)'}
                      </span>
                      <p className={`text-sm font-semibold truncate ${darkMode ? 'text-slate-200' : 'text-slate-800'}`}>{file2 ? file2.fileName : "Klik atau seret file Excel/CSV..."}</p>
                      
                      {file2 && (
                        <div className="mt-3 relative z-20 animate-fade-in">
                          <label className={`text-[11px] font-medium block mb-1 ${darkMode ? 'text-slate-400' : 'text-slate-600'}`}>Pilih Lembar Sheet:</label>
                          <select value={sheet2} onChange={(e) => setSheet2(e.target.value)} className={`w-full text-xs border rounded-xl px-3 py-2 font-medium ${darkMode ? 'bg-slate-900 border-slate-700 text-slate-100' : 'bg-white border-slate-300'}`}>
                            {file2.sheetNames.map(s => <option key={s} value={s}>{s}</option>)}
                          </select>
                        </div>
                      )}
                    </div>
                  </div>
                </div>
              </div>
            </section>

            <section className={`rounded-3xl border shadow-sm p-6 space-y-6 transition-all ${darkMode ? 'bg-slate-900/90 border-slate-800 shadow-xl shadow-black/10' : 'bg-white border-slate-200/90 shadow-sm'}`}>
              <div className={`flex items-center justify-between pb-4 border-b ${darkMode ? 'border-slate-800' : 'border-slate-100'}`}>
                <h2 className="text-base font-bold flex items-center gap-2">
                  <div className="p-1.5 bg-indigo-500/10 text-indigo-500 rounded-xl">
                    <Sliders className="w-4 h-4" />
                  </div>
                  Konfigurasi Primary Key & Parameter Komparasi
                </h2>
              </div>

              <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
                <div className="space-y-3 md:col-span-1 flex flex-col">
                  <label className={`text-xs font-bold flex items-center gap-1.5 ${darkMode ? 'text-slate-300' : 'text-slate-700'}`}>
                    <Key className="w-4 h-4 text-amber-500" />
                    Pilih Primary Key Utama:
                  </label>

                  <div className="relative">
                    <Search className="w-3.5 h-3.5 text-slate-400 absolute left-3.5 top-3" />
                    <input 
                      type="text" 
                      placeholder="Cari kolom kunci..." 
                      value={pkSearchQuery}
                      onChange={(e) => setPkSearchQuery(e.target.value)}
                      disabled={!file1 && !file2}
                      className={`w-full text-xs rounded-xl border pl-9 pr-3 py-2 disabled:opacity-50 ${darkMode ? 'bg-slate-950 border-slate-700 text-slate-100' : 'bg-slate-50 border-slate-300 text-slate-800'}`}
                    />
                  </div>

                  <div className={`border rounded-2xl p-3 h-40 overflow-y-auto space-y-1.5 ${darkMode ? 'bg-slate-950/60 border-slate-800' : 'bg-slate-50/80 border-slate-200'}`}>
                    {(!file1 && !file2) ? (
                      <p className="text-[11px] text-slate-400 italic p-2 text-center">Unggah file terlebih dahulu.</p>
                    ) : filteredPkHeaders.length === 0 ? (
                      <p className="text-[11px] text-slate-400 italic p-2 text-center">Kolom tidak ditemukan.</p>
                    ) : (
                      filteredPkHeaders.map(col => {
                        const isSelected = config.primaryKeys.includes(col);
                        return (
                          <div 
                            key={col} 
                            onClick={() => togglePrimaryKey(col)}
                            className={`flex items-center justify-between p-2.5 rounded-xl cursor-pointer text-xs transition-all ${
                              isSelected 
                                ? (darkMode ? 'bg-indigo-950/80 border border-indigo-700 text-indigo-300 font-bold shadow-xs' : 'bg-indigo-50 border border-indigo-200 text-indigo-900 font-bold shadow-2xs') 
                                : (darkMode ? 'hover:bg-slate-800/80 text-slate-300' : 'hover:bg-slate-100/80 text-slate-700')
                            }`}
                          >
                            <span className="truncate">{col}</span>
                            {isSelected ? <CheckSquare className="w-4 h-4 text-indigo-500 flex-shrink-0 ml-2" /> : <Square className="w-4 h-4 text-slate-500 flex-shrink-0 ml-2" />}
                          </div>
                        );
                      })
                    )}
                  </div>
                </div>

                <div className="space-y-4 md:col-span-1">
                  <div className="space-y-3">
                    <label className={`text-xs font-bold flex items-center gap-1.5 ${darkMode ? 'text-slate-300' : 'text-slate-700'}`}>
                      <Type className="w-4 h-4 text-indigo-400" />
                      Aturan Pemrosesan Teks & Toleransi:
                    </label>
                    <div className="space-y-2.5 pt-1">
                      <label className={`flex items-center gap-2.5 text-xs cursor-pointer p-2.5 rounded-xl border transition-colors ${darkMode ? 'border-slate-800 hover:bg-slate-950 text-slate-300' : 'border-slate-200/80 hover:bg-slate-50 text-slate-700'}`}>
                        <input type="checkbox" checked={config.ignoreCase} onChange={(e) => setConfig({ ...config, ignoreCase: e.target.checked })} className="w-4 h-4 text-indigo-600 rounded border-slate-600 accent-indigo-600" />
                        <span className="font-medium">Abaikan Kapitalisasi Huruf</span>
                      </label>
                      <label className={`flex items-center gap-2.5 text-xs cursor-pointer p-2.5 rounded-xl border transition-colors ${darkMode ? 'border-slate-800 hover:bg-slate-950 text-slate-300' : 'border-slate-200/80 hover:bg-slate-50 text-slate-700'}`}>
                        <input type="checkbox" checked={config.trimWhitespace} onChange={(e) => setConfig({ ...config, trimWhitespace: e.target.checked })} className="w-4 h-4 text-indigo-600 rounded border-slate-600 accent-indigo-600" />
                        <span className="font-medium">Abaikan Spasi Berlebih</span>
                      </label>
                    </div>
                  </div>
                </div>

                <div className={`space-y-2 md:col-span-1 border p-5 rounded-2xl flex flex-col justify-between shadow-2xs ${darkMode ? 'bg-indigo-950/20 border-indigo-900/50' : 'bg-indigo-50/40 border-indigo-100'}`}>
                  <div>
                    <span className={`text-xs font-bold block mb-1.5 flex items-center gap-1.5 ${darkMode ? 'text-indigo-300' : 'text-indigo-900'}`}>
                      <Sparkles className="w-3.5 h-3.5 text-indigo-500" /> Informasi Sistem:
                    </span>
                    <p className={`text-[11px] leading-relaxed ${darkMode ? 'text-indigo-200/80' : 'text-indigo-800/90'}`}>
                      {operationMode === 'COMPARE' 
                        ? 'Sistem otomatis mendeteksi baris yang identik, termodifikasi, data baru, atau data yang dihapus berdasarkan Primary Key.' 
                        : 'Sistem siap memproses rumus matematika kustom Anda dengan akurasi tinggi pada baris data terpilih.'}
                    </p>
                  </div>
                </div>
              </div>

              {operationMode === 'CALCULATE' && (
                <div className={`pt-4 border-t space-y-4 p-5 rounded-2xl border animate-fade-in ${darkMode ? 'border-slate-800 bg-emerald-950/20 border-emerald-900/40' : 'border-slate-100 bg-emerald-50/30 border-emerald-200/80'}`}>
                  <div className="flex items-center justify-between">
                    <label className={`text-xs font-bold flex items-center gap-2 ${darkMode ? 'text-emerald-300' : 'text-emerald-950'}`}>
                      <Target className="w-4 h-4 text-emerald-500" />
                      Pengaturan Rumus Kalkulasi Baris:
                    </label>
                    <label className={`flex items-center gap-2 text-xs font-bold cursor-pointer px-3 py-1.5 rounded-xl border transition-colors ${darkMode ? 'bg-emerald-950/50 border-emerald-800 text-emerald-200' : 'bg-white border-emerald-300 text-emerald-900 shadow-2xs'}`}>
                      <input type="checkbox" checked={enableFormula} onChange={(e) => setEnableFormula(e.target.checked)} className="w-4 h-4 text-emerald-600 rounded border-emerald-400 accent-emerald-600" />
                      <span>Aktifkan Kalkulasi</span>
                    </label>
                  </div>

                  {enableFormula && (
                    <div className="space-y-4 pt-1">
                      <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                        <div>
                          <label className={`text-xs font-semibold block mb-1 ${darkMode ? 'text-slate-300' : 'text-slate-700'}`}>Pilih Kolom Hasil / Target:</label>
                          <select 
                            value={formulaTargetColumn} 
                            onChange={(e) => setFormulaTargetColumn(e.target.value)} 
                            className={`w-full text-xs border rounded-xl px-3 py-2.5 font-medium ${darkMode ? 'bg-slate-950 border-emerald-800 text-slate-100' : 'bg-white border-emerald-300'}`}
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
                            <label className={`text-xs font-semibold block mb-1 ${darkMode ? 'text-slate-300' : 'text-slate-700'}`}>Nama Kolom Baru:</label>
                            <input 
                              type="text" 
                              placeholder="Contoh: Total_Take_Home_Pay" 
                              value={customTargetName} 
                              onChange={(e) => setCustomTargetName(e.target.value)} 
                              className={`w-full text-xs border rounded-xl px-3 py-2.5 font-medium ${darkMode ? 'bg-slate-950 border-emerald-800 text-slate-100' : 'bg-white border-emerald-300'}`}
                            />
                          </div>
                        )}
                      </div>

                      <div>
                        <label className={`text-xs font-semibold block mb-1 ${darkMode ? 'text-slate-300' : 'text-slate-700'}`}>Formula Hitung (Gunakan tanda Kurung Siku `[...]`):</label>
                        <input type="text" placeholder="Contoh: [Gaji_Pokok] + [Tunjangan] - [Potongan]" value={formulaExpr} onChange={(e) => setFormulaExpr(e.target.value)} className={`w-full text-xs font-mono rounded-xl border px-4 py-3 shadow-sm ${darkMode ? 'bg-slate-950 border-emerald-800 text-slate-100' : 'bg-white border-emerald-300 text-slate-800'}`} />
                      </div>

                      <div>
                        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 mb-2">
                          <span className={`text-[11px] font-semibold ${darkMode ? 'text-slate-400' : 'text-slate-600'}`}>Klik untuk menyisipkan kolom ke rumus:</span>
                          
                          <div className="relative w-full sm:w-64">
                            <Search className="w-3.5 h-3.5 text-slate-400 absolute left-3 top-2.5" />
                            <input 
                              type="text" 
                              placeholder="Cari nama kolom..." 
                              value={formulaColumnSearch}
                              onChange={(e) => setFormulaColumnSearch(e.target.value)}
                              className={`w-full text-[11px] rounded-xl border pl-8 pr-3 py-1.5 ${darkMode ? 'bg-slate-950 border-emerald-800 text-slate-100' : 'bg-white border-emerald-300 text-slate-800'}`}
                            />
                            {formulaColumnSearch && (
                              <button onClick={() => setFormulaColumnSearch('')} className="absolute right-2.5 top-1.5 text-slate-400 hover:text-slate-600 text-xs font-bold">×</button>
                            )}
                          </div>
                        </div>

                        <div className={`flex flex-wrap gap-1.5 max-h-36 overflow-y-auto p-3 border rounded-2xl ${darkMode ? 'bg-slate-950/60 border-emerald-900/50' : 'bg-white/80 border-emerald-200'}`}>
                          {filteredFormulaHeaders.length === 0 ? (
                            <p className="text-[11px] text-slate-400 italic p-1">Kolom tidak ditemukan.</p>
                          ) : (
                            filteredFormulaHeaders.map(col => (
                              <button key={col} onClick={() => insertColumnTagToFormula(col)} className={`px-3 py-1.5 border text-xs rounded-xl font-mono transition-all shadow-2xs ${darkMode ? 'bg-slate-900 border-slate-700 hover:bg-emerald-950 hover:border-emerald-700 text-slate-200' : 'bg-white border-slate-200 hover:border-emerald-600 hover:bg-emerald-50 text-slate-700'}`}>
                                +{col}
                              </button>
                            ))
                          )}
                        </div>
                      </div>
                    </div>
                  )}
                </div>
              )}
            </section>

            {comparisonResult && (
              <section className="space-y-4 animate-fade-in">
                {operationMode === 'COMPARE' ? (
                  <div className="grid grid-cols-2 md:grid-cols-5 gap-3.5">
                    <div className={`p-4 rounded-3xl border shadow-sm flex flex-col justify-between transition-transform hover:scale-[1.01] ${darkMode ? 'bg-slate-900 border-slate-800' : 'bg-white border-slate-200'}`}>
                      <div className="flex items-center justify-between text-slate-400">
                        <span className="text-[10px] font-bold uppercase tracking-widest">Total Record</span>
                        <Database className="w-4 h-4 text-slate-400" />
                      </div>
                      <div className="mt-3"><div className={`text-2xl font-black ${darkMode ? 'text-slate-100' : 'text-slate-900'}`}>{comparisonResult.counts.total}</div></div>
                    </div>

                    <div className={`p-4 rounded-3xl border shadow-sm flex flex-col justify-between transition-transform hover:scale-[1.01] ${darkMode ? 'bg-slate-900 border-slate-800' : 'bg-white border-slate-200'}`}>
                      <div className="flex items-center justify-between text-emerald-500">
                        <span className="text-[10px] font-bold uppercase tracking-widest">Identik</span>
                        <CheckCircle className="w-4 h-4" />
                      </div>
                      <div className="mt-3"><div className="text-2xl font-black text-emerald-500">{comparisonResult.counts.matched}</div></div>
                    </div>

                    <div className={`p-4 rounded-3xl border shadow-sm flex flex-col justify-between transition-transform hover:scale-[1.01] ${darkMode ? 'bg-slate-900 border-slate-800' : 'bg-white border-slate-200'}`}>
                      <div className="flex items-center justify-between text-amber-500">
                        <span className="text-[10px] font-bold uppercase tracking-widest">Perbedaan</span>
                        <AlertTriangle className="w-4 h-4" />
                      </div>
                      <div className="mt-3"><div className="text-2xl font-black text-amber-500">{comparisonResult.counts.modified}</div></div>
                    </div>

                    <div className={`p-4 rounded-3xl border shadow-sm flex flex-col justify-between transition-transform hover:scale-[1.01] ${darkMode ? 'bg-slate-900 border-slate-800' : 'bg-white border-slate-200'}`}>
                      <div className="flex items-center justify-between text-blue-500">
                        <span className="text-[10px] font-bold uppercase tracking-widest">Baru</span>
                        <PlusCircle className="w-4 h-4" />
                      </div>
                      <div className="mt-3"><div className="text-2xl font-black text-blue-500">{comparisonResult.counts.added}</div></div>
                    </div>

                    <div className={`p-4 rounded-3xl border shadow-sm col-span-2 md:col-span-1 flex flex-col justify-between transition-transform hover:scale-[1.01] ${darkMode ? 'bg-slate-900 border-slate-800' : 'bg-white border-slate-200'}`}>
                      <div className="flex items-center justify-between text-rose-500">
                        <span className="text-[10px] font-bold uppercase tracking-widest">Dihapus</span>
                        <MinusCircle className="w-4 h-4" />
                      </div>
                      <div className="mt-3"><div className="text-2xl font-black text-rose-500">{comparisonResult.counts.deleted}</div></div>
                    </div>
                  </div>
                ) : (
                  <div className={`p-5 rounded-3xl border flex items-center justify-between shadow-sm ${darkMode ? 'bg-emerald-950/30 border-emerald-900/60 text-emerald-200' : 'bg-emerald-50 border-emerald-200 text-emerald-950'}`}>
                    <div>
                      <span className="text-xs font-bold uppercase tracking-widest block mb-0.5">Kalkulasi Formula Berhasil</span>
                      <p className="text-xs">Berhasil memproses total **{comparisonResult.counts.total} baris data** secara akurat.</p>
                    </div>
                    <div className="text-2xl font-black font-mono">{comparisonResult.counts.total} Baris</div>
                  </div>
                )}

                <div className={`p-4 rounded-3xl border shadow-sm flex flex-col lg:flex-row items-center justify-between gap-4 ${darkMode ? 'bg-slate-900 border-slate-800' : 'bg-white border-slate-200'}`}>
                  {operationMode === 'COMPARE' ? (
                    <div className="flex items-center gap-1.5 overflow-x-auto w-full lg:w-auto p-1">
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
                          className={`px-4 py-2 rounded-2xl text-xs font-bold transition-all whitespace-nowrap ${
                            filter === tab.key 
                              ? 'bg-indigo-600 text-white shadow-md shadow-indigo-500/20' 
                              : (darkMode ? 'text-slate-400 hover:bg-slate-800 hover:text-white' : 'text-slate-600 hover:bg-slate-100')
                          }`}
                        >
                          {tab.label}
                        </button>
                      ))}
                    </div>
                  ) : (
                    <div className={`text-xs font-bold px-2 ${darkMode ? 'text-slate-200' : 'text-slate-700'}`}>Pratinjau Tabel Hasil Perhitungan</div>
                  )}

                  <div className="flex flex-col sm:flex-row items-center gap-2.5 w-full lg:w-auto">
                    <div className="relative w-full sm:w-72">
                      <Filter className="w-3.5 h-3.5 text-slate-400 absolute left-3.5 top-3" />
                      <input 
                        type="text" 
                        placeholder="Cari ID / Nama spesifik..." 
                        value={nameSearchFilter}
                        onChange={(e) => setNameSearchFilter(e.target.value)}
                        className={`w-full text-xs rounded-2xl border pl-9 pr-3.5 py-2.5 font-medium ${darkMode ? 'bg-slate-950 border-slate-700 text-slate-100' : 'bg-slate-50 border-slate-300 text-slate-800'}`}
                      />
                      {nameSearchFilter && (
                        <button onClick={() => setNameSearchFilter('')} className="absolute right-3 top-2.5 text-slate-400 hover:text-slate-600 text-xs font-bold">×</button>
                      )}
                    </div>

                    <button 
                      onClick={exportToExcel} 
                      disabled={!comparisonResult || filteredResults.length === 0} 
                      className="flex items-center gap-2 px-5 py-2.5 bg-emerald-600 hover:bg-emerald-500 disabled:opacity-40 disabled:cursor-not-allowed text-white font-bold text-xs rounded-2xl shadow-md shadow-emerald-600/20 transition-all whitespace-nowrap w-full sm:w-auto justify-center"
                    >
                      <Download className="w-4 h-4" />
                      <span>Ekspor Excel ({selectedRowIds.length > 0 ? `${selectedRowIds.length} Dipilih` : `Semua`})</span>
                    </button>
                  </div>
                </div>

                <div className={`rounded-3xl border shadow-sm overflow-hidden ${darkMode ? 'bg-slate-900 border-slate-800' : 'bg-white border-slate-200'}`}>
                  <div className="overflow-x-auto max-h-[520px]">
                    <table className="w-full text-left text-xs border-collapse">
                      <thead className={`uppercase font-bold tracking-wider sticky top-0 z-20 border-b ${darkMode ? 'bg-slate-950/90 text-slate-300 border-slate-800 backdrop-blur-md' : 'bg-slate-100/90 text-slate-700 border-slate-200 backdrop-blur-md'}`}>
                        <tr>
                          <th className="p-3.5 w-12 text-center">
                            <input 
                              type="checkbox" 
                              onChange={toggleSelectAll}
                              checked={filteredResults.length > 0 && filteredResults.every(r => selectedRowIds.includes(r.id))}
                              className="w-4 h-4 text-indigo-600 rounded border-slate-600 cursor-pointer accent-indigo-600"
                            />
                          </th>
                          {operationMode === 'COMPARE' && <th className="p-3.5 w-14 text-center">Audit</th>}
                          {operationMode === 'COMPARE' && <th className="p-3.5 w-32">Status</th>}
                          <th className="p-3.5 w-52">Composite Primary Key</th>
                          {comparisonResult.headers.slice(0, 3).map(h => (
                            <th key={h} className="p-3.5 min-w-[160px]">{h}</th>
                          ))}
                        </tr>
                      </thead>
                      <tbody className={`divide-y ${darkMode ? 'divide-slate-800/80' : 'divide-slate-100'}`}>
                        {filteredResults.length === 0 ? (
                          <tr>
                            <td colSpan={comparisonResult.headers.length + 4} className="p-12 text-center text-slate-400 italic">
                              Tidak ada data yang sesuai dengan kriteria filter.
                            </td>
                          </tr>
                        ) : (
                          filteredResults.map(row => {
                            const isChecked = selectedRowIds.includes(row.id);
                            return (
                              <tr key={row.id} className={`transition-colors ${isChecked ? (darkMode ? 'bg-indigo-950/60' : 'bg-indigo-50/60') : (darkMode ? 'hover:bg-slate-800/40' : 'hover:bg-slate-50/80')}`}>
                                <td className="p-3.5 text-center">
                                  <input 
                                    type="checkbox" 
                                    checked={isChecked}
                                    onChange={() => toggleRowSelection(row.id)}
                                    className="w-4 h-4 text-indigo-600 rounded border-slate-600 cursor-pointer accent-indigo-600"
                                  />
                                </td>
                                {operationMode === 'COMPARE' && (
                                  <td className="p-3.5 text-center">
                                    <button
                                      onClick={() => setSelectedRowDetail(row)}
                                      className={`p-2 rounded-xl transition-all shadow-2xs ${darkMode ? 'text-indigo-400 hover:bg-indigo-950 hover:text-indigo-300' : 'text-indigo-600 hover:bg-indigo-50 hover:text-indigo-700'}`}
                                      title="Lihat Detail Audit"
                                    >
                                      <Eye className="w-4 h-4" />
                                    </button>
                                  </td>
                                )}
                                {operationMode === 'COMPARE' && (
                                  <td className="p-3.5">
                                    {row.status === 'MATCHED' && <span className="inline-flex items-center gap-1.5 px-3 py-1 rounded-full bg-emerald-500/10 text-emerald-400 font-bold text-[11px]"><CheckCircle2 className="w-3.5 h-3.5" /> Sesuai</span>}
                                    {row.status === 'MODIFIED' && <span className="inline-flex items-center gap-1.5 px-3 py-1 rounded-full bg-amber-500/10 text-amber-400 font-bold text-[11px]"><AlertCircle className="w-3.5 h-3.5" /> Beda</span>}
                                    {row.status === 'ADDED' && <span className="inline-flex items-center gap-1.5 px-3 py-1 rounded-full bg-blue-500/10 text-blue-400 font-bold text-[11px]"><FilePlus className="w-3.5 h-3.5" /> Baru</span>}
                                    {row.status === 'DELETED' && <span className="inline-flex items-center gap-1.5 px-3 py-1 rounded-full bg-rose-500/10 text-rose-400 font-bold text-[11px]"><FileMinus className="w-3.5 h-3.5" /> Hapus</span>}
                                  </td>
                                )}
                                <td className={`p-3.5 font-mono font-bold tracking-tight ${darkMode ? 'text-indigo-400 bg-slate-950/40' : 'text-indigo-900 bg-slate-50/80'}`}>{row.key}</td>
                                {comparisonResult.headers.slice(0, 3).map(h => {
                                  const val = operationMode === 'CALCULATE' 
                                    ? (row.row2 ? row.row2[h] : "") 
                                    : (row.status === 'MATCHED' ? (row.row1 ? row.row1[h] : "") : (row.row2 ? row.row2[h] : ""));
                                  const isTargetCol = h === comparisonResult.targetColName;
                                  return (
                                    <td key={h} className={`p-3.5 font-mono ${isTargetCol ? (darkMode ? 'bg-emerald-950/40 text-emerald-300 font-bold border-l-2 border-emerald-500' : 'bg-emerald-50 text-emerald-900 font-bold border-l-2 border-emerald-400') : (darkMode ? 'text-slate-300' : 'text-slate-600')}`}>
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
          </>
        )}

        {appMode === 'WORD' && (
          <WordModule 
            darkMode={darkMode}
            resetAllData={resetAllData}
            handleUploadWord1={handleUploadWord1}
            handleUploadWord2={handleUploadWord2}
            wordFile1={wordFile1}
            wordFile2={wordFile2}
            wordText1={wordText1}
            wordText2={wordText2}
            wordDiffResult={wordDiffResult}
          />
        )}
      </main>

      <DetailModal 
        selectedRowDetail={selectedRowDetail}
        comparisonResult={comparisonResult}
        darkMode={darkMode}
        setSelectedRowDetail={setSelectedRowDetail}
      />

      <footer className={`border-t mt-auto py-6 text-center transition-colors ${darkMode ? 'bg-slate-900 border-slate-800 text-slate-400' : 'bg-slate-900 border-slate-800 text-slate-400'}`}>
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 space-y-1">
          <p className="font-semibold text-slate-300 text-xs">
            &copy; {new Date().getFullYear()} <span className="text-indigo-400 font-bold">Taufiq Ismail</span>. All rights reserved.
          </p>
          <p className="text-[10px] text-slate-400">
            Dibangun dengan React.js, Vite, Tailwind CSS, SheetJS (xlsx), & Mammoth.
          </p>
        </div>
      </footer>
    </div>
  );
}