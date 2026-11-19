import React, { useState, useMemo, useEffect } from 'react';
import * as XLSX from 'xlsx';
import {
  TableProperties,
  Download,
  Upload,
  FileSpreadsheet,
  Trash2,
  Layers,
  Plus,
  X,
  ChevronDown,
  ChevronUp,
  ArrowRight,
  ArrowLeft,
  CheckCircle2,
  AlertCircle,
  BookOpen,
  Type,
  Check,
  Info
} from 'lucide-react';
import { readFileData, formatCellValue, parseIndonesianNumber } from '../utils/fileHelpers';

const COLUMN_ALIASES = {
  'ID / Kode': {
    hint: 'Menyatukan kolom kode unik/identitas (misal: NIK, Kode Barang, No)',
    keywords: ['id', 'kode', 'kode barang', 'nik', 'nip', 'nrk', 'no', 'nomor']
  },
  'Nama / Item': {
    hint: 'Menyatukan kolom nama orang, barang, item, atau uraian',
    keywords: ['nama', 'nama barang', 'item', 'barang', 'produk', 'uraian', 'deskripsi']
  },
  'Kategori / Divisi': {
    hint: 'Menyatukan kolom pengelompokan seperti Divisi, Cabang, Client, atau Status',
    keywords: ['kategori', 'divisi', 'departemen', 'bagian', 'unit', 'cabang', 'klien', 'client', 'status']
  },
  'Jumlah / Qty': {
    hint: 'Menyatukan kolom kuantitas/jumlah unit (misal: Qty, Jumlah, Volume)',
    keywords: ['jumlah', 'qty', 'quantity', 'kuantitas', 'total unit', 'volume']
  },
  'Nilai / Nominal': {
    hint: 'Menyatukan kolom angka nilai/rupiah (misal: Total, Nominal, Harga, Biaya)',
    keywords: ['nilai', 'nominal', 'harga', 'subtotal', 'total', 'amount', 'gaji', 'biaya']
  }
};

export default function PivotModule({ darkMode }) {
  const [currentStep, setCurrentStep] = useState(1);
  const [uploadedFiles, setUploadedFiles] = useState([]);
  const [aliasMapping, setAliasMapping] = useState({});
  const [showColumnMapper, setShowColumnMapper] = useState(false);

  // Aturan Pemrosesan Teks & Toleransi
  const [ignoreCase, setIgnoreCase] = useState(true);
  const [ignoreExtraSpaces, setIgnoreExtraSpaces] = useState(true);

  // Masing-masing Row dan Column wajib ditambahkan minimal 1 kolom via tombol +
  const [rowFields, setRowFields] = useState([]);
  const [rowCandidate, setRowCandidate] = useState('');
  const [hasAddedRowViaPlus, setHasAddedRowViaPlus] = useState(false);

  const [colFields, setColFields] = useState([]);
  const [colCandidate, setColCandidate] = useState('');
  const [hasAddedColViaPlus, setHasAddedColViaPlus] = useState(false);

  const [valueField, setValueField] = useState('');
  const [aggType, setAggType] = useState('COUNT'); // 'COUNT' | 'SUM' | 'AVG'
  const [sortBy, setSortBy] = useState('TOTAL_DESC');

  // Helper pemrosesan teks pada isi data sesuai toggle toleransi
  const normalizeText = (val) => {
    if (val === null || val === undefined) return '';
    let str = String(val);

    if (ignoreExtraSpaces) {
      str = str.trim().replace(/\s+/g, ' ');
    }

    if (ignoreCase) {
      str = str.toUpperCase();
    }

    return str;
  };

  const handleAddFiles = async (fileList) => {
    const filesArray = Array.from(fileList || []);
    if (!filesArray.length) return;

    const validFiles = [];
    for (const file of filesArray) {
      const lower = file.name.toLowerCase();
      if (!lower.endsWith('.xlsx') && !lower.endsWith('.xls') && !lower.endsWith('.csv')) {
        alert(`File "${file.name}" diabaikan. Gunakan format .xlsx, .xls, atau .csv.`);
        continue;
      }
      try {
        const res = await readFileData(file);
        validFiles.push({
          id: `${file.name}-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
          fileName: file.name,
          sheetNames: res.sheetNames,
          selectedSheet: res.sheetNames[0] || '',
          sheetsData: res.sheetsData
        });
      } catch (err) {
        alert(`Gagal membaca file "${file.name}".`);
      }
    }

    if (validFiles.length > 0) {
      setUploadedFiles((prev) => [...prev, ...validFiles]);
    }
  };

  const handleChangeSheet = (fileId, newSheet) => {
    setUploadedFiles((prev) =>
      prev.map((f) => (f.id === fileId ? { ...f, selectedSheet: newSheet } : f))
    );
  };

  const handleRemoveFile = (fileId) => {
    setUploadedFiles((prev) => prev.filter((f) => f.id !== fileId));
  };

  const handleClearAllFiles = () => {
    setUploadedFiles([]);
    setAliasMapping({});
    setShowColumnMapper(false);
    setRowFields([]);
    setRowCandidate('');
    setHasAddedRowViaPlus(false);
    setColFields([]);
    setColCandidate('');
    setHasAddedColViaPlus(false);
    setValueField('');
    setCurrentStep(1);
  };

  const fileHeadersMap = useMemo(() => {
    const map = {};
    uploadedFiles.forEach((f) => {
      const rawRows = f.sheetsData?.[f.selectedSheet] || [];
      const hSet = new Set();
      rawRows.forEach((r) => {
        Object.keys(r || {}).forEach((k) => {
          if (!k.startsWith('__EMPTY')) hSet.add(k);
        });
      });
      map[f.id] = Array.from(hSet);
    });
    return map;
  }, [uploadedFiles]);

  // Auto-mapping cerdas di latar belakang agar kolom mirip otomatis menyatu
  useEffect(() => {
    if (uploadedFiles.length === 0) return;

    setAliasMapping((prev) => {
      const next = { ...prev };
      Object.keys(COLUMN_ALIASES).forEach((stdCol) => {
        const keywords = COLUMN_ALIASES[stdCol].keywords;
        next[stdCol] = next[stdCol] || {};

        uploadedFiles.forEach((f) => {
          if (next[stdCol][f.id]) return;
          const headers = fileHeadersMap[f.id] || [];
          const matched = headers.find((h) => {
            const lh = h.toLowerCase().trim().replace(/\s+/g, ' ');
            return keywords.some((kw) => lh === kw || lh.includes(kw));
          });
          if (matched) next[stdCol][f.id] = matched;
        });
      });
      return next;
    });
  }, [uploadedFiles, fileHeadersMap]);

  const handleManualAliasChange = (stdCol, fileId, sourceHeader) => {
    setAliasMapping((prev) => ({
      ...prev,
      [stdCol]: {
        ...(prev[stdCol] || {}),
        [fileId]: sourceHeader
      }
    }));
  };

  // Satukan otomatis semua file menjadi 1 dataset gabungan
  const mergedDataset = useMemo(() => {
    if (!uploadedFiles.length) return [];

    const combined = [];
    const activeStandardCols = Object.keys(COLUMN_ALIASES).filter((stdCol) => {
      const perFile = aliasMapping[stdCol] || {};
      return Object.values(perFile).some((v) => Boolean(v));
    });

    uploadedFiles.forEach((f) => {
      const rawRows = f.sheetsData?.[f.selectedSheet] || [];
      rawRows.forEach((raw, idx) => {
        const cleanRow = {
          'Sumber File': f.fileName,
          'Sheet Asal': f.selectedSheet,
          'Baris Sumber': idx + 1
        };

        Object.keys(raw || {}).forEach((k) => {
          if (!k.startsWith('__EMPTY')) {
            const v = raw[k];
            cleanRow[k] = typeof v === 'number' && !Number.isNaN(v) ? v : normalizeText(v);
          }
        });

        activeStandardCols.forEach((stdCol) => {
          const mappedCol = aliasMapping[stdCol]?.[f.id];
          if (mappedCol && mappedCol in raw) {
            const v = raw[mappedCol];
            cleanRow[stdCol] = typeof v === 'number' && !Number.isNaN(v) ? v : normalizeText(v);
          } else if (!(stdCol in cleanRow)) {
            cleanRow[stdCol] = '';
          }
        });

        combined.push(cleanRow);
      });
    });

    return combined;
  }, [uploadedFiles, aliasMapping, ignoreCase, ignoreExtraSpaces]);

  const mergedFields = useMemo(() => {
    if (!mergedDataset.length) return [];
    const set = new Set();
    mergedDataset.forEach((r) => Object.keys(r).forEach((k) => set.add(k)));
    return Array.from(set);
  }, [mergedDataset]);

  // Bersihkan pilihan yang sudah tidak ada di mergedFields & siapkan kandidat awal
  useEffect(() => {
    if (!mergedFields.length) {
      setRowFields([]);
      setColFields([]);
      setHasAddedRowViaPlus(false);
      setHasAddedColViaPlus(false);
      return;
    }

    setRowFields((prev) => {
      const valid = prev.filter((f) => mergedFields.includes(f));
      if (valid.length === 0) setHasAddedRowViaPlus(false);
      return valid;
    });

    setColFields((prev) => {
      const valid = prev.filter((f) => mergedFields.includes(f));
      if (valid.length === 0) setHasAddedColViaPlus(false);
      return valid;
    });

    if (!rowCandidate || !mergedFields.includes(rowCandidate)) {
      const preferredRow = mergedFields.find((f) =>
        ['client', 'klien', 'nama / item', 'barang', 'item', 'cabang'].includes(f.toLowerCase())
      );
      setRowCandidate(preferredRow || mergedFields[0]);
    }

    if (!colCandidate || !mergedFields.includes(colCandidate)) {
      const preferredCol = mergedFields.find((f) =>
        ['status', 'kategori / divisi', 'kategori', 'divisi', 'unit', 'departemen'].includes(f.toLowerCase())
      );
      setColCandidate(preferredCol || mergedFields[1] || mergedFields[0]);
    }

    if (!valueField || !mergedFields.includes(valueField)) {
      const numCandidate = mergedFields.find((f) =>
        ['jumlah / qty', 'nilai / nominal', 'jumlah', 'qty', 'total', 'nilai', 'nominal', 'biaya', 'tagihan'].includes(
          f.toLowerCase()
        )
      );
      setValueField(numCandidate || '');
    }
  }, [mergedFields]);

  const handleAddRowViaPlus = () => {
    const target = rowCandidate || mergedFields.find((f) => !rowFields.includes(f)) || '';
    if (!target) return;
    if (rowFields.includes(target)) {
      setHasAddedRowViaPlus(true);
      return;
    }
    const nextRows = [...rowFields, target];
    setRowFields(nextRows);
    setHasAddedRowViaPlus(true);

    const nextAvailable = mergedFields.find((f) => !nextRows.includes(f));
    setRowCandidate(nextAvailable || '');
  };

  const handleRemoveRowFieldAt = (index) => {
    setRowFields((prev) => {
      const updated = prev.filter((_, idx) => idx !== index);
      if (updated.length === 0) {
        setHasAddedRowViaPlus(false);
      }
      return updated;
    });
  };

  const handleAddColViaPlus = () => {
    const target = colCandidate || mergedFields.find((f) => !colFields.includes(f)) || '';
    if (!target) return;
    if (colFields.includes(target)) {
      setHasAddedColViaPlus(true);
      return;
    }
    const nextCols = [...colFields, target];
    setColFields(nextCols);
    setHasAddedColViaPlus(true);

    const nextAvailable = mergedFields.find((f) => !nextCols.includes(f));
    setColCandidate(nextAvailable || '');
  };

  const handleRemoveColFieldAt = (index) => {
    setColFields((prev) => {
      const updated = prev.filter((_, idx) => idx !== index);
      if (updated.length === 0) {
        setHasAddedColViaPlus(false);
      }
      return updated;
    });
  };

  const rowLabelHeader = useMemo(() => (rowFields.length ? rowFields.join(' & ') : 'Kelompok Baris'), [rowFields]);

  const pivotData = useMemo(() => {
    if (!mergedDataset.length || !rowFields.length || !colFields.length) {
      return { rowKeys: [], rowPartsMap: {}, colKeys: [], matrix: {}, rowTotals: {}, colTotals: {}, grandTotal: 0 };
    }

    const rowMap = new Map();
    const rowPartsMap = {};
    const colSet = new Set();

    mergedDataset.forEach((row) => {
      const rParts = rowFields.map((f) => normalizeText(row[f]) || '(Kosong)');
      const rKey = rParts.join(' › ');

      rowPartsMap[rKey] = rParts;

      const cKey = colFields.map((f) => normalizeText(row[f]) || '(Kosong)').join(' › ');
      colSet.add(cKey);

      let val = 1;
      if (aggType !== 'COUNT') {
        const rawNum =
          typeof row[valueField] === 'number' ? row[valueField] : parseIndonesianNumber(row[valueField]);
        val = Number.isNaN(rawNum) ? 0 : rawNum;
      }

      if (!rowMap.has(rKey)) rowMap.set(rKey, new Map());
      const cMap = rowMap.get(rKey);
      if (!cMap.has(cKey)) cMap.set(cKey, []);
      cMap.get(cKey).push(val);
    });

    const aggregate = (arr) => {
      if (!arr || !arr.length) return 0;
      if (aggType === 'COUNT') return arr.length;
      const sum = arr.reduce((a, b) => a + b, 0);
      if (aggType === 'AVG') return Number((sum / arr.length).toFixed(2));
      return Number(sum.toFixed(2));
    };

    const colKeys = Array.from(colSet);
    const rowKeys = Array.from(rowMap.keys());

    const matrix = {};
    const rowTotals = {};
    const colBuckets = {};
    let grandBucket = [];

    colKeys.forEach((c) => {
      colBuckets[c] = [];
    });

    rowKeys.forEach((r) => {
      matrix[r] = {};
      const cMap = rowMap.get(r);
      let rBucket = [];

      colKeys.forEach((c) => {
        const bucket = cMap.get(c) || [];
        matrix[r][c] = aggregate(bucket);
        if (bucket.length) {
          rBucket = rBucket.concat(bucket);
          colBuckets[c] = colBuckets[c].concat(bucket);
          grandBucket = grandBucket.concat(bucket);
        }
      });

      rowTotals[r] = aggregate(rBucket);
    });

    if (sortBy === 'TOTAL_DESC') {
      rowKeys.sort((a, b) => (rowTotals[b] || 0) - (rowTotals[a] || 0));
    } else if (sortBy === 'TOTAL_ASC') {
      rowKeys.sort((a, b) => (rowTotals[a] || 0) - (rowTotals[b] || 0));
    } else {
      rowKeys.sort((a, b) => a.localeCompare(b, 'id'));
    }

    const colTotals = {};
    colKeys.forEach((c) => {
      colTotals[c] = aggregate(colBuckets[c]);
    });

    const grandTotal = aggregate(grandBucket);

    return { rowKeys, rowPartsMap, colKeys, matrix, rowTotals, colTotals, grandTotal };
  }, [mergedDataset, rowFields, colFields, valueField, aggType, sortBy, ignoreCase, ignoreExtraSpaces]);

  const formatMetric = (val) => {
    if (aggType === 'COUNT') return Number(val || 0).toLocaleString('id-ID');
    return formatCellValue(valueField || 'Nilai', val);
  };

  const cleanExportDataRows = (rows) =>
    rows.map((r) => {
      const out = {};
      Object.keys(r || {}).forEach((k) => {
        const v = r[k];
        out[k] = typeof v === 'number' && !Number.isNaN(v) ? v : normalizeText(v);
      });
      return out;
    });

  const exportMergedOnlyToExcel = () => {
    if (!mergedDataset.length) return;
    const wb = XLSX.utils.book_new();
    const wsMerged = XLSX.utils.json_to_sheet(cleanExportDataRows(mergedDataset));
    XLSX.utils.book_append_sheet(wb, wsMerged, 'Data_Gabungan');
    XLSX.writeFile(wb, `Gabungan_Excel_${uploadedFiles.length}_File.xlsx`);
  };

  const exportPivotAndMergedToExcel = () => {
    if (!pivotData.rowKeys.length) return;

    const isCurrency =
      aggType !== 'COUNT' &&
      /nilai|nominal|harga|subtotal|total|amount|gaji|biaya|tagihan|rupiah|rp/i.test(valueField || '');

    const numFormat = aggType === 'COUNT' ? '#,##0' : isCurrency ? '"Rp"#,##0' : '#,##0.##';

    const dataRows = pivotData.rowKeys.map((rKey) => {
      const rParts = pivotData.rowPartsMap[rKey] || [rKey];
      const obj = {};

      rowFields.forEach((rf, idx) => {
        obj[rf] = normalizeText(rParts[idx] ?? '');
      });

      pivotData.colKeys.forEach((cKey) => {
        obj[cKey] = Number(pivotData.matrix[rKey][cKey] ?? 0);
      });

      if (pivotData.colKeys.length > 1) {
        obj['Total'] = Number(pivotData.rowTotals[rKey] ?? 0);
      }

      return obj;
    });

    const grandTotalObj = {};
    rowFields.forEach((rf, idx) => {
      grandTotalObj[rf] = idx === 0 ? 'Grand Total' : '';
    });

    pivotData.colKeys.forEach((cKey) => {
      grandTotalObj[cKey] = Number(pivotData.colTotals[cKey] ?? 0);
    });

    if (pivotData.colKeys.length > 1) {
      grandTotalObj['Total'] = Number(pivotData.grandTotal ?? 0);
    }

    const allHeaders = [
      ...rowFields,
      ...pivotData.colKeys,
      ...(pivotData.colKeys.length > 1 ? ['Total'] : [])
    ];

    const sheetMatrix = [
      [`LAPORAN PIVOT — ${rowLabelHeader.toUpperCase()}`],
      [
        `Perhitungan: ${aggType}${valueField && aggType !== 'COUNT' ? ` (${valueField})` : ''} | Grand Total: ${formatMetric(pivotData.grandTotal)}`
      ],
      [],
      allHeaders,
      ...dataRows.map((row) => allHeaders.map((h) => row[h] ?? '')),
      allHeaders.map((h) => grandTotalObj[h] ?? '')
    ];

    const wsPivot = XLSX.utils.aoa_to_sheet(sheetMatrix);

    const firstValueColIdx = rowFields.length;
    const lastValueColIdx = allHeaders.length - 1;
    const startDataRowIdx = 4;
    const endGrandTotalRowIdx = 4 + dataRows.length;

    for (let r = startDataRowIdx; r <= endGrandTotalRowIdx; r++) {
      for (let c = firstValueColIdx; c <= lastValueColIdx; c++) {
        const cellRef = XLSX.utils.encode_cell({ r, c });
        const cell = wsPivot[cellRef];
        if (cell && typeof cell.v === 'number') {
          cell.t = 'n';
          cell.z = numFormat;
        }
      }
    }

    wsPivot['!cols'] = allHeaders.map((h) => {
      const maxLen = Math.max(
        String(h).length,
        ...dataRows.map((r) => String(r[h] ?? '').length),
        String(grandTotalObj[h] ?? '').length
      );
      return { wch: Math.min(Math.max(maxLen + 4, 14), 42) };
    });

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, wsPivot, 'Ringkasan_Pivot');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(cleanExportDataRows(mergedDataset)), 'Data_Gabungan');

    XLSX.writeFile(wb, `Pivot_${rowLabelHeader.replace(/[^a-zA-Z0-9_-]/g, '_')}.xlsx`);
  };

  const hasData = mergedDataset.length > 0;
  // Row dan Column masing-masing wajib minimal 1 kolom yang ditambahkan lewat tombol +
  const canProceedFromStep2 =
    hasData &&
    hasAddedRowViaPlus &&
    rowFields.length >= 1 &&
    hasAddedColViaPlus &&
    colFields.length >= 1;

  const hasMetric = canProceedFromStep2 && (aggType === 'COUNT' || Boolean(valueField));
  const hasResult = pivotData.rowKeys.length > 0;

  const stepTabs = [
    { step: 1, label: 'Step 1: Unggah File', unlocked: true, done: hasData },
    { step: 2, label: 'Step 2: Baris & Kolom', unlocked: hasData, done: canProceedFromStep2 },
    { step: 3, label: 'Step 3: Cara Hitung', unlocked: canProceedFromStep2, done: hasMetric },
    { step: 4, label: 'Step 4: Tabel Pivot', unlocked: hasMetric, done: hasResult }
  ];

  const remainingRowOptions = mergedFields.filter((f) => !rowFields.includes(f));
  const remainingColOptions = mergedFields.filter((f) => !colFields.includes(f));

  return (
    <div className="space-y-6 animate-fade-in">
      {/* NAVIGASI STEP BY STEP (1 -> 2 -> 3 -> 4) */}
      <div className={`rounded-3xl border p-3 shadow-sm ${darkMode ? 'bg-slate-900/90 border-slate-800' : 'bg-white border-slate-200'}`}>
        <div className="grid grid-cols-2 lg:grid-cols-4 gap-2">
          {stepTabs.map((t) => {
            const isActive = currentStep === t.step;
            return (
              <button
                key={t.step}
                type="button"
                disabled={!t.unlocked}
                onClick={() => setCurrentStep(t.step)}
                className={`px-3.5 py-2.5 rounded-2xl text-left transition-all flex items-center justify-between gap-2 disabled:opacity-40 disabled:cursor-not-allowed ${
                  isActive
                    ? 'bg-purple-600 text-white shadow-sm'
                    : darkMode
                      ? 'bg-slate-950/60 text-slate-300 hover:bg-slate-800'
                      : 'bg-slate-50 text-slate-700 hover:bg-slate-100'
                }`}
              >
                <span className="text-xs font-bold truncate">{t.label}</span>
                {t.done && (
                  <CheckCircle2 className={`w-4 h-4 shrink-0 ${isActive ? 'text-white' : 'text-emerald-400'}`} />
                )}
              </button>
            );
          })}
        </div>
      </div>

      {/* STEP 1: UNGGAH FILE */}
      {currentStep === 1 && (
        <section className={`rounded-3xl border shadow-sm p-6 space-y-5 ${darkMode ? 'bg-slate-900/90 border-slate-800' : 'bg-white border-slate-200'}`}>
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 pb-4 border-b border-slate-800/70">
            <div>
              <span className="text-[10px] font-bold uppercase tracking-widest text-purple-400">
                STEP 1 DARI 4
              </span>
              <h2 className="text-base font-bold flex items-center gap-2 mt-0.5">
                <TableProperties className="w-4 h-4 text-purple-400" />
                Unggah & Satukan File Data
              </h2>
            </div>

            <div className="flex flex-wrap items-center gap-2">
              {mergedDataset.length > 0 && (
                <button
                  type="button"
                  onClick={exportMergedOnlyToExcel}
                  className="flex items-center gap-1.5 px-3.5 py-2 bg-indigo-600 hover:bg-indigo-500 text-white font-bold text-xs rounded-2xl shadow-sm transition-all"
                >
                  <Download className="w-3.5 h-3.5" />
                  <span>Unduh Gabungan ({mergedDataset.length} baris)</span>
                </button>
              )}

              {uploadedFiles.length > 0 && (
                <button
                  type="button"
                  onClick={handleClearAllFiles}
                  className="px-3 py-2 rounded-2xl border border-slate-700 hover:bg-rose-950/40 hover:text-rose-300 text-xs font-semibold transition-all"
                >
                  Reset
                </button>
              )}
            </div>
          </div>

          {/* PENJELASAN STEP 1 */}
          <div className={`p-4 rounded-2xl border flex items-start gap-3 text-xs leading-relaxed ${darkMode ? 'bg-purple-950/20 border-purple-500/30 text-slate-300' : 'bg-purple-50/70 border-purple-200 text-slate-700'}`}>
            <BookOpen className="w-4 h-4 text-purple-400 shrink-0 mt-0.5" />
            <div className="space-y-1">
              <p className="font-bold text-purple-300">Penjelasan Step 1:</p>
              <p>
                Tahap ini membaca file spreadsheet (.xlsx, .xls, atau .csv) langsung di browser tanpa menyimpannya ke database. Gunakan opsi toleransi teks di bawah untuk menyatukan penulisan huruf kapital dan merapikan spasi berlebih tanpa menghapus spasi antar kata.
              </p>
            </div>
          </div>

          {/* ATURAN PEMROSESAN TEKS & TOLERANSI (BERSEBELAHAN) */}
          <div className="space-y-2.5">
            <div className="flex items-center gap-2 text-xs font-bold text-slate-300">
              <Type className="w-4 h-4 text-indigo-400" />
              <span>Aturan Pemrosesan Teks & Toleransi:</span>
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
              <button
                type="button"
                onClick={() => setIgnoreCase((prev) => !prev)}
                className={`w-full px-4 py-3 rounded-2xl border text-left flex items-center gap-3 transition-all ${
                  darkMode
                    ? 'bg-slate-950/70 border-slate-800 hover:border-slate-700 text-slate-100'
                    : 'bg-slate-50 border-slate-200 hover:border-slate-300 text-slate-800'
                }`}
              >
                <span
                  className={`w-4 h-4 rounded flex items-center justify-center border shrink-0 transition-colors ${
                    ignoreCase
                      ? 'bg-indigo-600 border-indigo-600 text-white'
                      : 'border-slate-500 bg-transparent'
                  }`}
                >
                  {ignoreCase && <Check className="w-3 h-3 stroke-[3]" />}
                </span>
                <span className="text-xs font-semibold">Abaikan Kapitalisasi Huruf</span>
              </button>

              <button
                type="button"
                onClick={() => setIgnoreExtraSpaces((prev) => !prev)}
                className={`w-full px-4 py-3 rounded-2xl border text-left flex items-center gap-3 transition-all ${
                  darkMode
                    ? 'bg-slate-950/70 border-slate-800 hover:border-slate-700 text-slate-100'
                    : 'bg-slate-50 border-slate-200 hover:border-slate-300 text-slate-800'
                }`}
              >
                <span
                  className={`w-4 h-4 rounded flex items-center justify-center border shrink-0 transition-colors ${
                    ignoreExtraSpaces
                      ? 'bg-indigo-600 border-indigo-600 text-white'
                      : 'border-slate-500 bg-transparent'
                  }`}
                >
                  {ignoreExtraSpaces && <Check className="w-3 h-3 stroke-[3]" />}
                </span>
                <span className="text-xs font-semibold">Abaikan Spasi Berlebih</span>
              </button>
            </div>
          </div>

          <div className={`border-2 border-dashed rounded-3xl p-5 transition-all relative group overflow-hidden ${darkMode ? 'border-slate-800 bg-slate-950/40 hover:border-purple-500/50' : 'border-slate-200 bg-slate-50/50 hover:border-purple-400'}`}>
            <input
              type="file"
              multiple
              accept=".xlsx, .xls, .csv"
              onChange={(e) => {
                handleAddFiles(e.target.files);
                e.target.value = '';
              }}
              className="absolute inset-0 w-full h-full opacity-0 cursor-pointer z-10"
            />
            <div className="flex items-center gap-4">
              <div className="p-3 bg-purple-500/10 text-purple-400 rounded-2xl flex-shrink-0">
                <Upload className="w-6 h-6" />
              </div>
              <div>
                <p className={`text-sm font-semibold ${darkMode ? 'text-slate-200' : 'text-slate-800'}`}>
                  Klik atau seret file Excel (.xlsx, .xls) / CSV di sini...
                </p>
                <span className="text-[11px] text-slate-400 block mt-0.5">
                  {uploadedFiles.length > 0
                    ? `${uploadedFiles.length} file aktif • Total ${mergedDataset.length} baris sudah disatukan`
                    : 'Bisa pilih 1 file atau banyak file sekaligus'}
                </span>
              </div>
            </div>
          </div>

          {uploadedFiles.length > 0 && (
            <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-3">
              {uploadedFiles.map((f) => {
                const rowCount = (f.sheetsData?.[f.selectedSheet] || []).length;
                return (
                  <div
                    key={f.id}
                    className={`p-3 rounded-2xl border flex items-center justify-between gap-2 ${darkMode ? 'bg-slate-950/60 border-slate-800' : 'bg-slate-50 border-slate-200'}`}
                  >
                    <div className="flex items-center gap-2 min-w-0 flex-1">
                      <FileSpreadsheet className="w-4 h-4 text-emerald-400 flex-shrink-0" />
                      <div className="min-w-0 flex-1">
                        <p className="text-xs font-bold truncate">{f.fileName}</p>
                        <div className="flex items-center gap-2 mt-1">
                          <select
                            value={f.selectedSheet}
                            onChange={(e) => handleChangeSheet(f.id, e.target.value)}
                            className={`text-[11px] rounded-lg border px-2 py-0.5 max-w-[140px] ${darkMode ? 'bg-slate-900 border-slate-700 text-slate-200' : 'bg-white border-slate-300 text-slate-800'}`}
                          >
                            {f.sheetNames.map((s) => (
                              <option key={s} value={s}>{s}</option>
                            ))}
                          </select>
                          <span className="text-[10px] text-slate-400">{rowCount} baris</span>
                        </div>
                      </div>
                    </div>
                    <button
                      type="button"
                      onClick={() => handleRemoveFile(f.id)}
                      className="p-1.5 rounded-lg text-slate-400 hover:text-rose-400 hover:bg-rose-500/10"
                      title="Hapus file"
                    >
                      <Trash2 className="w-3.5 h-3.5" />
                    </button>
                  </div>
                );
              })}
            </div>
          )}

          {/* PENYESUAIAN KOLOM BEDA NAMA ANTAR FILE (DIPERJELAS FUNGSINYA) */}
          {uploadedFiles.length >= 2 && (
            <div className="pt-1">
              <button
                type="button"
                onClick={() => setShowColumnMapper((prev) => !prev)}
                className="flex items-center gap-1.5 text-xs font-semibold text-indigo-400 hover:text-indigo-300"
              >
                <Layers className="w-3.5 h-3.5" />
                <span>Penyesuaian Kolom Beda Nama Antar File (Opsional)</span>
                {showColumnMapper ? <ChevronUp className="w-3.5 h-3.5" /> : <ChevronDown className="w-3.5 h-3.5" />}
              </button>

              {showColumnMapper && (
                <div className={`mt-3 p-4 rounded-2xl border space-y-4 ${darkMode ? 'bg-slate-950/60 border-slate-800' : 'bg-indigo-50/40 border-indigo-200'}`}>
                  {/* Kotak penjelasan fungsi yang diperjelas */}
                  <div className={`p-3.5 rounded-xl border flex items-start gap-2.5 text-xs leading-relaxed ${darkMode ? 'bg-indigo-950/30 border-indigo-500/30 text-slate-300' : 'bg-white border-indigo-200 text-slate-700'}`}>
                    <Info className="w-4 h-4 text-indigo-400 shrink-0 mt-0.5" />
                    <div className="space-y-1">
                      <p className="font-bold text-indigo-300">
                        Fungsi Fitur: Menyatukan Kolom Berbeda Nama Menjadi Satu Kolom Acuan
                      </p>
                      <p>
                        Jika file yang diunggah memiliki nama header berbeda untuk maksud data yang sama (misalnya di File 1 tertulis <strong>Divisi</strong> dan di File 2 tertulis <strong>Bagian</strong>), pilih nama kolom dari tiap file pada kelompok acuan di bawah. Sistem sudah mendeteksi otomatis, dan Anda bisa mengubahnya manual bila ada yang belum tepat.
                      </p>
                    </div>
                  </div>

                  <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-3">
                    {Object.entries(COLUMN_ALIASES).map(([stdCol, cfg]) => (
                      <div
                        key={stdCol}
                        className={`p-3.5 rounded-xl border space-y-2.5 ${darkMode ? 'bg-slate-900 border-slate-800' : 'bg-white border-slate-200'}`}
                      >
                        <div>
                          <div className="text-xs font-bold text-indigo-400">Kelompok: {stdCol}</div>
                          <p className="text-[11px] text-slate-400 mt-0.5 leading-snug">{cfg.hint}</p>
                        </div>

                        <div className="space-y-2 pt-1 border-t border-slate-800/60">
                          {uploadedFiles.map((f) => {
                            const headers = fileHeadersMap[f.id] || [];
                            const currentVal = aliasMapping[stdCol]?.[f.id] || '';
                            return (
                              <div key={f.id} className="space-y-1 text-[11px]">
                                <div className="flex items-center justify-between gap-2">
                                  <span className="truncate font-medium text-slate-300" title={f.fileName}>
                                    File: {f.fileName}
                                  </span>
                                  <span className="text-[10px] text-emerald-400 shrink-0">
                                    {currentVal ? 'Terhubung' : 'Otomatis'}
                                  </span>
                                </div>
                                <select
                                  value={currentVal}
                                  onChange={(e) => handleManualAliasChange(stdCol, f.id, e.target.value)}
                                  className={`w-full text-[11px] rounded-lg border px-2.5 py-1.5 ${darkMode ? 'bg-slate-950 border-slate-700 text-slate-200' : 'bg-slate-50 border-slate-300 text-slate-800'}`}
                                >
                                  <option value="">-- Deteksi Otomatis --</option>
                                  {headers.map((h) => (
                                    <option key={h} value={h}>{h}</option>
                                  ))}
                                </select>
                              </div>
                            );
                          })}
                        </div>
                      </div>
                    ))}
                  </div>
                </div>
              )}
            </div>
          )}

          <div className="pt-3 border-t border-slate-800/70 flex justify-end">
            <button
              type="button"
              disabled={!hasData}
              onClick={() => setCurrentStep(2)}
              className="flex items-center gap-1.5 px-4 py-2.5 bg-purple-600 hover:bg-purple-500 disabled:opacity-40 disabled:cursor-not-allowed text-white font-bold text-xs rounded-2xl shadow-sm transition-all"
            >
              <span>Lanjut ke Step 2: Baris & Kolom</span>
              <ArrowRight className="w-4 h-4" />
            </button>
          </div>
        </section>
      )}

      {/* STEP 2: PILIH KELOMPOK ROW & COLUMN (MASING-MASING WAJIB MINIMAL 1 KOLOM VIA TOMBOL +) */}
      {currentStep === 2 && mergedFields.length > 0 && (
        <section className={`rounded-3xl border shadow-sm p-6 space-y-5 ${darkMode ? 'bg-slate-900/90 border-slate-800' : 'bg-white border-slate-200'}`}>
          <div className="pb-4 border-b border-slate-800/70">
            <span className="text-[10px] font-bold uppercase tracking-widest text-indigo-400">
              STEP 2 DARI 4
            </span>
            <h3 className="text-base font-bold mt-0.5">Menentukan Struktur Baris (Rows) & Kolom (Columns)</h3>
          </div>

          {/* PENJELASAN STEP 2 */}
          <div className={`p-4 rounded-2xl border flex items-start gap-3 text-xs leading-relaxed ${darkMode ? 'bg-indigo-950/20 border-indigo-500/30 text-slate-300' : 'bg-indigo-50/70 border-indigo-200 text-slate-700'}`}>
            <BookOpen className="w-4 h-4 text-indigo-400 shrink-0 mt-0.5" />
            <div className="space-y-1">
              <p className="font-bold text-indigo-300">Penjelasan Step 2:</p>
              <p>
                Pada tabel pivot, <strong>Rows</strong> menentukan kategori yang disusun ke bawah sebagai baris, sedangkan <strong>Columns</strong> menentukan kategori yang disebarkan ke samping sebagai kolom silang. Masing-masing <strong>Rows</strong> dan <strong>Columns</strong> wajib ditambahkan minimal satu kolom dengan memilih kolomnya lalu menekan tombol <strong>+</strong> sebelum lanjut ke Step 3.
              </p>
            </div>
          </div>

          <div className="grid grid-cols-1 lg:grid-cols-2 gap-5 items-start">
            {/* ROWS — WAJIB MINIMAL 1 VIA + ROW */}
            <div className={`p-4 rounded-2xl border space-y-3 ${darkMode ? 'bg-slate-950/50 border-slate-800' : 'bg-slate-50/70 border-slate-200'}`}>
              <div className="flex items-center justify-between">
                <label className="text-xs font-bold text-purple-400">
                  1. Kelompok Baris (Rows) — Wajib min. 1 via +
                </label>
                <span className="text-[10px] font-semibold text-slate-400">{rowFields.length} kolom</span>
              </div>

              <div className="flex items-center gap-2">
                <select
                  value={rowCandidate}
                  onChange={(e) => setRowCandidate(e.target.value)}
                  className={`flex-1 text-xs rounded-xl border px-3 py-2 font-medium ${darkMode ? 'bg-slate-900 border-slate-700 text-slate-100' : 'bg-white border-slate-300 text-slate-800'}`}
                >
                  <option value="">-- Pilih kolom baris --</option>
                  {remainingRowOptions.map((f) => (
                    <option key={f} value={f}>{f}</option>
                  ))}
                </select>
                <button
                  type="button"
                  onClick={handleAddRowViaPlus}
                  disabled={!rowCandidate && remainingRowOptions.length === 0}
                  className="flex items-center gap-1 px-3 py-2 rounded-xl bg-purple-600 hover:bg-purple-500 disabled:opacity-40 text-white text-xs font-bold shrink-0 transition-all"
                >
                  <Plus className="w-3.5 h-3.5" />
                  <span>+ Row</span>
                </button>
              </div>

              {rowFields.length > 0 ? (
                <div className="space-y-2 pt-1">
                  {rowFields.map((rf, idx) => (
                    <div
                      key={`${rf}-${idx}`}
                      className={`flex items-center justify-between gap-2 px-3 py-2 rounded-xl border text-xs ${darkMode ? 'bg-slate-900 border-slate-800 text-slate-100' : 'bg-white border-slate-200 text-slate-800'}`}
                    >
                      <div className="flex items-center gap-2 min-w-0">
                        <span className="text-[10px] font-mono font-bold px-2 py-0.5 rounded-md bg-purple-500/15 text-purple-300">
                          R{idx + 1}
                        </span>
                        <span className="font-semibold truncate">{rf}</span>
                      </div>
                      <button
                        type="button"
                        onClick={() => handleRemoveRowFieldAt(idx)}
                        className="p-1 rounded-lg text-slate-400 hover:text-rose-400"
                        title="Hapus baris"
                      >
                        <X className="w-3.5 h-3.5" />
                      </button>
                    </div>
                  ))}
                </div>
              ) : (
                <p className="text-[11px] text-amber-400 flex items-center gap-1.5">
                  <AlertCircle className="w-3.5 h-3.5 shrink-0" />
                  <span>Pilih kolom lalu klik <strong>+ Row</strong> (minimal 1 kolom).</span>
                </p>
              )}
            </div>

            {/* COLUMNS — WAJIB MINIMAL 1 VIA + COL */}
            <div className={`p-4 rounded-2xl border space-y-3 ${darkMode ? 'bg-slate-950/50 border-slate-800' : 'bg-slate-50/70 border-slate-200'}`}>
              <div className="flex items-center justify-between">
                <label className="text-xs font-bold text-indigo-400">
                  2. Pemecah Kolom (Columns) — Wajib min. 1 via +
                </label>
                <span className="text-[10px] font-semibold text-slate-400">{colFields.length} kolom</span>
              </div>

              <div className="flex items-center gap-2">
                <select
                  value={colCandidate}
                  onChange={(e) => setColCandidate(e.target.value)}
                  className={`flex-1 text-xs rounded-xl border px-3 py-2 font-medium ${darkMode ? 'bg-slate-900 border-slate-700 text-slate-100' : 'bg-white border-slate-300 text-slate-800'}`}
                >
                  <option value="">-- Pilih kolom pemecah --</option>
                  {remainingColOptions.map((f) => (
                    <option key={f} value={f}>{f}</option>
                  ))}
                </select>
                <button
                  type="button"
                  onClick={handleAddColViaPlus}
                  disabled={!colCandidate && remainingColOptions.length === 0}
                  className="flex items-center gap-1 px-3 py-2 rounded-xl bg-indigo-600 hover:bg-indigo-500 disabled:opacity-40 text-white text-xs font-bold shrink-0 transition-all"
                >
                  <Plus className="w-3.5 h-3.5" />
                  <span>+ Col</span>
                </button>
              </div>

              {colFields.length > 0 ? (
                <div className="space-y-2 pt-1">
                  {colFields.map((cf, idx) => (
                    <div
                      key={`${cf}-${idx}`}
                      className={`flex items-center justify-between gap-2 px-3 py-2 rounded-xl border text-xs ${darkMode ? 'bg-slate-900 border-slate-800 text-slate-100' : 'bg-white border-slate-200 text-slate-800'}`}
                    >
                      <div className="flex items-center gap-2 min-w-0">
                        <span className="text-[10px] font-mono font-bold px-2 py-0.5 rounded-md bg-indigo-500/15 text-indigo-300">
                          C{idx + 1}
                        </span>
                        <span className="font-semibold truncate">{cf}</span>
                      </div>
                      <button
                        type="button"
                        onClick={() => handleRemoveColFieldAt(idx)}
                        className="p-1 rounded-lg text-slate-400 hover:text-rose-400"
                        title="Hapus kolom"
                      >
                        <X className="w-3.5 h-3.5" />
                      </button>
                    </div>
                  ))}
                </div>
              ) : (
                <p className="text-[11px] text-amber-400 flex items-center gap-1.5">
                  <AlertCircle className="w-3.5 h-3.5 shrink-0" />
                  <span>Pilih kolom lalu klik <strong>+ Col</strong> (minimal 1 kolom).</span>
                </p>
              )}
            </div>
          </div>

          <div className="pt-3 border-t border-slate-800/70 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
            <button
              type="button"
              onClick={() => setCurrentStep(1)}
              className="flex items-center gap-1.5 px-4 py-2 rounded-2xl border border-slate-700 hover:bg-slate-800 text-xs font-semibold self-start"
            >
              <ArrowLeft className="w-3.5 h-3.5" />
              <span>Kembali ke Step 1</span>
            </button>

            <div className="flex flex-col sm:flex-row sm:items-center gap-2.5">
              {!canProceedFromStep2 && (
                <span className="text-[11px] text-amber-400 font-medium">
                  Tambahkan minimal 1 kolom <strong>+ Row</strong> dan 1 kolom <strong>+ Col</strong> untuk lanjut.
                </span>
              )}
              <button
                type="button"
                disabled={!canProceedFromStep2}
                onClick={() => setCurrentStep(3)}
                className="flex items-center justify-center gap-1.5 px-4 py-2.5 bg-purple-600 hover:bg-purple-500 disabled:opacity-40 disabled:cursor-not-allowed text-white font-bold text-xs rounded-2xl shadow-sm transition-all"
              >
                <span>Lanjut ke Step 3: Cara Hitung</span>
                <ArrowRight className="w-4 h-4" />
              </button>
            </div>
          </div>
        </section>
      )}

      {/* STEP 3: PILIH CARA HITUNG */}
      {currentStep === 3 && mergedFields.length > 0 && canProceedFromStep2 && (
        <section className={`rounded-3xl border shadow-sm p-6 space-y-5 ${darkMode ? 'bg-slate-900/90 border-slate-800' : 'bg-white border-slate-200'}`}>
          <div className="pb-4 border-b border-slate-800/70">
            <span className="text-[10px] font-bold uppercase tracking-widest text-amber-400">
              STEP 3 DARI 4
            </span>
            <h3 className="text-base font-bold mt-0.5">Menentukan Jenis Perhitungan Nilai</h3>
          </div>

          {/* PENJELASAN STEP 3 */}
          <div className={`p-4 rounded-2xl border flex items-start gap-3 text-xs leading-relaxed ${darkMode ? 'bg-amber-950/20 border-amber-500/30 text-slate-300' : 'bg-amber-50/70 border-amber-200 text-slate-700'}`}>
            <BookOpen className="w-4 h-4 text-amber-400 shrink-0 mt-0.5" />
            <div className="space-y-1">
              <p className="font-bold text-amber-300">Penjelasan Step 3:</p>
              <p>
                Di tahap ini Anda menentukan bagaimana isi tiap sel pertemuan antara baris dan kolom dihitung. Pilih <strong>COUNT</strong> bila ingin mengetahui jumlah catatan/baris data yang muncul, <strong>SUM</strong> bila ingin menjumlahkan nilai pada kolom angka tertentu, atau <strong>AVG</strong> bila ingin melihat nilai rata-ratanya.
              </p>
            </div>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
            <div>
              <label className="text-[11px] font-bold uppercase tracking-wider text-slate-400 block mb-1.5">
                Jenis Perhitungan
              </label>
              <select
                value={aggType}
                onChange={(e) => setAggType(e.target.value)}
                className={`w-full text-xs rounded-xl border px-3 py-2.5 font-medium ${darkMode ? 'bg-slate-950 border-slate-800 text-slate-100' : 'bg-slate-50 border-slate-300 text-slate-800'}`}
              >
                <option value="COUNT">Hitung Jumlah Data (COUNT)</option>
                <option value="SUM">Jumlahkan Angka (SUM)</option>
                <option value="AVG">Rata-rata Angka (AVG)</option>
              </select>
            </div>

            <div>
              <label className="text-[11px] font-bold uppercase tracking-wider text-slate-400 block mb-1.5">
                Kolom Angka (Jika SUM / AVG)
              </label>
              <select
                value={valueField}
                onChange={(e) => setValueField(e.target.value)}
                disabled={aggType === 'COUNT'}
                className={`w-full text-xs rounded-xl border px-3 py-2.5 font-medium disabled:opacity-40 ${darkMode ? 'bg-slate-950 border-slate-800 text-slate-100' : 'bg-slate-50 border-slate-300 text-slate-800'}`}
              >
                <option value="">-- Pilih Kolom Angka --</option>
                {mergedFields.map((f) => (
                  <option key={f} value={f}>{f}</option>
                ))}
              </select>
            </div>
          </div>

          <div className="pt-3 border-t border-slate-800/70 flex items-center justify-between">
            <button
              type="button"
              onClick={() => setCurrentStep(2)}
              className="flex items-center gap-1.5 px-4 py-2 rounded-2xl border border-slate-700 hover:bg-slate-800 text-xs font-semibold"
            >
              <ArrowLeft className="w-3.5 h-3.5" />
              <span>Kembali ke Step 2</span>
            </button>

            <button
              type="button"
              disabled={!hasMetric}
              onClick={() => setCurrentStep(4)}
              className="flex items-center gap-1.5 px-4 py-2.5 bg-emerald-600 hover:bg-emerald-500 disabled:opacity-40 text-white font-bold text-xs rounded-2xl shadow-sm transition-all"
            >
              <span>Lanjut ke Step 4: Lihat Tabel Pivot</span>
              <ArrowRight className="w-4 h-4" />
            </button>
          </div>
        </section>
      )}

      {/* STEP 4: HASIL TABEL PIVOT & EKSPOR */}
      {currentStep === 4 && pivotData.rowKeys.length > 0 && canProceedFromStep2 && (
        <section className={`rounded-3xl border shadow-sm overflow-hidden ${darkMode ? 'bg-slate-900 border-slate-800' : 'bg-white border-slate-200'}`}>
          <div className="px-6 py-4 border-b border-slate-800/70 space-y-3">
            <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-4">
              <div>
                <span className="text-[10px] font-bold uppercase tracking-widest text-emerald-400">
                  STEP 4 DARI 4 • HASIL TABEL PIVOT
                </span>
                <h3 className="text-sm font-bold mt-0.5">Ringkasan Pivot ({rowLabelHeader})</h3>
                <span className="text-xs font-mono font-bold text-emerald-400 block mt-0.5">
                  Grand Total: {formatMetric(pivotData.grandTotal)}
                </span>
              </div>

              <div className="flex flex-col sm:flex-row items-stretch sm:items-center gap-2.5">
                <select
                  value={sortBy}
                  onChange={(e) => setSortBy(e.target.value)}
                  className={`text-xs rounded-xl border px-3 py-2 font-medium ${darkMode ? 'bg-slate-950 border-slate-800 text-slate-100' : 'bg-slate-50 border-slate-300 text-slate-800'}`}
                >
                  <option value="TOTAL_DESC">Terbesar ke Terkecil</option>
                  <option value="TOTAL_ASC">Terkecil ke Terbesar</option>
                  <option value="LABEL_ASC">Urut Nama (A-Z)</option>
                </select>

                <button
                  type="button"
                  onClick={exportPivotAndMergedToExcel}
                  className="flex items-center justify-center gap-1.5 px-4 py-2 bg-emerald-600 hover:bg-emerald-500 text-white font-bold text-xs rounded-2xl shadow-sm transition-all"
                >
                  <Download className="w-3.5 h-3.5" />
                  <span>Ekspor Excel (.xlsx)</span>
                </button>
              </div>
            </div>

            {/* PENJELASAN STEP 4 */}
            <div className={`p-3.5 rounded-2xl border flex items-start gap-3 text-xs leading-relaxed ${darkMode ? 'bg-emerald-950/20 border-emerald-500/30 text-slate-300' : 'bg-emerald-50/70 border-emerald-200 text-slate-700'}`}>
              <BookOpen className="w-4 h-4 text-emerald-400 shrink-0 mt-0.5" />
              <p>
                <strong>Penjelasan Step 4:</strong> Tabel di bawah menampilkan rekapitulasi akhir berdasarkan baris, kolom silang, dan metode hitung yang telah dipilih, lengkap dengan <strong>Grand Total</strong> di bagian bawah. Anda dapat mengubah urutan baris secara langsung, lalu mengekspor hasilnya ke file <strong>.xlsx</strong> yang memuat sheet <em>Ringkasan_Pivot</em> dan <em>Data_Gabungan</em>.
              </p>
            </div>
          </div>

          <div className="overflow-x-auto max-h-[480px]">
            <table className="w-full text-left text-xs border-collapse">
              <thead className={`uppercase font-bold tracking-wider sticky top-0 z-10 border-b ${darkMode ? 'bg-slate-950 text-slate-300 border-slate-800' : 'bg-slate-100 text-slate-700 border-slate-200'}`}>
                <tr>
                  {rowFields.map((rf) => (
                    <th key={rf} className="p-3.5">{rf}</th>
                  ))}
                  {pivotData.colKeys.map((cKey) => (
                    <th key={cKey} className="p-3.5 text-right">{cKey}</th>
                  ))}
                  {pivotData.colKeys.length > 1 && (
                    <th className="p-3.5 text-right">Total Baris</th>
                  )}
                </tr>
              </thead>
              <tbody className={`divide-y ${darkMode ? 'divide-slate-800/80' : 'divide-slate-100'}`}>
                {pivotData.rowKeys.map((rKey) => {
                  const parts = pivotData.rowPartsMap[rKey] || [rKey];
                  return (
                    <tr key={rKey} className={darkMode ? 'hover:bg-slate-800/40' : 'hover:bg-slate-50'}>
                      {rowFields.map((rf, idx) => (
                        <td key={`${rKey}-${rf}-${idx}`} className="p-3.5 font-mono text-slate-300">
                          {parts[idx] ?? ''}
                        </td>
                      ))}
                      {pivotData.colKeys.map((cKey) => (
                        <td key={cKey} className="p-3.5 text-right font-mono">
                          {formatMetric(pivotData.matrix[rKey][cKey])}
                        </td>
                      ))}
                      {pivotData.colKeys.length > 1 && (
                        <td className="p-3.5 text-right font-mono font-bold text-indigo-400">
                          {formatMetric(pivotData.rowTotals[rKey])}
                        </td>
                      )}
                    </tr>
                  );
                })}
              </tbody>
              <tfoot className={`font-bold border-t-2 sticky bottom-0 z-10 ${darkMode ? 'bg-slate-950 text-slate-100 border-slate-700' : 'bg-slate-100 text-slate-900 border-slate-300'}`}>
                <tr>
                  <td colSpan={rowFields.length} className="p-3.5">
                    Grand Total
                  </td>
                  {pivotData.colKeys.map((cKey) => (
                    <td key={cKey} className="p-3.5 text-right font-mono text-emerald-400">
                      {formatMetric(pivotData.colTotals[cKey])}
                    </td>
                  ))}
                  {pivotData.colKeys.length > 1 && (
                    <td className="p-3.5 text-right font-mono text-emerald-400">
                      {formatMetric(pivotData.grandTotal)}
                    </td>
                  )}
                </tr>
              </tfoot>
            </table>
          </div>

          <div className="px-6 py-4 border-t border-slate-800/70 flex items-center justify-between">
            <button
              type="button"
              onClick={() => setCurrentStep(3)}
              className="flex items-center gap-1.5 px-4 py-2 rounded-2xl border border-slate-700 hover:bg-slate-800 text-xs font-semibold"
            >
              <ArrowLeft className="w-3.5 h-3.5" />
              <span>Ubah Cara Hitung (Step 3)</span>
            </button>

            <button
              type="button"
              onClick={() => setCurrentStep(2)}
              className="px-4 py-2 rounded-2xl border border-slate-700 hover:bg-slate-800 text-xs font-semibold text-purple-300"
            >
              Ubah Baris / Kolom (Step 2)
            </button>
          </div>
        </section>
      )}
    </div>
  );
}