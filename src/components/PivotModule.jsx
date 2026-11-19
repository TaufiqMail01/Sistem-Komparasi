import React, { useState, useMemo, useEffect, useRef } from 'react';
import * as XLSX from 'xlsx';
import {
  TableProperties,
  BarChart3,
  LayoutGrid,
  Download,
  Upload,
  FileSpreadsheet,
  Search,
  Trash2,
  Layers,
  GitMerge,
  Plus,
  X
} from 'lucide-react';
import { readFileData, formatCellValue, parseIndonesianNumber } from '../utils/fileHelpers';

// Sinonim umum untuk membantu auto-mapping kolom berbeda antar file
const COLUMN_ALIASES = {
  'ID / Kode': ['id', 'kode', 'kode barang', 'nik', 'nip', 'nrk', 'no', 'nomor'],
  'Nama / Item': ['nama', 'nama barang', 'item', 'barang', 'produk', 'uraian', 'deskripsi'],
  'Kategori / Divisi': ['kategori', 'divisi', 'departemen', 'bagian', 'unit', 'cabang', 'klien', 'client', 'status'],
  'Jumlah / Qty': ['jumlah', 'qty', 'quantity', 'kuantitas', 'total unit', 'volume'],
  'Nilai / Nominal': ['nilai', 'nominal', 'harga', 'subtotal', 'total', 'amount', 'gaji', 'biaya']
};

export default function PivotModule({ darkMode }) {
  // Daftar file yang diunggah: [{ id, fileName, sheetNames, selectedSheet, sheetsData }]
  const [uploadedFiles, setUploadedFiles] = useState([]);

  // Pemetaan kolom opsional: { [kolomStandar]: { [fileId]: namaKolomAsliDiFile } }
  const [aliasMapping, setAliasMapping] = useState({});
  const [enableAutoAlias, setEnableAutoAlias] = useState(true);

  // Multi-row & Multi-column (keduanya bisa lebih dari 1 kolom)
  const [rowFields, setRowFields] = useState([]);
  const [extraRowCandidate, setExtraRowCandidate] = useState('');

  const [colFields, setColFields] = useState([]);
  const [extraColCandidate, setExtraColCandidate] = useState('');

  const [valueField, setValueField] = useState('');
  const [aggType, setAggType] = useState('COUNT'); // 'COUNT' | 'SUM' | 'AVG'
  const [viewMode, setViewMode] = useState('BOTH'); // 'BOTH' | 'TABLE' | 'CHART'
  const [searchRow, setSearchRow] = useState('');
  const [sortBy, setSortBy] = useState('TOTAL_DESC'); // 'TOTAL_DESC' | 'TOTAL_ASC' | 'LABEL_ASC'

  const chartSvgRef = useRef(null);

  // Unggah 1 atau beberapa file sekaligus
  const handleAddFiles = async (fileList) => {
    const filesArray = Array.from(fileList || []);
    if (!filesArray.length) return;

    const validFiles = [];
    for (const file of filesArray) {
      const lower = file.name.toLowerCase();
      if (!lower.endsWith('.xlsx') && !lower.endsWith('.xls') && !lower.endsWith('.csv')) {
        alert(`⚠️ File "${file.name}" diabaikan karena bukan format .xlsx, .xls, atau .csv.`);
        continue;
      }
      try {
        const res = await readFileData(file);
        const firstSheet = res.sheetNames[0] || '';
        validFiles.push({
          id: `${file.name}-${Date.now()}-${Math.random().toString(36).slice(2, 7)}`,
          fileName: file.name,
          sheetNames: res.sheetNames,
          selectedSheet: firstSheet,
          sheetsData: res.sheetsData
        });
      } catch (err) {
        alert(`Gagal membaca file "${file.name}". Pastikan file valid.`);
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
    setRowFields([]);
    setExtraRowCandidate('');
    setColFields([]);
    setExtraColCandidate('');
    setValueField('');
  };

  // Header asli per file pada sheet yang dipilih
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

  // Sinkronkan auto-alias saat file/sheet berubah
  useEffect(() => {
    if (!enableAutoAlias || uploadedFiles.length === 0) return;

    setAliasMapping((prev) => {
      const next = { ...prev };
      Object.keys(COLUMN_ALIASES).forEach((stdCol) => {
        const keywords = COLUMN_ALIASES[stdCol];
        next[stdCol] = next[stdCol] || {};

        uploadedFiles.forEach((f) => {
          if (next[stdCol][f.id]) return;
          const headers = fileHeadersMap[f.id] || [];
          const matched = headers.find((h) => {
            const lh = h.toLowerCase().trim();
            return keywords.some((kw) => lh === kw || lh.includes(kw));
          });
          if (matched) {
            next[stdCol][f.id] = matched;
          }
        });
      });
      return next;
    });
  }, [uploadedFiles, fileHeadersMap, enableAutoAlias]);

  const handleManualAliasChange = (stdCol, fileId, sourceHeader) => {
    setAliasMapping((prev) => ({
      ...prev,
      [stdCol]: {
        ...(prev[stdCol] || {}),
        [fileId]: sourceHeader
      }
    }));
  };

  // TAHAP PENYATUAN OTOMATIS: Gabungkan semua file menjadi 1 dataset terpadu
  const mergedDataset = useMemo(() => {
    if (!uploadedFiles.length) return [];

    const combined = [];
    const activeStandardCols = enableAutoAlias
      ? Object.keys(COLUMN_ALIASES).filter((stdCol) => {
          const perFile = aliasMapping[stdCol] || {};
          return Object.values(perFile).some((v) => Boolean(v));
        })
      : [];

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
            cleanRow[k] = raw[k];
          }
        });

        if (enableAutoAlias) {
          activeStandardCols.forEach((stdCol) => {
            const mappedCol = aliasMapping[stdCol]?.[f.id];
            if (mappedCol) {
              cleanRow[stdCol] = raw[mappedCol] ?? '';
            } else if (!(stdCol in cleanRow)) {
              cleanRow[stdCol] = '';
            }
          });
        }

        combined.push(cleanRow);
      });
    });

    return combined;
  }, [uploadedFiles, aliasMapping, enableAutoAlias]);

  // Unduh khusus file Excel hasil penyatuan multi-file (1 dataset gabungan)
  const exportMergedOnlyToExcel = () => {
    if (!mergedDataset.length) {
      alert('Belum ada data gabungan yang bisa diunduh.');
      return;
    }

    const wb = XLSX.utils.book_new();
    const wsMerged = XLSX.utils.json_to_sheet(mergedDataset);
    XLSX.utils.book_append_sheet(wb, wsMerged, 'Data_Gabungan');

    const timestamp = new Date().toISOString().slice(0, 10);
    XLSX.writeFile(wb, `Gabungan_Excel_${uploadedFiles.length}_File_${timestamp}.xlsx`);
  };

  // Daftar kolom tersedia dari dataset yang sudah disatukan
  const mergedFields = useMemo(() => {
    if (!mergedDataset.length) return [];
    const set = new Set();
    mergedDataset.forEach((r) => Object.keys(r).forEach((k) => set.add(k)));
    return Array.from(set);
  }, [mergedDataset]);

  useEffect(() => {
    if (!mergedFields.length) return;

    setRowFields((prev) => {
      const validExisting = prev.filter((f) => mergedFields.includes(f));
      if (validExisting.length > 0) return validExisting;

      const preferred = mergedFields.find((f) =>
        ['client', 'klien', 'kategori / divisi', 'nama / item', 'sumber file', 'divisi', 'kategori', 'status', 'cabang'].includes(
          f.toLowerCase()
        )
      );
      return [preferred || mergedFields[0]];
    });

    setColFields((prev) => prev.filter((f) => mergedFields.includes(f)));

    if (valueField && !mergedFields.includes(valueField)) {
      const numCandidate = mergedFields.find((f) =>
        ['jumlah / qty', 'nilai / nominal', 'jumlah', 'qty', 'total', 'nilai'].includes(f.toLowerCase())
      );
      setValueField(numCandidate || '');
    }
  }, [mergedFields]);

  // Helper Multi-Row
  const handleUpdateRowFieldAt = (index, nextCol) => {
    setRowFields((prev) => {
      const copy = [...prev];
      copy[index] = nextCol;
      return Array.from(new Set(copy.filter(Boolean)));
    });
  };

  const handleAddExtraRowField = () => {
    const candidate = extraRowCandidate || mergedFields.find((f) => !rowFields.includes(f)) || '';
    if (!candidate || rowFields.includes(candidate)) return;
    setRowFields((prev) => [...prev, candidate]);
    setExtraRowCandidate('');
  };

  const handleRemoveRowFieldAt = (index) => {
    setRowFields((prev) => {
      if (prev.length <= 1) return prev;
      return prev.filter((_, idx) => idx !== index);
    });
  };

  // Helper Multi-Column
  const handleUpdateColFieldAt = (index, nextCol) => {
    if (!nextCol) {
      handleRemoveColFieldAt(index);
      return;
    }
    setColFields((prev) => {
      const copy = [...prev];
      copy[index] = nextCol;
      return Array.from(new Set(copy.filter(Boolean)));
    });
  };

  const handleAddExtraColField = () => {
    const candidate =
      extraColCandidate ||
      mergedFields.find((f) => !colFields.includes(f) && !rowFields.includes(f)) ||
      mergedFields.find((f) => !colFields.includes(f)) ||
      '';
    if (!candidate || colFields.includes(candidate)) return;
    setColFields((prev) => [...prev, candidate]);
    setExtraColCandidate('');
  };

  const handleRemoveColFieldAt = (index) => {
    setColFields((prev) => prev.filter((_, idx) => idx !== index));
  };

  const rowLabelHeader = useMemo(() => {
    if (!rowFields.length) return 'Kelompok Baris';
    return rowFields.join(' › ');
  }, [rowFields]);

  const colLabelHeader = useMemo(() => {
    if (!colFields.length) return '';
    return colFields.join(' › ');
  }, [colFields]);

  // Kalkulasi Pivot Table dari mergedDataset (mendukung multi-row & multi-column)
  const pivotData = useMemo(() => {
    if (!mergedDataset.length || !rowFields.length) {
      return { rowKeys: [], rowPartsMap: {}, colKeys: [], matrix: {}, rowTotals: {}, colTotals: {}, grandTotal: 0 };
    }

    const rowMap = new Map();
    const rowPartsMap = {};
    const colSet = new Set();

    mergedDataset.forEach((row) => {
      const rParts = rowFields.map((f) => String(row[f] ?? '').trim() || '(Kosong)');
      const rKey = rParts.join(' › ');

      if (searchRow.trim() && !rKey.toLowerCase().includes(searchRow.trim().toLowerCase())) {
        return;
      }

      rowPartsMap[rKey] = rParts;

      const cKey = colFields.length > 0
        ? colFields.map((f) => String(row[f] ?? '').trim() || '(Kosong)').join(' › ')
        : 'Nilai';

      if (colFields.length > 0) colSet.add(cKey);

      let val = 1;
      if (aggType !== 'COUNT') {
        const parsed = valueField ? parseIndonesianNumber(row[valueField]) : NaN;
        val = isNaN(parsed) ? 0 : parsed;
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

    const colKeys = colFields.length > 0 ? Array.from(colSet) : ['Nilai'];
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
        const bucket = colFields.length > 0 ? (cMap.get(c) || []) : (cMap.get('Nilai') || []);
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
  }, [mergedDataset, rowFields, colFields, valueField, aggType, searchRow, sortBy]);

  const formatMetric = (val) => {
    if (aggType === 'COUNT') return Number(val || 0).toLocaleString('id-ID');
    return formatCellValue(valueField || 'Nilai', val);
  };

  // Ekspor Pivot (multi-row & multi-column) + Data Gabungan ke 1 file Excel (.xlsx)
  const exportPivotAndMergedToExcel = () => {
    if (!pivotData.rowKeys.length) {
      alert('Belum ada data pivot untuk diekspor.');
      return;
    }

    const pivotSheetRows = pivotData.rowKeys.map((rKey) => {
      const rParts = pivotData.rowPartsMap[rKey] || [rKey];
      const obj = {};
      rowFields.forEach((rf, idx) => {
        obj[rf] = rParts[idx] ?? '';
      });
      if (rowFields.length > 1) {
        obj[rowLabelHeader] = rKey;
      }
      pivotData.colKeys.forEach((cKey) => {
        obj[cKey] = pivotData.matrix[rKey][cKey];
      });
      if (colFields.length > 0) {
        obj['Total Baris'] = pivotData.rowTotals[rKey];
      }
      return obj;
    });

    const grandTotalRow = {};
    rowFields.forEach((rf, idx) => {
      grandTotalRow[rf] = idx === 0 ? 'Grand Total' : '';
    });
    if (rowFields.length > 1) {
      grandTotalRow[rowLabelHeader] = 'Grand Total';
    }
    pivotData.colKeys.forEach((cKey) => {
      grandTotalRow[cKey] = pivotData.colTotals[cKey];
    });
    if (colFields.length > 0) {
      grandTotalRow['Total Baris'] = pivotData.grandTotal;
    }
    pivotSheetRows.push(grandTotalRow);

    const wb = XLSX.utils.book_new();
    const wsPivot = XLSX.utils.json_to_sheet(pivotSheetRows);
    XLSX.utils.book_append_sheet(wb, wsPivot, 'Pivot_Hasil');

    const wsMerged = XLSX.utils.json_to_sheet(mergedDataset);
    XLSX.utils.book_append_sheet(wb, wsMerged, 'Data_Gabungan_1_File');

    const safeCol = rowLabelHeader.replace(/[^a-zA-Z0-9_-]/g, '_');
    XLSX.writeFile(wb, `Pivot_Gabungan_${safeCol}.xlsx`);
  };

  const downloadChartPng = () => {
    const svg = chartSvgRef.current;
    if (!svg || !pivotData.rowKeys.length) return;

    const serializer = new XMLSerializer();
    const svgString = serializer.serializeToString(svg);
    const svgBlob = new Blob([svgString], { type: 'image/svg+xml;charset=utf-8' });
    const url = URL.createObjectURL(svgBlob);

    const img = new Image();
    img.onload = () => {
      const canvas = document.createElement('canvas');
      canvas.width = svg.clientWidth * 2 || 1200;
      canvas.height = svg.clientHeight * 2 || 700;
      const ctx = canvas.getContext('2d');
      ctx.fillStyle = darkMode ? '#090d16' : '#ffffff';
      ctx.fillRect(0, 0, canvas.width, canvas.height);
      ctx.drawImage(img, 0, 0, canvas.width, canvas.height);

      const link = document.createElement('a');
      link.download = `Grafik_Pivot_${rowLabelHeader.replace(/\s+/g, '_')}.png`;
      link.href = canvas.toDataURL('image/png');
      link.click();
      URL.revokeObjectURL(url);
    };
    img.src = url;
  };

  const maxBarValue = useMemo(() => {
    if (!pivotData.rowKeys.length) return 1;
    let maxV = 0;
    pivotData.rowKeys.forEach((r) => {
      const v = colFields.length > 0
        ? Number(pivotData.rowTotals[r] || 0)
        : Number(pivotData.matrix[r]?.['Nilai'] || 0);
      if (v > maxV) maxV = v;
    });
    return maxV > 0 ? maxV : 1;
  }, [pivotData, colFields]);

  const chartPalette = ['#6366f1', '#10b981', '#f59e0b', '#3b82f6', '#ec4899', '#8b5cf6'];
  const remainingRowOptions = mergedFields.filter((f) => !rowFields.includes(f));
  const remainingColOptions = mergedFields.filter((f) => !colFields.includes(f));

  return (
    <div className="space-y-6 animate-fade-in">
      {/* Langkah 1: Unggah Beberapa File Excel/CSV */}
      <section className={`rounded-3xl border shadow-sm p-6 space-y-5 ${darkMode ? 'bg-slate-900/90 border-slate-800' : 'bg-white border-slate-200'}`}>
        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 pb-4 border-b border-slate-800/70">
          <div>
            <h2 className="text-base font-bold flex items-center gap-2">
              <div className="p-1.5 bg-purple-500/10 text-purple-400 rounded-xl">
                <GitMerge className="w-4 h-4" />
              </div>
              Multi-File Auto-Merge & Pivot Table (Multi-Row & Multi-Column)
            </h2>
            <p className={`text-xs mt-0.5 ${darkMode ? 'text-slate-400' : 'text-slate-500'}`}>
              Unggah beberapa file Excel/CSV (meski format kolomnya berbeda). Sistem menyatukannya otomatis menjadi 1 data gabungan sebelum diproses Pivot.
            </p>
          </div>

          <div className="flex flex-wrap items-center gap-2">
            {uploadedFiles.length > 0 && (
              <button
                onClick={handleClearAllFiles}
                className="px-3.5 py-2 rounded-xl border border-slate-700 hover:bg-rose-950/40 hover:text-rose-300 text-xs font-semibold transition-all"
              >
                Hapus Semua File
              </button>
            )}

            {mergedDataset.length > 0 && (
              <button
                onClick={exportMergedOnlyToExcel}
                className="flex items-center gap-2 px-4 py-2 bg-indigo-600 hover:bg-indigo-500 text-white font-bold text-xs rounded-2xl shadow-md transition-all"
              >
                <Download className="w-4 h-4" />
                <span>Unduh Gabungan ({mergedDataset.length} baris)</span>
              </button>
            )}

            <button
              onClick={exportPivotAndMergedToExcel}
              disabled={!pivotData.rowKeys.length}
              className="flex items-center gap-2 px-4 py-2 bg-emerald-600 hover:bg-emerald-500 disabled:opacity-40 disabled:cursor-not-allowed text-white font-bold text-xs rounded-2xl shadow-md transition-all"
            >
              <Download className="w-4 h-4" />
              <span>Ekspor Pivot (.xlsx)</span>
            </button>
          </div>
        </div>

        {/* Dropzone Multiple Files */}
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
            <div className="p-3.5 bg-purple-500/10 text-purple-400 rounded-2xl flex-shrink-0">
              <Upload className="w-6 h-6" />
            </div>
            <div>
              <span className="text-[10px] font-bold text-purple-400 uppercase tracking-widest block mb-1">
                UNGGAH BEBERAPA FILE SEKALIGUS (MULTI-SELECT)
              </span>
              <p className={`text-sm font-semibold ${darkMode ? 'text-slate-200' : 'text-slate-800'}`}>
                Klik atau seret 2 atau lebih file Excel (.xlsx, .xls) / CSV untuk digabung otomatis...
              </p>
              <span className="text-[11px] text-slate-400 block mt-0.5">
                {uploadedFiles.length > 0
                  ? `${uploadedFiles.length} file terunggah • Total ${mergedDataset.length} baris setelah disatukan menjadi 1 data`
                  : 'Belum ada file yang diunggah'}
              </span>
            </div>
          </div>
        </div>

        {/* Daftar File Terunggah & Pilih Sheet */}
        {uploadedFiles.length > 0 && (
          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-3">
            {uploadedFiles.map((f) => {
              const rowCount = (f.sheetsData?.[f.selectedSheet] || []).length;
              return (
                <div
                  key={f.id}
                  className={`p-3.5 rounded-2xl border flex flex-col justify-between gap-2.5 ${darkMode ? 'bg-slate-950/60 border-slate-800' : 'bg-slate-50 border-slate-200'}`}
                >
                  <div className="flex items-start justify-between gap-2">
                    <div className="flex items-center gap-2 min-w-0">
                      <FileSpreadsheet className="w-4 h-4 text-emerald-400 flex-shrink-0" />
                      <div className="min-w-0">
                        <p className="text-xs font-bold truncate">{f.fileName}</p>
                        <p className="text-[10px] text-slate-400">{rowCount} baris</p>
                      </div>
                    </div>
                    <button
                      onClick={() => handleRemoveFile(f.id)}
                      className="p-1.5 rounded-lg text-slate-400 hover:text-rose-400 hover:bg-rose-500/10"
                      title="Hapus file"
                    >
                      <Trash2 className="w-3.5 h-3.5" />
                    </button>
                  </div>

                  <div>
                    <label className="text-[10px] text-slate-400 block mb-1">Sheet:</label>
                    <select
                      value={f.selectedSheet}
                      onChange={(e) => handleChangeSheet(f.id, e.target.value)}
                      className={`w-full text-xs rounded-xl border px-2.5 py-1.5 ${darkMode ? 'bg-slate-900 border-slate-700 text-slate-100' : 'bg-white border-slate-300 text-slate-800'}`}
                    >
                      {f.sheetNames.map((s) => (
                        <option key={s} value={s}>{s}</option>
                      ))}
                    </select>
                  </div>
                </div>
              );
            })}
          </div>
        )}

        {/* Pemetaan Kolom Beda Format Antar File */}
        {uploadedFiles.length >= 2 && (
          <div className={`p-4 rounded-2xl border space-y-3 ${darkMode ? 'bg-slate-950/50 border-slate-800' : 'bg-indigo-50/40 border-indigo-200'}`}>
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2">
              <div className="flex items-center gap-2">
                <Layers className="w-4 h-4 text-indigo-400" />
                <span className="text-xs font-bold">Penyatuan Kolom Beda Nama / Format Antar File</span>
              </div>
              <label className="flex items-center gap-2 text-xs font-semibold cursor-pointer">
                <input
                  type="checkbox"
                  checked={enableAutoAlias}
                  onChange={(e) => setEnableAutoAlias(e.target.checked)}
                  className="w-4 h-4 accent-indigo-500 rounded"
                />
                <span>Aktifkan Kolom Gabungan Standar</span>
              </label>
            </div>

            {enableAutoAlias && (
              <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-3 pt-1">
                {Object.keys(COLUMN_ALIASES).map((stdCol) => (
                  <div
                    key={stdCol}
                    className={`p-3 rounded-xl border space-y-2 ${darkMode ? 'bg-slate-900/90 border-slate-800' : 'bg-white border-slate-200'}`}
                  >
                    <div className="text-[11px] font-bold text-indigo-400">{stdCol}</div>
                    {uploadedFiles.map((f) => {
                      const headers = fileHeadersMap[f.id] || [];
                      const currentVal = aliasMapping[stdCol]?.[f.id] || '';
                      return (
                        <div key={f.id} className="flex items-center justify-between gap-2 text-[11px]">
                          <span className="truncate max-w-[105px] text-slate-400" title={f.fileName}>
                            {f.fileName}
                          </span>
                          <select
                            value={currentVal}
                            onChange={(e) => handleManualAliasChange(stdCol, f.id, e.target.value)}
                            className={`flex-1 text-[11px] rounded-lg border px-2 py-1 ${darkMode ? 'bg-slate-950 border-slate-700 text-slate-200' : 'bg-slate-50 border-slate-300 text-slate-800'}`}
                          >
                            <option value="">-- Tidak dipetakan --</option>
                            {headers.map((h) => (
                              <option key={h} value={h}>{h}</option>
                            ))}
                          </select>
                        </div>
                      );
                    })}
                  </div>
                ))}
              </div>
            )}
          </div>
        )}

        {/* Konfigurasi Pivot: Multi-Rows + Multi-Columns + Values */}
        {mergedFields.length > 0 && (
          <div className="grid grid-cols-1 lg:grid-cols-12 gap-4 pt-2 items-start">
            {/* ROWS: Mendukung > 1 kolom */}
            <div className="lg:col-span-4 space-y-2">
              <div className="flex items-center justify-between">
                <label className="text-[11px] font-bold uppercase tracking-wider text-slate-400 block">
                  ROWS (BISA &gt; 1 KOLOM)
                </label>
                <span className="text-[10px] font-semibold px-2 py-0.5 rounded-full bg-purple-500/15 text-purple-300">
                  {rowFields.length} kolom
                </span>
              </div>

              <div className="space-y-2">
                {rowFields.map((rf, idx) => (
                  <div key={`${rf}-${idx}`} className="flex items-center gap-1.5">
                    <span className="text-[10px] font-mono font-bold px-2 py-1.5 rounded-lg bg-purple-500/15 text-purple-300 shrink-0">
                      R{idx + 1}
                    </span>
                    <select
                      value={rf}
                      onChange={(e) => handleUpdateRowFieldAt(idx, e.target.value)}
                      className={`flex-1 text-xs rounded-xl border px-2.5 py-2 font-medium ${darkMode ? 'bg-slate-950 border-slate-800 text-slate-100' : 'bg-slate-50 border-slate-300 text-slate-800'}`}
                    >
                      {mergedFields.map((f) => (
                        <option key={f} value={f}>{f}</option>
                      ))}
                    </select>
                    {rowFields.length > 1 && (
                      <button
                        type="button"
                        onClick={() => handleRemoveRowFieldAt(idx)}
                        className="p-2 rounded-xl border border-slate-800 text-slate-400 hover:text-rose-400 hover:bg-rose-500/10"
                        title="Hapus row ini"
                      >
                        <X className="w-3.5 h-3.5" />
                      </button>
                    )}
                  </div>
                ))}
              </div>

              {remainingRowOptions.length > 0 && (
                <div className="flex items-center gap-1.5 pt-0.5">
                  <select
                    value={extraRowCandidate}
                    onChange={(e) => setExtraRowCandidate(e.target.value)}
                    className={`flex-1 text-xs rounded-xl border px-2.5 py-1.5 ${darkMode ? 'bg-slate-950 border-slate-800 text-slate-200' : 'bg-slate-50 border-slate-300 text-slate-800'}`}
                  >
                    <option value="">+ Pilih kolom Row...</option>
                    {remainingRowOptions.map((f) => (
                      <option key={f} value={f}>{f}</option>
                    ))}
                  </select>
                  <button
                    type="button"
                    onClick={handleAddExtraRowField}
                    className="flex items-center gap-1 px-2.5 py-1.5 rounded-xl bg-purple-600 hover:bg-purple-500 text-white text-xs font-bold transition-all shrink-0"
                  >
                    <Plus className="w-3.5 h-3.5" />
                    <span>Row</span>
                  </button>
                </div>
              )}
            </div>

            {/* COLUMNS: Mendukung > 1 kolom */}
            <div className="lg:col-span-4 space-y-2">
              <div className="flex items-center justify-between">
                <label className="text-[11px] font-bold uppercase tracking-wider text-slate-400 block">
                  COLUMNS (BISA &gt; 1 KOLOM)
                </label>
                <span className="text-[10px] font-semibold px-2 py-0.5 rounded-full bg-indigo-500/15 text-indigo-300">
                  {colFields.length} kolom
                </span>
              </div>

              <div className="space-y-2">
                {colFields.map((cf, idx) => (
                  <div key={`${cf}-${idx}`} className="flex items-center gap-1.5">
                    <span className="text-[10px] font-mono font-bold px-2 py-1.5 rounded-lg bg-indigo-500/15 text-indigo-300 shrink-0">
                      C{idx + 1}
                    </span>
                    <select
                      value={cf}
                      onChange={(e) => handleUpdateColFieldAt(idx, e.target.value)}
                      className={`flex-1 text-xs rounded-xl border px-2.5 py-2 font-medium ${darkMode ? 'bg-slate-950 border-slate-800 text-slate-100' : 'bg-slate-50 border-slate-300 text-slate-800'}`}
                    >
                      {mergedFields.map((f) => (
                        <option key={f} value={f}>{f}</option>
                      ))}
                    </select>
                    <button
                      type="button"
                      onClick={() => handleRemoveColFieldAt(idx)}
                      className="p-2 rounded-xl border border-slate-800 text-slate-400 hover:text-rose-400 hover:bg-rose-500/10"
                      title="Hapus column ini"
                    >
                      <X className="w-3.5 h-3.5" />
                    </button>
                  </div>
                ))}
              </div>

              {remainingColOptions.length > 0 && (
                <div className="flex items-center gap-1.5 pt-0.5">
                  <select
                    value={extraColCandidate}
                    onChange={(e) => setExtraColCandidate(e.target.value)}
                    className={`flex-1 text-xs rounded-xl border px-2.5 py-1.5 ${darkMode ? 'bg-slate-950 border-slate-800 text-slate-200' : 'bg-slate-50 border-slate-300 text-slate-800'}`}
                  >
                    <option value="">+ Pilih kolom Column...</option>
                    {remainingColOptions.map((f) => (
                      <option key={f} value={f}>{f}</option>
                    ))}
                  </select>
                  <button
                    type="button"
                    onClick={handleAddExtraColField}
                    className="flex items-center gap-1 px-2.5 py-1.5 rounded-xl bg-indigo-600 hover:bg-indigo-500 text-white text-xs font-bold transition-all shrink-0"
                  >
                    <Plus className="w-3.5 h-3.5" />
                    <span>Col</span>
                  </button>
                </div>
              )}
            </div>

            {/* VALUES / AGREGASI & KOLOM ANGKA */}
            <div className="lg:col-span-4 space-y-3">
              <div>
                <label className="text-[11px] font-bold uppercase tracking-wider text-slate-400 block mb-1.5">
                  VALUES / AGREGASI
                </label>
                <select
                  value={aggType}
                  onChange={(e) => setAggType(e.target.value)}
                  className={`w-full text-xs rounded-xl border px-3 py-2 font-medium ${darkMode ? 'bg-slate-950 border-slate-800 text-slate-100' : 'bg-slate-50 border-slate-300 text-slate-800'}`}
                >
                  <option value="COUNT">COUNT (Jumlah Baris)</option>
                  <option value="SUM">SUM (Total Nilai)</option>
                  <option value="AVG">AVG (Rata-rata)</option>
                </select>
              </div>

              <div>
                <label className="text-[11px] font-bold uppercase tracking-wider text-slate-400 block mb-1.5">
                  KOLOM ANGKA (SUM / AVG)
                </label>
                <select
                  value={valueField}
                  onChange={(e) => setValueField(e.target.value)}
                  disabled={aggType === 'COUNT'}
                  className={`w-full text-xs rounded-xl border px-3 py-2 font-medium disabled:opacity-40 ${darkMode ? 'bg-slate-950 border-slate-800 text-slate-100' : 'bg-slate-50 border-slate-300 text-slate-800'}`}
                >
                  <option value="">-- Pilih Kolom Numerik --</option>
                  {mergedFields.map((f) => (
                    <option key={f} value={f}>{f}</option>
                  ))}
                </select>
              </div>
            </div>
          </div>
        )}
      </section>

      {/* Toolbar Tampilan Output (Tabel & Grafik dari Data Gabungan) */}
      {pivotData.rowKeys.length > 0 && (
        <div className={`p-4 rounded-3xl border shadow-sm flex flex-col lg:flex-row items-center justify-between gap-4 ${darkMode ? 'bg-slate-900 border-slate-800' : 'bg-white border-slate-200'}`}>
          <div className="flex items-center gap-1.5 p-1 rounded-2xl border bg-slate-950/50 border-slate-800">
            <button
              onClick={() => setViewMode('BOTH')}
              className={`px-3.5 py-2 rounded-xl text-xs font-bold flex items-center gap-1.5 transition-all ${
                viewMode === 'BOTH' ? 'bg-purple-600 text-white' : 'text-slate-400 hover:text-white'
              }`}
            >
              <LayoutGrid className="w-3.5 h-3.5" />
              <span>Tabel & Grafik</span>
            </button>
            <button
              onClick={() => setViewMode('TABLE')}
              className={`px-3.5 py-2 rounded-xl text-xs font-bold flex items-center gap-1.5 transition-all ${
                viewMode === 'TABLE' ? 'bg-purple-600 text-white' : 'text-slate-400 hover:text-white'
              }`}
            >
              <TableProperties className="w-3.5 h-3.5" />
              <span>Hanya Tabel</span>
            </button>
            <button
              onClick={() => setViewMode('CHART')}
              className={`px-3.5 py-2 rounded-xl text-xs font-bold flex items-center gap-1.5 transition-all ${
                viewMode === 'CHART' ? 'bg-purple-600 text-white' : 'text-slate-400 hover:text-white'
              }`}
            >
              <BarChart3 className="w-3.5 h-3.5" />
              <span>Hanya Grafik</span>
            </button>
          </div>

          <div className="flex flex-col sm:flex-row items-center gap-2.5 w-full lg:w-auto">
            <div className="relative w-full sm:w-56">
              <Search className="w-3.5 h-3.5 text-slate-400 absolute left-3 top-2.5" />
              <input
                type="text"
                placeholder={`Filter ${rowLabelHeader}...`}
                value={searchRow}
                onChange={(e) => setSearchRow(e.target.value)}
                className={`w-full text-xs rounded-xl border pl-8 pr-3 py-2 ${darkMode ? 'bg-slate-950 border-slate-800 text-slate-100' : 'bg-slate-50 border-slate-300 text-slate-800'}`}
              />
            </div>

            <select
              value={sortBy}
              onChange={(e) => setSortBy(e.target.value)}
              className={`w-full sm:w-auto text-xs rounded-xl border px-3 py-2 font-medium ${darkMode ? 'bg-slate-950 border-slate-800 text-slate-100' : 'bg-slate-50 border-slate-300 text-slate-800'}`}
            >
              <option value="TOTAL_DESC">Urutkan: Terbesar ke Terkecil</option>
              <option value="TOTAL_ASC">Urutkan: Terkecil ke Terbesar</option>
              <option value="LABEL_ASC">Urutkan: Nama (A-Z)</option>
            </select>
          </div>
        </div>
      )}

      {/* Tabel Pivot Hasil Penyatuan File (Mendukung Multi-Row & Multi-Column) */}
      {pivotData.rowKeys.length > 0 && (viewMode === 'BOTH' || viewMode === 'TABLE') && (
        <section className={`rounded-3xl border shadow-sm overflow-hidden ${darkMode ? 'bg-slate-900 border-slate-800' : 'bg-white border-slate-200'}`}>
          <div className="px-6 py-4 border-b border-slate-800/70 flex items-center justify-between">
            <div>
              <h3 className="text-sm font-bold">Pivot Table dari {mergedDataset.length} Baris Data Gabungan</h3>
              <p className="text-[11px] text-slate-400">
                Rows: <span className="text-purple-400 font-semibold">{rowLabelHeader}</span>
                {colLabelHeader && (
                  <>
                    {' '}• Columns: <span className="text-indigo-400 font-semibold">{colLabelHeader}</span>
                  </>
                )}
              </p>
            </div>
            <span className="text-xs font-mono font-bold text-emerald-400">
              Grand Total: {formatMetric(pivotData.grandTotal)}
            </span>
          </div>

          <div className="overflow-x-auto max-h-[450px]">
            <table className="w-full text-left text-xs border-collapse">
              <thead className={`uppercase font-bold tracking-wider sticky top-0 z-10 border-b ${darkMode ? 'bg-slate-950 text-slate-300 border-slate-800' : 'bg-slate-100 text-slate-700 border-slate-200'}`}>
                <tr>
                  {rowFields.map((rf) => (
                    <th key={rf} className="p-3.5">{rf}</th>
                  ))}
                  {rowFields.length > 1 && <th className="p-3.5">Kombinasi Row</th>}
                  {pivotData.colKeys.map((cKey) => (
                    <th key={cKey} className="p-3.5 text-right">{cKey}</th>
                  ))}
                  {colFields.length > 0 && <th className="p-3.5 text-right">Total Baris</th>}
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
                      {rowFields.length > 1 && (
                        <td className="p-3.5 font-semibold font-mono text-purple-400">{rKey}</td>
                      )}
                      {pivotData.colKeys.map((cKey) => (
                        <td key={cKey} className="p-3.5 text-right font-mono">
                          {formatMetric(pivotData.matrix[rKey][cKey])}
                        </td>
                      ))}
                      {colFields.length > 0 && (
                        <td className="p-3.5 text-right font-mono font-bold text-indigo-400">
                          {formatMetric(pivotData.rowTotals[rKey])}
                        </td>
                      )}
                    </tr>
                  );
                })}
              </tbody>
              <tfoot className={`font-bold border-t-2 ${darkMode ? 'bg-slate-950 text-slate-100 border-slate-700' : 'bg-slate-100 text-slate-900 border-slate-300'}`}>
                <tr>
                  <td colSpan={rowFields.length > 1 ? rowFields.length + 1 : rowFields.length} className="p-3.5">
                    Grand Total
                  </td>
                  {pivotData.colKeys.map((cKey) => (
                    <td key={cKey} className="p-3.5 text-right font-mono text-emerald-400">
                      {formatMetric(pivotData.colTotals[cKey])}
                    </td>
                  ))}
                  {colFields.length > 0 && (
                    <td className="p-3.5 text-right font-mono text-emerald-400">
                      {formatMetric(pivotData.grandTotal)}
                    </td>
                  )}
                </tr>
              </tfoot>
            </table>
          </div>
        </section>
      )}

      {/* Grafik Evaluasi Atasan */}
      {pivotData.rowKeys.length > 0 && (viewMode === 'BOTH' || viewMode === 'CHART') && (
        <section className={`rounded-3xl border shadow-sm p-6 space-y-4 ${darkMode ? 'bg-slate-900 border-slate-800' : 'bg-white border-slate-200'}`}>
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 pb-3 border-b border-slate-800/70">
            <div>
              <h3 className="text-sm font-bold flex items-center gap-2">
                <BarChart3 className="w-4 h-4 text-indigo-400" />
                Grafik Visualisasi Evaluasi per {rowLabelHeader}
              </h3>
              <p className="text-[11px] text-slate-400">
                Dihitung dari gabungan {uploadedFiles.length} file ({aggType}{aggType !== 'COUNT' && valueField ? ` dari ${valueField}` : ''}).
              </p>
            </div>
            <button
              onClick={downloadChartPng}
              className="px-3.5 py-2 rounded-xl border border-slate-700 hover:bg-slate-800 text-xs font-semibold flex items-center gap-1.5 self-start sm:self-auto"
            >
              <Download className="w-3.5 h-3.5 text-indigo-400" />
              <span>Simpan Grafik (.png)</span>
            </button>
          </div>

          <div className="space-y-3 max-h-[460px] overflow-y-auto pr-1">
            {pivotData.rowKeys.map((rKey, idx) => {
              const displayTotal = colFields.length > 0
                ? Number(pivotData.rowTotals[rKey] || 0)
                : Number(pivotData.matrix[rKey]?.['Nilai'] || 0);
              const pct = Math.min(100, Math.round((displayTotal / maxBarValue) * 100));
              const barColor = chartPalette[idx % chartPalette.length];

              return (
                <div key={rKey} className="space-y-1">
                  <div className="flex items-center justify-between text-xs gap-3">
                    <span className="font-mono font-semibold truncate max-w-[65%]">{rKey}</span>
                    <span className="font-mono font-bold text-slate-300">{formatMetric(displayTotal)}</span>
                  </div>
                  <div className={`w-full h-3 rounded-full overflow-hidden ${darkMode ? 'bg-slate-950' : 'bg-slate-100'}`}>
                    <div
                      className="h-full rounded-full transition-all duration-500"
                      style={{ width: `${pct}%`, backgroundColor: barColor }}
                    />
                  </div>
                </div>
              );
            })}
          </div>

          <div className="hidden">
            <svg
              ref={chartSvgRef}
              width="900"
              height={Math.max(260, pivotData.rowKeys.length * 42 + 80)}
              xmlns="http://www.w3.org/2000/svg"
            >
              <text x="30" y="36" fill={darkMode ? '#f8fafc' : '#0f172a'} fontSize="18" fontWeight="bold">
                Grafik Evaluasi Pivot Gabungan: {rowLabelHeader} ({aggType})
              </text>
              {pivotData.rowKeys.map((rKey, idx) => {
                const val = colFields.length > 0
                  ? Number(pivotData.rowTotals[rKey] || 0)
                  : Number(pivotData.matrix[rKey]?.['Nilai'] || 0);
                const barWidth = Math.max(4, Math.round((val / maxBarValue) * 520));
                const y = 65 + idx * 40;
                return (
                  <g key={rKey}>
                    <text x="30" y={y + 16} fill={darkMode ? '#cbd5e1' : '#334155'} fontSize="13">
                      {String(rKey).slice(0, 32)}
                    </text>
                    <rect
                      x="260"
                      y={y}
                      width={barWidth}
                      height="22"
                      rx="6"
                      fill={chartPalette[idx % chartPalette.length]}
                    />
                    <text x={270 + barWidth} y={y + 16} fill={darkMode ? '#e2e8f0' : '#0f172a'} fontSize="13" fontWeight="bold">
                      {formatMetric(val)}
                    </text>
                  </g>
                );
              })}
            </svg>
          </div>
        </section>
      )}
    </div>
  );
}