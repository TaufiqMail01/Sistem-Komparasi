import React from 'react';
import { Upload, FileSpreadsheet, CheckCircle2 } from 'lucide-react';

export default function FileUploader({ label, fileData, onFileUpload, sheetSelected, onSheetChange }) {
  const handleFileChange = (e) => {
    const file = e.target.files[0];
    if (file) onFileUpload(file);
  };

  return (
    <div className="bg-white p-5 rounded-xl border border-slate-200 shadow-sm hover:border-indigo-300 transition-all">
      <div className="flex items-center gap-2 mb-3">
        <FileSpreadsheet className="w-5 h-5 text-indigo-600" />
        <h3 className="font-semibold text-slate-800">{label}</h3>
      </div>

      {!fileData ? (
        <label className="flex flex-col items-center justify-center p-6 border-2 border-dashed border-slate-300 rounded-lg cursor-pointer bg-slate-50 hover:bg-slate-100 transition-colors">
          <Upload className="w-8 h-8 text-slate-400 mb-2" />
          <span className="text-sm font-medium text-slate-600">Klik untuk upload Excel / CSV</span>
          <span className="text-xs text-slate-400 mt-1">Format: .xlsx, .xls, .csv</span>
          <input type="file" accept=".xlsx, .xls, .csv" onChange={handleFileChange} className="hidden" />
        </label>
      ) : (
        <div className="space-y-3">
          <div className="flex items-center justify-between bg-emerald-50 p-3 rounded-lg border border-emerald-200">
            <div className="flex items-center gap-2 overflow-hidden">
              <CheckCircle2 className="w-5 h-5 text-emerald-600 shrink-0" />
              <span className="text-sm font-medium text-emerald-900 truncate">{fileData.fileName}</span>
            </div>
            <label className="text-xs text-indigo-600 font-semibold cursor-pointer hover:underline shrink-0">
              Ganti
              <input type="file" accept=".xlsx, .xls, .csv" onChange={handleFileChange} className="hidden" />
            </label>
          </div>

          <div>
            <label className="block text-xs font-medium text-slate-600 mb-1">Pilih Sheet:</label>
            <select
              value={sheetSelected}
              onChange={(e) => onSheetChange(e.target.value)}
              className="w-full text-sm border border-slate-300 rounded-lg p-2 bg-white focus:ring-2 focus:ring-indigo-500"
            >
              {fileData.sheetNames.map((sheet) => (
                <option key={sheet} value={sheet}>{sheet}</option>
              ))}
            </select>
          </div>
        </div>
      )}
    </div>
  );
}