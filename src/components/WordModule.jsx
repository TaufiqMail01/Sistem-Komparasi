import React from 'react';
import { FileText, RotateCcw } from 'lucide-react';

export default function WordModule({ 
  darkMode, resetAllData, handleUploadWord1, handleUploadWord2, 
  wordFile1, wordFile2, wordText1, wordText2, wordDiffResult 
}) {
  return (
    <section className={`rounded-3xl border shadow-sm p-6 space-y-6 transition-all ${darkMode ? 'bg-slate-900/90 border-slate-800 shadow-xl' : 'bg-white border-slate-200/90 shadow-sm'}`}>
      <div className={`flex flex-col md:flex-row md:items-center justify-between gap-4 pb-4 border-b ${darkMode ? 'border-slate-800' : 'border-slate-100'}`}>
        <div>
          <h2 className="text-base font-bold flex items-center gap-2">
            <div className="p-1.5 bg-indigo-500/10 text-indigo-500 rounded-xl">
              <FileText className="w-4 h-4" />
            </div>
            Modul Review & Komparasi Dokumen Naskah Word (.docx)
          </h2>
          <p className={`text-xs mt-0.5 ${darkMode ? 'text-slate-400' : 'text-slate-500'}`}>
            Unggah dua dokumen Word (.docx) untuk melihat penyorotan perbedaan teks secara berdampingan.
          </p>
        </div>
        <button 
          onClick={resetAllData} 
          className={`text-xs font-semibold flex items-center gap-1.5 px-3.5 py-2 rounded-xl border transition-all ${
            darkMode ? 'border-slate-700 hover:bg-slate-800 text-slate-300' : 'border-slate-200 hover:bg-rose-50 hover:border-rose-200 text-slate-600 hover:text-rose-600'
          }`}
        >
          <RotateCcw className="w-3.5 h-3.5" /> Reset Dokumen
        </button>
      </div>

      <div className="grid grid-cols-1 md:grid-cols-2 gap-5">
        <div className={`border-2 border-dashed rounded-3xl p-5 relative transition-all group ${darkMode ? 'border-slate-800 bg-slate-950/40 hover:border-indigo-500/50' : 'border-slate-200 bg-slate-50/50 hover:border-indigo-400 hover:bg-indigo-50/20'}`}>
          <input type="file" accept=".docx" onChange={(e) => e.target.files[0] && handleUploadWord1(e.target.files[0])} className="absolute inset-0 w-full h-full opacity-0 cursor-pointer z-10" />
          <div className="flex items-start gap-4">
            <div className="p-3.5 bg-indigo-500/10 text-indigo-400 rounded-2xl group-hover:scale-110 transition-transform"><FileText className="w-6 h-6" /></div>
            <div className="flex-1 min-w-0">
              <span className="text-[10px] font-bold text-indigo-400 uppercase tracking-widest block mb-1">DOKUMEN 1 (ASLI / LAMA)</span>
              <p className={`text-sm font-semibold truncate ${darkMode ? 'text-slate-200' : 'text-slate-800'}`}>{wordFile1 ? wordFile1.fileName : "Klik atau seret Dokumen 1 (.docx)..."}</p>
            </div>
          </div>
        </div>

        <div className={`border-2 border-dashed rounded-3xl p-5 relative transition-all group ${darkMode ? 'border-slate-800 bg-slate-950/40 hover:border-purple-500/50' : 'border-slate-200 bg-slate-50/50 hover:border-purple-400 hover:bg-purple-50/20'}`}>
          <input type="file" accept=".docx" onChange={(e) => e.target.files[0] && handleUploadWord2(e.target.files[0])} className="absolute inset-0 w-full h-full opacity-0 cursor-pointer z-10" />
          <div className="flex items-start gap-4">
            <div className="p-3.5 bg-purple-500/10 text-purple-400 rounded-2xl group-hover:scale-110 transition-transform"><FileText className="w-6 h-6" /></div>
            <div className="flex-1 min-w-0">
              <span className="text-[10px] font-bold text-purple-400 uppercase tracking-widest block mb-1">DOKUMEN 2 (REVISI / BARU)</span>
              <p className={`text-sm font-semibold truncate ${darkMode ? 'text-slate-200' : 'text-slate-800'}`}>{wordFile2 ? wordFile2.fileName : "Klik atau seret Dokumen 2 (.docx)..."}</p>
            </div>
          </div>
        </div>
      </div>

      {wordText1 && wordText2 && (
        <div className={`space-y-5 pt-4 border-t animate-fade-in ${darkMode ? 'border-slate-800' : 'border-slate-100'}`}>
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
            <h3 className={`text-sm font-bold ${darkMode ? 'text-slate-200' : 'text-slate-900'}`}>Pratinjau Penyorotan Perubahan Dokumen:</h3>
            <div className="flex items-center gap-2.5 text-xs">
              <span className="flex items-center gap-1.5 px-3 py-1 bg-rose-500/10 text-rose-400 rounded-xl font-bold border border-rose-500/20">Teks Dihapus (Kiri)</span>
              <span className="flex items-center gap-1.5 px-3 py-1 bg-emerald-500/10 text-emerald-400 rounded-xl font-bold border border-emerald-500/20">Teks Ditambahkan (Kanan)</span>
            </div>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 gap-5">
            <div className={`border rounded-3xl p-5 h-96 overflow-y-auto shadow-inner ${darkMode ? 'bg-slate-950 border-slate-800' : 'bg-slate-50 border-slate-200'}`}>
              <span className={`text-xs font-bold uppercase tracking-wider block mb-3 pb-2 border-b ${darkMode ? 'text-slate-400 border-slate-800' : 'text-slate-500 border-slate-200'}`}>
                Dokumen 1 (Asli dengan Highlight Perubahan)
              </span>
              <div className={`text-xs whitespace-pre-wrap leading-relaxed font-mono ${darkMode ? 'text-slate-300' : 'text-slate-700'}`}>
                {wordDiffResult && wordDiffResult.map((part, index) => {
                  if (part.added) return null;
                  const style = part.removed 
                    ? 'bg-rose-500/20 text-rose-300 px-1.5 py-0.5 rounded-md line-through decoration-rose-500 font-medium' 
                    : '';
                  return <span key={index} className={style}>{part.value}</span>;
                })}
              </div>
            </div>

            <div className={`border rounded-3xl p-5 h-96 overflow-y-auto shadow-inner ${darkMode ? 'bg-slate-950 border-slate-800' : 'bg-slate-50 border-slate-200'}`}>
              <span className={`text-xs font-bold uppercase tracking-wider block mb-3 pb-2 border-b ${darkMode ? 'text-slate-400 border-slate-800' : 'text-slate-500 border-slate-200'}`}>
                Dokumen 2 (Revisi dengan Highlight Penambahan)
              </span>
              <div className={`text-xs whitespace-pre-wrap leading-relaxed font-mono ${darkMode ? 'text-slate-300' : 'text-slate-700'}`}>
                {wordDiffResult && wordDiffResult.map((part, index) => {
                  if (part.removed) return null;
                  const style = part.added 
                    ? 'bg-emerald-500/20 text-emerald-300 px-1.5 py-0.5 rounded-md font-bold' 
                    : '';
                  return <span key={index} className={style}>{part.value}</span>;
                })}
              </div>
            </div>
          </div>

          <div className={`rounded-3xl p-6 space-y-3 border shadow-md ${darkMode ? 'bg-slate-950 border-slate-800 text-white' : 'bg-slate-900 text-white border-slate-800'}`}>
            <span className="text-xs font-bold text-amber-400 uppercase tracking-widest block">Panel Review Gabungan (Format Track Changes):</span>
            <div className={`p-4 rounded-2xl border max-h-80 overflow-y-auto text-xs leading-relaxed font-mono ${darkMode ? 'bg-slate-900 border-slate-800' : 'bg-slate-950 border-slate-800'}`}>
              {wordDiffResult && wordDiffResult.map((part, index) => {
                const colorClass = part.added 
                  ? 'bg-emerald-500/30 text-emerald-300 border-b-2 border-emerald-500 px-1.5 py-0.5 rounded-sm font-bold' 
                  : part.removed 
                  ? 'bg-rose-500/30 text-rose-300 line-through px-1.5 py-0.5 rounded-sm' 
                  : 'text-slate-300';
                return <span key={index} className={colorClass}>{part.value}</span>;
              })}
            </div>
          </div>
        </div>
      )}
    </section>
  );
}