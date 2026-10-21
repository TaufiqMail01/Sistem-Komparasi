import React, { useState } from 'react';
import { FileSpreadsheet, FileText, Sun, Moon, Layers, Info, X, CheckCircle2 } from 'lucide-react';

export default function Header({ appMode, setAppMode, darkMode, setDarkMode }) {
  const [showInstructionsModal, setShowInstructionsModal] = useState(false);

  return (
    <>
      <header className={`border-b sticky top-0 z-40 backdrop-blur-md transition-colors duration-300 ${darkMode ? 'bg-slate-900/90 border-slate-800 text-white shadow-xl shadow-black/20' : 'bg-white/90 border-slate-200 text-slate-900 shadow-sm'}`}>
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-3">
          <div className="flex flex-col md:flex-row md:items-center md:justify-between gap-3">
            
            {/* Logo & Judul Sistem */}
            <div className="flex items-center gap-3">
              <div className="p-2.5 bg-gradient-to-br from-blue-600 to-indigo-600 rounded-2xl shadow-md shadow-indigo-500/30 text-white">
                <Layers className="w-5 h-5" />
              </div>
              <div>
                <div className="flex items-center gap-2">
                  <h1 className="text-base font-bold tracking-tight">ReconDoc</h1>
                  <span className="text-indigo-400 font-medium text-xs">| Sistem Rekonsiliasi & Audit Dokumen</span>
                </div>
                <p className={`text-[11px] ${darkMode ? 'text-slate-400' : 'text-slate-500'}`}>
                  Universitas Bina Sarana Informatika &bull; Divisi Admin Kontrak PT Permata Indo Sejahtera
                </p>
              </div>
            </div>

            {/* Navigasi Kanan */}
            <div className="flex items-center gap-2.5">
              <div className={`p-1 rounded-2xl flex items-center gap-1 border shadow-inner ${darkMode ? 'bg-slate-950 border-slate-800' : 'bg-slate-100 border-slate-200'}`}>
                <button 
                  onClick={() => setAppMode('EXCEL')}
                  className={`px-3 py-1.5 rounded-xl text-xs font-bold transition-all flex items-center gap-1.5 ${
                    appMode === 'EXCEL' 
                      ? 'bg-blue-600 text-white shadow-md shadow-blue-500/30' 
                      : (darkMode ? 'text-slate-400 hover:text-white' : 'text-slate-600 hover:text-slate-900')
                  }`}
                >
                  <FileSpreadsheet className="w-3.5 h-3.5" /> Mode Excel & CSV
                </button>
                <button 
                  onClick={() => setAppMode('WORD')}
                  className={`px-3 py-1.5 rounded-xl text-xs font-bold transition-all flex items-center gap-1.5 ${
                    appMode === 'WORD' 
                      ? 'bg-indigo-600 text-white shadow-md shadow-indigo-500/30' 
                      : (darkMode ? 'text-slate-400 hover:text-white' : 'text-slate-600 hover:text-slate-900')
                  }`}
                >
                  <FileText className="w-3.5 h-3.5" /> Mode Review Word
                </button>
              </div>

              {/* Tombol Cara Penggunaan */}
              <button 
                onClick={() => setShowInstructionsModal(true)}
                className={`px-3.5 py-2 rounded-2xl border text-xs font-semibold flex items-center gap-1.5 transition-all shadow-sm ${
                  darkMode ? 'bg-slate-800 border-slate-700 text-slate-200 hover:bg-slate-700' : 'bg-white border-slate-200 text-slate-700 hover:bg-slate-50'
                }`}
                title="Cara Penggunaan Aplikasi"
              >
                <Info className="w-3.5 h-3.5 text-indigo-400" />
                <span>Cara Penggunaan</span>
              </button>

              <button 
                onClick={() => setDarkMode(!darkMode)}
                className={`p-2 rounded-2xl border transition-all shadow-sm ${
                  darkMode ? 'bg-slate-800 border-slate-700 text-amber-400 hover:bg-slate-700' : 'bg-white border-slate-200 text-indigo-600 hover:bg-slate-50'
                }`}
                title="Ganti Tema"
              >
                {darkMode ? <Sun className="w-4 h-4" /> : <Moon className="w-4 h-4" />}
              </button>
            </div>

          </div>
        </div>
      </header>

      {/* Modal Instruksi / Cara Penggunaan */}
      {showInstructionsModal && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-slate-950/80 backdrop-blur-sm animate-fade-in">
          <div className={`max-w-xl w-full rounded-3xl border p-6 shadow-2xl space-y-5 ${darkMode ? 'bg-slate-900 border-slate-800 text-slate-100' : 'bg-white border-slate-200 text-slate-900'}`}>
            <div className="flex items-center justify-between pb-3 border-b border-slate-700/50">
              <div className="flex items-center gap-2.5">
                <div className="p-2 bg-indigo-500/10 text-indigo-400 rounded-xl">
                  <Info className="w-5 h-5" />
                </div>
                <h3 className="text-base font-bold">Instruksi & Cara Kerja Sistem ReconDoc</h3>
              </div>
              <button 
                onClick={() => setShowInstructionsModal(false)}
                className="p-2 rounded-xl text-slate-400 hover:bg-slate-800 hover:text-white transition-colors"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <div className="space-y-3.5 text-xs leading-relaxed text-slate-300">
              <div className="p-3.5 rounded-2xl bg-slate-950/50 border border-slate-800 space-y-1">
                <div className="flex items-center gap-2 font-bold text-blue-400">
                  <CheckCircle2 className="w-4 h-4" /> 1. Mode Excel & CSV (Data Reconciliation)
                </div>
                <ul className="list-disc list-inside space-y-1 pl-1 text-slate-400">
                  <li><strong>Unggah Dokumen:</strong> Masukkan File Lama (Referensi) dan File Baru (Pembanding).</li>
                  <li><strong>Pencocokan Baris:</strong> Sistem mendeteksi baris Identik, Perubahan Data, Data Baru, dan Data Dihapus berdasarkan *Primary Key*.</li>
                </ul>
              </div>

              <div className="p-3.5 rounded-2xl bg-slate-950/50 border border-slate-800 space-y-1">
                <div className="flex items-center gap-2 font-bold text-indigo-400">
                  <CheckCircle2 className="w-4 h-4" /> 2. Mode Review Word (.docx) & Unduh Laporan
                </div>
                <p className="text-slate-400">
                  Unggah dua dokumen Word untuk membandingkan perbedaan teks secara langsung. Hasil perbandingan dapat diunduh dalam bentuk laporan teks (.txt).
                </p>
              </div>
            </div>

            <div className="pt-2 flex justify-end">
              <button 
                onClick={() => setShowInstructionsModal(false)}
                className="px-5 py-2.5 bg-indigo-600 hover:bg-indigo-500 text-white font-bold text-xs rounded-2xl shadow-md transition-all"
              >
                Mengerti & Tutup
              </button>
            </div>
          </div>
        </div>
      )}
    </>
  );
}