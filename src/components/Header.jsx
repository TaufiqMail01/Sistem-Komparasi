import React from 'react';
import { FileSpreadsheet, FileText, Sun, Moon, HelpCircle, Layers } from 'lucide-react';

export default function Header({ appMode, setAppMode, darkMode, setDarkMode }) {
  return (
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
                <span className="text-indigo-400 font-medium text-xs">| Sistem Rekonsiliasi & Pencocokan Dokumen</span>
              </div>
              <p className={`text-[11px] ${darkMode ? 'text-slate-400' : 'text-slate-500'}`}>
                Universitas Bina Sarana Informatika &bull; Pusat Komparasi & Rekonsiliasi Dokumen Resmi
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

            <button 
              onClick={() => alert("Panduan: 1. Unggah dokumen pada Langkah 1. 2. Atur Primary Key pada Langkah 2. 3. Lihat hasil audit di bawah.")}
              className={`px-3 py-2 rounded-2xl border text-xs font-semibold flex items-center gap-1.5 transition-all ${
                darkMode ? 'bg-slate-800 border-slate-700 text-slate-200 hover:bg-slate-700' : 'bg-white border-slate-200 text-slate-700 hover:bg-slate-50'
              }`}
            >
              <HelpCircle className="w-3.5 h-3.5 text-blue-500" /> Panduan
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
  );
}