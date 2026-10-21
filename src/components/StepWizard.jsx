import React from 'react';
import { Check } from 'lucide-react';

export default function StepWizard({ step, darkMode }) {
  return (
    <div className={`p-5 rounded-3xl border shadow-sm ${darkMode ? 'bg-slate-900/60 border-slate-800' : 'bg-white border-slate-200'}`}>
      <div className="relative flex items-center justify-between max-w-4xl mx-auto px-4">
        
        <div className={`absolute left-16 right-16 top-5 h-0.5 -z-0 ${darkMode ? 'bg-slate-800' : 'bg-slate-200'}`}></div>

        {/* Step 1 */}
        <div className="relative z-10 flex flex-col items-center text-center">
          <div className={`w-10 h-10 rounded-full flex items-center justify-center font-bold text-sm shadow-md transition-all ${
            step >= 1 ? 'bg-emerald-500 text-white shadow-emerald-500/30' : (darkMode ? 'bg-slate-800 text-slate-400 border border-slate-700' : 'bg-slate-200 text-slate-600')
          }`}>
            {step > 1 ? <Check className="w-5 h-5" /> : '1'}
          </div>
          <span className={`text-xs font-bold mt-2 ${step >= 1 ? (darkMode ? 'text-slate-100' : 'text-slate-900') : 'text-slate-400'}`}>
            1. Unggah Dokumen
          </span>
          <span className="text-[10px] text-slate-400 mt-0.5">
            {step > 1 ? '2 File Terverifikasi' : 'Menunggu Berkas'}
          </span>
        </div>

        {/* Step 2 */}
        <div className="relative z-10 flex flex-col items-center text-center">
          <div className={`w-10 h-10 rounded-full flex items-center justify-center font-bold text-sm shadow-md transition-all ${
            step >= 2 ? 'bg-blue-600 text-white shadow-blue-600/30' : (darkMode ? 'bg-slate-800 text-slate-400 border border-slate-700' : 'bg-slate-200 text-slate-600')
          }`}>
            {step > 3 ? <Check className="w-5 h-5" /> : '2'}
          </div>
          <span className={`text-xs font-bold mt-2 ${step >= 2 ? (darkMode ? 'text-slate-100' : 'text-slate-900') : 'text-slate-400'}`}>
            2. Konfigurasi Kunci & Aturan
          </span>
          <span className="text-[10px] text-slate-400 mt-0.5">
            {step === 2 ? 'Sedang Dikonfigurasi' : step > 2 ? 'Selesai' : 'Tertunda'}
          </span>
        </div>

        {/* Step 3 */}
        <div className="relative z-10 flex flex-col items-center text-center">
          <div className={`w-10 h-10 rounded-full flex items-center justify-center font-bold text-sm shadow-md transition-all ${
            step >= 3 ? 'bg-indigo-600 text-white shadow-indigo-600/30' : (darkMode ? 'bg-slate-800 text-slate-400 border border-slate-700' : 'bg-slate-200 text-slate-600')
          }`}>
            3
          </div>
          <span className={`text-xs font-bold mt-2 ${step >= 3 ? (darkMode ? 'text-slate-100' : 'text-slate-900') : 'text-slate-400'}`}>
            3. Hasil Audit
          </span>
          <span className="text-[10px] text-slate-400 mt-0.5">Laporan & Unduh</span>
        </div>

      </div>
    </div>
  );
}