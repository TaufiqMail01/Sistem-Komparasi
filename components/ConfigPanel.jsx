import React from 'react';
import { Settings, Key, Sliders } from 'lucide-react';

export default function ConfigPanel({ commonHeaders, config, onChange }) {
  return (
    <div className="bg-white p-5 rounded-xl border border-slate-200 shadow-sm space-y-4">
      <div className="flex items-center gap-2 border-b border-slate-100 pb-3">
        <Settings className="w-5 h-5 text-indigo-600" />
        <h3 className="font-semibold text-slate-800">Pengaturan Komparasi</h3>
      </div>

      <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
        {/* Primary Key Selection */}
        <div>
          <label className="block text-xs font-medium text-slate-600 mb-1 flex items-center gap-1">
            <Key className="w-3.5 h-3.5 text-amber-500" />
            Kunci Utama (Primary Key):
          </label>
          <select
            value={config.primaryKey}
            onChange={(e) => onChange({ ...config, primaryKey: e.target.value })}
            className="w-full text-sm border border-slate-300 rounded-lg p-2 bg-white focus:ring-2 focus:ring-indigo-500"
          >
            <option value="">-- Tanpa Kunci (Bandingkan Urutan Baris) --</option>
            {commonHeaders.map((col) => (
              <option key={col} value={col}>{col}</option>
            ))}
          </select>
          <p className="text-[11px] text-slate-400 mt-1">
            Gunakan ID Unik seperti NIK, NRK, atau Kode Unik untuk hasil akurat meskipun urutan baris acak.
          </p>
        </div>

        {/* Toleransi Numerik */}
        <div>
          <label className="block text-xs font-medium text-slate-600 mb-1 flex items-center gap-1">
            <Sliders className="w-3.5 h-3.5 text-indigo-500" />
            Toleransi Beda Selisih Angka:
          </label>
          <input
            type="number"
            value={config.numericTolerance}
            onChange={(e) => onChange({ ...config, numericTolerance: parseFloat(e.target.value) || 0 })}
            placeholder="0"
            className="w-full text-sm border border-slate-300 rounded-lg p-2 focus:ring-2 focus:ring-indigo-500"
          />
          <p className="text-[11px] text-slate-400 mt-1">
            Abaikan selisih kecil (misal: pembulatan desimal atau selisih Rp 1 - Rp 10).
          </p>
        </div>
      </div>

      {/* Checkboxes Opsi Tambahan */}
      <div className="flex flex-wrap gap-6 pt-2 border-t border-slate-100">
        <label className="flex items-center gap-2 cursor-pointer text-xs font-medium text-slate-700">
          <input
            type="checkbox"
            checked={config.ignoreCase}
            onChange={(e) => onChange({ ...config, ignoreCase: e.target.checked })}
            className="rounded text-indigo-600 focus:ring-indigo-500"
          />
          Abaikan Huruf Besar/Kecil (Case-Insensitive)
        </label>

        <label className="flex items-center gap-2 cursor-pointer text-xs font-medium text-slate-700">
          <input
            type="checkbox"
            checked={config.trimWhitespace}
            onChange={(e) => onChange({ ...config, trimWhitespace: e.target.checked })}
            className="rounded text-indigo-600 focus:ring-indigo-500"
          />
          Hapus Spasi Tambahan di Awal/Akhir
        </label>
      </div>
    </div>
  );
}