import React, { useState } from 'react';
import {
  Sun,
  Moon,
  Layers,
  Info,
  X,
  CheckCircle2,
  ShieldCheck,
  Home
} from 'lucide-react';

export default function Header({
  darkMode,
  setDarkMode,
  viewMode = 'HOME',
  onBackToHome
}) {
  const [showInstructionsModal, setShowInstructionsModal] = useState(false);

  return (
    <>
      <header
        className={`border-b sticky top-0 z-40 backdrop-blur-md transition-colors duration-300 ${
          darkMode
            ? 'bg-slate-900/90 border-slate-800 text-white shadow-xl shadow-black/20'
            : 'bg-white/90 border-slate-200 text-slate-900 shadow-sm'
        }`}
      >
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-3">
          <div className="flex flex-col md:flex-row md:items-center md:justify-between gap-3">
            {/* Logo & Judul Sistem */}
            <div className="flex items-center gap-3">
              <div className="p-2.5 bg-gradient-to-br from-blue-600 to-indigo-600 rounded-2xl shadow-md shadow-indigo-500/30 text-white shrink-0">
                <Layers className="w-5 h-5" />
              </div>
              <div>
                <div className="flex flex-wrap items-center gap-2">
                  <h1 className="text-base font-bold tracking-tight">ReconDoc</h1>
                  <span className="text-indigo-400 font-medium text-xs">
                    | Sistem Rekonsiliasi & Audit Dokumen
                  </span>
                  <span
                    className={`inline-flex items-center gap-1 text-[10px] font-semibold px-2 py-0.5 rounded-full border ${
                      darkMode
                        ? 'bg-emerald-950/50 border-emerald-500/30 text-emerald-300'
                        : 'bg-emerald-50 border-emerald-200 text-emerald-700'
                    }`}
                    title="Seluruh file diproses langsung di browser tanpa disimpan ke server"
                  >
                    <ShieldCheck className="w-3 h-3" />
                    <span>Diproses Lokal di Browser</span>
                  </span>
                </div>
                <p className={`text-[11px] mt-0.5 ${darkMode ? 'text-slate-400' : 'text-slate-500'}`}>
                  Bandingkan data Excel/CSV, hitung rumus otomatis, buat Pivot Table, atau cek perbedaan dokumen Word
                </p>
              </div>
            </div>

            {/* Navigasi Kanan: Switcher Mode Excel/Word dihapus di beranda; Kembali ke Beranda muncul saat di workspace */}
            <div className="flex flex-wrap items-center gap-2.5">
              {viewMode === 'WORKSPACE' && (
                <button
                  type="button"
                  onClick={onBackToHome}
                  className={`px-3.5 py-2 rounded-2xl border text-xs font-bold flex items-center gap-1.5 transition-all shadow-sm ${
                    darkMode
                      ? 'bg-indigo-600/20 border-indigo-500/40 text-indigo-200 hover:bg-indigo-600/35'
                      : 'bg-indigo-50 border-indigo-200 text-indigo-700 hover:bg-indigo-100'
                  }`}
                  title="Kembali ke halaman beranda kartu"
                >
                  <Home className="w-3.5 h-3.5 text-indigo-400" />
                  <span>Kembali ke Beranda</span>
                </button>
              )}

              {/* Tombol Cara Penggunaan */}
              <button
                type="button"
                onClick={() => setShowInstructionsModal(true)}
                className={`px-3.5 py-2 rounded-2xl border text-xs font-semibold flex items-center gap-1.5 transition-all shadow-sm ${
                  darkMode
                    ? 'bg-slate-800 border-slate-700 text-slate-200 hover:bg-slate-700'
                    : 'bg-white border-slate-200 text-slate-700 hover:bg-slate-50'
                }`}
                title="Cara Penggunaan Aplikasi"
              >
                <Info className="w-3.5 h-3.5 text-indigo-400" />
                <span>Cara Penggunaan</span>
              </button>

              <button
                type="button"
                onClick={() => setDarkMode(!darkMode)}
                className={`p-2 rounded-2xl border transition-all shadow-sm ${
                  darkMode
                    ? 'bg-slate-800 border-slate-700 text-amber-400 hover:bg-slate-700'
                    : 'bg-white border-slate-200 text-indigo-600 hover:bg-slate-50'
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
          <div
            className={`max-w-2xl w-full max-h-[85vh] overflow-y-auto rounded-3xl border p-6 shadow-2xl space-y-5 ${
              darkMode
                ? 'bg-slate-900 border-slate-800 text-slate-100'
                : 'bg-white border-slate-200 text-slate-900'
            }`}
          >
            <div className="flex items-center justify-between pb-3 border-b border-slate-700/50">
              <div className="flex items-center gap-2.5">
                <div className="p-2 bg-indigo-500/10 text-indigo-400 rounded-xl">
                  <Info className="w-5 h-5" />
                </div>
                <h3 className="text-base font-bold">Instruksi & Cara Kerja Sistem ReconDoc</h3>
              </div>
              <button
                type="button"
                onClick={() => setShowInstructionsModal(false)}
                className="p-2 rounded-xl text-slate-400 hover:bg-slate-800 hover:text-white transition-colors"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <div
              className={`space-y-3.5 text-xs leading-relaxed ${
                darkMode ? 'text-slate-300' : 'text-slate-600'
              }`}
            >
              <div
                className={`p-3.5 rounded-2xl border space-y-1.5 ${
                  darkMode ? 'bg-slate-950/50 border-slate-800' : 'bg-slate-50 border-slate-200'
                }`}
              >
                <div className="flex items-center gap-2 font-bold text-blue-400">
                  <CheckCircle2 className="w-4 h-4" /> 1. Mode Excel & CSV — Komparasi File
                </div>
                <ul
                  className={`list-disc list-inside space-y-1 pl-1 ${
                    darkMode ? 'text-slate-400' : 'text-slate-600'
                  }`}
                >
                  <li>
                    <strong>Unggah 2 Dokumen:</strong> Masukkan File 1 (Referensi/Lama) dan File 2 (Pembanding/Baru), lalu pilih lembar <em>sheet</em>.
                  </li>
                  <li>
                    <strong>Pencocokan Baris:</strong> Pilih <em>Primary Key</em> utama untuk mendeteksi baris Identik, Beda (<em>Modified</em>), Baru (<em>Added</em>), dan Dihapus (<em>Deleted</em>), lalu ekspor hasilnya ke Excel (<code>.xlsx</code>).
                  </li>
                </ul>
              </div>

              <div
                className={`p-3.5 rounded-2xl border space-y-1.5 ${
                  darkMode ? 'bg-slate-950/50 border-slate-800' : 'bg-slate-50 border-slate-200'
                }`}
              >
                <div className="flex items-center gap-2 font-bold text-emerald-400">
                  <CheckCircle2 className="w-4 h-4" /> 2. Mode Excel & CSV — Kalkulasi Rumus
                </div>
                <ul
                  className={`list-disc list-inside space-y-1 pl-1 ${
                    darkMode ? 'text-slate-400' : 'text-slate-600'
                  }`}
                >
                  <li>
                    <strong>Susun Formula:</strong> Aktifkan kalkulasi, pilih kolom target hasil, dan tulis rumus menggunakan tanda kurung siku seperti <code className="text-emerald-400">[Gaji_Pokok] + [Tunjangan]</code>.
                  </li>
                  <li>
                    <strong>Ekspor Hasil:</strong> Unduh tabel hasil perhitungan langsung ke format Excel (<code>.xlsx</code>).
                  </li>
                </ul>
              </div>

              <div
                className={`p-3.5 rounded-2xl border space-y-1.5 ${
                  darkMode ? 'bg-slate-950/50 border-slate-800' : 'bg-slate-50 border-slate-200'
                }`}
              >
                <div className="flex items-center gap-2 font-bold text-purple-400">
                  <CheckCircle2 className="w-4 h-4" /> 3. Mode Excel & CSV — Pivot Table & Rekap
                </div>
                <ul
                  className={`list-disc list-inside space-y-1 pl-1 ${
                    darkMode ? 'text-slate-400' : 'text-slate-600'
                  }`}
                >
                  <li>
                    <strong>Penyatuan Otomatis Multi-File:</strong> Unggah satu atau beberapa file Excel/CSV dan satukan menjadi data gabungan (<code>.xlsx</code>).
                  </li>
                  <li>
                    <strong>Baris, Kolom & Agregasi:</strong> Atur baris (<strong>Rows</strong>) dan kolom (<strong>Columns</strong>), lalu pilih metode hitung (<code>COUNT</code>, <code>SUM</code>, <code>AVG</code>) dan ekspor ke Excel (<code>.xlsx</code>).
                  </li>
                </ul>
              </div>

              <div
                className={`p-3.5 rounded-2xl border space-y-1 ${
                  darkMode ? 'bg-slate-950/50 border-slate-800' : 'bg-slate-50 border-slate-200'
                }`}
              >
                <div className="flex items-center gap-2 font-bold text-indigo-400">
                  <CheckCircle2 className="w-4 h-4" /> 4. Mode Review Word (.docx)
                </div>
                <p className={darkMode ? 'text-slate-400' : 'text-slate-600'}>
                  Unggah dua dokumen Word (<code>.docx</code>) untuk membandingkan perbedaan teks secara langsung dan mengunduh laporannya.
                </p>
              </div>
            </div>

            <div className="pt-2 flex justify-end">
              <button
                type="button"
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