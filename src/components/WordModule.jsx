import React from 'react';
import { FileText, Download, Upload } from 'lucide-react';

export default function WordModule({ 
  darkMode, 
  wordFile1, 
  wordFile2, 
  handleUploadWord1, 
  handleUploadWord2, 
  wordText1,
  wordText2,
  wordDiffResult 
}) {

  // Fungsi untuk mengunduh laporan review komparasi teks Word
  const exportWordResults = () => {
    if (!wordDiffResult) {
      alert("Belum ada teks dokumen yang dikomparasi!");
      return;
    }

    let reportContent = "=== LAPORAN REVIEW & KOMPARASI DOKUMEN WORD ===\n\n";
    reportContent += `Dokumen 1 (Referensi): ${wordFile1 ? wordFile1.fileName : 'File 1'}\n`;
    reportContent += `Dokumen 2 (Pembanding): ${wordFile2 ? wordFile2.fileName : 'File 2'}\n\n`;
    reportContent += "----------------------------------------\n\n";

    wordDiffResult.forEach((part) => {
      const prefix = part.added ? "[TAMBAHAN BARU]: " : part.removed ? "[DIHAPUS/LAMA]: " : "[TETAP/SAMA]: ";
      reportContent += `${prefix} ${part.value}\n`;
    });

    const blob = new Blob([reportContent], { type: 'text/plain;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = 'Laporan_Review_Word.txt';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
  };

  return (
    <div className="space-y-6 animate-fade-in">
      {/* Bagian Unggah Dokumen Word */}
      <section className={`rounded-3xl border shadow-sm p-6 space-y-6 ${darkMode ? 'bg-slate-900/90 border-slate-800' : 'bg-white border-slate-200'}`}>
        <div className="flex items-center justify-between pb-4 border-b border-slate-700/50">
          <div>
            <h2 className="text-base font-bold flex items-center gap-2">
              <div className="p-1.5 bg-indigo-500/10 text-indigo-400 rounded-xl">
                <FileText className="w-4 h-4" />
              </div>
              Mode Review & Komparasi Dokumen Word (.docx)
            </h2>
            <p className={`text-xs mt-0.5 ${darkMode ? 'text-slate-400' : 'text-slate-500'}`}>
              Unggah dua dokumen Word untuk mendeteksi perbedaan secara berdampingan.
            </p>
          </div>
        </div>

        <div className="grid grid-cols-1 md:grid-cols-2 gap-5">
          {/* Upload Dokumen 1 */}
          <div className={`border-2 border-dashed rounded-3xl p-5 relative group overflow-hidden ${darkMode ? 'border-slate-800 bg-slate-950/40 hover:border-indigo-500/50' : 'border-slate-200 bg-slate-50/50 hover:border-indigo-400'}`}>
            <input 
              type="file" 
              accept=".docx" 
              onChange={(e) => e.target.files[0] && handleUploadWord1(e.target.files[0])} 
              className="absolute inset-0 w-full h-full opacity-0 cursor-pointer z-10" 
            />
            <div className="flex items-start gap-4">
              <div className="p-3.5 bg-indigo-500/10 text-indigo-400 rounded-2xl">
                <Upload className="w-6 h-6" />
              </div>
              <div className="flex-1 min-w-0">
                <span className="text-[10px] font-bold text-indigo-400 uppercase tracking-widest block mb-1">DOKUMEN WORD 1 (REFERENSI)</span>
                <p className={`text-sm font-semibold truncate ${darkMode ? 'text-slate-200' : 'text-slate-800'}`}>
                  {wordFile1 ? wordFile1.fileName : "Pilih file .docx..."}
                </p>
              </div>
            </div>
          </div>

          {/* Upload Dokumen 2 */}
          <div className={`border-2 border-dashed rounded-3xl p-5 relative group overflow-hidden ${darkMode ? 'border-slate-800 bg-slate-950/40 hover:border-indigo-500/50' : 'border-slate-200 bg-slate-50/50 hover:border-indigo-400'}`}>
            <input 
              type="file" 
              accept=".docx" 
              onChange={(e) => e.target.files[0] && handleUploadWord2(e.target.files[0])} 
              className="absolute inset-0 w-full h-full opacity-0 cursor-pointer z-10" 
            />
            <div className="flex items-start gap-4">
              <div className="p-3.5 bg-indigo-500/10 text-indigo-400 rounded-2xl">
                <Upload className="w-6 h-6" />
              </div>
              <div className="flex-1 min-w-0">
                <span className="text-[10px] font-bold text-indigo-400 uppercase tracking-widest block mb-1">DOKUMEN WORD 2 (PEMBANDING)</span>
                <p className={`text-sm font-semibold truncate ${darkMode ? 'text-slate-200' : 'text-slate-800'}`}>
                  {wordFile2 ? wordFile2.fileName : "Pilih file .docx..."}
                </p>
              </div>
            </div>
          </div>
        </div>
      </section>

      {/* Hasil Review Berdampingan dengan Penyorotan Warna Masing-Masing */}
      {wordDiffResult && (
        <section className={`rounded-3xl border shadow-sm p-6 space-y-4 ${darkMode ? 'bg-slate-900/90 border-slate-800' : 'bg-white border-slate-200'}`}>
          <div className="flex flex-col sm:flex-row items-center justify-between gap-4 pb-3 border-b border-slate-700/50">
            <div>
              <h3 className="text-sm font-bold">Perbandingan Teks Berdampingan (Side-by-Side Review)</h3>
              <p className="text-[11px] text-slate-400">
                Kolom kiri menyoroti bagian yang <span className="text-rose-400 font-bold">dihapus/berbeda</span>, dan kolom kanan menyoroti bagian <span className="text-emerald-400 font-bold">tambahan baru</span>.
              </p>
            </div>
            
            <button 
              onClick={exportWordResults}
              className="flex items-center gap-2 px-4 py-2 bg-emerald-600 hover:bg-emerald-500 text-white font-bold text-xs rounded-2xl shadow-md transition-all whitespace-nowrap"
            >
              <Download className="w-4 h-4" />
              <span>Unduh Laporan Review (.txt)</span>
            </button>
          </div>

          <div className="grid grid-cols-1 md:grid-cols-2 gap-5">
            {/* Sisi Kiri: Dokumen 1 (Menandai teks yang dihapus/removed dengan warna merah) */}
            <div className={`p-4 rounded-2xl border text-xs leading-loose max-h-[450px] overflow-y-auto whitespace-pre-wrap font-sans ${darkMode ? 'bg-slate-950 border-slate-800 text-slate-300' : 'bg-slate-50 border-slate-200 text-slate-800'}`}>
              <div className="text-[11px] font-bold uppercase tracking-wider text-rose-400 pb-2 mb-2 border-b border-slate-700/50">
                Dokumen 1: {wordFile1?.fileName} (Teks Asal / Referensi)
              </div>
              {wordDiffResult.map((part, index) => {
                if (part.added) return null;
                const styleClass = part.removed 
                  ? 'bg-rose-500/20 text-rose-300 line-through px-1.5 py-0.5 rounded font-medium border border-rose-500/30 mx-0.5 inline-block' 
                  : 'text-slate-300';
                return (
                  <span key={index} className={styleClass}>
                    {part.value}
                  </span>
                );
              })}
            </div>

            {/* Sisi Kanan: Dokumen 2 (Menandai teks baru/added dengan warna hijau) */}
            <div className={`p-4 rounded-2xl border text-xs leading-loose max-h-[450px] overflow-y-auto whitespace-pre-wrap font-sans ${darkMode ? 'bg-slate-950 border-slate-800 text-slate-300' : 'bg-slate-50 border-slate-200 text-slate-800'}`}>
              <div className="text-[11px] font-bold uppercase tracking-wider text-emerald-400 pb-2 mb-2 border-b border-slate-700/50">
                Dokumen 2: {wordFile2?.fileName} (Teks Baru / Pembanding)
              </div>
              {wordDiffResult.map((part, index) => {
                if (part.removed) return null;
                const styleClass = part.added 
                  ? 'bg-emerald-500/20 text-emerald-300 px-1.5 py-0.5 rounded font-medium border border-emerald-500/30 mx-0.5 inline-block' 
                  : 'text-slate-300';
                return (
                  <span key={index} className={styleClass}>
                    {part.value}
                  </span>
                );
              })}
            </div>
          </div>
        </section>
      )}
    </div>
  );
}