import React from 'react';
import { X, CheckCircle, XCircle } from 'lucide-react';

export default function DetailModal({ row, headers, onClose }) {
  if (!row) return null;

  return (
    <div className="fixed inset-0 z-50 bg-slate-900/50 backdrop-blur-sm flex items-center justify-center p-4">
      <div className="bg-white rounded-2xl shadow-xl max-w-3xl w-full max-h-[85vh] flex flex-col overflow-hidden">
        {/* Header Modal */}
        <div className="p-4 border-b border-slate-100 flex items-center justify-between bg-slate-50">
          <div>
            <h3 className="font-bold text-slate-800 text-base">Detail Audit Baris: {row.key}</h3>
            <p className="text-xs text-slate-500">Status Komparasi: <span className="font-semibold text-indigo-600">{row.status}</span></p>
          </div>
          <button onClick={onClose} className="p-1 text-slate-400 hover:text-slate-700 rounded-lg">
            <X className="w-5 h-5" />
          </button>
        </div>

        {/* Content Table */}
        <div className="p-4 overflow-y-auto space-y-2">
          <table className="w-full text-xs">
            <thead className="bg-slate-100 text-slate-600 font-semibold sticky top-0">
              <tr>
                <th className="p-2 text-left w-1/3">Nama Kolom</th>
                <th className="p-2 text-left w-1/3">Nilai File 1 (Lama)</th>
                <th className="p-2 text-left w-1/3">Nilai File 2 (Baru)</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-slate-100">
              {headers.map((h) => {
                const diff = row.diffs[h];
                const isEqual = diff ? diff.isEqual : true;
                const val1 = row.row1 ? row.row1[h] : "";
                const val2 = row.row2 ? row.row2[h] : "";

                return (
                  <tr key={h} className={!isEqual ? 'bg-amber-50/70 font-medium' : 'hover:bg-slate-50'}>
                    <td className="p-2 font-medium text-slate-700 border-r border-slate-100">{h}</td>
                    <td className={`p-2 border-r border-slate-100 ${!isEqual ? 'text-rose-700 bg-rose-50/50' : 'text-slate-600'}`}>
                      {val1 !== "" ? String(val1) : <span className="text-slate-300 italic">(kosong)</span>}
                    </td>
                    <td className={`p-2 ${!isEqual ? 'text-emerald-700 bg-emerald-50/50 font-bold' : 'text-slate-600'}`}>
                      {val2 !== "" ? String(val2) : <span className="text-slate-300 italic">(kosong)</span>}
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>

        {/* Footer */}
        <div className="p-4 border-t border-slate-100 bg-slate-50 flex justify-end">
          <button
            onClick={onClose}
            className="px-4 py-2 bg-slate-800 text-white text-xs font-semibold rounded-lg hover:bg-slate-900 transition-colors"
          >
            Tutup
          </button>
        </div>
      </div>
    </div>
  );
}