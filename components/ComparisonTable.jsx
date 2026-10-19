import React from 'react';
import { Eye, CheckCircle2, AlertCircle, FilePlus, FileMinus } from 'lucide-react';

export default function ComparisonTable({ results, headers, onSelectRow }) {
  const getBadge = (status) => {
    switch (status) {
      case 'MATCHED':
        return <span className="inline-flex items-center gap-1 text-xs font-medium px-2.5 py-1 rounded-full bg-emerald-100 text-emerald-800"><CheckCircle2 className="w-3.5 h-3.5" /> Sesuai</span>;
      case 'MODIFIED':
        return <span className="inline-flex items-center gap-1 text-xs font-medium px-2.5 py-1 rounded-full bg-amber-100 text-amber-800"><AlertCircle className="w-3.5 h-3.5" /> Beda</span>;
      case 'ADDED':
        return <span className="inline-flex items-center gap-1 text-xs font-medium px-2.5 py-1 rounded-full bg-blue-100 text-blue-800"><FilePlus className="w-3.5 h-3.5" /> Baru (F2)</span>;
      case 'DELETED':
        return <span className="inline-flex items-center gap-1 text-xs font-medium px-2.5 py-1 rounded-full bg-rose-100 text-rose-800"><FileMinus className="w-3.5 h-3.5" /> Hapus (F1)</span>;
      default:
        return null;
    }
  };

  return (
    <div className="bg-white rounded-xl border border-slate-200 shadow-sm overflow-hidden">
      <div className="overflow-x-auto max-h-[500px]">
        <table className="w-full text-left text-xs">
          <thead className="bg-slate-100 text-slate-700 uppercase font-semibold sticky top-0 z-10 border-b border-slate-200">
            <tr>
              <th className="p-3 w-12 text-center">Aksi</th>
              <th className="p-3 w-28">Status</th>
              <th className="p-3 w-36">Kunci / Identifier</th>
              {headers.slice(0, 5).map((h) => (
                <th key={h} className="p-3 min-w-[140px]">{h}</th>
              ))}
              {headers.length > 5 && <th className="p-3">+ {headers.length - 5} Kolom Lainnya</th>}
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-100">
            {results.length === 0 ? (
              <tr>
                <td colSpan={headers.length + 3} className="p-8 text-center text-slate-400">
                  Tidak ada data yang cocok dengan filter saat ini.
                </td>
              </tr>
            ) : (
              results.map((row) => (
                <tr key={row.id} className="hover:bg-slate-50 transition-colors">
                  <td className="p-3 text-center">
                    <button
                      onClick={() => onSelectRow(row)}
                      className="p-1.5 text-indigo-600 hover:bg-indigo-50 rounded-lg transition-colors"
                      title="Lihat Detail Komparasi Side-by-Side"
                    >
                      <Eye className="w-4 h-4" />
                    </button>
                  </td>
                  <td className="p-3">{getBadge(row.status)}</td>
                  <td className="p-3 font-medium text-slate-800">{row.key}</td>
                  {headers.slice(0, 5).map((h) => {
                    const diff = row.diffs[h];
                    const isDiff = diff && !diff.isEqual;
                    return (
                      <td
                        key={h}
                        className={`p-3 ${isDiff ? 'bg-amber-50 font-semibold text-amber-900 border-l-2 border-amber-400' : 'text-slate-600'}`}
                      >
                        {row.status === 'MATCHED' ? (
                          row.row1 ? row.row1[h] : "-"
                        ) : (
                          <div className="flex flex-col gap-0.5">
                            <span className="text-slate-700">{row.row1 ? row.row1[h] : <i className="text-slate-300">-</i>}</span>
                            {isDiff && (
                              <span className="text-rose-600 font-bold text-[11px] border-t border-dashed border-amber-200 pt-0.5">
                                ➔ {row.row2 ? row.row2[h] : <i className="text-slate-300">-</i>}
                              </span>
                            )}
                          </div>
                        )}
                      </td>
                    );
                  })}
                  {headers.length > 5 && (
                    <td className="p-3 text-slate-400 italic">Klik tombol mata untuk melihat semua kolom</td>
                  )}
                </tr>
              ))
            )}
          </tbody>
        </table>
      </div>
    </div>
  );
}