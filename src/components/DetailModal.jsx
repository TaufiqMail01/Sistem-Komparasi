import React from 'react';
import { Database, X } from 'lucide-react';
import { formatCellValue } from '../utils/fileHelpers';

export default function DetailModal({ selectedRowDetail, comparisonResult, darkMode, setSelectedRowDetail }) {
  if (!selectedRowDetail || !comparisonResult) return null;

  return (
    <div className="fixed inset-0 z-50 bg-slate-950/70 backdrop-blur-sm flex items-center justify-center p-4 animate-fade-in">
      <div className={`rounded-3xl shadow-2xl max-w-4xl w-full max-h-[90vh] flex flex-col overflow-hidden border ${darkMode ? 'bg-slate-900 border-slate-800 text-slate-100' : 'bg-white border-slate-200 text-slate-800'}`}>
        <div className={`p-5 border-b flex items-center justify-between ${darkMode ? 'bg-slate-950 border-slate-800 text-white' : 'bg-slate-900 text-white border-slate-100'}`}>
          <div className="flex items-center gap-3.5">
            <div className="p-2.5 bg-indigo-600 rounded-2xl text-white shadow-md">
              <Database className="w-5 h-5" />
            </div>
            <div>
              <h3 className="font-bold text-base tracking-tight">Detail Audit Baris / Record</h3>
              <p className="text-xs text-slate-300 mt-0.5">Composite Primary Key: <span className="font-mono text-amber-400 font-bold">{selectedRowDetail.key}</span></p>
            </div>
          </div>
          <button onClick={() => setSelectedRowDetail(null)} className="p-2 rounded-xl bg-slate-800 hover:bg-slate-700 text-slate-300 hover:text-white transition-colors">
            <X className="w-5 h-5" />
          </button>
        </div>
        
        <div className={`p-6 overflow-y-auto space-y-4 flex-1 ${darkMode ? 'bg-slate-950/40' : 'bg-slate-50/50'}`}>
          <div className={`p-4 rounded-2xl text-xs flex items-center justify-between shadow-2xs border ${darkMode ? 'bg-amber-950/20 border-amber-900/50 text-amber-200' : 'bg-amber-50/80 border-amber-200/80 text-amber-900'}`}>
            <span className="font-medium">Perbandingan atribut data secara presisi berdasarkan Primary Key.</span>
            <span className={`font-bold px-3 py-1 rounded-xl ${darkMode ? 'bg-amber-900/50 text-amber-100' : 'bg-amber-200 text-amber-950'}`}>{selectedRowDetail.diffCount} Kolom Berbeda</span>
          </div>

          <div className={`rounded-2xl border shadow-2xs overflow-hidden ${darkMode ? 'bg-slate-900 border-slate-800' : 'bg-white border-slate-200/80'}`}>
            <table className="w-full text-xs">
              <thead className={`font-bold uppercase tracking-wider border-b ${darkMode ? 'bg-slate-800 text-slate-300 border-slate-700' : 'bg-slate-100/90 text-slate-700 border-slate-200'}`}>
                <tr>
                  <th className="p-3.5 text-left w-1/3">Nama Kolom / Atribut</th>
                  <th className={`p-3.5 text-left w-1/3 ${darkMode ? 'text-slate-400' : 'text-slate-600'}`}>File 1 (Data Lama / Acuan)</th>
                  <th className={`p-3.5 text-left w-1/3 ${darkMode ? 'text-slate-200' : 'text-slate-800'}`}>File 2 (Data Baru / Pembanding)</th>
                </tr>
              </thead>
              <tbody className={`divide-y ${darkMode ? 'divide-slate-800' : 'divide-slate-100'}`}>
                {comparisonResult.headers.map(h => {
                  const diff = selectedRowDetail.diffs[h];
                  const isEqual = diff ? diff.isEqual : true;
                  return (
                    <tr key={h} className={!isEqual ? (darkMode ? 'bg-amber-950/30 font-semibold' : 'bg-amber-50/60 font-semibold') : (darkMode ? 'hover:bg-slate-800/40' : 'hover:bg-slate-50/50')}>
                      <td className={`p-3.5 font-medium border-r flex items-center gap-2.5 ${darkMode ? 'border-slate-800 text-slate-200' : 'border-slate-100 text-slate-800'}`}>
                        <span className={`w-2 h-2 rounded-full ${!isEqual ? 'bg-amber-500' : 'bg-slate-500'}`}></span>
                        {h}
                      </td>
                      <td className={`p-3.5 border-r font-mono ${darkMode ? 'border-slate-800' : 'border-slate-100'} ${!isEqual ? (darkMode ? 'text-rose-400 bg-rose-950/30' : 'text-rose-700 bg-rose-50/40') : (darkMode ? 'text-slate-400' : 'text-slate-600')}`}>
                        {formatCellValue(h, selectedRowDetail.row1 ? selectedRowDetail.row1[h] : "")}
                      </td>
                      <td className={`p-3.5 font-mono ${!isEqual ? (darkMode ? 'text-emerald-400 bg-emerald-950/30 font-bold' : 'text-emerald-700 bg-emerald-50/40 font-bold') : (darkMode ? 'text-slate-400' : 'text-slate-600')}`}>
                        {formatCellValue(h, selectedRowDetail.row2 ? selectedRowDetail.row2[h] : "")}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </div>

        <div className={`p-4 border-t flex justify-end ${darkMode ? 'bg-slate-900 border-slate-800' : 'bg-white border-slate-100'}`}>
          <button onClick={() => setSelectedRowDetail(null)} className="px-5 py-2.5 bg-slate-800 hover:bg-slate-700 text-white text-xs font-bold rounded-2xl transition-all shadow-sm">
            Tutup Jendela Detail
          </button>
        </div>
      </div>
    </div>
  );
}