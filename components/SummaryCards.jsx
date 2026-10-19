import React from 'react';
import { Layers, CheckCircle, AlertTriangle, PlusCircle, MinusCircle } from 'lucide-react';

export default function SummaryCards({ counts, filter, setFilter }) {
  const cards = [
    { key: 'ALL', label: 'Total Baris', count: counts.total, icon: Layers, color: 'border-slate-300 text-slate-700 bg-slate-50' },
    { key: 'MATCHED', label: 'Cocok Sesuai', count: counts.matched, icon: CheckCircle, color: 'border-emerald-300 text-emerald-700 bg-emerald-50' },
    { key: 'MODIFIED', label: 'Ada Perbedaan', count: counts.modified, icon: AlertTriangle, color: 'border-amber-300 text-amber-700 bg-amber-50' },
    { key: 'ADDED', label: 'Hanya di File 2', count: counts.added, icon: PlusCircle, color: 'border-blue-300 text-blue-700 bg-blue-50' },
    { key: 'DELETED', label: 'Hanya di File 1', count: counts.deleted, icon: MinusCircle, color: 'border-rose-300 text-rose-700 bg-rose-50' },
  ];

  return (
    <div className="grid grid-cols-2 md:grid-cols-5 gap-3">
      {cards.map((c) => {
        const Icon = c.icon;
        const isActive = filter === c.key;
        return (
          <button
            key={c.key}
            onClick={() => setFilter(c.key)}
            className={`p-4 rounded-xl border transition-all text-left flex flex-col justify-between ${c.color} ${
              isActive ? 'ring-2 ring-indigo-600 shadow-md scale-102' : 'hover:opacity-90 opacity-80'
            }`}
          >
            <div className="flex items-center justify-between mb-2">
              <span className="text-xs font-semibold uppercase tracking-wider">{c.label}</span>
              <Icon className="w-4 h-4" />
            </div>
            <div className="text-2xl font-bold">{c.count}</div>
          </button>
        );
      })}
    </div>
  );
}