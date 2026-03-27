import { useState, useMemo } from 'react';
import TEMPLATES, { type TemplateItem } from '../data/templates';

type InsertField = 'initialText' | 'procedureText' | 'postText';

type Props = {
  onInsert: (field: InsertField, text: string, mode: 'append' | 'replace') => void;
  onUndo: () => void;
  canUndo: boolean;
  undoLabel?: string;
  templates?: TemplateItem[];
};

export default function TemplatePicker({ onInsert, onUndo, canUndo, undoLabel, templates: templatesProp }: Props) {
  const [q, setQ] = useState('');
  const [mode, setMode] = useState<'append' | 'replace'>('append');
  const templates = typeof templatesProp !== 'undefined' ? templatesProp : TEMPLATES;
  const filtered = useMemo(() => {
    const ql = q.trim().toLowerCase();
    if (!ql) return templates;
    return templates.filter(t =>
      t.title.toLowerCase().includes(ql) || t.text.toLowerCase().includes(ql) || (t.tags || []).some(tag => tag.includes(ql))
    );
  }, [q, templates]);

  return (
    <div className="mt-4 bg-white p-4 rounded-xl border border-slate-200">
      <div className="flex items-center gap-3 mb-3">
        <input value={q} onChange={e => setQ(e.target.value)} placeholder="キーワードで検索（例：抜歯、術後）"
          className="flex-1 border rounded px-3 py-2 text-sm outline-none" />
        <div className="text-xs text-slate-500">候補: {filtered.length}</div>
        <button onClick={onUndo} disabled={!canUndo} className="ml-2 px-3 py-1 bg-rose-100 text-rose-700 rounded text-xs disabled:opacity-40">
          元に戻す{undoLabel ? `: ${undoLabel}` : ''}
        </button>
      </div>

      <div className="flex items-center gap-3 text-xs mb-3">
        <label className="flex items-center gap-2"><input type="radio" checked={mode === 'append'} onChange={() => setMode('append')} /> 末尾追加</label>
        <label className="flex items-center gap-2"><input type="radio" checked={mode === 'replace'} onChange={() => setMode('replace')} /> 置き換え</label>
      </div>

      <div className="space-y-2 max-h-56 overflow-y-auto pr-2">
        {filtered.map((t: TemplateItem) => (
          <div key={t.id} className="p-3 border rounded-lg flex justify-between items-start gap-3">
            <div>
              <div className="font-semibold text-sm">{t.title}</div>
              <div className="text-xs text-slate-600 mt-1">{t.text}</div>
            </div>
            <div className="flex flex-col gap-1">
              <button onClick={() => onInsert('initialText', t.text, mode)} className="px-3 py-1 bg-slate-100 rounded text-xs">初診に挿入</button>
              <button onClick={() => onInsert('procedureText', t.text, mode)} className="px-3 py-1 bg-slate-100 rounded text-xs">処置に挿入</button>
              <button onClick={() => onInsert('postText', t.text, mode)} className="px-3 py-1 bg-slate-100 rounded text-xs">術後に挿入</button>
            </div>
          </div>
        ))}
      </div>
    </div>
  );
}
