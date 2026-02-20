import React from 'react';
import type { LayoutOptions } from '../types';




interface Props {
  options: LayoutOptions;
  setOptions: React.Dispatch<React.SetStateAction<LayoutOptions>>;
}

const LayoutControls: React.FC<Props> = ({ options, setOptions }) => {
  const handleChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const { name, value } = e.target;
    setOptions(prev => ({
      ...prev,
      [name]: name === 'backgroundColor' ? value : Number(value)
    }));
  };

  return (
    <div className="bg-white p-6 rounded-2xl shadow-sm border border-slate-200 space-y-6">
      <div className="flex items-center justify-between">
        <h3 className="font-bold text-slate-800 text-base">黒枠フィット設定</h3>
        <span className="px-2 py-1 bg-indigo-50 text-indigo-600 text-[10px] font-bold rounded">Auto-Scale</span>
      </div>
      
      <div className="space-y-4">
        <div>
          <div className="flex justify-between items-center mb-2">
            <label className="text-[10px] font-black text-slate-500 uppercase tracking-wider">
              基本の画像間隔 (px)
            </label>
            <span className="text-xs font-bold text-indigo-600">{options.spacing}px</span>
          </div>
          <input
            type="range"
            name="spacing"
            min="0"
            max="40"
            step="1"
            value={options.spacing}
            onChange={handleChange}
            className="w-full h-1.5 bg-slate-100 rounded-lg appearance-none cursor-pointer accent-indigo-600"
          />
          <p className="mt-2 text-[9px] text-slate-400 leading-tight">
            ※枠から溢れる場合は、この間隔も自動的に縮小調整されます。
          </p>
        </div>

        <div>
          <label className="block text-[10px] font-black text-slate-500 uppercase tracking-wider mb-2">
            背景色
          </label>
          <div className="flex items-center gap-3 p-2 bg-slate-50 rounded-xl border border-slate-100">
            <input
              type="color"
              name="backgroundColor"
              value={options.backgroundColor}
              onChange={handleChange}
              className="h-8 w-12 border-none cursor-pointer rounded bg-transparent"
            />
            <span className="text-xs font-mono font-bold uppercase text-slate-500">{options.backgroundColor}</span>
          </div>
        </div>
      </div>

      <div className="p-3 bg-indigo-50 rounded-xl border border-indigo-100">
        <h4 className="text-[10px] font-black text-indigo-600 uppercase mb-1">現在のターゲット領域</h4>
        <p className="text-[9px] text-indigo-400 font-medium leading-relaxed">
          スライド右端から 0.32cm / 下端から 1.4cm の余白。画像が領域を超える場合は、アスペクト比を維持したまま全体が自動縮小されます。
        </p>
      </div>
    </div>
  );
};

export default LayoutControls;
