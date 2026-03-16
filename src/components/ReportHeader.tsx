import React from 'react';

export type ReportHeaderProps = {
  handleClearReportFields: () => void;
  reportFields: {
    reportDate: string;
  };
  getEmptyFieldToneClass: (value: unknown) => string;
  openCalendar: (field: 'reportDate') => void;
};

const ReportHeader: React.FC<ReportHeaderProps> = ({
  handleClearReportFields,
  reportFields,
  getEmptyFieldToneClass,
  openCalendar,
}) => (
  <div className="flex items-center justify-between gap-3 mb-4 pb-2 border-b border-slate-200">
    <div>
      <h3 className="text-lg font-semibold text-slate-800 tracking-tight">報告書データ入力</h3>
    </div>
    <div className="flex items-center gap-3">
      <label className="text-xs font-semibold text-slate-700 uppercase tracking-widest whitespace-nowrap">報告日</label>
      <div className="w-48 relative" data-date-field="reportDate">
        <input
          className={`w-full h-11 border rounded-xl px-3 py-2 text-base focus:ring-2 focus:ring-orange-500 outline-none transition-all cursor-pointer ${getEmptyFieldToneClass(reportFields.reportDate)} bg-white`}
          placeholder="202X年XX月XX日"
          value={reportFields.reportDate}
          readOnly
          onClick={() => openCalendar('reportDate')}
        />
      </div>
      <button
        type="button"
        onClick={handleClearReportFields}
        className="h-11 px-3 rounded-xl border border-slate-200 bg-white text-sm font-semibold text-slate-700 hover:bg-slate-50"
      >
        全ての入力クリア
      </button>
    </div>
  </div>
);

export default ReportHeader;
