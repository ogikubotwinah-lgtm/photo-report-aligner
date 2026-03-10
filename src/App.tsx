import React, { useState, useRef, useCallback, useMemo, useEffect } from 'react';
import type { CSSProperties } from 'react';
import type { ImageData, LayoutOptions } from './types';
import LayoutControls from './components/LayoutControls';
import pptxgen from 'pptxgenjs';
import html2canvas from 'html2canvas';
import jsPDF from 'jspdf';
import { LAYOUT } from './layout';
import TemplatePicker from './components/TemplatePicker';
import { buildSvgTextParts, addPptxText } from './reportTextRenderer';
import RowBoard from './components/RowBoard';
import { fetchSuggestions, addRefHospital } from "./serverApi";
import { createPortal } from "react-dom";


type AppSuggestions = {
  refHospitals: string[];
  doctors: string[];
  refHospitalEmails: Record<string, string>;
};

const REPORT_FIELDS_STORAGE_KEY = "photo-report-aligner:reportFields";
const SERVER_BASE_URL = (import.meta as any).env?.VITE_SERVER_BASE_URL?.trim() || "http://localhost:8787";
const DOCTOR_SUFFIX = "先生";
const DEFAULT_CLOSING_MESSAGE = `添付の通り、治療報告書をお送りします。ご確認よろしくお願いいたします。

---
荻窪ツイン動物病院
（住所などは今は不要。後で追加）`;
// Tailwind v4 uses OKLCH-based CSS variables for its default palette.
// html2canvas may fail when computed styles contain `oklch(...)`.
// We override palette variables (HEX) only for the PDF export area (#print-area).
const PRINT_SAFE_CSS_VARS: CSSProperties = {
  ['--color-white' as any]: '#ffffff',
  ['--color-black' as any]: '#000000',

  // slate palette (commonly used in this UI)
  ['--color-slate-100' as any]: '#f1f5f9',
  ['--color-slate-200' as any]: '#e2e8f0',
  ['--color-slate-300' as any]: '#cbd5e1',
  ['--color-slate-400' as any]: '#94a3b8',
  ['--color-slate-600' as any]: '#475569',
  ['--color-slate-700' as any]: '#334155',
  ['--color-slate-800' as any]: '#1e293b',
  ['--color-slate-900' as any]: '#0f172a',
};
function getInitialReportFields() {
  return {
    reportDate: new Date().toLocaleDateString("ja-JP", {
      year: "numeric",
      month: "long",
      day: "numeric",
    }),
    refHospitalName: '',
    refHospital: '',
    refHospitalEmail: '',
    refDoctor: '',
    ownerLastName: '',
    petName: '',
    firstVisitDate: '',
    anesthesiaDate: '',
    sedationDate: '',
    attendingVet: '',
    initialText: '',
    procedureText: '',
    postText: '',
    page3Text: '',
    chiefComplaint: '',
    closingMessageText: DEFAULT_CLOSING_MESSAGE,
  };
}

function normalizeHospitalKey(name: string): string {
  return String(name)
    .trim()
    .replace(/\u3000/g, ' ')
    .replace(/\s+/g, ' ');
}

// Minimal date normalizer used by onBlur handlers in this file.
function normalizeJapaneseDate(input: string): string {
  if (!input) return input;
  const s = String(input).trim();

  if (s.includes('年')) return s;

  const onlyDigits = s.replace(/\D+/g, '');

  if (onlyDigits.length === 8) {
    const y = onlyDigits.slice(0, 4);
    const m = onlyDigits.slice(4, 6);
    const d = onlyDigits.slice(6, 8);
    return `${y}年${m}月${d}日`;
  }

  const match = s.match(/^(\d{4})[\/\-.](\d{1,2})[\/\-.](\d{1,2})$/);
  if (match) {
    const y = match[1];
    const mm = match[2].padStart(2, '0');
    const dd = match[3].padStart(2, '0');
    return `${y}年${mm}月${dd}日`;
  }

  return s;
}

function toDateInputValue(value: string): string {
  const s = String(value || '').trim();
  if (!s) return '';
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;

  const jp = s.match(/^(\d{4})年(\d{1,2})月(\d{1,2})日$/);
  if (jp) {
    const y = jp[1];
    const m = jp[2].padStart(2, '0');
    const d = jp[3].padStart(2, '0');
    return `${y}-${m}-${d}`;
  }

  const slash = s.match(/^(\d{4})[\/\-.](\d{1,2})[\/\-.](\d{1,2})$/);
  if (slash) {
    const y = slash[1];
    const m = slash[2].padStart(2, '0');
    const d = slash[3].padStart(2, '0');
    return `${y}-${m}-${d}`;
  }

  return '';
}

function fromDateInputValue(value: string): string {
  if (!value) return '';
  const m = value.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return value;
  return `${m[1]}年${m[2]}月${m[3]}日`;
}

function escapeSvgText(value: string): string {
  return String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

const PREVIEW_TEXT_Y_OFFSETS_CM = {
  initialText: 0,
  procedureText: 0,
  postText: 0,
  page3Text: 0,
  closingMessage: 0,
} as const;

type Page2PhotoLabelOption = '' | 'treatment-post' | 'inspection';

type PreviewTextOffsetPx = {
  initialText: number;
  procedureText: number;
  postText: number;
  page3Text: number;
  closingMessage: number;
};

function shiftSvgTextY(part: string, deltaPx: number): string {
  if (!deltaPx) return part;
  return part.replace(/ y="([^"]+)"/, (_m, y) => ` y="${Number(y) + deltaPx}"`);
}

function applyPreviewTextOffsets(pageNum: 1 | 2 | 3, parts: string[], offsets: PreviewTextOffsetPx): string[] {
  if (pageNum === 1) {
    let waitingInitialBody = false;
    return parts.map((part) => {
      if (part.includes('【初診時】')) {
        waitingInitialBody = true;
        return part;
      }
      if (waitingInitialBody && part.includes('<tspan') && !part.includes('【')) {
        waitingInitialBody = false;
        return shiftSvgTextY(part, offsets.initialText);
      }
      if (part.includes('という主訴の為、拝見いたしました。')) {
        return shiftSvgTextY(part, offsets.closingMessage);
      }
      return part;
    });
  }

  if (pageNum === 2) {
    let waitingProcedureBody = false;
    let waitingPostBody = false;
    return parts.map((part) => {
      if (part.includes('【検査・処置内容】')) {
        waitingProcedureBody = true;
        return part;
      }
      if (part.includes('【術後経過】')) {
        waitingPostBody = true;
        return part;
      }
      if (waitingProcedureBody && part.includes('<tspan')) {
        waitingProcedureBody = false;
        return shiftSvgTextY(part, offsets.procedureText);
      }
      if (waitingPostBody && part.includes('<tspan')) {
        waitingPostBody = false;
        return shiftSvgTextY(part, offsets.postText);
      }
      return part;
    });
  }

  let shiftedPage3 = false;
  return parts.map((part) => {
    if (!shiftedPage3 && part.includes('<tspan')) {
      shiftedPage3 = true;
      return shiftSvgTextY(part, offsets.page3Text);
    }
    return part;
  });
}

function estimateWrappedLineCount(text: string, maxWidthPx: number, fontPx: number, maxLines: number): number {
  const source = String(text || '').trim();
  if (!source) return 1;

  const canvas = document.createElement('canvas');
  const ctx = canvas.getContext('2d');
  if (!ctx) return 1;
  ctx.font = `${fontPx}px Meiryo, 'MS PGothic', 'Noto Sans JP', sans-serif`;

  let lines = 0;
  source.split('\n').forEach((paragraph) => {
    if (!paragraph) {
      lines += 1;
      return;
    }
    let current = '';
    for (const ch of paragraph) {
      const candidate = current + ch;
      if (ctx.measureText(candidate).width <= maxWidthPx || current.length === 0) {
        current = candidate;
      } else {
        lines += 1;
        current = ch;
      }
      if (lines >= maxLines) break;
    }
    if (lines < maxLines) lines += 1;
  });

  return Math.min(Math.max(lines, 1), maxLines);
}

type PageSwitcherProps = {
  currentPage: number;
  onChange: (page: number) => void;
  pages: number[];
  size?: "large" | "small";
};

const PageSwitcher: React.FC<PageSwitcherProps> = ({
  currentPage,
  onChange,
  pages,
  size = "small",
}) => {
  const isLarge = size === "large";

  return (
    <div
      className={
        isLarge
          ? "inline-flex items-center justify-center bg-slate-100 p-1.5 rounded-2xl border border-slate-200 space-x-2"
          : "inline-flex items-center justify-center space-x-2"
      }
    >
      {pages.map((p) => (
        <button
          key={p}
          onClick={() => onChange(p)}
          className={
            isLarge
              ? `w-[112px] h-[36px] inline-flex items-center justify-center gap-2 rounded-xl text-xs font-semibold transition-colors border ${
                  currentPage === p
                    ? "bg-violet-600 text-white border-violet-600"
                    : "bg-white text-slate-600 border-slate-300 hover:bg-slate-50"
                }`
              : `w-[88px] h-[36px] inline-flex items-center justify-center rounded-lg text-sm font-semibold transition-colors border ${
                  currentPage === p
                    ? "bg-violet-600 text-white border-violet-500"
                    : "bg-white text-slate-600 border-slate-300 hover:bg-slate-50"
                }`
          }
        >
          {isLarge ? (
            <>
              <svg
                className={`w-4 h-4 ${
                  currentPage === p ? "text-white" : "text-slate-300"
                }`}
                fill="none"
                stroke="currentColor"
                viewBox="0 0 24 24"
              >
                <path
                  strokeLinecap="round"
                  strokeLinejoin="round"
                  strokeWidth="2.5"
                  d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"
                />
              </svg>
              <span className="uppercase tracking-[0.2em]">PAGE {p}</span>
            </>
          ) : (
            <>PAGE {p}</>
          )}
        </button>
      ))}
    </div>
  );
};


        
        


const App: React.FC = () => {
  type DateFieldKey = 'firstVisitDate' | 'sedationDate' | 'anesthesiaDate';

  const [showPage3, setShowPage3] = useState(false);

  // ページ管理
  const [currentPage, setCurrentPage] = useState<number>(1);

  const [reportFieldsHydrated, setReportFieldsHydrated] = useState(false);
  const [calendarOpenField, setCalendarOpenField] = useState<DateFieldKey | null>(null);
  const [calendarViewDate, setCalendarViewDate] = useState<Date>(() => {
    const now = new Date();
    return new Date(now.getFullYear(), now.getMonth(), 1);
  });
  const firstVisitDateInputRef = useRef<HTMLInputElement | null>(null);
  const sedationDateInputRef = useRef<HTMLInputElement | null>(null);
  const anesthesiaDateInputRef = useRef<HTMLInputElement | null>(null);
  const page1CutLineBottomBaseYRef = useRef<number>((LAYOUT.PAGE1.TEXT as any).CUT_LINE_BOTTOM.y);
  const page1FixedClosingBaseYRef = useRef<number>((LAYOUT.PAGE1.TEXT as any).FIXED_CLOSING_TEXT.y);

  // PageSwitcher には常に PAGE1-3 を表示
  const availablePages = useMemo<number[]>(() => [1, 2, 3], []);

  const handlePageChange = useCallback((page: number) => {
    if (page === 3 && !showPage3) {
      setShowPage3(true);
    }
    setCurrentPage(page);
  }, [showPage3]);

  // showPage3 をOFFにした瞬間に 3ページ目に居たら 2へ戻す
  useEffect(() => {
    if (!showPage3 && currentPage === 3) setCurrentPage(2);
  }, [showPage3, currentPage]);

  // 既存：参照や他state
  const rowBoardRef = useRef<HTMLDivElement | null>(null);
  const fileInputRef = useRef<HTMLInputElement | null>(null);

  // ...この下に既存の state / useEffect / handlers が続く

  // ===== 参照病院（候補取得＆保存）=====
  const [refHospitalInput, setRefHospitalInput] = useState<string>("");
  const [refHospitalError, setRefHospitalError] = useState<string | null>(null);

  const [suggestions, setSuggestions] = useState<AppSuggestions>({
    refHospitals: [],
    doctors: [],
    refHospitalEmails: {},
  });
  // 初回ロード：候補一覧を取得
  useEffect(() => {
    fetchSuggestions()
      .then((data) => {
        setSuggestions({
          refHospitals: Array.isArray(data?.refHospitals) ? data.refHospitals : [],
          doctors: Array.isArray(data?.doctors) ? data.doctors : [],
          refHospitalEmails: data?.refHospitalEmails ?? {},
        });
      })
      .catch((e) => {
        console.error(e);
        setRefHospitalError("候補の取得に失敗しました（サーバ起動を確認）");
      });
  }, []);

  const normalizedRefHospitalEmails = useMemo(() => {
    const map: Record<string, string> = {};
    Object.entries(suggestions.refHospitalEmails || {}).forEach(([k, v]) => {
      map[normalizeHospitalKey(k)] = v;
    });
    return map;
  }, [suggestions.refHospitalEmails]);

  const applyRefHospitalSelection = useCallback((hospitalName: string) => {
    const normalizedName = normalizeHospitalKey(hospitalName);
    const mappedEmail = normalizedName ? normalizedRefHospitalEmails[normalizedName] : "";

    setRefHospitalInput(hospitalName);
    setReportFields((prev) => {
      return {
        ...prev,
        refHospitalName: hospitalName,
        refHospital: hospitalName,
        refHospitalEmail: mappedEmail || "",
      };
    });
  }, [normalizedRefHospitalEmails]);

  // 参照病院を保存
  const handleAddRefHospital = useCallback(
    async (nameArg?: string) => {
      const name = (nameArg ?? refHospitalInput).trim();
      if (!name) return;

      setRefHospitalError(null);

      try {
        await addRefHospital(name);

        // 最新候補を再取得
        const latest = await fetchSuggestions();
        setSuggestions({
          refHospitals: Array.isArray(latest?.refHospitals) ? latest.refHospitals : [],
          doctors: Array.isArray(latest?.doctors) ? latest.doctors : [],
          refHospitalEmails: latest?.refHospitalEmails ?? {},
        });

        // 入力欄とreportFieldsを正規化して揃える
        applyRefHospitalSelection(name);
      } catch (e) {
        console.error(e);
        setRefHospitalError("保存に失敗しました（CORS/サーバ/URL確認）");
      }
    },
    [refHospitalInput, applyRefHospitalSelection]
  );

  // ===== 画像（ページ別）=====
  const [allPagesImages, setAllPagesImages] = useState<Record<number, ImageData[]>>({
    1: [],
    2: [],
    3: [],
  });
  const [allPagesHistory, setAllPagesHistory] = useState<Record<number, ImageData[][]>>({
    1: [],
    2: [],
    3: [],
  });
  const [previewId, setPreviewId] = useState<string | null>(null);

  // 報告書テキスト入力ステート（この下に既存の reportFields を続けてOK）
  const [reportFields, setReportFields] = useState(getInitialReportFields);
  const page1DateOutputItems = useMemo(() => {
    return [
      { label: '初診日', value: String(reportFields.firstVisitDate || '').trim() },
      { label: '鎮静日', value: String(reportFields.sedationDate || '').trim() },
      { label: '全身麻酔日', value: String(reportFields.anesthesiaDate || '').trim() },
    ].filter((item) => item.value !== '');
  }, [reportFields.firstVisitDate, reportFields.sedationDate, reportFields.anesthesiaDate]);

  const filledDateCount = useMemo(() => {
    return [
      reportFields.firstVisitDate,
      reportFields.sedationDate,
      reportFields.anesthesiaDate,
    ].reduce((count, value) => (String(value || '').trim() !== '' ? count + 1 : count), 0);
  }, [reportFields.firstVisitDate, reportFields.sedationDate, reportFields.anesthesiaDate]);

  useEffect(() => {
    const page1Text = LAYOUT.PAGE1.TEXT as any;
    if (filledDateCount === 1) {
      page1Text.CUT_LINE_BOTTOM.y = 9.35;
    } else if (filledDateCount === 2) {
      page1Text.CUT_LINE_BOTTOM.y = 9.79;
    } else if (filledDateCount === 3) {
      page1Text.CUT_LINE_BOTTOM.y = 10.2;
    } else {
      page1Text.CUT_LINE_BOTTOM.y = page1CutLineBottomBaseYRef.current;
    }

    page1Text.FIXED_CLOSING_TEXT.y =
      filledDateCount === 3 ? 10.3 : page1FixedClosingBaseYRef.current;

    return () => {
      page1Text.CUT_LINE_BOTTOM.y = page1CutLineBottomBaseYRef.current;
      page1Text.FIXED_CLOSING_TEXT.y = page1FixedClosingBaseYRef.current;
    };
  }, [filledDateCount]);

  const getInputToneClass = useCallback(
    (value: string) => (value.trim() === '' ? 'border-amber-200 bg-amber-50/60' : 'border-slate-200 bg-white'),
    []
  );
  const getTextareaToneClass = useCallback(
    (value: string) => (value.trim() === '' ? 'border-amber-200 bg-white' : 'border-slate-200 bg-white'),
    []
  );

  useEffect(() => {
    const initial = getInitialReportFields();
    if (typeof window === "undefined") {
      setReportFieldsHydrated(true);
      return;
    }

    try {
      const raw = window.localStorage.getItem(REPORT_FIELDS_STORAGE_KEY);
      if (!raw) {
        setReportFieldsHydrated(true);
        return;
      }

      const saved = JSON.parse(raw);
      if (!saved || typeof saved !== "object") {
        setReportFieldsHydrated(true);
        return;
      }

      const restored = { ...initial, ...saved };
      if (!restored.sedationDate && restored.anesthesiaDate) {
        restored.sedationDate = restored.anesthesiaDate;
      }
      if (!restored.anesthesiaDate && restored.sedationDate) {
        restored.anesthesiaDate = restored.sedationDate;
      }
      if (!restored.closingMessageText) {
        restored.closingMessageText = DEFAULT_CLOSING_MESSAGE;
      }
      setReportFields(restored);
      setRefHospitalInput(String(restored.refHospitalName || restored.refHospital || ""));
    } catch {
      // JSON parse failure: ignore and keep initial values
    } finally {
      setReportFieldsHydrated(true);
    }
  }, []);

  useEffect(() => {
    if (!reportFieldsHydrated || typeof window === "undefined") return;
    try {
      window.localStorage.setItem(REPORT_FIELDS_STORAGE_KEY, JSON.stringify(reportFields));
    } catch {
      // storage failure should not block app usage
    }
  }, [reportFields, reportFieldsHydrated]);

  const handleClearReportFields = useCallback(() => {
    const initial = getInitialReportFields();
    setReportFields(initial);
    setRefHospitalInput("");
    if (typeof window !== "undefined") {
      window.localStorage.removeItem(REPORT_FIELDS_STORAGE_KEY);
    }
  }, []);

  const handleEnterToNextField = useCallback((
    e: React.KeyboardEvent<HTMLInputElement | HTMLSelectElement>,
    options?: { targetSelector?: string }
  ) => {
    if (e.key !== 'Enter' || e.nativeEvent.isComposing) return;
    e.preventDefault();

    const scope = e.currentTarget.closest('[data-enter-scope="report-fields"]') as HTMLElement | null;
    const root = scope ?? document.body;
    if (options?.targetSelector) {
      const target = root.querySelector<HTMLElement>(options.targetSelector);
      target?.focus();
      return;
    }
    const fields = Array.from(
      root.querySelectorAll<HTMLElement>(
        'input:not([type="hidden"]):not([type="file"]):not([type="checkbox"]):not([type="radio"]):not([disabled]), select:not([disabled])'
      )
    ).filter((el) => el.tabIndex !== -1);

    const currentIndex = fields.indexOf(e.currentTarget);
    if (currentIndex === -1) return;

    const next = fields[currentIndex + 1] as HTMLInputElement | HTMLSelectElement | undefined;
    next?.focus();
  }, []);

  const parseStoredDate = useCallback((value: string): Date | null => {
    const iso = toDateInputValue(value);
    if (!iso) return null;
    const [y, m, d] = iso.split('-').map((x) => parseInt(x, 10));
    if (!y || !m || !d) return null;
    return new Date(y, m - 1, d);
  }, []);

  const openDatePicker = useCallback(
    (key: DateFieldKey, ref: React.RefObject<HTMLInputElement | null>) => {
      setCalendarOpenField(key);
      const base = parseStoredDate(reportFields[key]) ?? new Date();
      setCalendarViewDate(new Date(base.getFullYear(), base.getMonth(), 1));
      ref.current?.focus();
    },
    [parseStoredDate, reportFields]
  );

  const weekdayLabels = ['日', '月', '火', '水', '木', '金', '土'];
  const calendarDays = useMemo(() => {
    const year = calendarViewDate.getFullYear();
    const month = calendarViewDate.getMonth();
    const firstDay = new Date(year, month, 1).getDay();
    const daysInMonth = new Date(year, month + 1, 0).getDate();
    const daysInPrevMonth = new Date(year, month, 0).getDate();
    const cells: Array<{ iso: string; day: number; inMonth: boolean }> = [];

    for (let i = firstDay - 1; i >= 0; i -= 1) {
      const day = daysInPrevMonth - i;
      const d = new Date(year, month - 1, day);
      const iso = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
      cells.push({ iso, day, inMonth: false });
    }

    for (let day = 1; day <= daysInMonth; day += 1) {
      const iso = `${year}-${String(month + 1).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
      cells.push({ iso, day, inMonth: true });
    }

    while (cells.length % 7 !== 0) {
      const nextDay = cells.length - (firstDay + daysInMonth) + 1;
      const d = new Date(year, month + 1, nextDay);
      const iso = `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(nextDay).padStart(2, '0')}`;
      cells.push({ iso, day: nextDay, inMonth: false });
    }

    return cells;
  }, [calendarViewDate]);

  const moveCalendarMonth = useCallback((delta: number) => {
    setCalendarViewDate((prev) => new Date(prev.getFullYear(), prev.getMonth() + delta, 1));
  }, []);

  const selectCalendarDate = useCallback((field: DateFieldKey, iso: string) => {
    setReportFields((prev) => ({ ...prev, [field]: fromDateInputValue(iso) }));
    setCalendarOpenField(null);
  }, []);

  const handleReportFieldsFocusCapture = useCallback((e: React.FocusEvent<HTMLDivElement>) => {
    const target = e.target as HTMLElement | null;
    if (!target) return;
    if (target.closest('[data-date-field-root="true"]')) return;
    setCalendarOpenField(null);
  }, []);

  // 現在のページのデータへのエイリアス
  const images = allPagesImages[currentPage];


  const history = allPagesHistory[currentPage];

  const setImages = useCallback((updater: any) => {
    setAllPagesImages(prevAll => {
      const currentImgs = prevAll[currentPage];
      const nextImgs = typeof updater === 'function' ? updater(currentImgs) : updater;
      return { ...prevAll, [currentPage]: nextImgs };
    });
  }, [currentPage]);

  const setHistory = useCallback((updater: any) => {
    setAllPagesHistory(prevAll => {
      const currentHist = prevAll[currentPage];
      const nextHist = typeof updater === 'function' ? updater(currentHist) : updater;
      return { ...prevAll, [currentPage]: nextHist };
    });
  }, [currentPage]);

  // ページごとの配置定義を取得するヘルパー
  const getPageDimensions = useCallback((page: number) => {
    if (page === 2) {
      const config = LAYOUT.PAGE2.IMAGES;
      return {
        left: config.LEFT,
        right: config.RIGHT,
        startY: config.START_Y,
        maxW: LAYOUT.SLIDE.WIDTH_CM - config.LEFT - config.RIGHT,
        maxH: config.END_Y - config.START_Y,
        alignLeft: config.ALIGN_LEFT
      };
    }
    if (page === 3) {
      const config = LAYOUT.PAGE3.IMAGES;
      return {
        left: config.LEFT,
        right: config.RIGHT,
        startY: config.START_Y,
        maxW: LAYOUT.SLIDE.WIDTH_CM - config.LEFT - config.RIGHT,
        maxH: config.END_Y - config.START_Y,
        alignLeft: config.ALIGN_LEFT
      };
    }
    const config = LAYOUT.PAGE1.IMAGES;
    return {
      left: config.LEFT,
      right: config.RIGHT,
      startY: config.START_Y,
      maxW: LAYOUT.SLIDE.WIDTH_CM - config.LEFT - config.RIGHT,
      maxH: LAYOUT.SLIDE.HEIGHT_CM - config.START_Y - config.MARGIN_BOTTOM,
      alignLeft: config.ALIGN_LEFT
    };
  }, []);


  const [options, setOptions] = useState<LayoutOptions>({
    spacing: 1,
    padding: 0,
    targetHeight: 250,
    containerWidth: 650,
    backgroundColor: '#ffffff'
  });

  const [pptxStatus, setPptxStatus] = useState<string>('');
  const [isSavingPptx, setIsSavingPptx] = useState(false);
  const [isCreatingDraft, setIsCreatingDraft] = useState(false);
  const [isPrintMode, setIsPrintMode] = useState(false);

  // テンプレ挿入の undo 用（直前の挿入を1回だけ戻す）
  const [lastInsert, setLastInsert] = useState<
    | { field: 'initialText' | 'procedureText' | 'postText'; prevValue: string }
    | null
  >(null);
  // --- ページ確定状態（既存のものがあるなら合わせて） ---
const [page1Confirmed, setPage1Confirmed] = useState(false);
const [page2Confirmed, setPage2Confirmed] = useState(false);
const [page3Confirmed, setPage3Confirmed] = useState(false);

// ページ順モード（あなたの既存仕様に合わせて）
  const [pageOrderMode] = useState<"page2-page3" | "page3-page2">("page2-page3");

// どこに「経過」を入れるか（既存がこれならOK）
const [postPlacement, setPostPlacement] = useState<"page2" | "page3">("page2");
const [isAttendingVetOpen, setIsAttendingVetOpen] = useState(false);
const [page2PhotoLabel, setPage2PhotoLabel] = useState<Page2PhotoLabelOption>('treatment-post');
const [isPage2PhotoLabelOpen, setIsPage2PhotoLabelOpen] = useState(false);
const [isPostPlacementOpen, setIsPostPlacementOpen] = useState(false);
const attendingVetDropdownRef = useRef<HTMLDivElement | null>(null);
const page2PhotoLabelDropdownRef = useRef<HTMLDivElement | null>(null);
const postPlacementDropdownRef = useRef<HTMLDivElement | null>(null);
const page2PhotoLabelText = useMemo(() => {
  if (page2PhotoLabel === 'treatment-post') return '【治療時・治療後写真】';
  if (page2PhotoLabel === 'inspection') return '【検査時写真】';
  return '';
}, [page2PhotoLabel]);

useEffect(() => {
  const closeIfOutside = (event: Event) => {
    const target = event.target as Node | null;
    if (!target) return;
    if (attendingVetDropdownRef.current?.contains(target)) return;
    if (page2PhotoLabelDropdownRef.current?.contains(target)) return;
    if (postPlacementDropdownRef.current?.contains(target)) return;
    setIsAttendingVetOpen(false);
    setIsPage2PhotoLabelOpen(false);
    setIsPostPlacementOpen(false);
  };

  document.addEventListener('mousedown', closeIfOutside);
  document.addEventListener('focusin', closeIfOutside);
  return () => {
    document.removeEventListener('mousedown', closeIfOutside);
    document.removeEventListener('focusin', closeIfOutside);
  };
}, []);


// 出力順（PDF/PPTXの並びなどで使う想定）
const outputPages = useMemo<number[]>(() => {
  if (!showPage3) return [1, 2];
  return pageOrderMode === "page2-page3" ? [1, 2, 3] : [1, 3, 2];
}, [showPage3, pageOrderMode]);

// 今いるページが確定済みか
const isCurrentPageConfirmed =
  (currentPage === 1 && page1Confirmed) ||
  (currentPage === 2 && page2Confirmed) ||
  (currentPage === 3 && page3Confirmed);

// Page3をOFFにしたら、ページ3に居座らない＆状態を整える
useEffect(() => {
  if (!showPage3) {
    if (currentPage === 3) setCurrentPage(2);
    setPostPlacement("page2");
    setPage3Confirmed(false);
  }
}, [showPage3, currentPage]);

  const recordHistory = useCallback(() => {
    setHistory((prev: ImageData[][]) => [...prev, [...images]].slice(-20));
  }, [images, setHistory]);

  const handleUndo = useCallback(() => {
    if (history.length === 0) return;
    const lastState = history[history.length - 1];
    setImages(lastState);
    setHistory((prev: ImageData[][]) => prev.slice(0, -1));
  }, [history, setImages, setHistory]);

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    console.log('handleFileUpload triggered');
    const input = e.currentTarget;
    const files = Array.from(input.files || []) as File[];
    if (!files.length) {
      input.value = '';
      return;
    }

    try {
      recordHistory();
      const newImages: ImageData[] = await Promise.all(
        files.map((file: File) => {
          return new Promise<ImageData>((resolve) => {
            const reader = new FileReader();
            reader.onload = (event) => {
              const dataUrl = event.target?.result as string;
              const img = new Image();
              img.onload = () => {
                resolve({
                  id: Math.random().toString(36).substr(2, 9),
                  name: file.name,
                  dataUrl,
                  width: img.width,
                  height: img.height,
                  mimeType: file.type,
                  row: 0,
                  orderConfirmed: false,
                  rotation: 0
                });
              };
              img.src = dataUrl;
            };
            reader.readAsDataURL(file);
          });
        })
      );
      setImages((prev: ImageData[]) => [...prev, ...newImages]);
      if (currentPage === 1) setPage1Confirmed(false);
      if (currentPage === 2) setPage2Confirmed(false);
      if (currentPage === 3) setPage3Confirmed(false);
    } finally {
      // 同一ファイル再選択でも onChange が必ず発火するようにリセット
      input.value = '';
    }
  };

  const rotateImage = (id: string, direction: 'left' | 'right') => {
    recordHistory();
    setImages((prev: ImageData[]) => prev.map(img => {
      if (img.id === id) {
        let newRotation = direction === 'right' ? img.rotation + 90 : img.rotation - 90;
        if (newRotation < 0) newRotation = 270;
        if (newRotation >= 360) newRotation = 0;
        return { ...img, rotation: newRotation };
      }
      return img;
    }));
  };

  const updateImageRow = (id: string, newRow: number) => {
    recordHistory();
    setImages((prev: ImageData[]) => {
      return prev.map(img =>
        img.id === id ? { ...img, row: newRow, orderConfirmed: false } : img
      );
    });
  };

  const removeImage = (id: string) => {
    recordHistory();
    setImages((prev: ImageData[]) => prev.filter(img => img.id !== id));
  };




  // 確定後に RowBoard へスクロール
  useEffect(() => {
    if (page1Confirmed || page2Confirmed || page3Confirmed) {
      rowBoardRef.current?.scrollIntoView({ behavior: 'smooth', block: 'start' });
    }
  }, [page1Confirmed, page2Confirmed, page3Confirmed]);

  const unassignedImages = useMemo(() => images.filter(img => img.row === 0), [images]);

  const calculateLayoutForAnyPage = useCallback((rowsData: ImageData[][], pageNum: number) => {
    const activeRows = rowsData.filter(row => row.length > 0);
    if (activeRows.length === 0) return { rowResults: [], finalHeight: 0 };
    const dims = getPageDimensions(pageNum);
    const containerW = options.containerWidth;
    const baseSpacing = options.spacing;
    const targetBoxHPx = (dims.maxH / dims.maxW) * containerW;
    let totalBlockHeight = 0;
    const rawData = activeRows.map(rowImages => {
      const count = rowImages.length;
      const ars = rowImages.map(img => {
        const isPortrait = img.rotation === 90 || img.rotation === 270;
        return isPortrait ? (img.height / img.width) : (img.width / img.height);
      });
      const totalAR = ars.reduce((a, b) => a + b, 0);
      const isFew = count <= 2;
      let rowH;
      if (isFew) {
        const avgAR = totalAR / count;
        const totalAR_for_5 = avgAR * 5;
        rowH = (containerW - (5 - 1) * baseSpacing) / totalAR_for_5;
      } else {
        rowH = (containerW - (count - 1) * baseSpacing) / totalAR;
      }

      totalBlockHeight += rowH;
      return { rowImages, rowH, ars, count, isFew };
    });
    totalBlockHeight += (rawData.length - 1) * baseSpacing;
    const fitScale = Math.min(1, targetBoxHPx / totalBlockHeight);
    const finalSpacing = baseSpacing * fitScale;
    const rowResults: any[] = [];
    let curY = 0;
    rawData.forEach(row => {
      const scaledH = row.rowH * fitScale;
      const imgs: any[] = [];
      const count = row.count;
      const widths = row.ars.map(ar => ar * scaledH);
      const rowTotalWidth = widths.reduce((a, b) => a + b, 0) + (count - 1) * finalSpacing;
      const isLeftAlign = dims.alignLeft && !row.isFew;
      if (!isLeftAlign) {
        let curX = (containerW - rowTotalWidth) / 2;
        row.rowImages.forEach((img, idx) => {
          const w = widths[idx];
          imgs.push({ img, x: curX, y: curY, w, h: scaledH });
          curX += w + finalSpacing;
        });
      } else {
        let curX = 0;
        const dynamicGap = count > 1 ? (containerW - widths.reduce((a, b) => a + b, 0)) / (count - 1) : 0;
        row.rowImages.forEach((img, idx) => {
          const w = widths[idx];
          imgs.push({ img, x: curX, y: curY, w, h: scaledH });
          curX += w + dynamicGap;
        });
      }
      rowResults.push({ images: imgs });
      curY += scaledH + finalSpacing;
    });
    return { rowResults, finalHeight: targetBoxHPx };
  }, [options.containerWidth, options.spacing, getPageDimensions]);

  const clampPage1ImageCm = useCallback((xCm: number, yCm: number, wCm: number, hCm: number) => {
    const leftBoundCm = LAYOUT.PAGE1.IMAGES.LEFT;
    const rightMarginCm = LAYOUT.PAGE1.IMAGES.RIGHT;
    const maxRightCm = LAYOUT.SLIDE.WIDTH_CM - rightMarginCm;

    let clampedX = xCm;
    let clampedW = wCm;
    let clampedH = hCm;

    if (clampedX < leftBoundCm) {
      const shift = leftBoundCm - clampedX;
      clampedX = leftBoundCm;
      clampedW = Math.max(0.01, clampedW - shift);
    }

    const overflow = clampedX + clampedW - maxRightCm;
    if (overflow > 0) {
      const movableLeft = Math.max(0, clampedX - leftBoundCm);
      const moveLeft = Math.min(overflow, movableLeft);
      clampedX -= moveLeft;

      const remainOverflow = clampedX + clampedW - maxRightCm;
      if (remainOverflow > 0) {
        const nextW = Math.max(0.01, clampedW - remainOverflow);
        const scale = nextW / clampedW;
        clampedW = nextW;
        clampedH = Math.max(0.01, clampedH * scale);
      }
    }

    return { xCm: clampedX, yCm, wCm: clampedW, hCm: clampedH };
  }, []);

  const calculateSvgDataForPage = useCallback((pageNum: 1 | 2 | 3, renderTarget: 'preview' | 'export' = 'preview') => {
    const page1Text = LAYOUT.PAGE1.TEXT as any;
    const page1Images = LAYOUT.PAGE1.IMAGES as any;
    const page2Text = LAYOUT.PAGE2.TEXT as any;
    const page3Text = LAYOUT.PAGE3.TEXT as any;
    const originalPage1ImageStartY = page1Images.START_Y;
    const originalPage1FixedClosingY = page1Text.FIXED_CLOSING_TEXT.y;
    const originalPage2ProcedureHeaderY = page2Text.SECTION_HEADER_PROCEDURE?.y;
    const originalPage3FreeTextX = page3Text.FREE_TEXT_PAGE3?.x;
    const originalPage3FreeTextY = page3Text.FREE_TEXT_PAGE3?.y;

    if (pageNum === 1) {
      const bodyStartYcm = 11.8;
      const fontPx = LAYOUT.FONTS.BODY_BASE * (96 / 72);
      const lineHeightPx = fontPx * 1.15;
      const page1MaxW = LAYOUT.SLIDE.WIDTH_CM - page1Images.LEFT - page1Images.RIGHT;
      const pxPerCmForPage1 = options.containerWidth / page1MaxW;
      const bodyCfg = page1Text.FREE_TEXT_INITIAL;
      const maxWidthPx = bodyCfg.w * pxPerCmForPage1;
      const maxLines = Math.max(1, Math.floor((bodyCfg.h * pxPerCmForPage1) / lineHeightPx));
      const lineCount = estimateWrappedLineCount(reportFields.initialText || '', maxWidthPx, fontPx, maxLines);
      const lineHeightCm = lineHeightPx / pxPerCmForPage1;
      const imageHeaderYcm = bodyStartYcm + lineCount * lineHeightCm + lineHeightCm;

      page1Text.FIXED_CLOSING_TEXT.y = 10.3;
      page1Images.START_Y = imageHeaderYcm + 0.5;
    }

    const pageImages = allPagesImages[pageNum] ?? [];
    const pageConfirmed = pageImages.filter(img => img.row > 0);
    const pageRows = [1, 2, 3, 4].map(r => pageConfirmed.filter(img => img.row === r));

    const { rowResults } = calculateLayoutForAnyPage(pageRows, pageNum);
    const pageDims = getPageDimensions(pageNum);

    const pxPerCm = options.containerWidth / pageDims.maxW;
    const fullSlideW = LAYOUT.SLIDE.WIDTH_CM * pxPerCm;
    const fullSlideH = LAYOUT.SLIDE.HEIGHT_CM * pxPerCm;

    // slideOffset positions the whole slide within the preview (usually 0)
    const slideOffsetX = 0;
    const slideOffsetY = 0;

    // image area offset: only images use the page's IMAGES.LEFT / START_Y
    const imgCfg = pageNum === 1 ? LAYOUT.PAGE1.IMAGES : pageNum === 2 ? LAYOUT.PAGE2.IMAGES : LAYOUT.PAGE3.IMAGES;

    const svgParts: string[] = [];

    rowResults.forEach((row) => {
      row.images.forEach((item: any) => {
        const rawXcm = imgCfg.LEFT + item.x / pxPerCm;
        const rawYcm = imgCfg.START_Y + item.y / pxPerCm;
        const rawWcm = item.w / pxPerCm;
        const rawHcm = item.h / pxPerCm;

        const placed =
          pageNum === 1
            ? clampPage1ImageCm(rawXcm, rawYcm, rawWcm, rawHcm)
            : { xCm: rawXcm, yCm: rawYcm, wCm: rawWcm, hCm: rawHcm };

        const absX = slideOffsetX + placed.xCm * pxPerCm;
        const absY = slideOffsetY + placed.yCm * pxPerCm;
        const renderW = placed.wCm * pxPerCm;
        const renderH = placed.hCm * pxPerCm;

        const cx = absX + renderW / 2;
        const cy = absY + renderH / 2;

        const isPortrait = item.img.rotation === 90 || item.img.rotation === 270;
        let drawW = renderW, drawH = renderH;
        if (isPortrait) { drawW = renderH; drawH = renderW; }

        const transform =
          item.img.rotation !== 0
            ? `transform="rotate(${item.img.rotation} ${cx} ${cy})"`
            : '';

        svgParts.push(
          `  <image x="${cx - drawW / 2}" y="${cy - drawH / 2}" width="${drawW}" height="${drawH}" href="${item.img.dataUrl}" ${transform} />`
        );
      });
    });

    if (pageNum === 2 && rowResults.length > 0 && typeof originalPage2ProcedureHeaderY === 'number') {
      let imagesBottomCm: number = imgCfg.START_Y;
      rowResults.forEach((row) => {
        row.images.forEach((item: any) => {
          const bottomCm = imgCfg.START_Y + (item.y + item.h) / pxPerCm;
          if (bottomCm > imagesBottomCm) imagesBottomCm = bottomCm;
        });
      });

      const bodyFontPx = LAYOUT.FONTS.BODY_BASE * (96 / 72);
      const oneLineCm = (bodyFontPx * 1.15) / pxPerCm;
      page2Text.SECTION_HEADER_PROCEDURE.y = imagesBottomCm + oneLineCm;
    }

    if (pageNum === 3 && showPage3 && postPlacement === 'page3' && page3Text.FREE_TEXT_PAGE3) {
      if (rowResults.length > 0) {
        let imagesBottomCm: number = imgCfg.START_Y;
        rowResults.forEach((row) => {
          row.images.forEach((item: any) => {
            const bottomCm = imgCfg.START_Y + (item.y + item.h) / pxPerCm;
            if (bottomCm > imagesBottomCm) imagesBottomCm = bottomCm;
          });
        });

        const bodyFontPx = LAYOUT.FONTS.BODY_BASE * (96 / 72);
        const oneLineCm = (bodyFontPx * 1.15) / pxPerCm;
        page3Text.FREE_TEXT_PAGE3.y = imagesBottomCm + oneLineCm;
      } else {
        page3Text.FREE_TEXT_PAGE3.x = 1.04;
        page3Text.FREE_TEXT_PAGE3.y = 1.9;
      }
    }

    // use shared helper to render all text portions so that SVG and PPTX stay in sync
    const originalCutLineBottomY = page1Text.CUT_LINE_BOTTOM.y;
    const originalFixedClosingY = page1Text.FIXED_CLOSING_TEXT.y;

    if (pageNum === 1) {
      if (filledDateCount === 1) {
        page1Text.CUT_LINE_BOTTOM.y = 9.35;
      } else if (filledDateCount === 2) {
        page1Text.CUT_LINE_BOTTOM.y = 9.79;
      } else if (filledDateCount === 3) {
        page1Text.CUT_LINE_BOTTOM.y = 10.2;
      } else {
        page1Text.CUT_LINE_BOTTOM.y = page1CutLineBottomBaseYRef.current;
      }
      page1Text.FIXED_CLOSING_TEXT.y = 10.3;
    }

    const builtTextParts = buildSvgTextParts(pageNum, reportFields, pxPerCm, slideOffsetX, slideOffsetY, {
      showPage3,
      postPlacement,
    });

    const previewOffsetsPx: PreviewTextOffsetPx = {
      initialText: PREVIEW_TEXT_Y_OFFSETS_CM.initialText * pxPerCm,
      procedureText: PREVIEW_TEXT_Y_OFFSETS_CM.procedureText * pxPerCm,
      postText: PREVIEW_TEXT_Y_OFFSETS_CM.postText * pxPerCm,
      page3Text: PREVIEW_TEXT_Y_OFFSETS_CM.page3Text * pxPerCm,
      closingMessage: PREVIEW_TEXT_Y_OFFSETS_CM.closingMessage * pxPerCm,
    };
    const resolvedTextParts =
      renderTarget === 'preview'
        ? applyPreviewTextOffsets(pageNum, builtTextParts, previewOffsetsPx)
        : builtTextParts;

    if (pageNum === 1) {
      page1Text.CUT_LINE_BOTTOM.y = originalCutLineBottomY;
      page1Text.FIXED_CLOSING_TEXT.y = originalFixedClosingY;
    }
    if (pageNum === 1) {
      svgParts.push(
        ...resolvedTextParts.filter(
          (part) => !part.includes('初診日：') && !part.includes('鎮静日：') && !part.includes('全身麻酔日：') && !part.includes('【初診時】')
        )
      );

      const sectionHeaderX = slideOffsetX + page1Text.SECTION_HEADER.x * pxPerCm;
      const sectionHeaderY = slideOffsetY + 11.07 * pxPerCm;
      svgParts.push(
        `  <text x="${sectionHeaderX}" y="${sectionHeaderY}" font-size="${LAYOUT.FONTS.SECTION_HEADER * (96 / 72)}" fill="#000" font-family="Meiryo, 'MS PGothic', 'Noto Sans JP', sans-serif" dominant-baseline="hanging" font-weight="bold">【初診時】</text>`
      );
    } else {
      svgParts.push(...resolvedTextParts);
      if (pageNum === 2 && page2PhotoLabelText) {
        const labelX = slideOffsetX + 1.04 * pxPerCm;
        const labelY = slideOffsetY + 1.9 * pxPerCm;
        svgParts.push(
          `  <text x="${labelX}" y="${labelY}" font-size="${LAYOUT.FONTS.SECTION_HEADER * (96 / 72)}" fill="#000" font-family="Meiryo, 'MS PGothic', 'Noto Sans JP', sans-serif" dominant-baseline="hanging" font-weight="bold">${escapeSvgText(page2PhotoLabelText)}</text>`
        );
      }
    }

    if (pageNum === 1) {
      const firstVisitDateCfg = page1Text.FIRST_VISIT_DATE;
      const anesthesiaDateCfg = page1Text.ANESTHESIA_DATE;
      if (!firstVisitDateCfg || !anesthesiaDateCfg) {
        page1Images.START_Y = originalPage1ImageStartY;
        page1Text.FIXED_CLOSING_TEXT.y = originalPage1FixedClosingY;
        return { svgCode: '', fullSlideW, fullSlideH };
      }

      const firstVisitX = slideOffsetX + firstVisitDateCfg.x * pxPerCm;
      const firstVisitY = slideOffsetY + firstVisitDateCfg.y * pxPerCm;
      const lineGap = (anesthesiaDateCfg.y - firstVisitDateCfg.y) * pxPerCm;
      const fontPx = LAYOUT.FONTS.BODY_BASE * (96 / 72);

      page1DateOutputItems.forEach((item, index) => {
        const y = firstVisitY + lineGap * index;
        svgParts.push(
          `  <text x="${firstVisitX}" y="${y}" font-size="${fontPx}" fill="#000" font-family="Meiryo, 'MS PGothic', 'Noto Sans JP', sans-serif" dominant-baseline="hanging">${escapeSvgText(
            `${item.label}：${item.value}`
          )}</text>`
        );
      });
    }

    const svgCode = `
<svg xmlns="http://www.w3.org/2000/svg" width="100%" height="100%" viewBox="0 0 ${fullSlideW} ${fullSlideH}">
  <rect width="100%" height="100%" fill="${options.backgroundColor}" />
${svgParts.join('\n')}
</svg>`.trim();

    if (pageNum === 1) {
      page1Images.START_Y = originalPage1ImageStartY;
      page1Text.FIXED_CLOSING_TEXT.y = originalPage1FixedClosingY;
    }
    if (pageNum === 2 && typeof originalPage2ProcedureHeaderY === 'number') {
      page2Text.SECTION_HEADER_PROCEDURE.y = originalPage2ProcedureHeaderY;
    }
    if (pageNum === 3 && page3Text.FREE_TEXT_PAGE3) {
      if (typeof originalPage3FreeTextX === 'number') page3Text.FREE_TEXT_PAGE3.x = originalPage3FreeTextX;
      if (typeof originalPage3FreeTextY === 'number') page3Text.FREE_TEXT_PAGE3.y = originalPage3FreeTextY;
    }

    return { svgCode, fullSlideW, fullSlideH };
  }, [
    allPagesImages,
    calculateLayoutForAnyPage,
    clampPage1ImageCm,
    options.backgroundColor,
    options.containerWidth,
    getPageDimensions,
    reportFields,
    showPage3,
    postPlacement,
    page2PhotoLabelText
  ]);


  const printPdf = useCallback(async () => {
  setIsPrintMode(true);

  // React再描画を確実に待つ
  await new Promise(requestAnimationFrame);
  await new Promise(requestAnimationFrame);

  window.print();
}, []);

  useEffect(() => {
    if (typeof window === 'undefined') return;

    const handleBeforePrint = () => setIsPrintMode(true);
    const handleAfterPrint = () => setIsPrintMode(false);

    window.addEventListener('beforeprint', handleBeforePrint);
    window.addEventListener('afterprint', handleAfterPrint);

    return () => {
      window.removeEventListener('beforeprint', handleBeforePrint);
      window.removeEventListener('afterprint', handleAfterPrint);
    };
  }, []);

  const openGmailDraft = useCallback(async () => {
    if (isCreatingDraft) return;

    const to = (reportFields.refHospitalEmail || '').trim();
    if (!to) {
      window.alert('紹介病院メールが未入力です。メールアドレスを入力してください。');
      return;
    }

    const owner = (reportFields.ownerLastName || '').trim();
    const pet = (reportFields.petName || '').trim();
    const hospital = (reportFields.refHospitalName || reportFields.refHospital || '').trim();
    const doctor = (reportFields.refDoctor || '').trim();
    const doctorWithSuffix = doctor ? `${doctor} ${DOCTOR_SUFFIX}` : DOCTOR_SUFFIX;
    const vet = (reportFields.attendingVet || '').trim();
    const closingMessage = (reportFields.closingMessageText || '').trim() || DEFAULT_CLOSING_MESSAGE;

    const subject = `治療報告書（${owner}様 ${pet}ちゃん）`;
    const body = `${hospital} 御中
  ${doctorWithSuffix}

いつもお世話になっております。荻窪ツイン動物病院の${vet}です。
  ${closingMessage}`;

    setIsCreatingDraft(true);
    try {
      setIsPrintMode(true);
      await new Promise(requestAnimationFrame);
      await new Promise(requestAnimationFrame);

      const pdf = new jsPDF({ orientation: 'portrait', unit: 'mm', format: 'a4' });

      for (let index = 0; index < outputPages.length; index += 1) {
        const pageNum = outputPages[index];
        const printPage = document.getElementById(`print-page-${pageNum}`);
        if (!printPage) {
          window.alert(`印刷対象プレビューが見つかりません（PAGE ${pageNum}）。`);
          return;
        }

        const canvas = await html2canvas(printPage, {
          scale: 2,
          useCORS: true,
          backgroundColor: '#ffffff',
        });

        const imageData = canvas.toDataURL('image/png');
        if (index > 0) {
          pdf.addPage('a4', 'portrait');
        }
        pdf.addImage(imageData, 'PNG', 0, 0, 210, 297, undefined, 'FAST');
      }

      const pdfBlob = pdf.output('blob');

      const formData = new FormData();
      formData.append('to', to);
      formData.append('subject', subject);
      formData.append('body', body);
      formData.append('file', new File([pdfBlob], 'Photo_Report_A4.pdf', { type: 'application/pdf' }));

      const response = await fetch(`${SERVER_BASE_URL}/api/gmail/create-draft`, {
        method: 'POST',
        body: formData,
      });

      const data = await response.json().catch(() => ({}));
      if (!response.ok || !data?.ok || !data?.draftUrl) {
        let detail = data?.error || `HTTP ${response.status}`;
        if (response.status >= 500) {
          detail = `${detail}\nサーバ起動状態またはGmail OAuth設定（gmail_oauth_client.json / gmail_token.json）を確認してください。`;
        }
        throw new Error(detail);
      }

      window.open(data.draftUrl, '_blank');
    } catch (error) {
      console.error('Gmail draft creation failed:', error);
      const message = error instanceof Error ? error.message : String(error);
      if (message.includes('Failed to fetch') || message.includes('NetworkError')) {
        window.alert('Gmail下書き作成に失敗しました: サーバに接続できませんでした。photo-report-server が起動しているか確認してください。');
      } else {
        window.alert(`Gmail下書き作成に失敗しました: ${message}`);
      }
    } finally {
      setIsPrintMode(false);
      setIsCreatingDraft(false);
    }
  }, [isCreatingDraft, reportFields, outputPages]);

  const downloadPptx = async (e: React.MouseEvent) => {
    e.preventDefault();
    console.log('downloadPptx clicked');
    if (isSavingPptx) return;

    setIsSavingPptx(true);
    setPptxStatus('保存中…');

    try {
      const pptx = new pptxgen();
      pptx.defineLayout({
        name: LAYOUT.SLIDE.NAME,
        width: LAYOUT.SLIDE.WIDTH_CM / 2.54,
        height: LAYOUT.SLIDE.HEIGHT_CM / 2.54,
      });
      pptx.layout = LAYOUT.SLIDE.NAME;

      const buildSlide = (slide: pptxgen.Slide, pageNum: number, imagesData: ImageData[]) => {
        slide.background = { fill: options.backgroundColor.replace('#', '') };

        const pageDims = getPageDimensions(pageNum);
        const cmPerPx = pageDims.maxW / options.containerWidth;

        const pageDisplayRows = [1, 2, 3, 4].map(r => imagesData.filter(img => img.row === r));
        const { rowResults } = calculateLayoutForAnyPage(pageDisplayRows, pageNum);

        rowResults.forEach(row => {
          row.images.forEach((item: any) => {
            const rawXcm = pageDims.left + item.x * cmPerPx;
            const rawYcm = pageDims.startY + item.y * cmPerPx;
            const rawWcm = item.w * cmPerPx;
            const rawHcm = item.h * cmPerPx;

            const placed =
              pageNum === 1
                ? clampPage1ImageCm(rawXcm, rawYcm, rawWcm, rawHcm)
                : { xCm: rawXcm, yCm: rawYcm, wCm: rawWcm, hCm: rawHcm };

            slide.addImage({
              data: item.img.dataUrl,
              x: placed.xCm / 2.54,
              y: placed.yCm / 2.54,
              w: placed.wCm / 2.54,
              h: placed.hCm / 2.54,
              rotate: item.img.rotation,
            });
          });
        });

        const page1Text = LAYOUT.PAGE1.TEXT as any;
        const page2Text = LAYOUT.PAGE2.TEXT as any;
        const page3Text = LAYOUT.PAGE3.TEXT as any;
        const originalFirstVisitDateCfg = page1Text.FIRST_VISIT_DATE;
        const originalAnesthesiaDateCfg = page1Text.ANESTHESIA_DATE;
        const originalPage2ProcedureHeaderY = page2Text.SECTION_HEADER_PROCEDURE?.y;
        const originalPage3FreeTextX = page3Text.FREE_TEXT_PAGE3?.x;
        const originalPage3FreeTextY = page3Text.FREE_TEXT_PAGE3?.y;
        try {
          if (pageNum === 1) {
            page1Text.FIRST_VISIT_DATE = undefined;
            page1Text.ANESTHESIA_DATE = undefined;
          }
          if (pageNum === 2 && rowResults.length > 0 && typeof originalPage2ProcedureHeaderY === 'number') {
            let imagesBottomCm: number = pageDims.startY;
            rowResults.forEach((row) => {
              row.images.forEach((item: any) => {
                const bottomCm = pageDims.startY + (item.y + item.h) * cmPerPx;
                if (bottomCm > imagesBottomCm) imagesBottomCm = bottomCm;
              });
            });

            const bodyFontPx = LAYOUT.FONTS.BODY_BASE * (96 / 72);
            const oneLineCm = (bodyFontPx * 1.15) * cmPerPx;
            page2Text.SECTION_HEADER_PROCEDURE.y = imagesBottomCm + oneLineCm;
          }
          if (pageNum === 3 && showPage3 && postPlacement === 'page3' && page3Text.FREE_TEXT_PAGE3) {
            if (rowResults.length > 0) {
              let imagesBottomCm: number = pageDims.startY;
              rowResults.forEach((row) => {
                row.images.forEach((item: any) => {
                  const bottomCm = pageDims.startY + (item.y + item.h) * cmPerPx;
                  if (bottomCm > imagesBottomCm) imagesBottomCm = bottomCm;
                });
              });

              const bodyFontPx = LAYOUT.FONTS.BODY_BASE * (96 / 72);
              const oneLineCm = (bodyFontPx * 1.15) * cmPerPx;
              page3Text.FREE_TEXT_PAGE3.y = imagesBottomCm + oneLineCm;
            } else {
              page3Text.FREE_TEXT_PAGE3.x = 1.04;
              page3Text.FREE_TEXT_PAGE3.y = 1.9;
            }
          }

          addPptxText(slide, pageNum, reportFields, {
            showPage3,
            postPlacement,
          });
        } finally {
          if (pageNum === 1) {
            page1Text.FIRST_VISIT_DATE = originalFirstVisitDateCfg;
            page1Text.ANESTHESIA_DATE = originalAnesthesiaDateCfg;
          }
          if (pageNum === 2 && typeof originalPage2ProcedureHeaderY === 'number') {
            page2Text.SECTION_HEADER_PROCEDURE.y = originalPage2ProcedureHeaderY;
          }
          if (pageNum === 3 && page3Text.FREE_TEXT_PAGE3) {
            if (typeof originalPage3FreeTextX === 'number') page3Text.FREE_TEXT_PAGE3.x = originalPage3FreeTextX;
            if (typeof originalPage3FreeTextY === 'number') page3Text.FREE_TEXT_PAGE3.y = originalPage3FreeTextY;
          }
        }

        if (pageNum === 1) {

          const lineGapCm = originalAnesthesiaDateCfg.y - originalFirstVisitDateCfg.y;
          page1DateOutputItems.forEach((item, index) => {
            slide.addText(`${item.label}：${item.value}`, {
              x: originalFirstVisitDateCfg.x / 2.54,
              y: (originalFirstVisitDateCfg.y + lineGapCm * index) / 2.54,
              w: originalFirstVisitDateCfg.w / 2.54,
              h: originalFirstVisitDateCfg.h / 2.54,
              fontSize: LAYOUT.FONTS.BODY_BASE,
            });
          });
        }

        if (pageNum === 2 && page2PhotoLabelText) {
          slide.addText(page2PhotoLabelText, {
            x: 1.04 / 2.54,
            y: 1.9 / 2.54,
            w: 8.0 / 2.54,
            h: 0.8 / 2.54,
            fontSize: LAYOUT.FONTS.SECTION_HEADER,
            bold: true,
          });
        }
      };

      outputPages.forEach(pageNum => {
        const pageImages = allPagesImages[pageNum] ?? [];
        const pageConfirmed = pageImages.filter(img => img.row > 0);

        const slide = pptx.addSlide();
        buildSlide(slide, pageNum, pageConfirmed);
      });

      const fileName = 'Photo_Report_A4.pptx';
      const ua = navigator.userAgent;
      const isIOSSafari = /iPad|iPhone|iPod/.test(ua) || (ua.includes('Mac') && navigator.maxTouchPoints > 1);

      if (isIOSSafari) {
        const blob = (await pptx.write({ outputType: 'blob' })) as Blob;
        const blobUrl = URL.createObjectURL(blob);
        const link = document.createElement('a');
        link.href = blobUrl;
        link.download = fileName;
        link.rel = 'noopener';
        document.body.appendChild(link);
        link.click();
        link.remove();
        setTimeout(() => URL.revokeObjectURL(blobUrl), 1000);
      } else {
        await pptx.writeFile({ fileName });
      }

      setPptxStatus('保存しました');
      setTimeout(() => setPptxStatus(''), 2500);
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      console.error('PPTX save failed:', error);
      setPptxStatus(`保存に失敗: ${message}`);
    } finally {
      setIsSavingPptx(false);
    }
  };

  const svgData = useMemo(() => calculateSvgDataForPage(currentPage as 1 | 2 | 3, 'preview'), [calculateSvgDataForPage, currentPage]);
  const svgPage1 = useMemo(() => calculateSvgDataForPage(1, 'export').svgCode, [calculateSvgDataForPage]);
  const svgPage2 = useMemo(() => calculateSvgDataForPage(2, 'export').svgCode, [calculateSvgDataForPage]);
  const svgPage3 = useMemo(() => calculateSvgDataForPage(3, 'export').svgCode, [calculateSvgDataForPage]);

  return (
    <div className="min-h-screen bg-slate-50 pb-32 font-sans">
      {/* プレビューモーダル */}
      {previewId && (
        <div className="fixed inset-0 z-[100] bg-slate-900/90 backdrop-blur flex items-center justify-center p-10" onClick={() => setPreviewId(null)}>
          <div className="relative max-w-full max-h-full" onClick={e => e.stopPropagation()}>
            {images.find(img => img.id === previewId) && (
              <img
                src={images.find(img => img.id === previewId)!.dataUrl}
                style={{ transform: `rotate(${images.find(img => img.id === previewId)!.rotation}deg)` }}
                className="max-w-full max-h-[85vh] object-contain rounded-xl shadow-2xl"
              />
            )}
            <button onClick={() => setPreviewId(null)} className="absolute -top-12 -right-12 text-white hover:text-orange-500 transition-colors">
              <svg className="w-10 h-10" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M6 18L18 6M6 6l12 12" /></svg>
            </button>
          </div>
        </div>
      )}

      <nav className="bg-white border-b border-slate-200 py-4 px-6 sticky top-0 z-50 shadow-sm">
        <div className="max-w-7xl mx-auto">
          <h1 className="text-xl font-black text-slate-900 tracking-tight leading-none">歯科治療報告書</h1>
        </div>
      </nav>

      <main className="max-w-7xl mx-auto px-6 mt-10 grid grid-cols-1 lg:grid-cols-12 gap-10">
        {/* 報告書データ入力フォーム */}
        <div className="lg:col-span-12 bg-white p-7 rounded-[2.5rem] shadow-sm border border-slate-200 space-y-6">
          <div className="flex items-start justify-between gap-3">
            <div>
              <h3 className="font-black text-slate-800 text-base mb-1">報告書データ入力</h3>
              <p className="text-[10px] text-slate-400 font-bold uppercase tracking-wider">現在の段落配置が自動で反映されます</p>
            </div>
            <button
              type="button"
              onClick={handleClearReportFields}
              className="h-8 px-3 rounded-lg border border-slate-300 bg-white text-xs font-semibold text-slate-600 hover:bg-slate-50"
            >
              入力をクリア
            </button>
          </div>

          <div className="space-y-6" data-enter-scope="report-fields" onFocusCapture={handleReportFieldsFocusCapture}>
            {/* 基本情報グリッド */}
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4">
              <div className="space-y-1">
                <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest">報告日</label>
                <input className={`w-full border rounded-xl px-3 py-2 text-sm focus:ring-2 focus:ring-orange-500 outline-none transition-all ${getInputToneClass(reportFields.reportDate)}`}
                  placeholder="2026年2月16日"
                  value={reportFields.reportDate}
                  onChange={e => setReportFields(v => ({ ...v, reportDate: e.target.value }))}
                  onBlur={e => setReportFields(v => ({ ...v, reportDate: normalizeJapaneseDate(e.target.value) }))}
                  onKeyDown={handleEnterToNextField}
                />
              </div>
              <div className="space-y-1">
                <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest">
                  紹介病院名
                </label>

                <input
                  className={`w-full max-w-[520px] h-9 px-3 rounded-lg border text-sm ${getInputToneClass(refHospitalInput)}`}
                  placeholder="例：中川動物病院"
                  value={refHospitalInput}
                  onChange={(e) => applyRefHospitalSelection(e.target.value)}
                  onBlur={(e) => applyRefHospitalSelection(e.target.value)}
                  onKeyDown={(e) => {
                    if (e.key === "Enter") {
                      e.preventDefault();
                      const v = e.currentTarget.value;
                      const mappedEmail = normalizedRefHospitalEmails[normalizeHospitalKey(v)] || "";
                      applyRefHospitalSelection(v);
                      handleAddRefHospital(v);
                      if (mappedEmail) {
                        handleEnterToNextField(e, { targetSelector: 'input[data-enter-field="doctor"]' });
                      } else {
                        handleEnterToNextField(e);
                      }
                    }
                  }}
                  list="refHospitalsList"
                />

                {/* 入力候補（予測変換） */}
                <datalist id="refHospitalsList">
                  {suggestions.refHospitals.map((h) => (
                    <option key={h} value={h} />
                  ))}
                </datalist>

                {/* 保存エラー（任意表示） */}
                {refHospitalError && (
                  <div className="text-xs text-red-600">{refHospitalError}</div>
                )}
              </div>
              <div className="space-y-1">
                <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest">紹介病院メール（Gmail）</label>
                <input className={`w-full border rounded-xl px-3 py-2 text-sm focus:ring-2 focus:ring-orange-500 outline-none transition-all ${getInputToneClass(reportFields.refHospitalEmail)}`}
                  placeholder="example@gmail.com"
                  value={reportFields.refHospitalEmail}
                  onChange={e => setReportFields(v => ({ ...v, refHospitalEmail: e.target.value }))}
                  onKeyDown={handleEnterToNextField}
                />
                {!refHospitalError &&
                  (reportFields.refHospitalName || reportFields.refHospital).trim() !== '' &&
                  reportFields.refHospitalEmail.trim() === '' && (
                    <div className="text-xs text-slate-500">
                      この紹介病院はメール未登録です。必要なら手入力してください。
                    </div>
                  )}
              </div>
              <div className="space-y-1">
                <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest">先生名</label>
                <div className="relative">
                  <input className={`w-full border rounded-xl px-3 pr-14 py-2 text-sm focus:ring-2 focus:ring-orange-500 outline-none transition-all ${getInputToneClass(reportFields.refDoctor)}`}
                    placeholder="△△"
                    value={reportFields.refDoctor}
                    onChange={e => setReportFields(v => ({ ...v, refDoctor: e.target.value }))}
                    onKeyDown={handleEnterToNextField}
                    data-enter-field="doctor"
                  />
                  <span className="pointer-events-none absolute inset-y-0 right-3 flex items-center text-sm text-slate-500">先生</span>
                </div>
              </div>
              <div className="space-y-1">
                <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest">飼い主姓</label>
                <input className={`w-full border rounded-xl px-3 py-2 text-sm focus:ring-2 focus:ring-orange-500 outline-none transition-all ${getInputToneClass(reportFields.ownerLastName)}`}
                  placeholder="山田"
                  value={reportFields.ownerLastName}
                  onChange={e => setReportFields(v => ({ ...v, ownerLastName: e.target.value }))}
                  onKeyDown={handleEnterToNextField}
                />
              </div>
              <div className="space-y-1">
                <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest">ペット名</label>
                <input className={`w-full border rounded-xl px-3 py-2 text-sm focus:ring-2 focus:ring-orange-500 outline-none transition-all ${getInputToneClass(reportFields.petName)}`}
                  placeholder="タロウ"
                  value={reportFields.petName}
                  onChange={e => setReportFields(v => ({ ...v, petName: e.target.value }))}
                  onKeyDown={handleEnterToNextField}
                />
              </div>
              <div className="space-y-1 relative" data-date-field-root="true">
                <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest">初診日</label>
                <div className="relative" onClick={() => openDatePicker('firstVisitDate', firstVisitDateInputRef)}>
                  <input
                    ref={firstVisitDateInputRef}
                    className={`w-full border rounded-xl px-3 pr-16 py-2 text-base focus:ring-2 focus:ring-orange-500 outline-none transition-all ${getInputToneClass(reportFields.firstVisitDate)}`}
                    type="text"
                    readOnly
                    placeholder="202X年XX月XX日"
                    value={reportFields.firstVisitDate}
                    onKeyDown={handleEnterToNextField}
                  />
                  <button
                    type="button"
                    className="absolute right-2 top-1/2 -translate-y-1/2 rounded-md px-2 py-1 text-[10px] font-semibold text-slate-500 hover:bg-slate-100"
                    onClick={(e) => {
                      e.stopPropagation();
                      setReportFields((v) => ({ ...v, firstVisitDate: '' }));
                    }}
                  >
                    クリア
                  </button>
                </div>
                {calendarOpenField === 'firstVisitDate' && (
                  <div className="absolute left-0 top-full mt-2 rounded-2xl border border-slate-300 bg-white p-3 shadow-xl w-[315px] max-w-[88vw] z-30">
                    <div className="mb-3 flex items-center justify-between">
                      <button type="button" onClick={() => moveCalendarMonth(-1)} className="h-8 w-8 rounded-lg border border-slate-300 text-base">◀</button>
                      <div className="text-base font-black text-slate-700">{calendarViewDate.getFullYear()}年{calendarViewDate.getMonth() + 1}月</div>
                      <button type="button" onClick={() => moveCalendarMonth(1)} className="h-8 w-8 rounded-lg border border-slate-300 text-base">▶</button>
                    </div>
                    <div className="grid grid-cols-7 gap-1.5">
                      {weekdayLabels.map((d) => (
                        <div key={d} className="text-center text-sm font-bold text-slate-500">{d}</div>
                      ))}
                      {calendarDays.map((d) => {
                        if (!d.inMonth) {
                          return <div key={d.iso} className="h-9" />;
                        }
                        const selected = toDateInputValue(reportFields.firstVisitDate) === d.iso;
                        return (
                          <button
                            key={d.iso}
                            type="button"
                            className={`h-9 rounded-lg text-base font-semibold ${selected ? 'bg-orange-500 text-white' : 'bg-slate-50 text-slate-800 hover:bg-orange-50'}`}
                            onClick={() => selectCalendarDate('firstVisitDate', d.iso)}
                          >
                            {d.day}
                          </button>
                        );
                      })}
                    </div>
                  </div>
                )}
              </div>
              <div className="space-y-1 relative" data-date-field-root="true">
                <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest">鎮静日</label>
                <div className="relative" onClick={() => openDatePicker('sedationDate', sedationDateInputRef)}>
                  <input
                    ref={sedationDateInputRef}
                    className={`w-full border rounded-xl px-3 pr-16 py-2 text-base focus:ring-2 focus:ring-orange-500 outline-none transition-all ${getInputToneClass(reportFields.sedationDate)}`}
                    type="text"
                    readOnly
                    placeholder="202X年XX月XX日"
                    value={reportFields.sedationDate}
                    onKeyDown={handleEnterToNextField}
                  />
                  <button
                    type="button"
                    className="absolute right-2 top-1/2 -translate-y-1/2 rounded-md px-2 py-1 text-[10px] font-semibold text-slate-500 hover:bg-slate-100"
                    onClick={(e) => {
                      e.stopPropagation();
                      setReportFields((v) => ({ ...v, sedationDate: '' }));
                    }}
                  >
                    クリア
                  </button>
                </div>
                {calendarOpenField === 'sedationDate' && (
                  <div className="absolute left-0 top-full mt-2 rounded-2xl border border-slate-300 bg-white p-3 shadow-xl w-[315px] max-w-[88vw] z-30">
                    <div className="mb-3 flex items-center justify-between">
                      <button type="button" onClick={() => moveCalendarMonth(-1)} className="h-8 w-8 rounded-lg border border-slate-300 text-base">◀</button>
                      <div className="text-base font-black text-slate-700">{calendarViewDate.getFullYear()}年{calendarViewDate.getMonth() + 1}月</div>
                      <button type="button" onClick={() => moveCalendarMonth(1)} className="h-8 w-8 rounded-lg border border-slate-300 text-base">▶</button>
                    </div>
                    <div className="grid grid-cols-7 gap-1.5">
                      {weekdayLabels.map((d) => (
                        <div key={d} className="text-center text-sm font-bold text-slate-500">{d}</div>
                      ))}
                      {calendarDays.map((d) => {
                        if (!d.inMonth) {
                          return <div key={d.iso} className="h-9" />;
                        }
                        const selected = toDateInputValue(reportFields.sedationDate) === d.iso;
                        return (
                          <button
                            key={d.iso}
                            type="button"
                            className={`h-9 rounded-lg text-base font-semibold ${selected ? 'bg-orange-500 text-white' : 'bg-slate-50 text-slate-800 hover:bg-orange-50'}`}
                            onClick={() => selectCalendarDate('sedationDate', d.iso)}
                          >
                            {d.day}
                          </button>
                        );
                      })}
                    </div>
                  </div>
                )}
              </div>
              <div className="space-y-1 relative" data-date-field-root="true">
                <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest">全身麻酔日</label>
                <div className="relative" onClick={() => openDatePicker('anesthesiaDate', anesthesiaDateInputRef)}>
                  <input
                    ref={anesthesiaDateInputRef}
                    className={`w-full border rounded-xl px-3 pr-16 py-2 text-base focus:ring-2 focus:ring-orange-500 outline-none transition-all ${getInputToneClass(reportFields.anesthesiaDate)}`}
                    type="text"
                    readOnly
                    placeholder="202X年XX月XX日"
                    value={reportFields.anesthesiaDate}
                    onKeyDown={handleEnterToNextField}
                  />
                  <button
                    type="button"
                    className="absolute right-2 top-1/2 -translate-y-1/2 rounded-md px-2 py-1 text-[10px] font-semibold text-slate-500 hover:bg-slate-100"
                    onClick={(e) => {
                      e.stopPropagation();
                      setReportFields((v) => ({ ...v, anesthesiaDate: '' }));
                    }}
                  >
                    クリア
                  </button>
                </div>
                {calendarOpenField === 'anesthesiaDate' && (
                  <div className="absolute left-0 top-full mt-2 rounded-2xl border border-slate-300 bg-white p-3 shadow-xl w-[315px] max-w-[88vw] z-30">
                    <div className="mb-3 flex items-center justify-between">
                      <button type="button" onClick={() => moveCalendarMonth(-1)} className="h-8 w-8 rounded-lg border border-slate-300 text-base">◀</button>
                      <div className="text-base font-black text-slate-700">{calendarViewDate.getFullYear()}年{calendarViewDate.getMonth() + 1}月</div>
                      <button type="button" onClick={() => moveCalendarMonth(1)} className="h-8 w-8 rounded-lg border border-slate-300 text-base">▶</button>
                    </div>
                    <div className="grid grid-cols-7 gap-1.5">
                      {weekdayLabels.map((d) => (
                        <div key={d} className="text-center text-sm font-bold text-slate-500">{d}</div>
                      ))}
                      {calendarDays.map((d) => {
                        if (!d.inMonth) {
                          return <div key={d.iso} className="h-9" />;
                        }
                        const selected = toDateInputValue(reportFields.anesthesiaDate) === d.iso;
                        return (
                          <button
                            key={d.iso}
                            type="button"
                            className={`h-9 rounded-lg text-base font-semibold ${selected ? 'bg-orange-500 text-white' : 'bg-slate-50 text-slate-800 hover:bg-orange-50'}`}
                            onClick={() => selectCalendarDate('anesthesiaDate', d.iso)}
                          >
                            {d.day}
                          </button>
                        );
                      })}
                    </div>
                  </div>
                )}
              </div>
              {/* 担当獣医師（新規：プルダウン） */}
              <div className="space-y-1 relative z-40" ref={attendingVetDropdownRef}>
                <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest">担当獣医師</label>
                <div className="relative">
                  <button
                    type="button"
                    className={`w-full border rounded-xl px-3 py-2 text-base text-left focus:ring-2 focus:ring-orange-500 outline-none transition-all ${getInputToneClass(reportFields.attendingVet)}`}
                    onClick={() => {
                      setIsPage2PhotoLabelOpen(false);
                      setIsPostPlacementOpen(false);
                      setIsAttendingVetOpen((v) => !v);
                    }}
                    onKeyDown={(e) => {
                      if (e.key === 'Enter' && !isAttendingVetOpen) {
                        handleEnterToNextField(e as any);
                      }
                    }}
                  >
                    {reportFields.attendingVet || '選択してください'}
                  </button>
                  {isAttendingVetOpen && (
                    <div className="absolute z-50 mt-1 w-full overflow-hidden rounded-xl border border-slate-200 bg-white shadow-lg">
                      {['', '町田健吾', '江成翔馬', '神田珠希', '小林嵩', '金田七海'].map((name) => (
                        <button
                          key={name || 'empty'}
                          type="button"
                          className="w-full px-3 py-2 text-left text-base hover:bg-slate-100"
                          onClick={() => {
                            setReportFields(v => ({ ...v, attendingVet: name }));
                            setIsAttendingVetOpen(false);
                          }}
                        >
                          {name || '選択してください'}
                        </button>
                      ))}
                    </div>
                  )}
                </div>
              </div>
              <div className="space-y-1 relative z-40" ref={page2PhotoLabelDropdownRef}>
                <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest">PAGE2写真区分ラベル</label>
                <div className="relative">
                  <button
                    type="button"
                    className={`w-full border rounded-xl px-3 py-2 text-base text-left focus:ring-2 focus:ring-orange-500 outline-none transition-all ${getInputToneClass(page2PhotoLabel)}`}
                    onClick={() => {
                      setIsAttendingVetOpen(false);
                      setIsPostPlacementOpen(false);
                      setIsPage2PhotoLabelOpen((v) => !v);
                    }}
                  >
                    {page2PhotoLabel === 'treatment-post'
                      ? '治療時・治療後写真'
                      : page2PhotoLabel === 'inspection'
                        ? '検査時写真'
                        : '空欄'}
                  </button>
                  {isPage2PhotoLabelOpen && (
                    <div className="absolute z-50 mt-1 w-full overflow-hidden rounded-xl border border-slate-200 bg-white shadow-lg">
                      <button
                        type="button"
                        className="w-full px-3 py-2 text-left text-base hover:bg-slate-100"
                        onClick={() => {
                          setPage2PhotoLabel('');
                          setIsPage2PhotoLabelOpen(false);
                        }}
                      >
                        空欄
                      </button>
                      <button
                        type="button"
                        className="w-full px-3 py-2 text-left text-base hover:bg-slate-100"
                        onClick={() => {
                          setPage2PhotoLabel('treatment-post');
                          setIsPage2PhotoLabelOpen(false);
                        }}
                      >
                        治療時・治療後写真
                      </button>
                      <button
                        type="button"
                        className="w-full px-3 py-2 text-left text-base hover:bg-slate-100"
                        onClick={() => {
                          setPage2PhotoLabel('inspection');
                          setIsPage2PhotoLabelOpen(false);
                        }}
                      >
                        検査時写真
                      </button>
                    </div>
                  )}
                </div>
              </div>
              <div className="space-y-1 relative z-20">
                <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest">PAGE3設定</label>
                <label className="h-9 px-3 border rounded-xl flex items-center gap-2 text-sm text-slate-700 font-semibold bg-white">
                  <input
                    type="checkbox"
                    checked={showPage3}
                    onChange={e => setShowPage3(e.target.checked)}
                  />
                  PAGE3を追加する
                </label>
              </div>
              {/* 主訴（新規：テキスト入力） */}
              <div className="space-y-1 relative z-20 lg:col-span-3">
                <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest">主訴</label>
                <input className={`w-full border rounded-xl px-3 py-2 text-sm focus:ring-2 focus:ring-orange-500 outline-none transition-all ${getInputToneClass(reportFields.chiefComplaint)}`}
                  placeholder="主な症状や主訴"
                  value={reportFields.chiefComplaint}
                  onChange={e => setReportFields(v => ({ ...v, chiefComplaint: e.target.value }))}
                  onKeyDown={handleEnterToNextField}
                />
              </div>
            </div>

            {/* PAGE3設定 */}
            <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
              {showPage3 && (
                <div className="space-y-1 relative z-40" ref={postPlacementDropdownRef}>
                  <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest">術後経過の配置</label>
                  <div className="relative">
                    <button
                      type="button"
                      className={`w-full border rounded-xl px-3 py-2 text-base text-left focus:ring-2 focus:ring-orange-500 outline-none transition-all ${getInputToneClass(postPlacement)}`}
                      onClick={() => {
                        setIsAttendingVetOpen(false);
                        setIsPage2PhotoLabelOpen(false);
                        setIsPostPlacementOpen((v) => !v);
                      }}
                      onKeyDown={(e) => {
                        if (e.key === 'Enter' && !isPostPlacementOpen) {
                          handleEnterToNextField(e as any);
                        }
                      }}
                    >
                      {postPlacement === 'page2' ? 'PAGE2に置く' : 'PAGE3に移す'}
                    </button>
                    {isPostPlacementOpen && (
                      <div className="absolute z-50 mt-1 w-full overflow-hidden rounded-xl border border-slate-200 bg-white shadow-lg">
                        <button
                          type="button"
                          className="w-full px-3 py-2 text-left text-base hover:bg-slate-100"
                          onClick={() => {
                            setPostPlacement('page2');
                            setIsPostPlacementOpen(false);
                          }}
                        >
                          PAGE2に置く
                        </button>
                        <button
                          type="button"
                          className="w-full px-3 py-2 text-left text-base hover:bg-slate-100"
                          onClick={() => {
                            setPostPlacement('page3');
                            setIsPostPlacementOpen(false);
                          }}
                        >
                          PAGE3に移す
                        </button>
                      </div>
                    )}
                  </div>
                </div>
              )}
            </div>

            {/* 自由記載エリア */}
            <div className="grid grid-cols-1 gap-4">
              <div className="space-y-1">
                <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest">【初診時】本文 (Page 1)</label>
                <textarea className={`w-full border rounded-xl px-3 py-2 text-sm min-h-[80px] focus:ring-2 focus:ring-orange-500 outline-none transition-all ${getTextareaToneClass(reportFields.initialText)}`}
                  placeholder="初診時の所見など..."
                  value={reportFields.initialText}
                  onChange={e => setReportFields(v => ({ ...v, initialText: e.target.value }))}
                />
              </div>
              <div className="space-y-1">
                <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest">【検査・処置内容】本文 (Page 2)</label>
                <textarea className={`w-full border rounded-xl px-3 py-2 text-sm min-h-[80px] focus:ring-2 focus:ring-orange-500 outline-none transition-all ${getTextareaToneClass(reportFields.procedureText)}`}
                  placeholder="実施した検査や処置の詳細..."
                  value={reportFields.procedureText}
                  onChange={e => setReportFields(v => ({ ...v, procedureText: e.target.value }))}
                />
              </div>
              <div className="space-y-1">
                <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest">【術後経過】本文 ({showPage3 && postPlacement === 'page3' ? 'Page 3' : 'Page 2'})</label>
                <textarea className={`w-full border rounded-xl px-3 py-2 text-sm min-h-[80px] focus:ring-2 focus:ring-orange-500 outline-none transition-all ${getTextareaToneClass(reportFields.postText)}`}
                  placeholder="術後の状態や今後の予定..."
                  value={reportFields.postText}
                  onChange={e => setReportFields(v => ({ ...v, postText: e.target.value }))}
                />
              </div>
              {showPage3 && (
                <div className="space-y-1">
                  <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest">【PAGE3】自由入力</label>
                  <textarea className={`w-full border rounded-xl px-3 py-2 text-sm min-h-[80px] focus:ring-2 focus:ring-orange-500 outline-none transition-all ${getTextareaToneClass(reportFields.page3Text || '')}`}
                    placeholder="PAGE3に出す補足テキスト..."
                    value={reportFields.page3Text || ''}
                    onChange={e => setReportFields(v => ({ ...v, page3Text: e.target.value }))}
                  />
                </div>
              )}
              <div className="space-y-1">
                <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest">メール締め文</label>
                <textarea className={`w-full border rounded-xl px-3 py-2 text-sm min-h-[80px] focus:ring-2 focus:ring-orange-500 outline-none transition-all ${getTextareaToneClass(reportFields.closingMessageText || '')}`}
                  placeholder="メール本文の締めメッセージ"
                  value={reportFields.closingMessageText || ''}
                  onChange={e => setReportFields(v => ({ ...v, closingMessageText: e.target.value }))}
                />
              </div>
            </div>
            {/* テンプレートピッカー */}
            <TemplatePicker
              onInsert={(field, text, mode) => {
                const prev = reportFields[field as 'initialText' | 'procedureText' | 'postText'] || '';
                // replace のときは既存テキストがある場合に確認
                if (mode === 'replace' && prev) {
                  const ok = window.confirm('既に本文があります。置き換えますか？（キャンセルで中止）');
                  if (!ok) return;
                }
                const next = mode === 'replace' ? text : (prev ? `${prev}\n${text}` : text);
                setReportFields(prevState => ({ ...prevState, [field]: next }));
                setLastInsert({ field, prevValue: prev });
              }}
              onUndo={() => {
                if (!lastInsert) return;
                setReportFields(prevState => ({ ...prevState, [lastInsert.field]: lastInsert.prevValue }));
                setLastInsert(null);
              }}
              canUndo={!!lastInsert}
              undoLabel={lastInsert ? (lastInsert.field === 'initialText' ? '初診' : lastInsert.field === 'procedureText' ? '処置' : '術後') : ''}
            />
          </div>
        </div>

        {/* 操作バー：ページ選択・画像追加 */}
        <div className="lg:col-span-12 relative w-full my-4 h-[48px]">
          <div className="absolute right-0 top-0 z-10">
            <button
              onClick={() => {
                console.log('image upload button clicked');
                if (fileInputRef.current) fileInputRef.current.value = '';
                fileInputRef.current?.click();
              }}
              className="bg-indigo-600 hover:bg-indigo-700 text-white px-6 py-3 rounded-2xl font-bold shadow-lg active:scale-95 flex items-center gap-2 transition-all"
            >
              <svg className="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2.5" d="M12 4v16m8-8H4" />
              </svg>
              画像を追加
            </button>
          </div>

          <div className="absolute left-1/2 -translate-x-1/2 top-0 pointer-events-none">
  <div className="pointer-events-auto">
    <PageSwitcher
      currentPage={currentPage}
      onChange={handlePageChange}
      size="small"
      pages={availablePages}
    />
  </div>
</div>
          <input
            type="file"
            ref={fileInputRef}
            onChange={handleFileUpload}
            multiple
            accept="image/*"
            className="hidden"
          />
        </div>

        {/* 左カラム - 確定前のみ表示 */}
        {!isCurrentPageConfirmed && (
          <div className="lg:col-span-5 lg:col-start-4 space-y-8">
            <LayoutControls options={options} setOptions={setOptions} />

            <div className="bg-white p-7 rounded-[2.5rem] shadow-sm border border-slate-200 overflow-hidden space-y-8 relative">
              <section>
                <div className="mb-6 flex justify-between items-start">
                  <div>
                    <h3 className="font-black text-slate-800 text-base mb-1">Page {currentPage} - STEP1:画像編集・段落選択</h3>
                  </div>
                  {history.length > 0 && (
                    <button onClick={handleUndo} className="px-3 py-1.5 bg-slate-50 text-slate-500 rounded-xl hover:bg-orange-50 hover:text-orange-600 transition-all border border-slate-200 flex items-center gap-1.5 shadow-sm active:scale-95">
                      <svg className="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="3" d="M3 10h10a8 8 0 018 8v2M3 10l6 6m-6-6l6-6" /></svg>
                      <span className="text-[10px] font-black uppercase tracking-widest">戻す</span>
                    </button>
                  )}
                </div>

                <div className="space-y-4 max-h-[40vh] overflow-y-auto pr-2 custom-scrollbar">
                  {unassignedImages.map(img => (
                    <div
                      key={img.id}
                      className="p-3 bg-slate-50/50 border border-slate-100 rounded-[2rem] hover:border-orange-200 transition-all shadow-sm flex items-start gap-3 group"
                    >
                      <div
                        style={{
                          flex: 1,
                          minWidth: 0,
                          height: "200px",
                          overflow: "hidden",
                          borderRadius: "12px",        // rounded-xl相当
                          border: "1px solid #e2e8f0", // border-slate-200相当（薄い枠）
                          background: "#ffffff",       // 白
                          boxShadow: "inset 0 1px 2px rgba(0,0,0,0.06)" // shadow-inner相当
                        }}
                        onClick={() => setPreviewId(img.id)}
                      >
                        <img
                          src={img.dataUrl}
                          alt={img.name}
                          style={{
                            width: "100%",
                            height: "100%",
                            objectFit: "cover",
                            transform: `rotate(${img.rotation}deg)`,
                            display: "block"
                          }}
                        />
                      </div>

                      <div className="w-[160px] flex-shrink-0">
                        <div className="flex flex-col gap-2">
                          <div className="flex gap-2">
                            <button
                              onClick={() => rotateImage(img.id, "left")}
                              className="w-12 h-12 rounded-lg bg-indigo-100 text-indigo-700 hover:bg-indigo-200 border border-indigo-200 transition-all shadow-sm flex items-center justify-center active:scale-95"
                            >
                              <svg className="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2.5" d="M4 4v5h.582m15.356 2A8.001 8.001 0 004.582 9m0 0H9" />
                              </svg>
                            </button>
                            <button
                              onClick={() => rotateImage(img.id, 'right')}
                              className="w-12 h-12 rounded-lg bg-indigo-100 text-indigo-700 hover:bg-indigo-200 border border-indigo-200 transition-all shadow-sm flex items-center justify-center active:scale-95"
                            >
                              <svg className="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2.5" d="M20 4v5h-.582m-15.356 2A8.001 8.001 0 0119.418 9m0 0H15" />
                              </svg>
                            </button>
                            <button
                              onClick={() => removeImage(img.id)}
                              className="w-12 h-12 rounded-lg bg-rose-100 text-rose-700 hover:bg-rose-200 border border-rose-200 transition-all shadow-sm flex items-center justify-center active:scale-95"
                            >
                              <svg className="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2.5" d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16" />
                              </svg>
                            </button>
                          </div>
                          <div className="grid grid-cols-2 gap-1.5">
                            {[1, 2, 3, 4].map(num => (
                              <button
                                key={num}
                                onClick={() => {
                                  const isLastUnassigned = unassignedImages.length === 1;
                                  updateImageRow(img.id, num);
                                  if (isLastUnassigned) {
                                    setImages((prev: ImageData[]) =>
                                      prev.map(i => (i.row > 0 && !i.orderConfirmed ? { ...i, orderConfirmed: true } : i))
                                    );
                                    if (currentPage === 1) setPage1Confirmed(true);
                                    if (currentPage === 2) setPage2Confirmed(true);
                                    if (currentPage === 3) setPage3Confirmed(true);
                                  }
                                }}
                                className={`h-12 rounded-lg text-lg font-semibold transition-all shadow-sm active:scale-95 border ${img.row === num
                                    ? 'bg-orange-500 text-white border-orange-400 shadow font-bold'
                                    : 'bg-white text-slate-800 border-slate-300 hover:bg-slate-100'
                                  }`}
                              >
                                {num}
                              </button>
                            ))}
                          </div>
                        </div>
                      </div>
                    </div>
                  ))}


                </div>
              </section>
            </div>
          </div>
        )}

        {/* 段落ドラッグ移動 */}
        <div ref={rowBoardRef} className="lg:col-span-12 bg-white p-7 rounded-[2.5rem] shadow-sm border border-slate-200 space-y-4">
          <div>
            <h3 className="font-black text-slate-800 text-base mb-1">
              {isCurrentPageConfirmed ? `Page ${currentPage} - STEP2:画像入替` : '段落ドラッグ移動'}
            </h3>
          </div>
          <RowBoard images={images} setImages={setImages} rows={4} />
        </div>

        {/* PAGE切替ボタン（段落エリアとプレビューの間） */}
        <div className="lg:col-span-12 grid grid-cols-3 items-center my-4">
  <div className="justify-self-start min-w-[180px]" />
  <div className="justify-self-center">
    <PageSwitcher
      currentPage={currentPage}
      onChange={handlePageChange}
      size="small"
      pages={availablePages}
    />
  </div>
  <div className="justify-self-end min-w-[180px]" />
</div>

        <div className="lg:col-span-12 bg-white p-6 rounded-[2.5rem] shadow-sm border border-slate-200 overflow-hidden">
          <div
            className="p-4 flex justify-center items-center overflow-hidden"
            style={{ backgroundColor: 'rgba(241,245,249,0.5)' }}
          >
            <div
              className="shadow-[0_15px_30px_rgba(0,0,0,0.1)] bg-white border border-slate-300 relative"
              style={{ height: '480px', aspectRatio: '210 / 297' }}
            >
              <div
                className="w-full h-full"
                style={{ color: "#0f172a" }}
                dangerouslySetInnerHTML={{ __html: svgData.svgCode }}
              />
            </div>
          </div>
        </div>
      </main>

      {isPrintMode &&
  createPortal(
    <div id="print-area">
      {outputPages.map((pageNum) => (
        <div
          key={pageNum}
          id={`print-page-${pageNum}`}
          className="print-page"
          style={{ ...PRINT_SAFE_CSS_VARS, background: "#fff", color: "#0f172a" }}
          dangerouslySetInnerHTML={{
            __html: pageNum === 1 ? svgPage1 : pageNum === 2 ? svgPage2 : svgPage3,
          }}
        />
      ))}
    </div>,
    document.body
  )
}

      {/* Sticky bottom bar */}
      <div className="sticky bottom-0 z-50 bg-white/90 backdrop-blur border-t border-slate-200 p-3">
    <div className="max-w-7xl mx-auto flex items-center justify-between px-6">
      <div>
        <h3 className="font-black text-slate-800 text-base">プレビュー</h3>
        <p className="text-[10px] text-slate-400 font-bold uppercase tracking-wider">確定済みの画像が反映されます</p>
        {pptxStatus && <p className="text-[11px] text-slate-600 font-bold mt-1">{pptxStatus}</p>}
      </div>
      <div className="flex flex-wrap gap-3">
        <button onClick={openGmailDraft}
          disabled={isCreatingDraft || !(reportFields.refHospitalEmail || '').trim()}
          title="Gmail下書きを作成し、PDFを添付します（送信はしません）"
          className="bg-slate-800 text-white border border-slate-700 px-5 py-2 rounded-xl text-[10px] font-black hover:bg-slate-700 transition-all flex items-center gap-2 disabled:opacity-50 disabled:cursor-not-allowed">
          {isCreatingDraft ? "作成中…" : "PDF / Gmail"}
        </button>
        <button onClick={printPdf}
          title="印刷ダイアログを開きます（PDF保存/プリンタ印刷）"
          className="bg-slate-800 text-white border border-slate-700 px-5 py-2 rounded-xl text-[10px] font-black hover:bg-slate-700 transition-all flex items-center gap-2">
          PDF 印刷
        </button>
        <button onClick={downloadPptx} disabled={isSavingPptx}
          className="bg-orange-600 text-white px-5 py-2 rounded-xl text-[10px] font-black hover:bg-orange-700 shadow-xl shadow-orange-900/20 flex items-center gap-2 transition-all disabled:opacity-50 disabled:cursor-not-allowed">
          {isSavingPptx ? '保存中…' : 'PPTX出力/編集'}
        </button>
      </div>
      </div>
      </div>
    </div>
  );
};


export default App;
