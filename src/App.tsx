import React, { useState, useRef, useCallback, useMemo, useEffect } from 'react';
import type { CSSProperties } from 'react';
import type { ImageData, LayoutOptions } from './types';
import LayoutControls from './components/LayoutControls';
import pptxgen from 'pptxgenjs';
import html2canvas from 'html2canvas';
import jsPDF from 'jspdf';
import { LAYOUT } from './layout';
import TemplatePicker from './components/TemplatePicker';
import { buildSvgTextParts, addPptxText, getPage1ImageStartYcm } from './reportTextRenderer';
import RowBoard from './components/RowBoard';
import { fetchSuggestions, addRefHospital } from "./serverApi";
import { createPortal } from "react-dom";


type AppSuggestions = {
  refHospitals: string[];
  doctors: string[];
  refHospitalEmails: Record<string, string>;
};

type PreviewYOffsetKey =
  | 'page1InitialHeader'
  | 'page1InitialBody'
  | 'page2FreeText'
  | 'page2ProcedureTitle'
  | 'page2ProcedureBody'
  | 'page2PostTitle'
  | 'page2PostBody'
  | 'page2ThanksBody'
  | 'page2PhotoCategoryTitle'
  | 'page3FreeTextArea'
  | 'page3FreeText'
  | 'page3PostopTitle'
  | 'page3PostopBody'
  | 'page3ThanksBody';

export type PreviewYOffsetMap = Record<PreviewYOffsetKey, number>;

type ImageCrop = {
  left: number;
  top: number;
  right: number;
  bottom: number;
};

const DEFAULT_IMAGE_CROP: ImageCrop = {
  left: 0,
  top: 0,
  right: 1,
  bottom: 1,
};

const IMAGE_CROP_MIN_SIZE = 0.12;
const CROP_UI_GUTTER_PX = 12;

type CropHandle = 'top-left' | 'top-right' | 'bottom-left' | 'bottom-right';
type CropDragMode = 'resize' | 'move';

type CropViewportRect = {
  left: number;
  top: number;
  width: number;
  height: number;
};

type CropPixelRect = {
  left: number;
  top: number;
  right: number;
  bottom: number;
};

function formatCaseDisplayId(caseId?: string): string {
  if (!caseId) return '';
  const parts = caseId.split('-');
  return parts[parts.length - 1] || caseId;
}

function formatDateShort(iso: string): string {
  if (!iso) return '';
  const d = new Date(iso);
  return `${d.getFullYear()}/${String(d.getMonth() + 1).padStart(2, '0')}/${String(d.getDate()).padStart(2, '0')}`;
}

function getDaysElapsed(registeredAt?: string): number | null {
  if (!registeredAt) return null;
  const d = new Date(registeredAt);
  if (Number.isNaN(d.getTime())) return null;
  return Math.floor((Date.now() - d.getTime()) / (1000 * 60 * 60 * 24));
}

function getDelayLabel(days: number | null): { text: string; className: string } {
  if (days == null) return { text: '-', className: 'bg-gray-100 text-gray-500' };
  if (days >= 8) return { text: '遅延', className: 'bg-red-100 text-red-700' };
  if (days >= 4) return { text: '注意', className: 'bg-yellow-100 text-yellow-700' };
  return { text: '正常', className: 'bg-green-100 text-green-700' };
}

const ATTENDING_VET_OPTIONS: string[] = ['町田健吾', '江成翔馬', '神田珠希', '小林嵩', '金田七海'];
const DELAY_APOLOGY_TEXT = 'ご報告が遅くなりましたことをお詫び申し上げます。';

function clampNumber(value: number, min: number, max: number): number {
  return Math.min(max, Math.max(min, value));
}

function normalizeImageCrop(crop?: Partial<ImageCrop> | null): ImageCrop {
  const left = Number.isFinite(Number((crop as any)?.left))
    ? Number((crop as any).left)
    : DEFAULT_IMAGE_CROP.left;
  const top = Number.isFinite(Number((crop as any)?.top))
    ? Number((crop as any).top)
    : DEFAULT_IMAGE_CROP.top;
  const right = Number.isFinite(Number((crop as any)?.right))
    ? Number((crop as any).right)
    : DEFAULT_IMAGE_CROP.right;
  const bottom = Number.isFinite(Number((crop as any)?.bottom))
    ? Number((crop as any).bottom)
    : DEFAULT_IMAGE_CROP.bottom;

  const normalizedLeft = clampNumber(left, 0, 1 - IMAGE_CROP_MIN_SIZE);
  const normalizedTop = clampNumber(top, 0, 1 - IMAGE_CROP_MIN_SIZE);
  const normalizedRight = clampNumber(right, normalizedLeft + IMAGE_CROP_MIN_SIZE, 1);
  const normalizedBottom = clampNumber(bottom, normalizedTop + IMAGE_CROP_MIN_SIZE, 1);

  return {
    left: normalizedLeft,
    top: normalizedTop,
    right: normalizedRight,
    bottom: normalizedBottom,
  };
}

function normalizeImageCropByHandle(crop: ImageCrop, handle: CropHandle): ImageCrop {
  let next = { ...crop };

  if (handle === 'top-left') {
    next.left = clampNumber(next.left, 0, next.right - IMAGE_CROP_MIN_SIZE);
    next.top = clampNumber(next.top, 0, next.bottom - IMAGE_CROP_MIN_SIZE);
  }

  if (handle === 'top-right') {
    next.right = clampNumber(next.right, next.left + IMAGE_CROP_MIN_SIZE, 1);
    next.top = clampNumber(next.top, 0, next.bottom - IMAGE_CROP_MIN_SIZE);
  }

  if (handle === 'bottom-left') {
    next.left = clampNumber(next.left, 0, next.right - IMAGE_CROP_MIN_SIZE);
    next.bottom = clampNumber(next.bottom, next.top + IMAGE_CROP_MIN_SIZE, 1);
  }

  if (handle === 'bottom-right') {
    next.right = clampNumber(next.right, next.left + IMAGE_CROP_MIN_SIZE, 1);
    next.bottom = clampNumber(next.bottom, next.top + IMAGE_CROP_MIN_SIZE, 1);
  }

  return normalizeImageCrop(next);
}

// 回転行列による座標変換
function rotatePoint(x: number, y: number, cx: number, cy: number, deg: number): { x: number, y: number } {
  const rad = (deg * Math.PI) / 180;
  const cos = Math.cos(rad);
  const sin = Math.sin(rad);
  const dx = x - cx;
  const dy = y - cy;
  return {
    x: cx + dx * cos - dy * sin,
    y: cy + dx * sin + dy * cos,
  };
}

function cropToPixels(
  crop: ImageCrop,
  viewport: Pick<CropViewportRect, 'width' | 'height'>,
  rotation: number = 0,
  flipX: boolean = false,
  flipY: boolean = false
): CropPixelRect {
  // 画像の中心
  const cx = viewport.width / 2;
  const cy = viewport.height / 2;
  // 非回転時の矩形
  let left = crop.left * viewport.width;
  let top = crop.top * viewport.height;
  let right = crop.right * viewport.width;
  let bottom = crop.bottom * viewport.height;
  // flipX/flipYを反映
  if (flipX) {
    const l = left;
    const r = right;
    left = viewport.width - r;
    right = viewport.width - l;
  }
  if (flipY) {
    const t = top;
    const b = bottom;
    top = viewport.height - b;
    bottom = viewport.height - t;
  }
  // 4隅を回転
  const p1 = rotatePoint(left, top, cx, cy, rotation);
  const p2 = rotatePoint(right, top, cx, cy, rotation);
  const p3 = rotatePoint(left, bottom, cx, cy, rotation);
  const p4 = rotatePoint(right, bottom, cx, cy, rotation);
  // 回転後の外接矩形
  const xs = [p1.x, p2.x, p3.x, p4.x];
  const ys = [p1.y, p2.y, p3.y, p4.y];
  return {
    left: Math.min(...xs),
    top: Math.min(...ys),
    right: Math.max(...xs),
    bottom: Math.max(...ys),
  };
}

function pixelsToCrop(
  pixelRect: CropPixelRect,
  viewport: Pick<CropViewportRect, 'width' | 'height'>,
  rotation: number = 0,
  flipX: boolean = false,
  flipY: boolean = false
): ImageCrop {
  // 画像の中心
  const cx = viewport.width / 2;
  const cy = viewport.height / 2;
  // 逆回転で4隅を元に戻す
  const p1 = rotatePoint(pixelRect.left, pixelRect.top, cx, cy, -rotation);
  const p2 = rotatePoint(pixelRect.right, pixelRect.top, cx, cy, -rotation);
  const p3 = rotatePoint(pixelRect.left, pixelRect.bottom, cx, cy, -rotation);
  const p4 = rotatePoint(pixelRect.right, pixelRect.bottom, cx, cy, -rotation);
  // 非回転時の外接矩形
  let xs = [p1.x, p2.x, p3.x, p4.x];
  let ys = [p1.y, p2.y, p3.y, p4.y];
  const w = viewport.width;
  const h = viewport.height;
  if (w <= 0 || h <= 0) return DEFAULT_IMAGE_CROP;
  // flipX/flipYを逆変換
  if (flipX) {
    xs = xs.map(x => w - x);
  }
  if (flipY) {
    ys = ys.map(y => h - y);
  }
  return normalizeImageCrop({
    left: Math.min(...xs) / w,
    top: Math.min(...ys) / h,
    right: Math.max(...xs) / w,
    bottom: Math.max(...ys) / h,
  });
}

function normalizeCropPixelRectByHandle(
  pixelRect: CropPixelRect,
  handle: CropHandle,
  viewport: Pick<CropViewportRect, 'width' | 'height'>
): CropPixelRect {
  const minW = IMAGE_CROP_MIN_SIZE * viewport.width;
  const minH = IMAGE_CROP_MIN_SIZE * viewport.height;
  const next = { ...pixelRect };

  if (handle === 'top-left') {
    next.left = clampNumber(next.left, 0, next.right - minW);
    next.top = clampNumber(next.top, 0, next.bottom - minH);
  }

  if (handle === 'top-right') {
    next.right = clampNumber(next.right, next.left + minW, viewport.width);
    next.top = clampNumber(next.top, 0, next.bottom - minH);
  }

  if (handle === 'bottom-left') {
    next.left = clampNumber(next.left, 0, next.right - minW);
    next.bottom = clampNumber(next.bottom, next.top + minH, viewport.height);
  }

  if (handle === 'bottom-right') {
    next.right = clampNumber(next.right, next.left + minW, viewport.width);
    next.bottom = clampNumber(next.bottom, next.top + minH, viewport.height);
  }

  return next;
}

const PREVIEW_Y_OFFSET_UI_GROUPS: Array<{
  section: 'PAGE1' | 'PAGE2' | 'PAGE3';
  items: Array<{ key: PreviewYOffsetKey; label: string }>;
}> = [
  {
    section: 'PAGE1',
    items: [
      { key: 'page1InitialHeader', label: '【初診時】タイトル' },
      { key: 'page1InitialBody', label: '【初診時】本文' },
    ],
  },
  {
    section: 'PAGE2',
    items: [
      { key: 'page2FreeText', label: '【自由記載欄】（PAGE2）' },
      { key: 'page2PhotoCategoryTitle', label: 'PAGE2 写真区分ラベル タイトル' },
      { key: 'page2ProcedureTitle', label: '【検査・処置内容】タイトル' },
      { key: 'page2ProcedureBody', label: '【検査・処置内容】本文' },
      { key: 'page2PostTitle', label: '【術後経過】タイトル' },
      { key: 'page2PostBody', label: '【術後経過】本文' },
      { key: 'page2ThanksBody', label: '【お礼文】本文' },
    ],
  },
  {
    section: 'PAGE3',
    items: [
      { key: 'page3FreeText', label: '【自由記載欄】' },
      { key: 'page3PostopTitle', label: '【術後経過】タイトル' },
      { key: 'page3PostopBody', label: '【術後経過】本文' },
      { key: 'page3ThanksBody', label: '【お礼文】本文' },
    ],
  },
];

const PREVIEW_Y_OFFSET_INITIAL: PreviewYOffsetMap = {
  page1InitialHeader: 0,
  page1InitialBody: 0,
  page2FreeText: 0,
  page2ProcedureTitle: 0,
  page2ProcedureBody: 0,
  page2PostTitle: 0,
  page2PostBody: 0,
  page2ThanksBody: 0,
  page2PhotoCategoryTitle: 0,
  page3FreeTextArea: 0,
  page3FreeText: 0,
  page3PostopTitle: 0,
  page3PostopBody: 0,
  page3ThanksBody: 0,
};

const REPORT_FIELDS_STORAGE_KEY = "photo-report-aligner:reportFields";
const SERVER_BASE_URL = (import.meta as any).env?.VITE_SERVER_BASE_URL?.trim() || "http://localhost:8787";
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
    sedationDate: '',
    anesthesiaDate: '',
    attendingVet: '町田健吾',
    initialText: '',
    procedureText: '',
    postText: '',
    thankYouTextType: 'first-time',
    page3Text: '',
    chiefComplaint: '',
    page2PhotoCategory: 'treatment-after',
    page3PhotoLabel: '',
  };
}

function normalizeHospitalKey(name: string): string {
  return String(name)
    .trim()
    .replace(/\u3000/g, ' ')
    .replace(/\s+/g, ' ');
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
              ? `w-[112px] h-[36px] inline-flex items-center justify-center gap-2 rounded-xl text-sm font-semibold transition-colors border ${
                  currentPage === p
                    ? "bg-violet-600 text-white border-violet-600"
                    : "bg-white text-slate-600 border-slate-300 hover:bg-slate-50"
                }`
              : `w-[88px] h-[36px] inline-flex items-center justify-center rounded-lg text-base font-semibold transition-colors border ${
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
    // 紹介病院名input用ref
    // --- 編集モード用state ---
      const [editingImageId, setEditingImageId] = useState<string | null>(null);
      // 編集モード用サブステート
      const [editStep, setEditStep] = useState<'orientation' | 'crop'>('orientation');
      const [isOrientationOpen, setIsOrientationOpen] = useState(false);
      const [isCropOpen, setIsCropOpen] = useState(false);
    const refHospitalInputRef = useRef<HTMLInputElement | null>(null);
  const [showPage3, setShowPage3] = useState(false);

  // ページ管理
  const [currentPage, setCurrentPage] = useState<number>(1);

  const [reportFieldsHydrated, setReportFieldsHydrated] = useState(false);

  const availablePages = useMemo<number[]>(() => (showPage3 ? [1, 2, 3] : [1, 2]), [showPage3]);

  const handlePageChange = useCallback((page: number) => {
    if (page === 3 && !showPage3) {
      setShowPage3(true);
    }
    setCurrentPage(page);
    closeEditUiState();
  }, [showPage3]);

  // showPage3 をOFFにした瞬間に 3ページ目に居たら 2へ戻す
  useEffect(() => {
    if (!showPage3 && currentPage === 3) setCurrentPage(2);
  }, [showPage3, currentPage]);

  // 編集系 state の終了処理ヘルパー
  const closeCropState = () => {
    setActiveCropImageId(null);
    setActiveCropViewportRect(null);
    setIsCropOpen(false);
  };

  const closeEditUiState = () => {
    setIsOrientationOpen(false);
    setIsCropOpen(false);
    setEditingImageId(null);
    setActiveCropImageId(null);
    setActiveCropViewportRect(null);
  };

  // editingImageId が null になったとき編集用 open state を強制クリア
  useEffect(() => {
    if (editingImageId === null) {
      setIsOrientationOpen(false);
      setIsCropOpen(false);
      setActiveCropImageId(null);
      setActiveCropViewportRect(null);
    }
  }, [editingImageId]);

  // 既存：参照や他state
  const rowBoardRef = useRef<HTMLDivElement | null>(null);
  const fileInputRef = useRef<HTMLInputElement | null>(null);
  const attendingVetDropdownRef = useRef<HTMLDivElement | null>(null);
  const [isAttendingVetDropdownOpen, setIsAttendingVetDropdownOpen] = useState(false);
  const page2PhotoCategoryDropdownRef = useRef<HTMLDivElement | null>(null);
  // PAGE2「検査・処置内容」textarea用ref
  const page2ProcedureTextareaRef = useRef<HTMLTextAreaElement | null>(null);
  const [isPage2PhotoCategoryDropdownOpen, setIsPage2PhotoCategoryDropdownOpen] = useState(false);
  const postPlacementDropdownRef = useRef<HTMLDivElement | null>(null);
  const [isPostPlacementDropdownOpen, setIsPostPlacementDropdownOpen] = useState(false);
  const thankYouTextTypeDropdownRef = useRef<HTMLDivElement | null>(null);
  const pageSwitcherRef = useRef<HTMLDivElement | null>(null);
  const imageToolbarRef = useRef<HTMLDivElement | null>(null);
  const [isThankYouTextTypeDropdownOpen, setIsThankYouTextTypeDropdownOpen] = useState(false);
  const [dropdownHighlight, setDropdownHighlight] = useState(-1);
  const shouldOpenAttendingVetOnFocusRef = useRef(false);
  const shouldOpenPostPlacementOnFocusRef = useRef(false);
  const shouldOpenPage2PhotoCategoryOnFocusRef = useRef(false);
  const shouldOpenThankYouTextTypeOnFocusRef = useRef(false);
  const holdTimeoutRef = useRef<number | null>(null);
  const holdIntervalRef = useRef<number | null>(null);
  const suppressNextClickRef = useRef(false);

  // ...この下に既存の state / useEffect / handlers が続く

  // ===== 参照病院（候補取得＆保存）=====
  const [refHospitalInput, setRefHospitalInput] = useState<string>("");
  const [refHospitalError, setRefHospitalError] = useState<string | null>(null);
  const [isSavingRefHospital, setIsSavingRefHospital] = useState(false);
  const [showHospitalSavedMessage, setShowHospitalSavedMessage] = useState(false);
  const hospitalSavedMessageTimeoutRef = useRef<number | null>(null);

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

  // --- 患者一覧 ---
  const [reportCases, setReportCases] = useState<any[]>([]);
  const [selectedCaseId, setSelectedCaseId] = useState<string | null>(null);
  const [selectedCaseStatus, setSelectedCaseStatus] = useState<string>('');

  useEffect(() => {
    fetch('http://localhost:8787/api/report-cases')
      .then(res => res.json())
      .then(data => setReportCases(data))
      .catch(err => console.error('[report-cases]', err));
  }, []);

  const handleSelectCase = useCallback(async (c: any) => {
    setSelectedCaseId(c.case_id);
    setSelectedCaseStatus(c.status || '');
    setRefHospitalInput(c.referring_hospital || '');
    setReportFields((prev) => ({
      ...prev,
      ownerLastName: c.owner_last_name || '',
      petName: c.pet_name || '',
      attendingVet: c.attending_vet || '',
      refHospitalName: c.referring_hospital || '',
      refHospital: c.referring_hospital || '',
      refDoctor: c.referring_doctor_name || '',
    }));

    // 下書き復元
    try {
      const res = await fetch(`http://localhost:8787/api/report-drafts/${c.case_id}`);
      if (res.ok) {
        const data = await res.json();
        const raw = data?.draft_data_json;
        const d = typeof raw === 'string' ? JSON.parse(raw) : raw;
        if (d && typeof d === 'object') {
          setRefHospitalInput(d.refHospitalName || d.refHospital || c.referring_hospital || '');
          setReportFields((prev) => ({ ...prev, ...d }));
        }
      }
    } catch (e) {
      console.log('[draft] load error', e);
    }
  }, []);

  // --- 新規患者登録 ---
  const [newCaseOwnerLastName, setNewCaseOwnerLastName] = useState('');
  const [newCasePetName, setNewCasePetName] = useState('');
  const [newCaseAttendingVet, setNewCaseAttendingVet] = useState('町田健吾');
  const [isCreatingCase, setIsCreatingCase] = useState(false);

  const handleCreateCase = useCallback(async () => {
    if (!newCaseOwnerLastName.trim() || !newCasePetName.trim()) {
      alert('飼い主姓とペット名を入力してください。');
      return;
    }

    try {
      setIsCreatingCase(true);

      const res = await fetch('http://localhost:8787/api/report-cases', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          attending_vet: newCaseAttendingVet || '未入力',
          owner_last_name: newCaseOwnerLastName.trim(),
          pet_name: newCasePetName.trim(),
          referring_hospital: '未入力',
          referring_doctor_name: '未入力',
        }),
      });

      if (!res.ok) {
        throw new Error('新規患者登録に失敗しました。');
      }

      const created = await res.json();

      setReportCases((prev) => [created, ...prev]);
      handleSelectCase(created);

      setNewCaseOwnerLastName('');
      setNewCasePetName('');
      setNewCaseAttendingVet('町田健吾');

      alert('新規患者を登録しました。');
    } catch (err: any) {
      alert(err?.message || '新規患者登録に失敗しました。');
    } finally {
      setIsCreatingCase(false);
    }
  }, [newCaseOwnerLastName, newCasePetName, newCaseAttendingVet, handleSelectCase]);

  const normalizedRefHospitalEmails = useMemo(() => {
    const map: Record<string, string> = {};
    Object.entries(suggestions.refHospitalEmails || {}).forEach(([k, v]) => {
      map[normalizeHospitalKey(k)] = v;
    });
    return map;
  }, [suggestions.refHospitalEmails]);

  const normalizedRefHospitalNames = useMemo(() => {
    return new Set((suggestions.refHospitals || []).map((name) => normalizeHospitalKey(name)));
  }, [suggestions.refHospitals]);

  const normalizedCurrentRefHospitalName = useMemo(
    () => normalizeHospitalKey(refHospitalInput),
    [refHospitalInput]
  );

  const shouldShowRegisterRefHospitalButton = useMemo(() => {
    if (!normalizedCurrentRefHospitalName) return false;
    if (normalizedRefHospitalNames.has(normalizedCurrentRefHospitalName)) return false;
    if (isSavingRefHospital) return false;
    return true;
  }, [normalizedCurrentRefHospitalName, normalizedRefHospitalNames, isSavingRefHospital]);

  useEffect(() => {
    return () => {
      if (hospitalSavedMessageTimeoutRef.current !== null) {
        window.clearTimeout(hospitalSavedMessageTimeoutRef.current);
      }
    };
  }, []);

  const getThankYouTextTypeByHospital = useCallback((hospitalName: string) => {
    const normalizedName = normalizeHospitalKey(hospitalName);
    return normalizedName && normalizedRefHospitalNames.has(normalizedName) ? 'existing' : 'first-time';
  }, [normalizedRefHospitalNames]);

  const applyRefHospitalSelection = useCallback((hospitalName: string) => {
    const normalizedName = normalizeHospitalKey(hospitalName);
    const mappedEmail = normalizedName ? normalizedRefHospitalEmails[normalizedName] : "";
    const nextThankYouTextType = getThankYouTextTypeByHospital(hospitalName);

    setRefHospitalInput(hospitalName);
    setReportFields((prev) => {
      return {
        ...prev,
        refHospitalName: hospitalName,
        refHospital: hospitalName,
        refHospitalEmail: mappedEmail || "",
        thankYouTextType: nextThankYouTextType,
      };
    });
  }, [normalizedRefHospitalEmails, getThankYouTextTypeByHospital]);

  // 参照病院を保存
  const handleAddRefHospital = useCallback(
    async (nameArg?: string, emailArg?: string) => {
      const name = (nameArg ?? refHospitalInput).trim();
      if (!name) return;

      if (normalizedRefHospitalNames.has(normalizeHospitalKey(name))) {
        return;
      }

      setRefHospitalError(null);
      setIsSavingRefHospital(true);

      try {
        await addRefHospital(name, emailArg);

        // 最新候補を再取得
        const latest = await fetchSuggestions();
        setSuggestions({
          refHospitals: Array.isArray(latest?.refHospitals) ? latest.refHospitals : [],
          doctors: Array.isArray(latest?.doctors) ? latest.doctors : [],
          refHospitalEmails: latest?.refHospitalEmails ?? {},
        });

        setShowHospitalSavedMessage(true);
        if (hospitalSavedMessageTimeoutRef.current !== null) {
          window.clearTimeout(hospitalSavedMessageTimeoutRef.current);
        }
        hospitalSavedMessageTimeoutRef.current = window.setTimeout(() => {
          setShowHospitalSavedMessage(false);
          hospitalSavedMessageTimeoutRef.current = null;
        }, 2500);

        // 入力欄とreportFieldsを正規化して揃える
        applyRefHospitalSelection(name);
      } catch (e) {
        console.error(e);
        setRefHospitalError("保存に失敗しました（CORS/サーバ/URL確認）");
      } finally {
        setIsSavingRefHospital(false);
      }
    },
    [
      refHospitalInput,
      applyRefHospitalSelection,
      normalizedRefHospitalNames,
    ]
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
  const [activeCropImageId, setActiveCropImageId] = useState<string | null>(null);
  const [activeCropViewportRect, setActiveCropViewportRect] = useState<CropViewportRect | null>(null);
  const cropDragStateRef = useRef<{
    mode: CropDragMode;
    imageId: string;
    handle: CropHandle | null;
    startX: number;
    startY: number;
    startCrop: ImageCrop;
  } | null>(null);

  // 報告書テキスト入力ステート（この下に既存の reportFields を続けてOK）
  const [reportFields, setReportFields] = useState(getInitialReportFields);
  const [previewYOffsets, setPreviewYOffsets] = useState<PreviewYOffsetMap>(PREVIEW_Y_OFFSET_INITIAL);

  const page2PhotoCategoryLabel = useMemo(() => {
    if (reportFields.page2PhotoCategory === 'treatment-after') return '【治療時・治療後写真】';
    if (reportFields.page2PhotoCategory === 'inspection') return '【検査時写真】';
    return '';
  }, [reportFields.page2PhotoCategory]);

  const page3PhotoLabelText = useMemo(() => {
    const label = String((reportFields as any).page3PhotoLabel || '').trim();
    if (!label) return '';
    return `【${label}】`;
  }, [(reportFields as any).page3PhotoLabel]);

  // --- 下書き保存 ---
  const handleSaveDraft = useCallback(async () => {
    if (!selectedCaseId) {
      alert('患者を選択してください');
      return;
    }
    try {
      const res = await fetch(`http://localhost:8787/api/report-drafts/${selectedCaseId}`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ draft_data_json: reportFields }),
      });
      if (!res.ok) throw new Error('保存失敗');

      // 未着手 → 報告書作成途中 に自動更新
      if (selectedCaseId && selectedCaseStatus === '未着手') {
        try {
          const patchRes = await fetch(`http://localhost:8787/api/report-cases/${selectedCaseId}/status`, {
            method: 'PATCH',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ status: '報告書作成途中' }),
          });
          if (patchRes.ok) {
            const updated = await patchRes.json();
            setSelectedCaseStatus(updated.status || '');
            setReportCases(prev => prev.map(c => c.case_id === selectedCaseId ? updated : c));
          }
        } catch (e) {
          console.log('[auto status] update failed', e);
        }
      }

      // 基本情報を report_cases に同期（副処理）
      try {
        const syncRes = await fetch(`http://localhost:8787/api/report-cases/${selectedCaseId}`, {
          method: 'PATCH',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({
            attending_vet: reportFields.attendingVet,
            owner_last_name: reportFields.ownerLastName,
            pet_name: reportFields.petName,
            referring_hospital: reportFields.refHospitalName || reportFields.refHospital,
            referring_doctor_name: reportFields.refDoctor,
          }),
        });
        if (syncRes.ok) {
          const updated = await syncRes.json();
          setReportCases(prev => prev.map(c => c.case_id === selectedCaseId ? updated : c));
        }
      } catch (e) {
        console.log('[case sync] failed', e);
      }

      alert('下書きを保存しました');
    } catch {
      alert('下書きの保存に失敗しました');
    }
  }, [selectedCaseId, selectedCaseStatus, reportFields]);

  // --- ステータス更新 ---
  const handleMarkMailSent = useCallback(async () => {
    if (!selectedCaseId) { alert('患者を選択してください'); return; }
    try {
      const res = await fetch(`http://localhost:8787/api/report-cases/${selectedCaseId}/status`, {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ status: 'メール送信済み', mail_sent_at: new Date().toISOString() }),
      });
      if (!res.ok) throw new Error('更新失敗');
      const updated = await res.json();
      setSelectedCaseStatus(updated.status || '');
      setReportCases(prev => prev.map(item => item.case_id === selectedCaseId ? updated : item));
      alert('ステータスをメール送信済みに更新しました');
    } catch { alert('ステータス更新に失敗しました'); }
  }, [selectedCaseId]);

  const handleMarkPostalSent = useCallback(async () => {
    if (!selectedCaseId) { alert('患者を選択してください'); return; }
    try {
      const res = await fetch(`http://localhost:8787/api/report-cases/${selectedCaseId}/status`, {
        method: 'PATCH',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ status: '印刷郵送済み', postal_sent_at: new Date().toISOString() }),
      });
      if (!res.ok) throw new Error('更新失敗');
      const updated = await res.json();
      setSelectedCaseStatus(updated.status || '');
      setReportCases(prev => prev.map(item => item.case_id === selectedCaseId ? updated : item));
      alert('ステータスを印刷郵送済みに更新しました');
    } catch { alert('ステータス更新に失敗しました'); }
  }, [selectedCaseId]);

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
    const page1ImageStartY = getPage1ImageStartYcm(reportFields);
    return {
      left: config.LEFT,
      right: config.RIGHT,
      startY: page1ImageStartY,
      maxW: LAYOUT.SLIDE.WIDTH_CM - config.LEFT - config.RIGHT,
      maxH: LAYOUT.SLIDE.HEIGHT_CM - page1ImageStartY - config.MARGIN_BOTTOM,
      alignLeft: config.ALIGN_LEFT
    };
  }, [reportFields]);


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
  const [isSendingGmail, setIsSendingGmail] = useState(false);
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

const handleUnassignImage = useCallback((id: string) => {
  setImages((prev: ImageData[]) =>
    prev.map(i => (i.id === id ? { ...i, row: 0 } : i))
  );

  closeCropState();

  if (currentPage === 1) setPage1Confirmed(false);
  if (currentPage === 2) setPage2Confirmed(false);
  if (currentPage === 3) setPage3Confirmed(false);
}, [currentPage, setImages]);

// PAGE2を出力に含めるか（PAGE3追加時のみ有効）
const [includePage2InExport, setIncludePage2InExport] = useState(true);

// どこに「経過」を入れるか（既存がこれならOK）
const [postPlacement, setPostPlacement] = useState<"page2" | "page3">("page2");


// 出力順（PDF/PPTXの並びなどで使う想定）
const outputPages = useMemo<number[]>(() => {
  if (!showPage3) return [1, 2];
  return includePage2InExport ? [1, 2, 3] : [1, 3];
}, [showPage3, includePage2InExport]);

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
    const inputEl = e.currentTarget;
    const files = Array.from(e.target.files || []) as File[];
    if (!files.length) {
      inputEl.value = '';
      return;
    }
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
                rotation: 0,
                flipX: false,
                flipY: false,
                originalDataUrl: dataUrl,
                originalWidth: img.width,
                originalHeight: img.height
              });
            };
            img.src = dataUrl;
          };
          reader.readAsDataURL(file);
        });
      })
    );
    setImages((prev: ImageData[]) => [...prev, ...newImages]);
    closeEditUiState();
    if (currentPage === 1) setPage1Confirmed(false);
    if (currentPage === 2) setPage2Confirmed(false);
    if (currentPage === 3) setPage3Confirmed(false);
    if (fileInputRef.current) fileInputRef.current.value = '';
    else inputEl.value = '';
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

  const flipImageX = (id: string) => {
    setImages((prev: ImageData[]) => prev.map(img => {
      if (img.id !== id) return img;
      let newCrop = (img as any).crop;
      if (newCrop && typeof newCrop === 'object' && activeCropViewportRect) {
        const cropPx = cropToPixels(
          normalizeImageCrop(newCrop),
          activeCropViewportRect,
          img.rotation,
          img.flipX,
          img.flipY
        );
        // X反転: flipXをトグルして新しい座標系に変換
        newCrop = pixelsToCrop(
          cropPx,
          activeCropViewportRect,
          img.rotation,
          !img.flipX,
          img.flipY
        );
      }
      return { ...img, flipX: !img.flipX, crop: newCrop };
    }));
  };

  const flipImageY = (id: string) => {
    setImages((prev: ImageData[]) => prev.map(img => {
      if (img.id !== id) return img;
      let newCrop = (img as any).crop;
      if (newCrop && typeof newCrop === 'object' && activeCropViewportRect) {
        const cropPx = cropToPixels(
          normalizeImageCrop(newCrop),
          activeCropViewportRect,
          img.rotation,
          img.flipX,
          img.flipY
        );
        // Y反転: flipYをトグルして新しい座標系に変換
        newCrop = pixelsToCrop(
          cropPx,
          activeCropViewportRect,
          img.rotation,
          img.flipX,
          !img.flipY
        );
      }
      return { ...img, flipY: !img.flipY, crop: newCrop };
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

  const getImageCrop = useCallback((img: ImageData): ImageCrop | null => {
    const crop = (img as any).crop;
    if (!crop || typeof crop !== 'object') return null;
    return normalizeImageCrop(crop as Partial<ImageCrop>);
  }, []);

  const updateImageCrop = useCallback((id: string, updater: (current: ImageCrop) => ImageCrop) => {
    setImages((prev: ImageData[]) => {
      return prev.map((img) => {
        if (img.id !== id) return img;
        const current = normalizeImageCrop((img as any).crop as Partial<ImageCrop>);
        const next = normalizeImageCrop(updater(current));
        return { ...img, crop: next } as ImageData;
      });
    });
  }, [setImages]);

  const resetImageCrop = useCallback((id: string) => {
    const target = images.find(i => i.id === id);
    console.log('RESET CROP', { imgId: id, beforeDataUrlLength: target?.dataUrl?.length, beforeOriginalDataUrlLength: target?.originalDataUrl?.length, beforeCrop: target?.crop });
    setImages((prev: ImageData[]) => {
      return prev.map((img) => {
        if (img.id !== id) return img;
        return { ...img, crop: undefined } as ImageData;
      });
    });
  }, [setImages, images]);

  const startCropDrag = useCallback((e: React.MouseEvent, id: string, handle: CropHandle) => {
    e.preventDefault();
    e.stopPropagation();
    if (activeCropImageId !== id) return;

    const target = images.find((img) => img.id === id);
    if (!target) return;

    recordHistory();
    cropDragStateRef.current = {
      mode: 'resize',
      imageId: id,
      handle,
      startX: e.clientX,
      startY: e.clientY,
      startCrop: getImageCrop(target) ?? DEFAULT_IMAGE_CROP,
    };
  }, [activeCropImageId, getImageCrop, images, recordHistory]);

  const startCropMove = useCallback((e: React.MouseEvent, id: string) => {
    e.preventDefault();
    e.stopPropagation();
    if (activeCropImageId !== id) return;

    const target = images.find((img) => img.id === id);
    if (!target) return;

    recordHistory();
    cropDragStateRef.current = {
      mode: 'move',
      imageId: id,
      handle: null,
      startX: e.clientX,
      startY: e.clientY,
      startCrop: getImageCrop(target) ?? DEFAULT_IMAGE_CROP,
    };
  }, [activeCropImageId, getImageCrop, images, recordHistory]);

  const stopCropDrag = useCallback(() => {
    cropDragStateRef.current = null;
  }, []);

  const getDisplayedImageRect = useCallback((img: ImageData): CropViewportRect | null => {
    const host = document.getElementById(`crop-host-${img.id}`);
    if (!host) return null;

    const hostRect = host.getBoundingClientRect();
    if (hostRect.width <= 0 || hostRect.height <= 0) return null;

    const hostStyle = window.getComputedStyle(host);
    const padLeft = Number.parseFloat(hostStyle.paddingLeft) || 0;
    const padRight = Number.parseFloat(hostStyle.paddingRight) || 0;
    const padTop = Number.parseFloat(hostStyle.paddingTop) || 0;
    const padBottom = Number.parseFloat(hostStyle.paddingBottom) || 0;
    const contentWidth = hostRect.width - padLeft - padRight;
    const contentHeight = hostRect.height - padTop - padBottom;
    if (contentWidth <= 0 || contentHeight <= 0) return null;

    const rotated = img.rotation === 90 || img.rotation === 270;
    const mediaW = rotated ? img.height : img.width;
    const mediaH = rotated ? img.width : img.height;
    if (!mediaW || !mediaH) return null;

    const hostAspect = contentWidth / contentHeight;
    const mediaAspect = mediaW / mediaH;

    if (!Number.isFinite(hostAspect) || !Number.isFinite(mediaAspect) || mediaAspect <= 0) return null;

    if (mediaAspect > hostAspect) {
      const width = contentWidth;
      const height = width / mediaAspect;
      return { left: padLeft, top: padTop + (contentHeight - height) / 2, width, height };
    }

    const height = contentHeight;
    const width = height * mediaAspect;
    return { left: padLeft + (contentWidth - width) / 2, top: padTop, width, height };
  }, []);

  useEffect(() => {
    if (!activeCropImageId) {
      setActiveCropViewportRect(null);
      return;
    }

    const target = images.find((img) => img.id === activeCropImageId);
    if (!target) {
      setActiveCropViewportRect(null);
      return;
    }

    const measure = () => {
      setActiveCropViewportRect(getDisplayedImageRect(target));
    };

    measure();
    window.addEventListener('resize', measure);

    return () => {
      window.removeEventListener('resize', measure);
    };
  }, [activeCropImageId, images, getDisplayedImageRect]);

  useEffect(() => {
    const onMouseMove = (event: MouseEvent) => {
      const dragState = cropDragStateRef.current;
      if (!dragState) return;

      const overlay = document.getElementById(`crop-overlay-${dragState.imageId}`);
      if (!overlay) return;

      const rect = overlay.getBoundingClientRect();
      if (rect.width <= 0 || rect.height <= 0) return;

      const viewport = { width: rect.width, height: rect.height };
      const dxPx = event.clientX - dragState.startX;
      const dyPx = event.clientY - dragState.startY;

      const base = dragState.startCrop;
      // 画像のrotationを考慮
      const img = images.find((img) => img.id === dragState.imageId);
      const rotation = img?.rotation ?? 0;
      const flipX = img?.flipX ?? false;
      const flipY = img?.flipY ?? false;
      const basePixels = cropToPixels(base, viewport, rotation, flipX, flipY);

      if (dragState.mode === 'move') {
        const width = basePixels.right - basePixels.left;
        const height = basePixels.bottom - basePixels.top;
        const nextLeft = clampNumber(basePixels.left + dxPx, 0, viewport.width - width);
        const nextTop = clampNumber(basePixels.top + dyPx, 0, viewport.height - height);
        const movedPixels: CropPixelRect = {
          left: nextLeft,
          top: nextTop,
          right: nextLeft + width,
          bottom: nextTop + height,
        };
        updateImageCrop(dragState.imageId, () => pixelsToCrop(movedPixels, viewport, rotation, flipX, flipY));
        return;
      }

      if (!dragState.handle) return;
      let nextPixels: CropPixelRect = { ...basePixels };

      if (dragState.handle === 'top-left') {
        nextPixels.left += dxPx;
        nextPixels.top += dyPx;
      }
      if (dragState.handle === 'top-right') {
        nextPixels.right += dxPx;
        nextPixels.top += dyPx;
      }
      if (dragState.handle === 'bottom-left') {
        nextPixels.left += dxPx;
        nextPixels.bottom += dyPx;
      }
      if (dragState.handle === 'bottom-right') {
        nextPixels.right += dxPx;
        nextPixels.bottom += dyPx;
      }

      const normalizedPixels = normalizeCropPixelRectByHandle(nextPixels, dragState.handle, viewport);
      const normalized = normalizeImageCropByHandle(pixelsToCrop(normalizedPixels, viewport, rotation, flipX, flipY), dragState.handle);
      updateImageCrop(dragState.imageId, () => normalized);
    };

    const onMouseUp = () => {
      if (!cropDragStateRef.current) return;
      stopCropDrag();
    };

    window.addEventListener('mousemove', onMouseMove);
    window.addEventListener('mouseup', onMouseUp);

    return () => {
      window.removeEventListener('mousemove', onMouseMove);
      window.removeEventListener('mouseup', onMouseUp);
    };
  }, [stopCropDrag, updateImageCrop]);

  const getImageDisplayStyle = useCallback((img: ImageData): CSSProperties => {
    const crop = getImageCrop(img);
    const isActiveCropTarget = activeCropImageId === img.id && !!activeCropViewportRect;
    console.log('DISPLAY STYLE', { imgId: img.id, crop: getImageCrop(img), isActiveCropTarget: activeCropImageId === img.id && !!activeCropViewportRect, activeCropViewportRect, rotation: img.rotation, flipX: img.flipX, flipY: img.flipY });

    const flipX = img.flipX ? -1 : 1;
    const flipY = img.flipY ? -1 : 1;
    if (isActiveCropTarget) {
      const insetTop = (crop ?? DEFAULT_IMAGE_CROP).top * 100;
      const insetRight = (1 - (crop ?? DEFAULT_IMAGE_CROP).right) * 100;
      const insetBottom = (1 - (crop ?? DEFAULT_IMAGE_CROP).bottom) * 100;
      const insetLeft = (crop ?? DEFAULT_IMAGE_CROP).left * 100;
      return {
        position: 'absolute',
        left: `${activeCropViewportRect!.left}px`,
        top: `${activeCropViewportRect!.top}px`,
        width: `${activeCropViewportRect!.width}px`,
        height: `${activeCropViewportRect!.height}px`,
        objectFit: 'contain',
        clipPath: `inset(${insetTop}% ${insetRight}% ${insetBottom}% ${insetLeft}%)`,
        transform: `rotate(${img.rotation}deg) scaleX(${flipX}) scaleY(${flipY})`,
        transformOrigin: 'center center',
        display: 'block',
      };
    }
    if (!crop) {
      return {
        width: '100%',
        height: '100%',
        objectFit: 'contain',
        transform: `rotate(${img.rotation}deg) scaleX(${flipX}) scaleY(${flipY})`,
        display: 'block',
      };
    }
    const insetTop = crop.top * 100;
    const insetRight = (1 - crop.right) * 100;
    const insetBottom = (1 - crop.bottom) * 100;
    const insetLeft = crop.left * 100;
    return {
      width: '100%',
      height: '100%',
      objectFit: 'contain',
      clipPath: `inset(${insetTop}% ${insetRight}% ${insetBottom}% ${insetLeft}%)`,
      transform: `rotate(${img.rotation}deg) scaleX(${flipX}) scaleY(${flipY})`,
      transformOrigin: 'center center',
      display: 'block',
    };
  }, [activeCropImageId, activeCropViewportRect, getImageCrop]);

  useEffect(() => {
    if (!activeCropImageId) return;
    if (images.some((img) => img.id === activeCropImageId)) return;
    setActiveCropImageId(null);
  }, [activeCropImageId, images]);




  // 第3段階スクロール（段落ドラッグ移動エリアへの自動スクロール）は無効化済み

  const unassignedImages = useMemo(() => images.filter(img => img.row === 0), [images]);

  const calculateLayoutForAnyPage = useCallback((rowsData: ImageData[][], pageNum: number) => {
    const activeRows = rowsData.filter(row => row.length > 0);
    if (activeRows.length === 0) return { rowResults: [], finalHeight: 0 };
    const dims = getPageDimensions(pageNum);
    const containerW = options.containerWidth;
    const baseSpacing = options.spacing;
    const targetBoxHPx = (dims.maxH / dims.maxW) * containerW;
    const isPage2 = pageNum === 2;
    let totalBlockHeight = 0;
    const rawData = activeRows.map(rowImages => {
      const count = rowImages.length;
      const ars = rowImages.map(img => {
        const isPortrait = img.rotation === 90 || img.rotation === 270;
        const baseAR = isPortrait ? (img.height / img.width) : (img.width / img.height);
        const crop = getImageCrop(img);
        if (!crop) return baseAR;
        const visW = crop.right - crop.left;
        const visH = crop.bottom - crop.top;
        if (visW <= 0 || visH <= 0) return baseAR;
        return baseAR * (isPortrait ? (visH / visW) : (visW / visH));
      });
      const totalAR = ars.reduce((a, b) => a + b, 0);
      const isFew = !isPage2 && count <= 1;
      let rowH;
      if (isPage2) {
        // PAGE2: always fill container width using post-crop AR, cap at full area height
        rowH = Math.min(
          (containerW - Math.max(0, count - 1) * baseSpacing) / totalAR,
          targetBoxHPx
        );
      } else if (isFew) {
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
    const fitScale = isPage2 ? 1.0 : Math.min(1, targetBoxHPx / totalBlockHeight);
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
  }, [options.containerWidth, options.spacing, getPageDimensions, getImageCrop]);

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

  const calculateSvgDataForPage = useCallback((
    pageNum: 1 | 2 | 3,
    renderOpts?: { applyPreviewOffsets?: boolean }
  ) => {
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
    const imageStartYcm = pageNum === 1 ? getPage1ImageStartYcm(reportFields) : imgCfg.START_Y;

    const svgParts: string[] = [];
    let page2ImagesBottomYcm: number | undefined;
    let page3ImagesBottomYcm: number | undefined;

    rowResults.forEach((row) => {
      row.images.forEach((item: any) => {
        const rawXcm = imgCfg.LEFT + item.x / pxPerCm;
        const rawYcm = imageStartYcm + item.y / pxPerCm;
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

        if (pageNum === 2) {
          const imageBottomYcm = placed.yCm + placed.hCm;
          page2ImagesBottomYcm =
            page2ImagesBottomYcm === undefined
              ? imageBottomYcm
              : Math.max(page2ImagesBottomYcm, imageBottomYcm);
        }

        if (pageNum === 3) {
          const imageBottomYcm = placed.yCm + placed.hCm;
          page3ImagesBottomYcm =
            page3ImagesBottomYcm === undefined
              ? imageBottomYcm
              : Math.max(page3ImagesBottomYcm, imageBottomYcm);
        }

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

    // use shared helper to render all text portions so that SVG and PPTX stay in sync
    let textParts = buildSvgTextParts(pageNum, reportFields, pxPerCm, slideOffsetX, slideOffsetY, {
      showPage3,
      postPlacement,
      indentPostOnPage3: true,
      page2ImagesBottomYcm: pageNum === 2 ? page2ImagesBottomYcm : undefined,
      page3ImagesBottomYcm: pageNum === 3 ? page3ImagesBottomYcm : undefined,
      previewYOffsets: (renderOpts?.applyPreviewOffsets && pageNum === 3) ? previewYOffsets : undefined,
    });

    const offsetTextY = (part: string, deltaPx: number): string => {
      if (!deltaPx || !part.includes('<text')) return part;
      return part.replace(/ y="([^"]+)"/, (_m, yStr: string) => {
        const baseY = Number(yStr);
        if (!Number.isFinite(baseY)) return _m;
        return ` y="${baseY + deltaPx}"`;
      });
    };

    const getTextX = (part: string): number | null => {
      const match = part.match(/<text[^>]* x="([^"]+)"/);
      if (!match) return null;
      const x = Number(match[1]);
      return Number.isFinite(x) ? x : null;
    };

    const nearlyEqual = (a: number, b: number) => Math.abs(a - b) < 0.1;

    // PAGE3かつ術後経過をPAGE2に置く設定のとき、PAGE3側の術後経過タイトル・本文を除去する
    if (pageNum === 3 && postPlacement === 'page2') {
      let skipNextBody = false;
      textParts = textParts.filter((part) => {
        if (part.includes('【術後経過】')) {
          skipNextBody = true;
          return false;
        }
        if (skipNextBody && part.includes('<text') && part.includes('<tspan') && !part.includes('【')) {
          skipNextBody = false;
          return false;
        }
        return true;
      });
    }

    if (renderOpts?.applyPreviewOffsets) {
      if (pageNum === 1) {
        const headerDelta = previewYOffsets.page1InitialHeader * pxPerCm;
        const bodyDelta = previewYOffsets.page1InitialBody * pxPerCm;
        const bodyX = slideOffsetX + LAYOUT.PAGE1.TEXT.FREE_TEXT_INITIAL.x * pxPerCm;

        let bodyOffsetApplied = false;
        textParts = textParts.map((part) => {
          if (headerDelta && part.includes('【初診時】')) {
            return offsetTextY(part, headerDelta);
          }

          if (!bodyOffsetApplied && bodyDelta && part.includes('<tspan') && !part.includes('【')) {
            const x = getTextX(part);
            if (x !== null && nearlyEqual(x, bodyX)) {
              bodyOffsetApplied = true;
              return offsetTextY(part, bodyDelta);
            }
          }

          return part;
        });
      }

      if (pageNum === 2) {
        const procTitleDelta = previewYOffsets.page2ProcedureTitle * pxPerCm;
        const procDelta = previewYOffsets.page2ProcedureBody * pxPerCm;
        const postTitleDelta = previewYOffsets.page2PostTitle * pxPerCm;
        const postDelta = previewYOffsets.page2PostBody * pxPerCm;
        const thanksDelta = previewYOffsets.page2ThanksBody * pxPerCm;

        let pendingBody: 'procedure' | 'post' | null = null;
        let bodyBlockIndex = 0;

        textParts = textParts.map((part) => {
          if (part.includes('【検査・処置内容】')) {
            pendingBody = 'procedure';
            return procTitleDelta ? offsetTextY(part, procTitleDelta) : part;
          }
          if (part.includes('【術後経過】')) {
            pendingBody = 'post';
            return postTitleDelta ? offsetTextY(part, postTitleDelta) : part;
          }

          if (part.includes('<text') && part.includes('<tspan') && !part.includes('【')) {
            if (pendingBody) {
              const delta = pendingBody === 'procedure' ? procDelta : postDelta;
              pendingBody = null;
              bodyBlockIndex += 1;
              return delta ? offsetTextY(part, delta) : part;
            }

            bodyBlockIndex += 1;
            if (bodyBlockIndex === 3 && thanksDelta) {
              return offsetTextY(part, thanksDelta);
            }
          }

          return part;
        });
      }

      // PAGE3 オフセットはrenderer側(buildSvgTextParts)で cursor-based 方式により適用済み
    }

    svgParts.push(...textParts);

    if (pageNum === 2 && page2PhotoCategoryLabel) {
      const labelX = slideOffsetX + 1.0 * pxPerCm;
      const labelBaseY = slideOffsetY + (LAYOUT.PAGE2.LINES.SEP_TOP.y + 0.3) * pxPerCm;
      const labelY = labelBaseY + previewYOffsets.page2PhotoCategoryTitle * pxPerCm;
      const labelFontSize = 0.42 * pxPerCm;
      svgParts.push(
        `  <text x="${labelX}" y="${labelY}" font-size="${labelFontSize}" font-weight="700" fill="#0f172a" dominant-baseline="hanging">${page2PhotoCategoryLabel}</text>`
      );
    }

    if (pageNum === 3 && page3PhotoLabelText && page3ImagesBottomYcm !== undefined) {
      const labelX = slideOffsetX + 1.0 * pxPerCm;
      const labelY = slideOffsetY + (LAYOUT.PAGE3.LINES.SEP_TOP.y + 0.3) * pxPerCm;
      const labelFontSize = 0.42 * pxPerCm;
      svgParts.push(
        `  <text x="${labelX}" y="${labelY}" font-size="${labelFontSize}" font-weight="700" fill="#0f172a" dominant-baseline="hanging">${page3PhotoLabelText}</text>`
      );
    }

    const svgCode = `
<svg xmlns="http://www.w3.org/2000/svg" width="100%" height="100%" viewBox="0 0 ${fullSlideW} ${fullSlideH}">
  <rect width="100%" height="100%" fill="${options.backgroundColor}" />
${svgParts.join('\n')}
</svg>`.trim();

    return { svgCode, fullSlideW, fullSlideH };
  }, [
    allPagesImages,
    calculateLayoutForAnyPage,
    clampPage1ImageCm,
    options.backgroundColor,
    options.containerWidth,
    getPageDimensions,
    reportFields,
    page2PhotoCategoryLabel,
    page3PhotoLabelText,
    showPage3,
    postPlacement,
    previewYOffsets
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
    const vet = (reportFields.attendingVet || '').trim();

    const subject = `治療報告書（${owner}様 ${pet}ちゃん）`;
    const body = `${hospital} 御中
${doctor} 先生

いつもお世話になっております。荻窪ツイン動物病院の${vet}です。
添付の通り、治療報告書をお送りします。ご確認よろしくお願いいたします。

---
荻窪ツイン動物病院
（住所などは今は不要。後で追加）`;

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

  const sendGmail = useCallback(async () => {
    if (isSendingGmail) return;

    const to = (reportFields.refHospitalEmail || '').trim();
    if (!to) {
      window.alert('紹介病院メールが未入力です。メールアドレスを入力してください。');
      return;
    }

    if (!window.confirm(`${to} にメールを送信します。よろしいですか？`)) return;

    const owner = (reportFields.ownerLastName || '').trim();
    const pet = (reportFields.petName || '').trim();
    const hospital = (reportFields.refHospitalName || reportFields.refHospital || '').trim();
    const doctor = (reportFields.refDoctor || '').trim();
    const vet = (reportFields.attendingVet || '').trim();

    const subject = `治療報告書（${owner}様 ${pet}ちゃん）`;

    const selectedCase = reportCases.find(c => c.case_id === selectedCaseId);
    const daysElapsed = getDaysElapsed(selectedCase?.registered_at);
    const isDelayed = daysElapsed != null && daysElapsed >= 7;

    const body = `${hospital} 御中
${doctor} 先生

いつもお世話になっております。荻窪ツイン動物病院の${vet}です。
${isDelayed ? DELAY_APOLOGY_TEXT + '\n' : ''}添付の通り、治療報告書をお送りします。ご確認よろしくお願いいたします。

---
荻窪ツイン動物病院
（住所などは今は不要。後で追加）`;

    setIsSendingGmail(true);
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

      const response = await fetch(`${SERVER_BASE_URL}/api/gmail/send`, {
        method: 'POST',
        body: formData,
      });

      const data = await response.json().catch(() => ({}));
      if (!response.ok || !data?.success) {
        let detail = data?.error || `HTTP ${response.status}`;
        if (response.status >= 500) {
          detail = `${detail}\nサーバ起動状態またはGmail OAuth設定を確認してください。`;
        }
        throw new Error(detail);
      }

      window.alert('メールを送信しました。');

      // 自動ステータス更新（失敗しても送信成功は維持）
      if (selectedCaseId) {
        try {
          const statusRes = await fetch(
            `http://localhost:8787/api/report-cases/${selectedCaseId}/status`,
            {
              method: 'PATCH',
              headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({ status: 'メール送信済み', mail_sent_at: new Date().toISOString() }),
            }
          );
          if (statusRes.ok) {
            const updated = await statusRes.json();
            setSelectedCaseStatus(updated.status || '');
            setReportCases(prev => prev.map(item => item.case_id === selectedCaseId ? updated : item));
          }
        } catch (e) {
          console.log('[auto status update] failed', e);
        }
      }
    } catch (error) {
      console.error('Gmail send failed:', error);
      const message = error instanceof Error ? error.message : String(error);
      if (message.includes('Failed to fetch') || message.includes('NetworkError')) {
        window.alert('Gmail送信に失敗しました: サーバに接続できませんでした。photo-report-server が起動しているか確認してください。');
      } else {
        window.alert(`Gmail送信に失敗しました: ${message}`);
      }
    } finally {
      setIsPrintMode(false);
      setIsSendingGmail(false);
    }
  }, [isSendingGmail, reportFields, outputPages, selectedCaseId]);

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

        let page2ImagesBottomYcm: number | undefined;
        let page3ImagesBottomYcm: number | undefined;

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

            if (pageNum === 2) {
              const imageBottomYcm = placed.yCm + placed.hCm;
              page2ImagesBottomYcm =
                page2ImagesBottomYcm === undefined
                  ? imageBottomYcm
                  : Math.max(page2ImagesBottomYcm, imageBottomYcm);
            }

            if (pageNum === 3) {
              const imageBottomYcm = placed.yCm + placed.hCm;
              page3ImagesBottomYcm =
                page3ImagesBottomYcm === undefined
                  ? imageBottomYcm
                  : Math.max(page3ImagesBottomYcm, imageBottomYcm);
            }
          });
        });

        addPptxText(slide, pageNum, reportFields, {
          showPage3,
          postPlacement,
          indentPostOnPage3: true,
          page2ImagesBottomYcm: pageNum === 2 ? page2ImagesBottomYcm : undefined,
          page3ImagesBottomYcm: pageNum === 3 ? page3ImagesBottomYcm : undefined,
        });

        if (pageNum === 2 && page2PhotoCategoryLabel) {
          slide.addText(page2PhotoCategoryLabel, {
            x: 1.04 / 2.54,
            y: Math.max(
              LAYOUT.PAGE2.LINES.SEP_TOP.y + 0.3,
              LAYOUT.PAGE2.TEXT.SECTION_HEADER_PROCEDURE.y - 0.6
            ) / 2.54,
            w: 8.0 / 2.54,
            h: 0.6 / 2.54,
            fontFace: 'Meiryo',
            fontSize: 13,
            bold: true,
            color: '0F172A',
          });
        }

        if (pageNum === 3 && page3PhotoLabelText && imagesData.length > 0) {
          slide.addText(page3PhotoLabelText, {
            x: 1.04 / 2.54,
            y: (LAYOUT.PAGE3.LINES.SEP_TOP.y + 0.3) / 2.54,
            w: 8.0 / 2.54,
            h: 0.6 / 2.54,
            fontFace: 'Meiryo',
            fontSize: 13,
            bold: true,
            color: '0F172A',
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

  const previewPage = useMemo<1 | 2 | 3>(() => {
    if (!showPage3 && currentPage === 3) return 2;
    return currentPage as 1 | 2 | 3;
  }, [showPage3, currentPage]);

  const activePreviewYOffsetGroup = useMemo(() => {
    const section = `PAGE${previewPage}` as 'PAGE1' | 'PAGE2' | 'PAGE3';
    const group = PREVIEW_Y_OFFSET_UI_GROUPS.find((g) => g.section === section) ?? PREVIEW_Y_OFFSET_UI_GROUPS[0];
    const items = group.items.filter((item) => {
      if (item.key === 'page2ThanksBody') return !showPage3;
      if (item.key === 'page3ThanksBody') return showPage3;
      return true;
    });
    return { ...group, items };
  }, [previewPage, showPage3]);

  const resetCurrentPagePreviewYOffsets = useCallback(() => {
    const targetKeys = Array.from(new Set(activePreviewYOffsetGroup.items.map((item) => item.key)));
    setPreviewYOffsets((prev) => {
      const next = { ...prev };
      targetKeys.forEach((key) => {
        next[key] = 0;
      });
      return next;
    });
  }, [activePreviewYOffsetGroup]);

  const nudgePreviewYOffset = useCallback((key: PreviewYOffsetKey, delta: number) => {
    setPreviewYOffsets((prev) => ({
      ...prev,
      [key]: Number((prev[key] + delta).toFixed(2)),
    }));
  }, []);

  const stopContinuousAdjust = useCallback(() => {
    if (holdTimeoutRef.current !== null) {
      window.clearTimeout(holdTimeoutRef.current);
      holdTimeoutRef.current = null;
    }
    if (holdIntervalRef.current !== null) {
      window.clearInterval(holdIntervalRef.current);
      holdIntervalRef.current = null;
    }
  }, []);

  const startContinuousAdjust = useCallback((adjustFn: () => void) => {
    stopContinuousAdjust();
    suppressNextClickRef.current = false;
    holdTimeoutRef.current = window.setTimeout(() => {
      suppressNextClickRef.current = true;
      holdIntervalRef.current = window.setInterval(adjustFn, 80);
    }, 350);
  }, [stopContinuousAdjust]);

  useEffect(() => {
    return () => {
      stopContinuousAdjust();
    };
  }, [stopContinuousAdjust]);

  const svgData = useMemo(() => calculateSvgDataForPage(previewPage, { applyPreviewOffsets: true }), [calculateSvgDataForPage, previewPage]);
  const svgPage1 = useMemo(() => calculateSvgDataForPage(1).svgCode, [calculateSvgDataForPage]);
  const svgPage2 = useMemo(() => calculateSvgDataForPage(2).svgCode, [calculateSvgDataForPage]);
  const svgPage3 = useMemo(() => calculateSvgDataForPage(3).svgCode, [calculateSvgDataForPage]);

  type DateFieldKey = 'reportDate' | 'firstVisitDate' | 'sedationDate' | 'anesthesiaDate';

  const [openDateField, setOpenDateField] = useState<DateFieldKey | null>(null);
  const [calendarMonth, setCalendarMonth] = useState<Date>(
    () => new Date(new Date().getFullYear(), new Date().getMonth(), 1)
  );

  const parseCalendarDate = useCallback((value: string): Date | null => {
    const text = String(value || '').trim();
    if (!text) return null;

    const jp = text.match(/^(\d{4})年(\d{1,2})月(\d{1,2})日$/);
    if (jp) {
      const y = Number(jp[1]);
      const m = Number(jp[2]);
      const d = Number(jp[3]);
      const date = new Date(y, m - 1, d);
      if (date.getFullYear() === y && date.getMonth() === m - 1 && date.getDate() === d) return date;
      return null;
    }

    const iso = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
    if (iso) {
      const y = Number(iso[1]);
      const m = Number(iso[2]);
      const d = Number(iso[3]);
      const date = new Date(y, m - 1, d);
      if (date.getFullYear() === y && date.getMonth() === m - 1 && date.getDate() === d) return date;
    }

    return null;
  }, []);

  const formatCalendarDate = useCallback((date: Date): string => {
    const y = date.getFullYear();
    const m = String(date.getMonth() + 1).padStart(2, '0');
    const d = String(date.getDate()).padStart(2, '0');
    return `${y}年${m}月${d}日`;
  }, []);

  const openCalendar = useCallback((field: DateFieldKey) => {
    const value = String(reportFields[field] || '');
    const parsed = parseCalendarDate(value);
    const base = parsed || new Date();
    setCalendarMonth(new Date(base.getFullYear(), base.getMonth(), 1));
    setOpenDateField(field);
  }, [parseCalendarDate, reportFields]);

  const closeCalendar = useCallback(() => {
    setOpenDateField(null);
  }, []);

  const selectCalendarDate = useCallback((day: number) => {
    if (!openDateField) return;
    const date = new Date(calendarMonth.getFullYear(), calendarMonth.getMonth(), day);
    const formatted = formatCalendarDate(date);
    setReportFields(prev => ({ ...prev, [openDateField]: formatted }));
    setOpenDateField(null);

    // 日付選択後の自動遷移
    if (openDateField === 'reportDate') {
      // 報告日→初診日カレンダー
      setTimeout(() => openCalendar('firstVisitDate'), 0);
    } else if (openDateField === 'firstVisitDate') {
      // 初診日→鎮静日カレンダー
      setTimeout(() => openCalendar('sedationDate'), 0);
    } else if (openDateField === 'sedationDate') {
      // 鎮静日→全身麻酔日カレンダー
      setTimeout(() => openCalendar('anesthesiaDate'), 0);
    } else if (openDateField === 'anesthesiaDate') {
      // 全身麻酔日→紹介病院名inputへフォーカス
      setTimeout(() => {
        if (refHospitalInputRef.current) {
          refHospitalInputRef.current.focus();
          refHospitalInputRef.current.scrollIntoView({ behavior: 'smooth', block: 'center', inline: 'nearest' });
        }
      }, 0);
    }
  }, [calendarMonth, formatCalendarDate, openDateField]);

  const clearCalendarDate = useCallback(() => {
    if (!openDateField) return;
    setReportFields(prev => ({ ...prev, [openDateField]: '' }));
    setOpenDateField(null);
  }, [openDateField]);

  const moveCalendarMonth = useCallback((offset: number) => {
    setCalendarMonth(prev => new Date(prev.getFullYear(), prev.getMonth() + offset, 1));
  }, []);

  const focusAndScroll = useCallback((el: HTMLElement | null) => {
    if (!el) return;
    el.focus();
    requestAnimationFrame(() => {
      const rect = el.getBoundingClientRect();
      const vh = window.innerHeight;
      // 要素が画面の安全範囲（上120px〜下から200px）内なら何もしない
      if (rect.top >= 120 && rect.bottom <= vh - 200) return;
      // 安全範囲外なら、要素が画面の約35%位置に来るようスクロール
      const targetY = rect.top + window.scrollY - vh * 0.35;
      window.scrollTo({ top: Math.max(0, targetY), behavior: 'smooth' });
    });
  }, []);

  const focusNextAfterRefHospitalSelection = useCallback((hospitalName: string) => {
    const normalizedName = normalizeHospitalKey(hospitalName);
    const selected = {
      email: normalizedName ? (normalizedRefHospitalEmails[normalizedName] ?? '') : '',
    };
    const nextEmail = selected.email ?? '';
    const nextTargetId = nextEmail.trim() ? 'ref-doctor-input' : 'ref-hospital-email';
    requestAnimationFrame(() => focusAndScroll(document.getElementById(nextTargetId)));
  }, [normalizedRefHospitalEmails, focusAndScroll]);

  const handleDropdownKeyDown = useCallback((e: React.KeyboardEvent, itemCount: number, onSelect: (index: number) => void, onClose: () => void, focusNext?: () => void) => {
    if (e.key === 'ArrowDown') {
      e.preventDefault();
      setDropdownHighlight(prev => (prev + 1) % itemCount);
    } else if (e.key === 'ArrowUp') {
      e.preventDefault();
      setDropdownHighlight(prev => (prev - 1 + itemCount) % itemCount);
    } else if (e.key === 'Enter') {
      e.preventDefault();
      if (dropdownHighlight >= 0 && dropdownHighlight < itemCount) {
        onSelect(dropdownHighlight);
        if (focusNext) {
          setTimeout(() => focusNext(), 0);
        }
      }
    } else if (e.key === 'Escape') {
      e.preventDefault();
      onClose();
    }
  }, [dropdownHighlight]);

  useEffect(() => {
    if (!openDateField) return;

    const closeWhenOutside = (event: Event) => {
      const target = event.target as HTMLElement | null;
      if (!target) return;
      if (target.closest('[data-date-field]')) return;
      setOpenDateField(null);
    };

    document.addEventListener('pointerdown', closeWhenOutside, true);
    document.addEventListener('focusin', closeWhenOutside, true);

    return () => {
      document.removeEventListener('pointerdown', closeWhenOutside, true);
      document.removeEventListener('focusin', closeWhenOutside, true);
    };
  }, [openDateField]);

  useEffect(() => {
    if (!isAttendingVetDropdownOpen) return;

    const closeWhenOutside = (event: Event) => {
      const target = event.target as Node | null;
      if (!target) return;
      if (attendingVetDropdownRef.current?.contains(target)) return;
      setIsAttendingVetDropdownOpen(false);
    };

    document.addEventListener('pointerdown', closeWhenOutside, true);
    document.addEventListener('focusin', closeWhenOutside, true);

    return () => {
      document.removeEventListener('pointerdown', closeWhenOutside, true);
      document.removeEventListener('focusin', closeWhenOutside, true);
    };
  }, [isAttendingVetDropdownOpen]);

  useEffect(() => {
    if (!isPage2PhotoCategoryDropdownOpen) return;

    const closeWhenOutside = (event: Event) => {
      const target = event.target as Node | null;
      if (!target) return;
      if (page2PhotoCategoryDropdownRef.current?.contains(target)) return;
      setIsPage2PhotoCategoryDropdownOpen(false);
    };

    document.addEventListener('pointerdown', closeWhenOutside, true);
    document.addEventListener('focusin', closeWhenOutside, true);

    return () => {
      document.removeEventListener('pointerdown', closeWhenOutside, true);
      document.removeEventListener('focusin', closeWhenOutside, true);
    };
  }, [isPage2PhotoCategoryDropdownOpen]);

  useEffect(() => {
    if (!isPostPlacementDropdownOpen) return;

    const closeWhenOutside = (event: Event) => {
      const target = event.target as Node | null;
      if (!target) return;
      if (postPlacementDropdownRef.current?.contains(target)) return;
      setIsPostPlacementDropdownOpen(false);
    };

    document.addEventListener('pointerdown', closeWhenOutside, true);
    document.addEventListener('focusin', closeWhenOutside, true);

    return () => {
      document.removeEventListener('pointerdown', closeWhenOutside, true);
      document.removeEventListener('focusin', closeWhenOutside, true);
    };
  }, [isPostPlacementDropdownOpen]);


  useEffect(() => {
    if (!isThankYouTextTypeDropdownOpen) return;

    const closeWhenOutside = (event: Event) => {
      const target = event.target as Node | null;
      if (!target) return;
      if (thankYouTextTypeDropdownRef.current?.contains(target)) return;
      setIsThankYouTextTypeDropdownOpen(false);
    };

    document.addEventListener('pointerdown', closeWhenOutside, true);
    document.addEventListener('focusin', closeWhenOutside, true);

    return () => {
      document.removeEventListener('pointerdown', closeWhenOutside, true);
      document.removeEventListener('focusin', closeWhenOutside, true);
    };
  }, [isThankYouTextTypeDropdownOpen]);

  const calendarFirstDay = new Date(calendarMonth.getFullYear(), calendarMonth.getMonth(), 1).getDay();
  const calendarDaysInMonth = new Date(calendarMonth.getFullYear(), calendarMonth.getMonth() + 1, 0).getDate();
  const calendarCells: Array<number | null> = [
    ...Array.from({ length: calendarFirstDay }, () => null),
    ...Array.from({ length: calendarDaysInMonth }, (_, idx) => idx + 1),
  ];

  const selectedCalendarDate = openDateField
    ? parseCalendarDate(String(reportFields[openDateField] || ''))
    : null;

  const filledDateCount = useMemo(() => {
    const values = [reportFields.firstVisitDate, reportFields.sedationDate, reportFields.anesthesiaDate];
    return values.filter(v => String(v || '').trim() !== '').length;
  }, [reportFields.anesthesiaDate, reportFields.firstVisitDate, reportFields.sedationDate]);

  const dateDividerOffsetClass = useMemo(() => {
    if (filledDateCount <= 0) return 'mt-1';
    if (filledDateCount === 1) return 'mt-2';
    if (filledDateCount === 2) return 'mt-3';
    return 'mt-4';
  }, [filledDateCount]);

  const getEmptyFieldToneClass = useCallback((value: unknown) => {
    const isEmpty = String(value ?? '').trim() === '';
    return isEmpty
      ? 'bg-emerald-50/50 border-emerald-100 text-slate-800'
      : 'bg-white border-slate-200 text-slate-900';
  }, []);

  const handleEnterFocusNextInput = useCallback((e: React.KeyboardEvent<HTMLElement>) => {
    if (e.key !== 'Enter' || (e.nativeEvent as KeyboardEvent).isComposing) return;

    const target = e.target as HTMLElement | null;
    if (!target) return;
    const tag = target.tagName.toLowerCase();
    if (tag !== 'input' && tag !== 'select') return;

    if (tag === 'input') {
      const inputEl = target as HTMLInputElement;
      const inputType = (inputEl.type || 'text').toLowerCase();
      if (['checkbox', 'radio', 'file', 'button', 'submit', 'reset'].includes(inputType)) return;

      if (inputEl.id === 'ref-doctor-input') {
        e.preventDefault();
        focusAndScroll(document.getElementById('chief-complaint-input'));
        return;
      }

      if (inputEl.id === 'chief-complaint-input') {
        e.preventDefault();
        const initialTextArea = document.getElementById('initial-textarea') as HTMLTextAreaElement | null;
        focusAndScroll(initialTextArea);
        return;
      }
    }

    e.preventDefault();

    const container = e.currentTarget as HTMLElement;
    const focusables = Array.from(
      container.querySelectorAll<HTMLElement>(
        'input:not([type="hidden"]):not([type="file"]):not([type="checkbox"]):not([type="radio"]):not([disabled]), select:not([disabled])'
      )
    );
    const idx = focusables.indexOf(target);
    if (idx >= 0 && idx < focusables.length - 1) {
      focusAndScroll(focusables[idx + 1]);
    }
  }, [focusAndScroll]);

  return (
    <div className="min-h-screen bg-slate-50 pb-4 font-sans">
      <nav className="bg-white border-b border-slate-200 py-4 px-6 sticky top-0 z-50 shadow-sm">
        <div className="max-w-7xl mx-auto">
          <h1 className="text-xl font-black text-slate-900 tracking-tight leading-none">歯科治療報告書</h1>
        </div>
      </nav>

      <main className="max-w-7xl mx-auto px-6 mt-10 grid grid-cols-1 lg:grid-cols-12 gap-10">
        {/* 新規患者登録 */}
        <div className="lg:col-span-12 mb-0">
          <div className="bg-white border border-slate-200 rounded-2xl shadow-sm p-4">
            <div className="font-bold text-sm text-slate-700 mb-2">新規患者登録</div>
            <div className="flex items-center gap-2 text-sm">
              <input value={newCaseOwnerLastName} onChange={(e) => setNewCaseOwnerLastName(e.target.value)} placeholder="飼い主姓" className="border border-slate-200 rounded-lg px-2 py-1.5" />
              <input value={newCasePetName} onChange={(e) => setNewCasePetName(e.target.value)} placeholder="ペット名" className="border border-slate-200 rounded-lg px-2 py-1.5" />
              <select value={newCaseAttendingVet} onChange={(e) => setNewCaseAttendingVet(e.target.value)} className="border border-slate-200 rounded-lg px-2 py-1.5 bg-white">
                {ATTENDING_VET_OPTIONS.map(name => (
                  <option key={name} value={name}>{name}</option>
                ))}
              </select>
              <button type="button" onClick={handleCreateCase} disabled={isCreatingCase} className="px-3 py-1.5 rounded-lg bg-green-600 text-white font-semibold disabled:opacity-50 hover:bg-green-700 transition-colors whitespace-nowrap">
                {isCreatingCase ? '登録中...' : '新規患者を登録'}
              </button>
            </div>
          </div>
        </div>

        {/* 患者一覧 */}
        {reportCases.length > 0 && (
          <div className="lg:col-span-12 mb-0">
            <div className="bg-white border border-slate-200 rounded-2xl shadow-sm p-4">
              <div className="font-bold text-sm text-slate-700 mb-2">患者一覧（{reportCases.length}件）</div>
              <div className="max-h-40 overflow-y-auto text-sm">
                {reportCases.map((c: any) => (
                  <div
                    key={c.case_id}
                    onClick={() => handleSelectCase(c)}
                    className={`cursor-pointer p-2 border-b border-slate-100 hover:bg-gray-100 ${
                      selectedCaseId === c.case_id ? 'bg-blue-100' : ''
                    }`}
                  >
                    {(() => { const days = getDaysElapsed(c.registered_at); const delay = getDelayLabel(days); return (<>
                    <span className="text-slate-400 mr-2">{formatCaseDisplayId(c.case_id)}</span>
                    <span className="text-slate-400 mr-2">{formatDateShort(c.registered_at)}</span>
                    <span className="text-slate-400 mr-3">{days ?? '-'}日</span>
                    <span className={`mr-3 text-xs px-1.5 py-0.5 rounded-full font-semibold ${delay.className}`}>{delay.text}</span>
                    {c.owner_last_name} / {c.pet_name} / {c.attending_vet || ''} / {c.referring_hospital}
                    <span className={`ml-2 text-xs px-1.5 py-0.5 rounded-full font-semibold ${
                      c.status === 'メール送信済み' ? 'bg-green-100 text-green-700' :
                      c.status === '印刷郵送済み' ? 'bg-orange-100 text-orange-700' :
                      c.status === '報告書作成途中' ? 'bg-blue-100 text-blue-700' :
                      'bg-slate-100 text-slate-500'
                    }`}>{c.status}</span>
                    </>); })()}
                  </div>
                ))}
              </div>
            </div>
          </div>
        )}

        {/* 選択中の患者情報 */}
        {selectedCaseId && (
          <div className="lg:col-span-12 mb-0 flex items-center gap-3 text-sm text-slate-600">
            {(() => { const sc = reportCases.find(r => r.case_id === selectedCaseId); const days = getDaysElapsed(sc?.registered_at); const delay = getDelayLabel(days); return (
            <span>選択中: <strong>{formatCaseDisplayId(selectedCaseId ?? '')}</strong> / {selectedCaseStatus || '-'} / {formatDateShort(sc?.registered_at || '')} / {days ?? '-'}日 <span className={`text-xs px-1.5 py-0.5 rounded-full font-semibold ${delay.className}`}>{delay.text}</span></span>
            ); })()}
            <button type="button" onClick={handleMarkMailSent} disabled={!selectedCaseId}
              className="px-3 py-1 rounded-lg bg-green-600 text-white text-xs font-semibold disabled:opacity-50 hover:bg-green-700 transition-colors">
              メール送信済みにする
            </button>
            <button type="button" onClick={handleMarkPostalSent} disabled={!selectedCaseId}
              className="px-3 py-1 rounded-lg bg-orange-500 text-white text-xs font-semibold disabled:opacity-50 hover:bg-orange-600 transition-colors">
              印刷郵送済みにする
            </button>
          </div>
        )}

        {/* 報告書データ入力フォーム */}
        <div className="lg:col-span-12 bg-white border border-slate-200 rounded-2xl shadow-sm p-4 md:p-5 space-y-4" onKeyDown={handleEnterFocusNextInput}>
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
      {openDateField === 'reportDate' && (
        <div className="absolute right-0 top-full mt-2 z-40 w-72 rounded-2xl border border-slate-200 bg-white p-3 shadow-xl">
          <div className="flex items-center justify-between mb-2">
            <button
              type="button"
              className="px-3 py-2 rounded transition-colors cursor-pointer text-gray-700 hover:text-orange-600"
              onClick={() => moveCalendarMonth(-1)}
            >
              <span
                style={{ fontSize: '22px', display: 'inline-block', lineHeight: '1' }}
              >{'<'}</span>
            </button>
            <span className="text-lg font-bold text-indigo-700">
              {calendarMonth.getFullYear()}年{calendarMonth.getMonth() + 1}月
            </span>
            <button
              type="button"
              className="px-3 py-2 rounded transition-colors cursor-pointer text-gray-700 hover:text-orange-600"
              onClick={() => moveCalendarMonth(1)}
            >
              <span
                style={{ fontSize: '22px', display: 'inline-block', lineHeight: '1' }}
              >{'>'}</span>
            </button>
          </div>
          <div className="grid grid-cols-7 gap-1">
            {calendarCells.map((day, idx) => {
              if (!day) return <span key={`empty-${idx}`} className="h-8" />;
              const isSelected =
                !!selectedCalendarDate &&
                selectedCalendarDate.getFullYear() === calendarMonth.getFullYear() &&
                selectedCalendarDate.getMonth() === calendarMonth.getMonth() &&
                selectedCalendarDate.getDate() === day;
              return (
                <button
                  key={day}
                  type="button"
                  onClick={() => selectCalendarDate(day)}
                  className={`h-8 rounded-lg text-base font-medium transition-colors ${
                    isSelected
                      ? 'bg-orange-500 text-white'
                      : 'text-slate-700 hover:bg-slate-100'
                  }`}
                >
                  {day}
                </button>
              );
            })}
          </div>
          <div className="flex justify-between mt-2">
            <button type="button" className="text-xs bg-gray-100 rounded px-2 py-1 text-gray-700 hover:bg-gray-200 transition-colors mr-2" onClick={clearCalendarDate}>クリア</button>
            <button type="button" className="text-xs bg-gray-100 rounded px-2 py-1 text-gray-700 hover:bg-gray-200 transition-colors" onClick={closeCalendar}>閉じる</button>
          </div>
        </div>
      )}
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

          <div className="space-y-4">
            {/* 基本情報グリッド */}
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4 border border-sky-200 bg-sky-50/60 rounded-xl p-4">
              <div className="sm:col-span-2 lg:col-span-3 text-base font-semibold text-slate-700">基本情報</div>
              <div className="sm:col-span-2 lg:col-span-3 rounded-xl border border-slate-200 bg-transparent p-3 md:p-4">
                <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4">
                  {/* 飼い主姓／ペット名／担当獣医師ブロック（先頭へ移動） */}
                  <div className="space-y-1">
                    <label className="text-xs font-semibold text-slate-700 uppercase tracking-widest">飼い主姓</label>
                    <input className={`w-full h-11 border rounded-xl px-3 py-2 text-base focus:ring-2 focus:ring-orange-500 outline-none transition-all ${getEmptyFieldToneClass(reportFields.ownerLastName)} bg-white`}
                      placeholder="山田"
                      value={reportFields.ownerLastName}
                      onChange={e => setReportFields(v => ({ ...v, ownerLastName: e.target.value }))}
                    />
                  </div>
                  <div className="space-y-1">
                    <label className="text-xs font-semibold text-slate-700 uppercase tracking-widest">ペット名</label>
                    <input className={`w-full h-11 border rounded-xl px-3 py-2 text-base focus:ring-2 focus:ring-orange-500 outline-none transition-all ${getEmptyFieldToneClass(reportFields.petName)} bg-white`}
                      placeholder="タロウ"
                      value={reportFields.petName}
                      onChange={e => setReportFields(v => ({ ...v, petName: e.target.value }))}
                      onKeyDown={e => {
                        if (e.key === 'Tab' && !e.shiftKey) {
                          shouldOpenAttendingVetOnFocusRef.current = true;
                        }
                        if (e.key === 'Enter') {
                          e.preventDefault();
                          shouldOpenAttendingVetOnFocusRef.current = true;
                          requestAnimationFrame(() => focusAndScroll(document.getElementById('attending-vet-btn')));
                        }
                      }}
                    />
                  </div>
                  {/* 担当獣医師（新規：プルダウン） */}
                  <div className="space-y-1">
                    <label className="text-xs font-semibold text-slate-700 uppercase tracking-widest">担当獣医師</label>
                    <div className="relative" ref={attendingVetDropdownRef}>
                      <button
                        id="attending-vet-btn"
                        type="button"
                        className={`w-full h-11 border rounded-xl px-3 py-2 text-base text-left focus:ring-2 focus:ring-orange-500 outline-none transition-all flex items-center ${getEmptyFieldToneClass(reportFields.attendingVet)} bg-white`}
                        aria-haspopup="listbox"
                        aria-expanded={isAttendingVetDropdownOpen}
                        onFocus={() => {
                          if (!shouldOpenAttendingVetOnFocusRef.current) return;
                          shouldOpenAttendingVetOnFocusRef.current = false;
                          setIsAttendingVetDropdownOpen(true);
                          setDropdownHighlight(0);
                        }}
                        onClick={() => { setIsAttendingVetDropdownOpen(v => !v); setDropdownHighlight(-1); }}
                        onKeyDown={isAttendingVetDropdownOpen ? (e) => {
                          const items = ATTENDING_VET_OPTIONS;
                          handleDropdownKeyDown(e, items.length, (idx) => {
                            const name = items[idx];
                            setReportFields(v => ({ ...v, attendingVet: name }));
                            setIsAttendingVetDropdownOpen(false);
                            (document.activeElement as HTMLElement | null)?.blur();
                            requestAnimationFrame(() => openCalendar('firstVisitDate'));
                          }, () => setIsAttendingVetDropdownOpen(false));
                        } : undefined}
                      >
                        <span className={reportFields.attendingVet ? 'text-slate-900' : 'text-slate-500'}>
                          {reportFields.attendingVet || '選択してください'}
                        </span>
                      </button>
                      {isAttendingVetDropdownOpen && (
                        <ul
                          role="listbox"
                          className="absolute z-40 mt-1 max-h-56 w-full overflow-auto rounded-xl border border-slate-200 bg-white py-1 shadow-lg"
                        >
                          {ATTENDING_VET_OPTIONS.map((name, idx) => {
                            const label = name;
                            const isSelected = reportFields.attendingVet === name;
                            const isHighlighted = dropdownHighlight === idx;
                            return (
                              <li key={label}>
                                <button
                                  type="button"
                                  className={`w-full px-3 py-2 text-left text-base transition-colors ${isHighlighted ? 'bg-orange-100 text-orange-800' : isSelected ? 'bg-orange-50 text-orange-700' : 'text-slate-800 hover:bg-slate-50'}`}
                                  onClick={() => {
                                    setReportFields(v => ({ ...v, attendingVet: name }));
                                    setIsAttendingVetDropdownOpen(false);
                                    if (name) {
                                      requestAnimationFrame(() => openCalendar('firstVisitDate'));
                                    }
                                  }}
                                >
                                  {label}
                                </button>
                              </li>
                            );
                          })}
                        </ul>
                      )}
                    </div>
                  </div>
                  <div className="space-y-1">
                    <label className="text-xs font-semibold text-slate-700 uppercase tracking-widest">初診日</label>
                    <div className="relative" data-date-field="firstVisitDate">
                      <input
                        className={`w-full h-11 border rounded-xl px-3 py-2 text-base focus:ring-2 focus:ring-orange-500 outline-none transition-all cursor-pointer ${getEmptyFieldToneClass(reportFields.firstVisitDate)} bg-white`}
                        placeholder="202X年XX月XX日"
                        value={reportFields.firstVisitDate}
                        readOnly
                        onClick={() => openCalendar("firstVisitDate")}
                        data-date-field
                      />
                      {openDateField === "firstVisitDate" && (
                        <div className="absolute z-20 mt-1 left-0">
                          <div className="bg-white border rounded shadow-lg p-2 w-64">
                            <div className="flex items-center justify-between mb-2">
                              <button type="button" className="font-bold text-gray-700 hover:text-orange-600 px-3 py-2 rounded transition-colors cursor-pointer" onClick={() => moveCalendarMonth(-1)}>
                                <span className="text-xl leading-none">{'<'}</span>
                              </button>
                              <span className="text-lg font-bold text-indigo-700">
                                {calendarMonth.getFullYear()}年{calendarMonth.getMonth() + 1}月
                              </span>
                              <button type="button" className="font-bold text-gray-700 hover:text-orange-600 px-3 py-2 rounded transition-colors cursor-pointer" onClick={() => moveCalendarMonth(1)}>
                                <span className="text-xl leading-none">{'>'}</span>
                              </button>
                            </div>
                            <div className="grid grid-cols-7 gap-1 text-xs mb-1">
                              {calendarCells.map((cell, idx) =>
                                cell === null ? (
                                  <div key={idx} />
                                ) : (
                                  <button
                                    key={idx}
                                    type="button"
                                    className="w-7 h-7 rounded hover:bg-blue-100 text-center"
                                    onClick={() => selectCalendarDate(cell)}
                                  >
                                    {cell}
                                  </button>
                                )
                              )}
                            </div>
                            <div className="flex justify-between mt-2">
                              <button type="button" className="text-xs bg-gray-100 rounded px-2 py-1 text-gray-700 hover:bg-gray-200 transition-colors mr-2" onClick={clearCalendarDate}>クリア</button>
                              <button type="button" className="text-xs bg-gray-100 rounded px-2 py-1 text-gray-700 hover:bg-gray-200 transition-colors" onClick={closeCalendar}>閉じる</button>
                            </div>
                          </div>
                        </div>
                      )}
                    </div>
                  </div>
                  <div className="space-y-1">
                    <label className="text-xs font-semibold text-slate-700 uppercase tracking-widest">鎮静日</label>
                    <div className="relative" data-date-field="sedationDate">
                      <input
                        className={`w-full h-11 border rounded-xl px-3 py-2 text-base focus:ring-2 focus:ring-orange-500 outline-none transition-all cursor-pointer ${getEmptyFieldToneClass(reportFields.sedationDate)} bg-white`}
                        placeholder="202X年XX月XX日"
                        value={reportFields.sedationDate}
                        readOnly
                        onClick={() => openCalendar("sedationDate")}
                        data-date-field
                      />
                      {openDateField === "sedationDate" && (
                        <div className="absolute z-20 mt-1 left-0">
                          <div className="bg-white border rounded shadow-lg p-2 w-64">
                            <div className="flex items-center justify-between mb-2">
                              <button type="button" className="font-bold text-gray-700 hover:text-orange-600 px-3 py-2 rounded transition-colors cursor-pointer" onClick={() => moveCalendarMonth(-1)}>
                                <span className="text-xl leading-none">{'<'}</span>
                              </button>
                              <span className="text-lg font-bold text-indigo-700">
                                {calendarMonth.getFullYear()}年{calendarMonth.getMonth() + 1}月
                              </span>
                              <button type="button" className="font-bold text-gray-700 hover:text-orange-600 px-3 py-2 rounded transition-colors cursor-pointer" onClick={() => moveCalendarMonth(1)}>
                                <span className="text-xl leading-none">{'>'}</span>
                              </button>
                            </div>
                            <div className="grid grid-cols-7 gap-1 text-xs mb-1">
                              {calendarCells.map((cell, idx) =>
                                cell === null ? (
                                  <div key={idx} />
                                ) : (
                                  <button
                                    key={idx}
                                    type="button"
                                    className="w-7 h-7 rounded hover:bg-blue-100 text-center"
                                    onClick={() => selectCalendarDate(cell)}
                                  >
                                    {cell}
                                  </button>
                                )
                              )}
                            </div>
                            <div className="flex justify-between mt-2">
                              <button type="button" className="text-xs bg-gray-100 rounded px-2 py-1 text-gray-700 hover:bg-gray-200 transition-colors mr-2" onClick={clearCalendarDate}>クリア</button>
                              <button type="button" className="text-xs bg-gray-100 rounded px-2 py-1 text-gray-700 hover:bg-gray-200 transition-colors" onClick={closeCalendar}>閉じる</button>
                            </div>
                          </div>
                        </div>
                      )}
                    </div>
                  </div>
                  <div className="space-y-1">
                    <label className="text-xs font-semibold text-slate-700 uppercase tracking-widest">全身麻酔日</label>
                    <div className="relative" data-date-field="anesthesiaDate">
                      <input
                        className={`w-full h-11 border rounded-xl px-3 py-2 text-base focus:ring-2 focus:ring-orange-500 outline-none transition-all cursor-pointer ${getEmptyFieldToneClass(reportFields.anesthesiaDate)} bg-white`}
                        placeholder="202X年XX月XX日"
                        value={reportFields.anesthesiaDate}
                        readOnly
                        onClick={() => openCalendar("anesthesiaDate")}
                        data-date-field
                      />
                      {openDateField === "anesthesiaDate" && (
                        <div className="absolute z-20 mt-1 left-0">
                          <div className="bg-white border rounded shadow-lg p-2 w-64">
                            <div className="flex items-center justify-between mb-2">
                              <button type="button" className="font-bold text-gray-700 hover:text-orange-600 px-3 py-2 rounded transition-colors cursor-pointer" onClick={() => moveCalendarMonth(-1)}>
                                <span className="text-xl leading-none">{'<'}</span>
                              </button>
                              <span className="text-lg font-bold text-indigo-700">
                                {calendarMonth.getFullYear()}年{calendarMonth.getMonth() + 1}月
                              </span>
                              <button type="button" className="font-bold text-gray-700 hover:text-orange-600 px-3 py-2 rounded transition-colors cursor-pointer" onClick={() => moveCalendarMonth(1)}>
                                <span className="text-xl leading-none">{'>'}</span>
                              </button>
                            </div>
                            <div className="grid grid-cols-7 gap-1 text-xs mb-1">
                              {calendarCells.map((cell, idx) =>
                                cell === null ? (
                                  <div key={idx} />
                                ) : (
                                  <button
                                    key={idx}
                                    type="button"
                                    className="w-7 h-7 rounded hover:bg-blue-100 text-center"
                                    onClick={() => selectCalendarDate(cell)}
                                  >
                                    {cell}
                                  </button>
                                )
                              )}
                            </div>
                            <div className="flex justify-between mt-2">
                              <button type="button" className="text-xs bg-gray-100 rounded px-2 py-1 text-gray-700 hover:bg-gray-200 transition-colors mr-2" onClick={clearCalendarDate}>クリア</button>
                              <button type="button" className="text-xs bg-gray-100 rounded px-2 py-1 text-gray-700 hover:bg-gray-200 transition-colors" onClick={closeCalendar}>閉じる</button>
                            </div>
                          </div>
                        </div>
                      )}
                    </div>
                  </div>
                  <div className="space-y-1">
                    <label className="text-xs font-semibold text-slate-700 uppercase tracking-widest">
                      紹介病院名
                    </label>

                    <input
                      id="ref-hospital-input"
                      ref={refHospitalInputRef}
                      className={`w-full max-w-[520px] h-11 px-3 py-2 rounded-xl border text-base ${getEmptyFieldToneClass(refHospitalInput)} bg-white`}
                      placeholder="例：中川動物病院"
                      value={refHospitalInput}
                      onChange={(e) => {
                        const v = e.target.value;
                        applyRefHospitalSelection(v);
                        const norm = normalizeHospitalKey(v);
                        if (norm && normalizedRefHospitalNames.has(norm)) {
                          focusNextAfterRefHospitalSelection(v);
                        }
                      }}
                      onBlur={(e) => applyRefHospitalSelection(e.target.value)}
                      onKeyDown={(e) => {
                        if (e.key === "Enter") {
                          e.preventDefault();
                          const v = e.currentTarget.value;
                          const norm = normalizeHospitalKey(v);
                          const isKnown = norm ? normalizedRefHospitalNames.has(norm) : false;
                          applyRefHospitalSelection(v);
                          if (isKnown) {
                            focusNextAfterRefHospitalSelection(v);
                          } else {
                            requestAnimationFrame(() => focusAndScroll(document.getElementById('ref-hospital-email')));
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

                    {shouldShowRegisterRefHospitalButton && (
                      <button
                        type="button"
                        onClick={() => handleAddRefHospital(refHospitalInput, reportFields.refHospitalEmail)}
                        className="text-sm px-3 py-1.5 rounded-lg border border-rose-200 bg-rose-50 text-rose-700 hover:bg-rose-100 transition-colors"
                        disabled={isSavingRefHospital}
                      >
                        {isSavingRefHospital ? '登録中...' : 'この病院を登録'}
                      </button>
                    )}

                    {showHospitalSavedMessage && (
                      <div className="text-sm text-emerald-600 mt-1">登録しました</div>
                    )}

                    {/* 保存エラー（任意表示） */}
                    {refHospitalError && (
                      <div className="text-sm text-red-600">{refHospitalError}</div>
                    )}
                  </div>
                  <div className="space-y-1">
                    <label className="text-xs font-semibold text-slate-700 uppercase tracking-widest">紹介病院メールアドレス</label>
                    <input id="ref-hospital-email" className={`w-full h-11 border rounded-xl px-3 py-2 text-base focus:ring-2 focus:ring-orange-500 outline-none transition-all ${getEmptyFieldToneClass(reportFields.refHospitalEmail)} bg-white`}
                      placeholder="example@gmail.com"
                      value={reportFields.refHospitalEmail}
                      onChange={e => setReportFields(v => ({ ...v, refHospitalEmail: e.target.value }))}
                    />
                    {!refHospitalError &&
                      (reportFields.refHospitalName || reportFields.refHospital).trim() !== '' &&
                      reportFields.refHospitalEmail.trim() === '' && (
                        <div className="text-sm text-slate-500">
                          この紹介病院はメール未登録です。必要なら手入力してください。
                        </div>
                      )}
                  </div>
                  <div className="space-y-1">
                    <label className="text-xs font-semibold text-slate-700 uppercase tracking-widest">先生名</label>
                    <div className="relative">
                      <input className={`w-full h-11 border rounded-xl px-3 pr-12 py-2 text-base focus:ring-2 focus:ring-orange-500 outline-none transition-all ${getEmptyFieldToneClass(reportFields.refDoctor)} bg-white`}
                        id="ref-doctor-input"
                        placeholder="△△"
                        value={reportFields.refDoctor}
                        onChange={e => setReportFields(v => ({ ...v, refDoctor: e.target.value }))}
                      />
                      <span className="pointer-events-none absolute inset-y-0 right-3 flex items-center text-base text-slate-500">先生</span>
                    </div>
                  </div>
                </div>
              </div>


            </div>


            {/* PAGE3設定 */}
            <div className="grid grid-cols-1 sm:grid-cols-3 gap-4 border border-amber-200 bg-amber-50/60 rounded-xl p-4">
              <div className="sm:col-span-3 text-base font-semibold text-slate-700">ページ設定</div>
              <label className="flex items-center gap-2 text-base text-slate-700 font-semibold">
                <input
                  type="checkbox"
                  className="bg-white"
                  checked={showPage3}
                  onChange={e => setShowPage3(e.target.checked)}
                />
                PAGE3を追加する
              </label>


              {showPage3 && (
                <div className="space-y-1">
                  <label className="text-xs font-semibold text-slate-700 uppercase tracking-widest">術後経過の配置</label>
                  <div className="relative" ref={postPlacementDropdownRef}>
                    <button
                      type="button"
                      className="w-full h-11 border border-slate-200 rounded-xl px-3 py-2 text-base text-left focus:ring-2 focus:ring-orange-500 outline-none transition-all bg-white text-slate-900 flex items-center"
                      aria-haspopup="listbox"
                      aria-expanded={isPostPlacementDropdownOpen}
                      onFocus={() => {
                        if (!shouldOpenPostPlacementOnFocusRef.current) return;
                        shouldOpenPostPlacementOnFocusRef.current = false;
                        setIsPostPlacementDropdownOpen(true);
                        setDropdownHighlight(0);
                      }}
                      onClick={() => { setIsPostPlacementDropdownOpen(v => !v); setDropdownHighlight(-1); }}
                      onKeyDown={isPostPlacementDropdownOpen ? (e) => {
                        const items = [{ value: 'page2' }, { value: 'page3' }];
                        handleDropdownKeyDown(e, items.length, (idx) => {
                          setPostPlacement(items[idx].value as 'page2' | 'page3');
                          setIsPostPlacementDropdownOpen(false);
                        }, () => setIsPostPlacementDropdownOpen(false));
                      } : undefined}
                    >
                      {postPlacement === 'page3' ? 'PAGE3に移す' : 'PAGE2に置く'}
                    </button>

                    {isPostPlacementDropdownOpen && (
                      <ul
                        role="listbox"
                        className="absolute z-40 mt-1 max-h-56 w-full overflow-auto rounded-xl border border-slate-200 bg-white py-1 shadow-lg"
                      >
                        {[
                          { value: 'page2', label: 'PAGE2に置く' },
                          { value: 'page3', label: 'PAGE3に移す' },
                        ].map((item, idx) => {
                          const isSelected = postPlacement === item.value;
                          const isHighlighted = dropdownHighlight === idx;
                          return (
                            <li key={item.value}>
                              <button
                                type="button"
                                className={`w-full px-3 py-2 text-left text-base transition-colors ${isHighlighted ? 'bg-orange-100 text-orange-800' : isSelected ? 'bg-orange-50 text-orange-700' : 'text-slate-800 hover:bg-slate-50'}`}
                                onClick={() => {
                                  setPostPlacement(item.value as 'page2' | 'page3');
                                  setIsPostPlacementDropdownOpen(false);
                                }}
                              >
                                {item.label}
                              </button>
                            </li>
                          );
                        })}
                      </ul>
                    )}
                  </div>
                </div>
              )}
            </div>

            {/* 自由記載エリア */}
            <div className="grid grid-cols-1 gap-4 border border-rose-200 bg-rose-50/40 rounded-xl p-4">
              <div className="flex items-center justify-between mb-4 pb-2 border-b border-slate-200">
                <div className="text-lg font-semibold text-slate-800 tracking-tight">報告内容</div>
              </div>

              {/* 主訴（新規：テキスト入力） */}
              <div className="space-y-1">
                <label className="text-xs font-semibold text-slate-700 uppercase tracking-widest">主訴</label>
                <input className={`w-full h-11 border rounded-xl px-3 py-2 text-base focus:ring-2 focus:ring-orange-500 outline-none transition-all ${getEmptyFieldToneClass(reportFields.chiefComplaint)} bg-white`}
                  style={{ wordBreak: 'break-word', overflowWrap: 'break-word' }}
                  id="chief-complaint-input"
                  placeholder="主な症状や主訴"
                  value={reportFields.chiefComplaint}
                  onChange={e => setReportFields(v => ({ ...v, chiefComplaint: e.target.value }))}
                />
              </div>
              <div className="space-y-1">
                <label className="text-xs font-semibold text-slate-700 uppercase tracking-widest">【初診時】</label>
                <textarea className="w-full border border-slate-200 rounded-xl px-3 py-2 text-base min-h-[80px] focus:ring-2 focus:ring-orange-500 outline-none transition-all bg-white"
                  style={{ whiteSpace: 'pre-wrap', wordBreak: 'break-word', overflowWrap: 'break-word', maxHeight: '18.26cm', overflowY: 'auto' }}
                  id="initial-textarea"
                  placeholder="初診時の所見など..."
                  value={reportFields.initialText}
                  onKeyDown={(e) => {
                    if (e.key === 'Tab' && !e.shiftKey) {
                      shouldOpenPage2PhotoCategoryOnFocusRef.current = true;
                    }
                  }}
                  onChange={e => setReportFields(v => ({ ...v, initialText: e.target.value }))}
                />
              </div>
              <div className="space-y-1">
                <label className="text-xs font-semibold text-slate-700 uppercase tracking-widest">PAGE2写真ラベル</label>
                <div className="relative" ref={page2PhotoCategoryDropdownRef}>
                  <button
                    type="button"
                    className="w-full h-11 border border-slate-200 rounded-xl px-3 py-2 text-base text-left focus:ring-2 focus:ring-orange-500 outline-none transition-all bg-white text-slate-900 flex items-center"
                    aria-haspopup="listbox"
                    aria-expanded={isPage2PhotoCategoryDropdownOpen}
                    onFocus={() => {
                      if (!shouldOpenPage2PhotoCategoryOnFocusRef.current) return;
                      shouldOpenPage2PhotoCategoryOnFocusRef.current = false;
                      setIsPage2PhotoCategoryDropdownOpen(true);
                      setDropdownHighlight(0);
                    }}
                    onClick={() => { setIsPage2PhotoCategoryDropdownOpen(v => !v); setDropdownHighlight(-1); }}
                    onKeyDown={isPage2PhotoCategoryDropdownOpen ? (e) => {
                      const items = [{ value: '' }, { value: 'treatment-after' }, { value: 'inspection' }];
                      handleDropdownKeyDown(
                        e,
                        items.length,
                        (idx) => {
                          setReportFields(v => ({ ...v, page2PhotoCategory: items[idx].value }));
                          setIsPage2PhotoCategoryDropdownOpen(false);
                        },
                        () => setIsPage2PhotoCategoryDropdownOpen(false),
                        () => focusAndScroll(page2ProcedureTextareaRef.current ?? null)
                      );
                    } : undefined}
                  >
                    <span className={reportFields.page2PhotoCategory ? 'text-slate-900' : 'text-slate-500'}>
                      {reportFields.page2PhotoCategory === 'treatment-after'
                        ? '治療時・治療後写真'
                        : reportFields.page2PhotoCategory === 'inspection'
                          ? '検査時写真'
                          : '空欄'}
                    </span>
                  </button>

                  {isPage2PhotoCategoryDropdownOpen && (
                    <ul
                      role="listbox"
                      className="absolute z-40 mt-1 max-h-56 w-full overflow-auto rounded-xl border border-slate-200 bg-white py-1 shadow-lg"
                    >
                      {[
                        { value: '', label: '空欄' },
                        { value: 'treatment-after', label: '治療時・治療後写真' },
                        { value: 'inspection', label: '検査時写真' },
                      ].map((item, idx) => {
                        const isSelected = reportFields.page2PhotoCategory === item.value;
                        const isHighlighted = dropdownHighlight === idx;
                        return (
                          <li key={item.value}>
                            <button
                              type="button"
                              className={`w-full px-3 py-2 text-left text-base transition-colors ${isHighlighted ? 'bg-orange-100 text-orange-800' : isSelected ? 'bg-orange-50 text-orange-700' : 'text-slate-800 hover:bg-slate-50'}`}
                              onClick={() => {
                                setReportFields(v => ({ ...v, page2PhotoCategory: item.value }));
                                setIsPage2PhotoCategoryDropdownOpen(false);
                              }}
                            >
                              {item.label}
                            </button>
                          </li>
                        );
                      })}
                    </ul>
                  )}
                </div>
              </div>
              <div className="space-y-1">
                <label className="text-xs font-semibold text-slate-700 uppercase tracking-widest">【検査・処置内容】</label>
                <textarea
                  ref={page2ProcedureTextareaRef}
                  className="w-full border border-slate-200 rounded-xl px-3 py-2 text-base min-h-[80px] focus:ring-2 focus:ring-orange-500 outline-none transition-all bg-white"
                  style={{ whiteSpace: 'pre-wrap', wordBreak: 'break-word', overflowWrap: 'break-word', maxHeight: '18.26cm', overflowY: 'auto' }}
                  placeholder=""
                  value={reportFields.procedureText}
                  onChange={e => setReportFields(v => ({ ...v, procedureText: e.target.value }))}
                />
              </div>
              <div className="space-y-1">
                <label className="text-xs font-semibold text-slate-700 uppercase tracking-widest">【術後経過】</label>
                <textarea className="w-full border border-slate-200 rounded-xl px-3 py-2 text-base min-h-[80px] focus:ring-2 focus:ring-orange-500 outline-none transition-all bg-white"
                  style={{ whiteSpace: 'pre-wrap', wordBreak: 'break-word', overflowWrap: 'break-word', maxHeight: '18.26cm', overflowY: 'auto' }}
                  placeholder=""
                  value={reportFields.postText}
                  onKeyDown={(e) => {
                    if (e.key === 'Tab' && !e.shiftKey && !showPage3) {
                      setTimeout(() => {
                        if (imageToolbarRef.current) {
                          const top = imageToolbarRef.current.getBoundingClientRect().top + window.scrollY - 60;
                          window.scrollTo({ top, behavior: 'smooth' });
                        }
                      }, 30);
                    }
                  }}
                  onChange={e => setReportFields(v => ({ ...v, postText: e.target.value }))}
                />
              </div>
              {!showPage3 && (
              <div className="space-y-1">
                <label className="text-xs font-semibold text-slate-700 uppercase tracking-widest">【お礼文】</label>
                <div className="relative" ref={thankYouTextTypeDropdownRef}>
                  <button
                    type="button"
                    tabIndex={-1}
                    className={`w-full h-11 border rounded-xl px-3 py-2 text-base text-left focus:ring-2 focus:ring-orange-500 outline-none transition-all flex items-center ${getEmptyFieldToneClass(reportFields.thankYouTextType || '')} bg-white`}
                    aria-haspopup="listbox"
                    aria-expanded={isThankYouTextTypeDropdownOpen}
                    onFocus={() => {
                      if (!shouldOpenThankYouTextTypeOnFocusRef.current) return;
                      shouldOpenThankYouTextTypeOnFocusRef.current = false;
                      setIsThankYouTextTypeDropdownOpen(true);
                      setDropdownHighlight(0);
                    }}
                    onClick={() => { setIsThankYouTextTypeDropdownOpen(v => !v); setDropdownHighlight(-1); }}
                    onKeyDown={isThankYouTextTypeDropdownOpen ? (e) => {
                      const items = [{ value: 'existing' }, { value: 'first-time' }];
                      handleDropdownKeyDown(e, items.length, (idx) => {
                        setReportFields(v => ({ ...v, thankYouTextType: items[idx].value }));
                        setIsThankYouTextTypeDropdownOpen(false);
                      }, () => setIsThankYouTextTypeDropdownOpen(false),
                      () => {
                        if (imageToolbarRef.current) {
                          const top = imageToolbarRef.current.getBoundingClientRect().top + window.scrollY - 60;
                          window.scrollTo({ top, behavior: 'smooth' });
                        }
                      });
                    } : undefined}
                  >
                    <span className={reportFields.thankYouTextType ? 'text-slate-900' : 'text-slate-500'}>
                      {reportFields.thankYouTextType === 'existing'
                        ? '① 既存紹介先向け'
                        : '② 初回紹介先向け'}
                    </span>
                  </button>

                  {isThankYouTextTypeDropdownOpen && (
                    <ul
                      role="listbox"
                      className="absolute z-40 mt-1 max-h-56 w-full overflow-auto rounded-xl border border-slate-200 bg-white py-1 shadow-lg"
                    >
                      {[
                        { value: 'existing', label: '① 既存紹介先向け' },
                        { value: 'first-time', label: '② 初回紹介先向け' },
                      ].map((item, idx) => {
                        const isSelected = (reportFields.thankYouTextType || 'first-time') === item.value;
                        const isHighlighted = dropdownHighlight === idx;
                        return (
                          <li key={item.value}>
                            <button
                              type="button"
                              className={`w-full px-3 py-2 text-left text-base transition-colors ${isHighlighted ? 'bg-orange-100 text-orange-800' : isSelected ? 'bg-orange-50 text-orange-700' : 'text-slate-800 hover:bg-slate-50'}`}
                              onClick={() => {
                                setReportFields(v => ({ ...v, thankYouTextType: item.value }));
                                setIsThankYouTextTypeDropdownOpen(false);
                              }}
                            >
                              {item.label}
                            </button>
                          </li>
                        );
                      })}
                    </ul>
                  )}
                </div>
              </div>
              )}
              {showPage3 && (
                <div className="bg-white p-3 rounded-xl space-y-2">
                  <div className="flex items-center gap-2">
                    <label className="text-xs font-semibold text-slate-700 uppercase tracking-widest">PAGE3写真ラベル</label>
                    <span className="text-xs text-slate-400">（自由入力）</span>
                  </div>
                  <div className="inline-flex items-center h-11 max-w-md rounded-xl border border-slate-200 bg-white overflow-hidden focus-within:ring-2 focus-within:ring-orange-500 transition-all">
                    <span className="pl-3 pr-0.5 text-slate-700 text-base font-medium shrink-0 select-none">【</span>
                    <input
                      type="text"
                      size={Math.max(((reportFields as any).page3PhotoLabel || '').length + 1, 10)}
                      className="border-0 outline-none bg-transparent text-base px-0 h-full min-w-[10ch] max-w-[30ch] placeholder:text-slate-400"
                      placeholder="術後口腔内写真"
                      value={(reportFields as any).page3PhotoLabel || ''}
                      onChange={e => setReportFields(v => ({ ...v, page3PhotoLabel: e.target.value }))}
                    />
                    <span className="pl-0.5 pr-3 text-slate-700 text-base font-medium shrink-0 select-none">】</span>
                  </div>
                </div>
              )}
              {showPage3 && (
                <div className="space-y-1">
                  <label className="text-xs font-semibold text-slate-700 uppercase tracking-widest">【自由入力】　PAGE3</label>
                  <textarea className="w-full border border-slate-200 rounded-xl px-3 py-2 text-base min-h-[80px] focus:ring-2 focus:ring-orange-500 outline-none transition-all bg-white"
                    style={{ whiteSpace: 'pre-wrap', wordBreak: 'break-word', overflowWrap: 'break-word', maxHeight: '18.26cm', overflowY: 'auto' }}
                    placeholder=""
                    value={reportFields.page3Text || ''}
                    onKeyDown={(e) => {
                      if (e.key === 'Tab' && !e.shiftKey) {
                        setTimeout(() => {
                          if (imageToolbarRef.current) {
                            const top = imageToolbarRef.current.getBoundingClientRect().top + window.scrollY - 60;
                            window.scrollTo({ top, behavior: 'smooth' });
                          }
                        }, 30);
                      }
                    }}
                    onChange={e => setReportFields(v => ({ ...v, page3Text: e.target.value }))}
                  />
                </div>
              )}
              {showPage3 && (
              <div className="space-y-1">
                <label className="text-xs font-semibold text-slate-700 uppercase tracking-widest">【お礼文】</label>
                <div className="relative" ref={thankYouTextTypeDropdownRef}>
                  <button
                    type="button"
                    tabIndex={-1}
                    className={`w-full h-11 border rounded-xl px-3 py-2 text-base text-left focus:ring-2 focus:ring-orange-500 outline-none transition-all flex items-center ${getEmptyFieldToneClass(reportFields.thankYouTextType || '')} bg-white`}
                    aria-haspopup="listbox"
                    aria-expanded={isThankYouTextTypeDropdownOpen}
                    onFocus={() => {
                      if (!shouldOpenThankYouTextTypeOnFocusRef.current) return;
                      shouldOpenThankYouTextTypeOnFocusRef.current = false;
                      setIsThankYouTextTypeDropdownOpen(true);
                      setDropdownHighlight(0);
                    }}
                    onClick={() => { setIsThankYouTextTypeDropdownOpen(v => !v); setDropdownHighlight(-1); }}
                    onKeyDown={isThankYouTextTypeDropdownOpen ? (e) => {
                      const items = [{ value: 'existing' }, { value: 'first-time' }];
                      handleDropdownKeyDown(e, items.length, (idx) => {
                        setReportFields(v => ({ ...v, thankYouTextType: items[idx].value }));
                        setIsThankYouTextTypeDropdownOpen(false);
                      }, () => setIsThankYouTextTypeDropdownOpen(false));
                    } : undefined}
                  >
                    <span className={reportFields.thankYouTextType ? 'text-slate-900' : 'text-slate-500'}>
                      {reportFields.thankYouTextType === 'existing'
                        ? '① 既存紹介先向け'
                        : '② 初回紹介先向け'}
                    </span>
                  </button>

                  {isThankYouTextTypeDropdownOpen && (
                    <ul
                      role="listbox"
                      className="absolute z-40 mt-1 max-h-56 w-full overflow-auto rounded-xl border border-slate-200 bg-white py-1 shadow-lg"
                    >
                      {[
                        { value: 'existing', label: '① 既存紹介先向け' },
                        { value: 'first-time', label: '② 初回紹介先向け' },
                      ].map((item, idx) => {
                        const isSelected = (reportFields.thankYouTextType || 'first-time') === item.value;
                        const isHighlighted = dropdownHighlight === idx;
                        return (
                          <li key={item.value}>
                            <button
                              type="button"
                              className={`w-full px-3 py-2 text-left text-base transition-colors ${isHighlighted ? 'bg-orange-100 text-orange-800' : isSelected ? 'bg-orange-50 text-orange-700' : 'text-slate-800 hover:bg-slate-50'}`}
                              onClick={() => {
                                setReportFields(v => ({ ...v, thankYouTextType: item.value }));
                                setIsThankYouTextTypeDropdownOpen(false);
                              }}
                            >
                              {item.label}
                            </button>
                          </li>
                        );
                      })}
                    </ul>
                  )}
                </div>
              </div>
              )}
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
        <div ref={imageToolbarRef} className="lg:col-span-12 relative w-full my-4 h-[48px]">
          <div className="absolute right-0 top-0 z-10">
            <button
              onClick={() => {
                console.log('image upload button clicked');
                if (fileInputRef.current) {
                  fileInputRef.current.value = '';
                }
                fileInputRef.current?.click();
              }}
              className="bg-indigo-600 hover:bg-indigo-700 text-white px-6 py-3 rounded-2xl font-bold shadow-lg active:scale-95 flex items-center gap-2 transition-all"
            >
              <svg className="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2.5" d="M12 4v16m8-8H4" />
              </svg>
              画像追加
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
            className="bg-white hidden"
          />
        </div>

        {/* 左カラム - 確定前のみ表示 */}
        {!isCurrentPageConfirmed && (
          <div className="lg:col-span-12 space-y-8 flex flex-col">
            <LayoutControls options={options} setOptions={setOptions} />

            <div className="w-full max-w-5xl mx-auto bg-white p-7 rounded-[2.5rem] shadow-sm border border-slate-200 overflow-hidden space-y-8 relative">
              <section>
                <div className="mb-4 flex items-center gap-3 flex-wrap justify-between">
                  <div>
                    <h3 className="inline-flex items-center gap-2 px-4 py-2 rounded-full border text-lg font-semibold shadow-sm bg-orange-50 border-orange-200 text-orange-700">画像編集・段落選択（Page {currentPage}）</h3>
                  </div>
                  {history.length > 0 && (
                    <button onClick={handleUndo} className="px-3 py-1.5 bg-slate-50 text-slate-500 rounded-xl hover:bg-orange-50 hover:text-orange-600 transition-all border border-slate-200 flex items-center gap-1.5 shadow-sm active:scale-95">
                      <svg className="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="3" d="M3 10h10a8 8 0 018 8v2M3 10l6 6m-6-6l6-6" /></svg>
                      <span className="text-sm font-black uppercase tracking-widest">戻す</span>
                    </button>
                  )}
                </div>

                <div className="space-y-4 flex-1 max-h-[500px] overflow-y-auto pr-2 custom-scrollbar">
                  {unassignedImages.map(img => {
                    // 画像表示エリアは常時表示
                    return (
                      <div key={img.id} className="p-3 bg-slate-50/50 border border-slate-100 rounded-[2rem] flex items-start gap-3 group">
                        <div
                          id={`crop-host-${img.id}`}
                          style={{
                            position: 'relative',
                            flex: 1,
                            minWidth: 0,
                            height: "200px",
                            padding: activeCropImageId === img.id ? `${CROP_UI_GUTTER_PX}px` : 0,
                            overflow: "hidden",
                            borderRadius: "12px",
                            border: "1px solid #e2e8f0",
                            background: "#ffffff",
                            boxShadow: "inset 0 1px 2px rgba(0,0,0,0.06)"
                          }}
                        >
                          <img
                            id={`crop-image-${img.id}`}
                            src={img.dataUrl}
                            alt="画像"
                            style={getImageDisplayStyle(img)}
                          />
                          {activeCropImageId === img.id && (
                            // ...既存 code for crop overlay ...
                            <div
                              id={`crop-overlay-${img.id}`}
                              className="absolute"
                              style={activeCropViewportRect && activeCropImageId === img.id ? {
                                left: `${activeCropViewportRect.left}px`,
                                top: `${activeCropViewportRect.top}px`,
                                width: `${activeCropViewportRect.width}px`,
                                height: `${activeCropViewportRect.height}px`,
                              } : {
                                left: 0,
                                top: 0,
                                width: '100%',
                                height: '100%',
                              }}
                              onClick={(e) => e.stopPropagation()}
                            >
                              <div className="absolute inset-0 bg-slate-900/10" />
                              {(() => {
                                const crop = getImageCrop(img) ?? DEFAULT_IMAGE_CROP;
                                const cropPixels = activeCropViewportRect && activeCropImageId === img.id
                                  ? cropToPixels(crop, activeCropViewportRect, img.rotation, img.flipX, img.flipY)
                                  : null;
                                return (
                                  <div
                                    className="absolute border-2 border-emerald-500 bg-transparent cursor-move"
                                    style={cropPixels ? {
                                      left: `${cropPixels.left}px`,
                                      top: `${cropPixels.top}px`,
                                      width: `${cropPixels.right - cropPixels.left}px`,
                                      height: `${cropPixels.bottom - cropPixels.top}px`,
                                    } : {
                                      left: `${crop.left * 100}%`,
                                      top: `${crop.top * 100}%`,
                                      width: `${(crop.right - crop.left) * 100}%`,
                                      height: `${(crop.bottom - crop.top) * 100}%`,
                                    }}
                                    onMouseDown={(e) => startCropMove(e, img.id)}
                                  >
                                    <div className="absolute inset-0 border border-white/80 pointer-events-none" />
                                    <button
                                      type="button"
                                      className="absolute -left-2.5 -top-2.5 h-5 w-5 rounded-full bg-white border-2 border-emerald-600 cursor-nwse-resize"
                                      onMouseDown={(e) => startCropDrag(e, img.id, 'top-left')}
                                      onClick={(e) => e.stopPropagation()}
                                      aria-label="左上ハンドル"
                                    />
                                    <button
                                      type="button"
                                      className="absolute -right-2.5 -top-2.5 h-5 w-5 rounded-full bg-white border-2 border-emerald-600 cursor-nesw-resize"
                                      onMouseDown={(e) => startCropDrag(e, img.id, 'top-right')}
                                      onClick={(e) => e.stopPropagation()}
                                      aria-label="右上ハンドル"
                                    />
                                    <button
                                      type="button"
                                      className="absolute -left-2.5 -bottom-2.5 h-5 w-5 rounded-full bg-white border-2 border-emerald-600 cursor-nesw-resize"
                                      onMouseDown={(e) => startCropDrag(e, img.id, 'bottom-left')}
                                      onClick={(e) => e.stopPropagation()}
                                      aria-label="左下ハンドル"
                                    />
                                    <button
                                      type="button"
                                      className="absolute -right-2.5 -bottom-2.5 h-5 w-5 rounded-full bg-white border-2 border-emerald-600 cursor-nwse-resize"
                                      onMouseDown={(e) => startCropDrag(e, img.id, 'bottom-right')}
                                      onClick={(e) => e.stopPropagation()}
                                      aria-label="右下ハンドル"
                                    />
                                  </div>
                                );
                              })()}
                            </div>
                          )}
                        </div>
                        {/* 右側の操作ボタン群のみ分岐 */}
                        <div className="w-56 flex-shrink-0">
                          <div className="grid grid-cols-3 gap-2 items-start w-full">
                            {editingImageId === img.id ? (
                              <>
                                {/* 操作パネル: 回転・反転 | トリミング ドロップダウン */}
                                <div className="col-span-3 flex flex-col gap-2 w-full">
                                  <div className="flex gap-2 w-full">
                                    {/* 左: 回転・反転ドロップダウン */}
                                    <div className="flex-1 flex flex-col gap-1">
                                      <button
                                        onClick={() => setIsOrientationOpen(v => !v)}
                                        disabled={isCropOpen || isOrientationOpen}
                                        className={`h-10 rounded-lg border text-sm font-semibold transition-all shadow-sm whitespace-nowrap bg-indigo-100 text-indigo-800 border-indigo-200 hover:bg-indigo-200 active:scale-95 ${isCropOpen || isOrientationOpen ? 'opacity-40 cursor-not-allowed' : ''}`}
                                      >回転・反転</button>
                                      {isOrientationOpen && (
                                        <div className="flex flex-col gap-1 mt-1">
                                          <div className="grid grid-cols-2 gap-1">
                                            <button onClick={() => rotateImage(img.id, 'left')} className="h-9 rounded-lg bg-indigo-50 text-indigo-700 hover:bg-indigo-100 border border-indigo-200 transition-all active:scale-95 flex items-center justify-center"><svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2.5" d="M4 4v5h.582m15.356 2A8.001 8.001 0 004.582 9m0 0H9" /></svg></button>
                                            <button onClick={() => rotateImage(img.id, 'right')} className="h-9 rounded-lg bg-indigo-50 text-indigo-700 hover:bg-indigo-100 border border-indigo-200 transition-all active:scale-95 flex items-center justify-center"><svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2.5" d="M20 4v5h-.582m-15.356 2A8.001 8.001 0 0119.418 9m0 0H15" /></svg></button>
                                            <button onClick={() => flipImageX(img.id)} className="h-9 rounded-lg bg-indigo-50 text-indigo-700 hover:bg-indigo-100 border border-indigo-200 transition-all active:scale-95 flex items-center justify-center"><svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2.5" d="M4 12h16M10 16l-4-4 4-4" /></svg></button>
                                            <button onClick={() => flipImageY(img.id)} className="h-9 rounded-lg bg-indigo-50 text-indigo-700 hover:bg-indigo-100 border border-indigo-200 transition-all active:scale-95 flex items-center justify-center"><svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2.5" d="M12 4v16M16 10l-4-4-4 4" /></svg></button>
                                          </div>
                                          <button
                                            onClick={async () => {
                                              recordHistory();
                                              const image = new window.Image();
                                              image.src = img.dataUrl;
                                              await new Promise(resolve => { image.onload = resolve; });
                                              const rotated = img.rotation === 90 || img.rotation === 270;
                                              const canvasW = rotated ? image.naturalHeight : image.naturalWidth;
                                              const canvasH = rotated ? image.naturalWidth : image.naturalHeight;
                                              const canvas = document.createElement('canvas');
                                              canvas.width = canvasW;
                                              canvas.height = canvasH;
                                              const ctx = canvas.getContext('2d');
                                              if (ctx) {
                                                ctx.save();
                                                ctx.translate(canvasW / 2, canvasH / 2);
                                                ctx.rotate((img.rotation * Math.PI) / 180);
                                                ctx.scale(img.flipX ? -1 : 1, img.flipY ? -1 : 1);
                                                ctx.drawImage(image, -image.naturalWidth / 2, -image.naturalHeight / 2, image.naturalWidth, image.naturalHeight);
                                                ctx.restore();
                                              }
                                              const bakedDataUrl = canvas.toDataURL();
                                              setImages((prev: ImageData[]) =>
                                                prev.map(i =>
                                                  i.id === img.id
                                                    ? { ...i, dataUrl: bakedDataUrl, rotation: 0, flipX: false, flipY: false, crop: undefined, width: canvasW, height: canvasH }
                                                    : i
                                                )
                                              );
                                              closeCropState();
                                              setIsOrientationOpen(false);
                                            }}
                                            className="h-9 rounded-lg bg-indigo-100 text-indigo-700 hover:bg-indigo-200 border border-indigo-200 text-sm font-semibold transition-all active:scale-95"
                                          >確定</button>
                                        </div>
                                      )}
                                    </div>
                                    {/* 右: トリミングドロップダウン */}
                                    <div className="flex-1 flex flex-col gap-1 border-l border-slate-200 pl-2">
                                      <button
                                        onClick={() => {
                                          console.log('TRIM START', { imgId: img.id, row: img.row, dataUrlLength: img.dataUrl?.length, crop: getImageCrop(img), activeCropImageId, activeCropViewportRect, rotation: img.rotation, flipX: img.flipX, flipY: img.flipY });
                                          // 古いトリミング状態を先にクリアして isActiveCropTarget を false にする
                                          setActiveCropImageId(null);
                                          setActiveCropViewportRect(null);
                                          if (!getImageCrop(img)) {
                                            updateImageCrop(img.id, () => DEFAULT_IMAGE_CROP);
                                          }
                                          setIsCropOpen(true);
                                          requestAnimationFrame(() => {
                                            const imageEl = document.getElementById(`crop-image-${img.id}`) as HTMLImageElement | null;
                                            if (imageEl) {
                                              const rect = imageEl.getBoundingClientRect();
                                              setActiveCropViewportRect({ left: rect.left, top: rect.top, width: rect.width, height: rect.height });
                                            } else {
                                              setActiveCropViewportRect(null);
                                            }
                                            setActiveCropImageId(img.id);
                                          });
                                        }}
                                        disabled={isOrientationOpen || img.rotation !== 0 || !!img.flipX || !!img.flipY}
                                        title={img.rotation !== 0 || !!img.flipX || !!img.flipY ? '向きを反映してからトリミングしてください' : undefined}
                                        className={`h-10 rounded-lg border text-sm font-semibold transition-all shadow-sm whitespace-nowrap ${isCropOpen ? 'bg-emerald-200 text-emerald-900 border-emerald-300' : 'bg-emerald-50 text-emerald-800 border-emerald-200 hover:bg-emerald-100'} ${isOrientationOpen || img.rotation !== 0 || !!img.flipX || !!img.flipY ? 'opacity-40 cursor-not-allowed' : 'active:scale-95'}`}
                                      >トリミング</button>
                                      {isCropOpen && (
                                        <div className="flex flex-col gap-1 mt-1">
                                          <button
                                            onClick={() => { recordHistory(); resetImageCrop(img.id); }}
                                            className="h-9 rounded-lg bg-emerald-50 text-emerald-700 hover:bg-emerald-100 border border-emerald-200 text-sm font-semibold transition-all active:scale-95"
                                          >リセット</button>
                                          <button
                                            onClick={async () => {
                                              const crop = getImageCrop(img);
                                              if (crop && activeCropViewportRect) {
                                                const image = new window.Image();
                                                image.src = img.dataUrl;
                                                await new Promise(resolve => { image.onload = resolve; });
                                                const canvas = document.createElement('canvas');
                                                const viewportW = activeCropViewportRect.width;
                                                const viewportH = activeCropViewportRect.height;
                                                const imageAspect = image.naturalWidth / image.naturalHeight;
                                                const viewportAspect = viewportW / viewportH;
                                                let drawWidth: number, drawHeight: number, offsetX: number, offsetY: number;
                                                if (imageAspect > viewportAspect) {
                                                  drawWidth = viewportW;
                                                  drawHeight = viewportW / imageAspect;
                                                  offsetX = 0;
                                                  offsetY = (viewportH - drawHeight) / 2;
                                                } else {
                                                  drawHeight = viewportH;
                                                  drawWidth = viewportH * imageAspect;
                                                  offsetY = 0;
                                                  offsetX = (viewportW - drawWidth) / 2;
                                                }
                                                const adjLeft   = Math.max(0, Math.min(1, (crop.left   * viewportW - offsetX) / drawWidth));
                                                const adjTop    = Math.max(0, Math.min(1, (crop.top    * viewportH - offsetY) / drawHeight));
                                                const adjRight  = Math.max(0, Math.min(1, (crop.right  * viewportW - offsetX) / drawWidth));
                                                const adjBottom = Math.max(0, Math.min(1, (crop.bottom * viewportH - offsetY) / drawHeight));
                                                const sx = Math.round(adjLeft   * image.naturalWidth);
                                                const sy = Math.round(adjTop    * image.naturalHeight);
                                                const sw = Math.round((adjRight  - adjLeft)  * image.naturalWidth);
                                                const sh = Math.round((adjBottom - adjTop)   * image.naturalHeight);
                                                canvas.width = sw;
                                                canvas.height = sh;
                                                const ctx = canvas.getContext('2d');
                                                if (ctx) {
                                                  ctx.drawImage(image, sx, sy, sw, sh, 0, 0, sw, sh);
                                                  const croppedDataUrl = canvas.toDataURL();
                                                  setImages((prev: ImageData[]) =>
                                                    prev.map(i =>
                                                      i.id === img.id
                                                        ? { ...i, dataUrl: croppedDataUrl, crop: undefined, width: sw, height: sh }
                                                        : i
                                                    )
                                                  );
                                                }
                                              }
                                              closeCropState();
                                            }}
                                            className="h-9 rounded-lg bg-emerald-100 text-emerald-800 hover:bg-emerald-200 border border-emerald-200 text-sm font-semibold transition-all active:scale-95"
                                          >確定</button>
                                        </div>
                                      )}
                                    </div>
                                  </div>
                                  {/* 編集を閉じる: 常に表示・独立 */}
                                  <button
                                    onClick={() => {
                                      closeEditUiState();
                                      setEditStep('orientation');
                                    }}
                                    className="h-10 rounded-lg border border-slate-300 bg-white text-slate-600 text-sm font-semibold hover:bg-slate-100 w-full transition-all active:scale-95"
                                  >編集を閉じる</button>
                                </div>
                              </>
                            ) : (
                              <>
                                {/* 通常モード用ボタン群 */}
                                {/* 右列: 段落番号ボタン */}
                                {editingImageId !== img.id && (
                                  <div className="flex flex-col justify-between items-stretch" style={{ minHeight: 'calc(100% - 4px)', gap: '10px', height: '100%' }}>
                                    {[1, 2, 3, 4].map(num => (
                                      <button
                                        key={num}
                                        onClick={() => {
                                          closeCropState();
                                          setEditingImageId(null);
                                          updateImageRow(img.id, num);
                                          const isLastUnassigned = unassignedImages.length === 1 && img.row === 0;
                                          if (isLastUnassigned) {
                                            if (currentPage === 1) setPage1Confirmed(true);
                                            if (currentPage === 2) setPage2Confirmed(true);
                                            if (currentPage === 3) setPage3Confirmed(true);
                                          }
                                        }}
                                        className={`rounded-lg text-lg font-semibold transition-all shadow-sm active:scale-95 border ${img.row === num ? 'bg-orange-500 text-white border-orange-400 shadow font-bold' : 'bg-white text-slate-800 border-slate-300 hover:bg-slate-100'}`}
                                        style={{ height: '44px', minHeight: '40px', padding: '0' }}
                                      >
                                        {num}
                                      </button>
                                    ))}
                                  </div>
                                )}
                                {/* 中央列: 削除ボタン */}
                                <div className="flex flex-col gap-2">
                                  <button onClick={() => removeImage(img.id)} className="w-full h-12 rounded-lg bg-rose-100 text-rose-700 hover:bg-rose-200 border border-rose-200 transition-all shadow-sm flex items-center justify-center active:scale-95"><svg className="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2.5" d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16" /></svg></button>
                                </div>
                                {/* 左列: 編集ボタン */}
                                <div className="flex flex-col gap-2">
                                  <button onClick={() => setEditingImageId(img.id)} className="w-full h-12 rounded-lg bg-indigo-100 text-indigo-700 hover:bg-indigo-200 border border-indigo-200 transition-all shadow-sm flex items-center justify-center active:scale-95">編集</button>
                                </div>
                              </>
                            )}
                          </div>
                        </div>
                      </div>
                    );
                  })}


                </div>
              </section>
            </div>
          </div>
        )}

        {/* 段落ドラッグ移動 */}
        <div ref={rowBoardRef} className="lg:col-span-12 bg-white p-7 rounded-[2.5rem] shadow-sm border border-slate-200 space-y-4">
          <div className="flex items-center gap-3 mb-4 flex-wrap">
            <h3 className="inline-flex items-center gap-2 px-4 py-2 rounded-full border text-lg font-semibold shadow-sm bg-sky-50 border-sky-200 text-sky-700">
              {isCurrentPageConfirmed ? `画像入れ替え（Page ${currentPage}）` : '段落ドラッグ移動'}
            </h3>
          </div>
          <RowBoard
  images={images}
  setImages={setImages}
  rows={4}
  setActiveCropImageId={setActiveCropImageId}
  onUnassignImage={handleUnassignImage}
/>
        </div>

        {/* PAGE切替ボタン（段落エリアとプレビューの間） */}
        <div ref={pageSwitcherRef} className="lg:col-span-12 grid grid-cols-3 items-center my-4">
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

        <div className="lg:col-span-12 flex flex-row items-start w-full">
          {/* 左: プレビュー */}
          <div className="flex-1 flex justify-center mt-12 mb-20">
            <div className="flex justify-center items-center w-full py-8" style={{ backgroundColor: 'rgba(241,245,249,0.5)' }}>
              <div className="relative" style={{ width: '560px', height: '560px', maxWidth: '98vw', aspectRatio: '210 / 297', transform: 'scale(1.3)' }}>
                <div
                  className="w-full h-full"
                  style={{ color: '#0f172a' }}
                  dangerouslySetInnerHTML={{ __html: svgData.svgCode }}
                />
              </div>
            </div>
          </div>
          {/* 右: Yオフセット調整 */}
          <div className="flex-1 flex justify-center">
            <div className="rounded-2xl border border-slate-200 bg-white shadow-sm px-4 py-3 flex flex-col w-full max-w-[520px]">
              <div className="flex items-center w-full text-left mb-2">
                <span className="text-base font-semibold text-slate-700">Yオフセット調整</span>
                <div className="flex-1" />
                <button
                  type="button"
                  onClick={resetCurrentPagePreviewYOffsets}
                  className="text-base px-3 py-1.5 rounded-lg border border-slate-200 bg-slate-50 hover:bg-slate-100 text-slate-700 transition-colors ml-2"
                >
                  このページをリセット
                </button>
              </div>
              <div>
                {activePreviewYOffsetGroup.items.map((item, idx) => (
                  <label key={`${item.key}-${idx}`} className="grid grid-cols-[1fr_auto] items-center gap-2 mb-2 last:mb-0">
                    <span className="text-base text-slate-700">{item.label}</span>
                    <div className="flex items-center gap-2">
                      <button
                        type="button"
                        onClick={() => {
                          if (suppressNextClickRef.current) {
                            suppressNextClickRef.current = false;
                            return;
                          }
                          nudgePreviewYOffset(item.key, -0.01);
                        }}
                        onMouseDown={() => startContinuousAdjust(() => nudgePreviewYOffset(item.key, -0.01))}
                        onMouseUp={stopContinuousAdjust}
                        onMouseLeave={stopContinuousAdjust}
                        onTouchStart={() => startContinuousAdjust(() => nudgePreviewYOffset(item.key, -0.01))}
                        onTouchEnd={stopContinuousAdjust}
                        onTouchCancel={stopContinuousAdjust}
                        className="h-10 px-3 rounded-lg border border-slate-200 bg-slate-50 hover:bg-slate-100 text-sm font-semibold text-slate-700 transition-colors"
                        aria-label={`${item.label} を上へ移動`}
                        title="上へ移動"
                      >
                        △
                      </button>
                      <input
                        type="number"
                        step="0.01"
                        className="w-20 h-10 rounded-lg border border-slate-200 bg-white px-2 text-base text-right focus:ring-2 focus:ring-orange-500 outline-none"
                        value={previewYOffsets[item.key]}
                        onChange={(e) => {
                          const raw = e.target.value;
                          const next = raw === '' ? 0 : Number(raw);
                          setPreviewYOffsets((prev) => ({
                            ...prev,
                            [item.key]: Number.isFinite(next) ? next : 0,
                          }));
                        }}
                      />
                      <span className="text-base text-slate-500">cm</span>
                      <button
                        type="button"
                        onClick={() => {
                          if (suppressNextClickRef.current) {
                            suppressNextClickRef.current = false;
                            return;
                          }
                          nudgePreviewYOffset(item.key, 0.01);
                        }}
                        onMouseDown={() => startContinuousAdjust(() => nudgePreviewYOffset(item.key, 0.01))}
                        onMouseUp={stopContinuousAdjust}
                        onMouseLeave={stopContinuousAdjust}
                        onTouchStart={() => startContinuousAdjust(() => nudgePreviewYOffset(item.key, 0.01))}
                        onTouchEnd={stopContinuousAdjust}
                        onTouchCancel={stopContinuousAdjust}
                        className="h-10 px-3 rounded-lg border border-slate-200 bg-slate-50 hover:bg-slate-100 text-sm font-semibold text-slate-700 transition-colors"
                        aria-label={`${item.label} を下へ移動`}
                        title="下へ移動"
                      >
                        ▽
                      </button>
                    </div>
                  </label>
                ))}
              </div>
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
        <h3 className="inline-flex items-center gap-2 px-4 py-2 rounded-full border text-lg font-semibold shadow-sm bg-emerald-50 border-emerald-200 text-emerald-700 mb-4">プレビュー</h3>
        
        {pptxStatus && <p className="text-base text-slate-600 font-bold mt-1">{pptxStatus}</p>}
      </div>
      <div className="flex flex-wrap gap-3">
        <button onClick={handleSaveDraft} disabled={!selectedCaseId}
          className="bg-slate-500 hover:bg-slate-600 text-white px-4 py-2 rounded-xl text-sm font-semibold shadow-md transition-all disabled:opacity-50 disabled:cursor-not-allowed">
          下書き保存
        </button>
        <button onClick={openGmailDraft}
          disabled={isCreatingDraft || !(reportFields.refHospitalEmail || '').trim()}
          title="Gmail下書きを作成し、PDFを添付します（送信はしません）"
          className="bg-slate-100 text-slate-700 border border-slate-200 px-4 py-2 rounded-xl text-sm font-semibold hover:bg-slate-200 transition-all flex items-center gap-2 disabled:opacity-50 disabled:cursor-not-allowed">
          {isCreatingDraft ? "作成中…" : "PDF / Gmail"}
        </button>
        <button onClick={sendGmail}
          disabled={isSendingGmail || !(reportFields.refHospitalEmail || '').trim()}
          title="PDFを添付してGmailで即時送信します"
          className="bg-blue-500 hover:bg-blue-600 text-white px-4 py-2 rounded-xl text-sm font-semibold shadow-md transition-all flex items-center gap-2 disabled:opacity-50 disabled:cursor-not-allowed">
          {isSendingGmail ? "送信中…" : "Gmail送信"}
        </button>
        <button onClick={printPdf}
          title="印刷ダイアログを開きます（PDF保存/プリンタ印刷）"
          className="bg-orange-500 hover:bg-orange-600 text-white font-semibold shadow-md rounded-xl px-4 py-2 transition-all focus:outline-none focus:ring-2 focus:ring-orange-400 text-sm flex items-center gap-2">
          PDF 印刷
        </button>
        <button onClick={downloadPptx} disabled={isSavingPptx}
          className="bg-slate-100 text-slate-700 border border-slate-200 px-4 py-2 rounded-xl text-sm font-semibold hover:bg-slate-200 flex items-center gap-2 transition-all disabled:opacity-50 disabled:cursor-not-allowed">
          {isSavingPptx ? '保存中…' : 'PPTX出力/編集'}
        </button>
      </div>
      </div>
      </div>
    </div>
  );
};


export default App;
