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
    attendingVet: '',
    initialText: '',
    procedureText: '',
    postText: '',
    page3Text: '',
    chiefComplaint: '',
    page2PhotoCategory: '',
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
  }, [showPage3]);

  // showPage3 をOFFにした瞬間に 3ページ目に居たら 2へ戻す
  useEffect(() => {
    if (!showPage3 && currentPage === 3) setCurrentPage(2);
  }, [showPage3, currentPage]);

  // 既存：参照や他state
  const rowBoardRef = useRef<HTMLDivElement | null>(null);
  const fileInputRef = useRef<HTMLInputElement | null>(null);
  const attendingVetDropdownRef = useRef<HTMLDivElement | null>(null);
  const [isAttendingVetDropdownOpen, setIsAttendingVetDropdownOpen] = useState(false);
  const page2PhotoCategoryDropdownRef = useRef<HTMLDivElement | null>(null);
  const [isPage2PhotoCategoryDropdownOpen, setIsPage2PhotoCategoryDropdownOpen] = useState(false);
  const pageOrderDropdownRef = useRef<HTMLDivElement | null>(null);
  const [isPageOrderDropdownOpen, setIsPageOrderDropdownOpen] = useState(false);
  const postPlacementDropdownRef = useRef<HTMLDivElement | null>(null);
  const [isPostPlacementDropdownOpen, setIsPostPlacementDropdownOpen] = useState(false);

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

  const page2PhotoCategoryLabel = useMemo(() => {
    if (reportFields.page2PhotoCategory === 'treatment-after') return '【治療時・治療後写真】';
    if (reportFields.page2PhotoCategory === 'inspection') return '【検査時写真】';
    return '';
  }, [reportFields.page2PhotoCategory]);

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
const [pageOrderMode, setPageOrderMode] = useState<"page2-page3" | "page3-page2">("page2-page3");

// どこに「経過」を入れるか（既存がこれならOK）
const [postPlacement, setPostPlacement] = useState<"page2" | "page3">("page2");


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
    const files = Array.from(e.target.files || []) as File[];
    if (!files.length) return;
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

  const calculateSvgDataForPage = useCallback((pageNum: 1 | 2 | 3) => {
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

    // use shared helper to render all text portions so that SVG and PPTX stay in sync
    svgParts.push(...buildSvgTextParts(pageNum, reportFields, pxPerCm, slideOffsetX, slideOffsetY, {
      showPage3,
      postPlacement,
    }));

    if (pageNum === 2 && page2PhotoCategoryLabel) {
      const labelX = slideOffsetX + 1.04 * pxPerCm;
      const labelY = slideOffsetY + 1.9 * pxPerCm;
      const labelFontSize = 0.42 * pxPerCm;
      svgParts.push(
        `  <text x="${labelX}" y="${labelY}" font-size="${labelFontSize}" font-weight="700" fill="#0f172a" dominant-baseline="hanging">${page2PhotoCategoryLabel}</text>`
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
    showPage3,
    postPlacement
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

        addPptxText(slide, pageNum, reportFields, {
          showPage3,
          postPlacement,
        });

        if (pageNum === 2 && page2PhotoCategoryLabel) {
          slide.addText(page2PhotoCategoryLabel, {
            x: 1.04 / 2.54,
            y: 1.9 / 2.54,
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

  const svgData = useMemo(() => calculateSvgDataForPage(previewPage), [calculateSvgDataForPage, previewPage]);
  const svgPage1 = useMemo(() => calculateSvgDataForPage(1).svgCode, [calculateSvgDataForPage]);
  const svgPage2 = useMemo(() => calculateSvgDataForPage(2).svgCode, [calculateSvgDataForPage]);
  const svgPage3 = useMemo(() => calculateSvgDataForPage(3).svgCode, [calculateSvgDataForPage]);

  type DateFieldKey = 'firstVisitDate' | 'sedationDate' | 'anesthesiaDate';

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
  }, [calendarMonth, formatCalendarDate, openDateField]);

  const clearCalendarDate = useCallback(() => {
    if (!openDateField) return;
    setReportFields(prev => ({ ...prev, [openDateField]: '' }));
    setOpenDateField(null);
  }, [openDateField]);

  const moveCalendarMonth = useCallback((offset: number) => {
    setCalendarMonth(prev => new Date(prev.getFullYear(), prev.getMonth() + offset, 1));
  }, []);

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
    if (!isPageOrderDropdownOpen) return;

    const closeWhenOutside = (event: Event) => {
      const target = event.target as Node | null;
      if (!target) return;
      if (pageOrderDropdownRef.current?.contains(target)) return;
      setIsPageOrderDropdownOpen(false);
    };

    document.addEventListener('pointerdown', closeWhenOutside, true);
    document.addEventListener('focusin', closeWhenOutside, true);

    return () => {
      document.removeEventListener('pointerdown', closeWhenOutside, true);
      document.removeEventListener('focusin', closeWhenOutside, true);
    };
  }, [isPageOrderDropdownOpen]);

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

      if (inputEl.id === 'chief-complaint-input') {
        e.preventDefault();
        const initialTextArea = document.getElementById('initial-textarea') as HTMLTextAreaElement | null;
        initialTextArea?.focus();
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
      focusables[idx + 1].focus();
    }
  }, []);

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
        <div className="lg:col-span-12 bg-white border border-slate-200 rounded-2xl shadow-sm p-4 md:p-5 space-y-4" onKeyDown={handleEnterFocusNextInput}>
          <div className="flex items-start justify-between gap-3">
            <div>
              <h3 className="font-semibold text-slate-700 text-base mb-3">報告書データ入力</h3>
            </div>
            <button
              type="button"
              onClick={handleClearReportFields}
              className="h-8 px-3 rounded-lg border border-slate-200 bg-white text-xs font-semibold text-slate-700 hover:bg-slate-50"
            >
              入力をクリア
            </button>
          </div>

          <div className="space-y-4">
            {/* 基本情報グリッド */}
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4 bg-white border border-slate-200 rounded-2xl shadow-sm p-4 md:p-5">
              <div className="space-y-1">
                <label className="text-[10px] font-semibold text-slate-700 uppercase tracking-widest">報告日</label>
                <input className={`w-full border rounded-xl px-3 py-2 text-sm focus:ring-2 focus:ring-orange-500 outline-none transition-all ${getEmptyFieldToneClass(reportFields.reportDate)}`}
                  placeholder="2026年2月16日"
                  value={reportFields.reportDate}
                  onChange={e => setReportFields(v => ({ ...v, reportDate: e.target.value }))}
                  onBlur={e => setReportFields(v => ({ ...v, reportDate: normalizeJapaneseDate(e.target.value) }))}
                />
              </div>
              <div className="space-y-1">
                <label className="text-[10px] font-semibold text-slate-700 uppercase tracking-widest">
                  紹介病院名
                </label>

                <input
                  id="ref-hospital-input"
                  className={`w-full max-w-[520px] h-9 px-3 rounded-lg border text-sm ${getEmptyFieldToneClass(refHospitalInput)}`}
                  placeholder="例：中川動物病院"
                  value={refHospitalInput}
                  onChange={(e) => applyRefHospitalSelection(e.target.value)}
                  onBlur={(e) => applyRefHospitalSelection(e.target.value)}
                  onKeyDown={(e) => {
                    if (e.key === "Enter") {
                      e.preventDefault();
                      const v = e.currentTarget.value;
                      const normalizedName = normalizeHospitalKey(v);
                      const hasMappedEmail = normalizedName ? !!normalizedRefHospitalEmails[normalizedName] : false;
                      applyRefHospitalSelection(v);
                      handleAddRefHospital(v);
                      if (hasMappedEmail) {
                        e.stopPropagation();
                        requestAnimationFrame(() => {
                          const doctorInput = document.getElementById('ref-doctor-input') as HTMLInputElement | null;
                          doctorInput?.focus();
                        });
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
                <label className="text-[10px] font-semibold text-slate-700 uppercase tracking-widest">紹介病院メール（Gmail）</label>
                <input className={`w-full border rounded-xl px-3 py-2 text-sm focus:ring-2 focus:ring-orange-500 outline-none transition-all ${getEmptyFieldToneClass(reportFields.refHospitalEmail)}`}
                  placeholder="example@gmail.com"
                  value={reportFields.refHospitalEmail}
                  onChange={e => setReportFields(v => ({ ...v, refHospitalEmail: e.target.value }))}
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
                <label className="text-[10px] font-semibold text-slate-700 uppercase tracking-widest">先生名</label>
                <div className="relative">
                  <input className={`w-full border rounded-xl px-3 pr-12 py-2 text-sm focus:ring-2 focus:ring-orange-500 outline-none transition-all ${getEmptyFieldToneClass(reportFields.refDoctor)}`}
                    id="ref-doctor-input"
                    placeholder="△△"
                    value={reportFields.refDoctor}
                    onChange={e => setReportFields(v => ({ ...v, refDoctor: e.target.value }))}
                  />
                  <span className="pointer-events-none absolute inset-y-0 right-3 flex items-center text-sm text-slate-500">先生</span>
                </div>
              </div>
              <div className="space-y-1">
                <label className="text-[10px] font-semibold text-slate-700 uppercase tracking-widest">飼い主姓</label>
                <input className={`w-full border rounded-xl px-3 py-2 text-sm focus:ring-2 focus:ring-orange-500 outline-none transition-all ${getEmptyFieldToneClass(reportFields.ownerLastName)}`}
                  placeholder="山田"
                  value={reportFields.ownerLastName}
                  onChange={e => setReportFields(v => ({ ...v, ownerLastName: e.target.value }))}
                />
              </div>
              <div className="space-y-1">
                <label className="text-[10px] font-semibold text-slate-700 uppercase tracking-widest">ペット名</label>
                <input className={`w-full border rounded-xl px-3 py-2 text-sm focus:ring-2 focus:ring-orange-500 outline-none transition-all ${getEmptyFieldToneClass(reportFields.petName)}`}
                  placeholder="タロウ"
                  value={reportFields.petName}
                  onChange={e => setReportFields(v => ({ ...v, petName: e.target.value }))}
                />
              </div>
              <div className="space-y-1">
                <label className="text-[10px] font-semibold text-slate-700 uppercase tracking-widest">初診日</label>
                <div className="relative" data-date-field="firstVisitDate">
                  <input className={`w-full border rounded-xl px-3 py-2 text-sm focus:ring-2 focus:ring-orange-500 outline-none transition-all cursor-pointer ${getEmptyFieldToneClass(reportFields.firstVisitDate)}`}
                    placeholder="202X年XX月XX日"
                    value={reportFields.firstVisitDate}
                    readOnly
                    onClick={() => openCalendar('firstVisitDate')}
                  />
                  {openDateField === 'firstVisitDate' && (
                    <div className="absolute left-0 top-full mt-2 z-40 w-72 rounded-2xl border border-slate-200 bg-white p-3 shadow-xl">
                      <div className="mb-2 flex items-center justify-between">
                        <div className="text-lg font-bold text-slate-800">{calendarMonth.getFullYear()}年 {calendarMonth.getMonth() + 1}月</div>
                        <div className="flex items-center gap-1">
                          <button type="button" onClick={() => moveCalendarMonth(-1)} className="h-7 w-7 rounded-lg border border-slate-200 text-slate-600 hover:bg-slate-50">‹</button>
                          <button type="button" onClick={() => moveCalendarMonth(1)} className="h-7 w-7 rounded-lg border border-slate-200 text-slate-600 hover:bg-slate-50">›</button>
                        </div>
                      </div>
                      <div className="mb-2 grid grid-cols-7 gap-1 text-center text-[11px] font-semibold text-slate-500">
                        {['日', '月', '火', '水', '木', '金', '土'].map(day => <span key={day}>{day}</span>)}
                      </div>
                      <div className="grid grid-cols-7 gap-1">
                        {calendarCells.map((day, idx) => {
                          if (!day) return <span key={`empty-${idx}`} className="h-8" />;
                          const isSelected = !!selectedCalendarDate
                            && selectedCalendarDate.getFullYear() === calendarMonth.getFullYear()
                            && selectedCalendarDate.getMonth() === calendarMonth.getMonth()
                            && selectedCalendarDate.getDate() === day;
                          return (
                            <button
                              key={day}
                              type="button"
                              onClick={() => selectCalendarDate(day)}
                              className={`h-8 rounded-lg text-sm font-medium transition-colors ${
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
                      <div className="mt-3 flex justify-between">
                        <button type="button" onClick={clearCalendarDate} className="rounded-lg border border-slate-200 px-2 py-1 text-xs font-semibold text-slate-400 hover:bg-slate-50">クリア</button>
                        <button type="button" onClick={closeCalendar} className="rounded-lg border border-slate-200 px-2 py-1 text-xs font-semibold text-slate-600 hover:bg-slate-50">閉じる</button>
                      </div>
                    </div>
                  )}
                </div>
              </div>
              <div className="space-y-1">
                <label className="text-[10px] font-semibold text-slate-700 uppercase tracking-widest">鎮静日</label>
                <div className="relative" data-date-field="sedationDate">
                  <input className={`w-full border rounded-xl px-3 py-2 text-sm focus:ring-2 focus:ring-orange-500 outline-none transition-all cursor-pointer ${getEmptyFieldToneClass(reportFields.sedationDate)}`}
                    placeholder="202X年XX月XX日"
                    value={reportFields.sedationDate || ''}
                    readOnly
                    onClick={() => openCalendar('sedationDate')}
                  />
                  {openDateField === 'sedationDate' && (
                    <div className="absolute left-0 top-full mt-2 z-40 w-72 rounded-2xl border border-slate-200 bg-white p-3 shadow-xl">
                      <div className="mb-2 flex items-center justify-between">
                        <div className="text-lg font-bold text-slate-800">{calendarMonth.getFullYear()}年 {calendarMonth.getMonth() + 1}月</div>
                        <div className="flex items-center gap-1">
                          <button type="button" onClick={() => moveCalendarMonth(-1)} className="h-7 w-7 rounded-lg border border-slate-200 text-slate-600 hover:bg-slate-50">‹</button>
                          <button type="button" onClick={() => moveCalendarMonth(1)} className="h-7 w-7 rounded-lg border border-slate-200 text-slate-600 hover:bg-slate-50">›</button>
                        </div>
                      </div>
                      <div className="mb-2 grid grid-cols-7 gap-1 text-center text-[11px] font-semibold text-slate-500">
                        {['日', '月', '火', '水', '木', '金', '土'].map(day => <span key={day}>{day}</span>)}
                      </div>
                      <div className="grid grid-cols-7 gap-1">
                        {calendarCells.map((day, idx) => {
                          if (!day) return <span key={`empty-${idx}`} className="h-8" />;
                          const isSelected = !!selectedCalendarDate
                            && selectedCalendarDate.getFullYear() === calendarMonth.getFullYear()
                            && selectedCalendarDate.getMonth() === calendarMonth.getMonth()
                            && selectedCalendarDate.getDate() === day;
                          return (
                            <button
                              key={day}
                              type="button"
                              onClick={() => selectCalendarDate(day)}
                              className={`h-8 rounded-lg text-sm font-medium transition-colors ${
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
                      <div className="mt-3 flex justify-between">
                        <button type="button" onClick={clearCalendarDate} className="rounded-lg border border-slate-200 px-2 py-1 text-xs font-semibold text-slate-400 hover:bg-slate-50">クリア</button>
                        <button type="button" onClick={closeCalendar} className="rounded-lg border border-slate-200 px-2 py-1 text-xs font-semibold text-slate-600 hover:bg-slate-50">閉じる</button>
                      </div>
                    </div>
                  )}
                </div>
              </div>
              <div className="space-y-1">
                <label className="text-[10px] font-semibold text-slate-700 uppercase tracking-widest">全身麻酔日</label>
                <div className="relative" data-date-field="anesthesiaDate">
                  <input className={`w-full border rounded-xl px-3 py-2 text-sm focus:ring-2 focus:ring-orange-500 outline-none transition-all cursor-pointer ${getEmptyFieldToneClass(reportFields.anesthesiaDate)}`}
                    placeholder="202X年XX月XX日"
                    value={reportFields.anesthesiaDate}
                    readOnly
                    onClick={() => openCalendar('anesthesiaDate')}
                  />
                  {openDateField === 'anesthesiaDate' && (
                    <div className="absolute left-0 top-full mt-2 z-40 w-72 rounded-2xl border border-slate-200 bg-white p-3 shadow-xl">
                      <div className="mb-2 flex items-center justify-between">
                        <div className="text-lg font-bold text-slate-800">{calendarMonth.getFullYear()}年 {calendarMonth.getMonth() + 1}月</div>
                        <div className="flex items-center gap-1">
                          <button type="button" onClick={() => moveCalendarMonth(-1)} className="h-7 w-7 rounded-lg border border-slate-200 text-slate-600 hover:bg-slate-50">‹</button>
                          <button type="button" onClick={() => moveCalendarMonth(1)} className="h-7 w-7 rounded-lg border border-slate-200 text-slate-600 hover:bg-slate-50">›</button>
                        </div>
                      </div>
                      <div className="mb-2 grid grid-cols-7 gap-1 text-center text-[11px] font-semibold text-slate-500">
                        {['日', '月', '火', '水', '木', '金', '土'].map(day => <span key={day}>{day}</span>)}
                      </div>
                      <div className="grid grid-cols-7 gap-1">
                        {calendarCells.map((day, idx) => {
                          if (!day) return <span key={`empty-${idx}`} className="h-8" />;
                          const isSelected = !!selectedCalendarDate
                            && selectedCalendarDate.getFullYear() === calendarMonth.getFullYear()
                            && selectedCalendarDate.getMonth() === calendarMonth.getMonth()
                            && selectedCalendarDate.getDate() === day;
                          return (
                            <button
                              key={day}
                              type="button"
                              onClick={() => selectCalendarDate(day)}
                              className={`h-8 rounded-lg text-sm font-medium transition-colors ${
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
                      <div className="mt-3 flex justify-between">
                        <button type="button" onClick={clearCalendarDate} className="rounded-lg border border-slate-200 px-2 py-1 text-xs font-semibold text-slate-400 hover:bg-slate-50">クリア</button>
                        <button type="button" onClick={closeCalendar} className="rounded-lg border border-slate-200 px-2 py-1 text-xs font-semibold text-slate-600 hover:bg-slate-50">閉じる</button>
                      </div>
                    </div>
                  )}
                </div>
              </div>
              {/* 担当獣医師（新規：プルダウン） */}
              <div className="space-y-1">
                <label className="text-[10px] font-semibold text-slate-700 uppercase tracking-widest">担当獣医師</label>
                <div className="relative" ref={attendingVetDropdownRef}>
                  <button
                    type="button"
                    className={`w-full border rounded-xl px-3 py-2 text-base text-left focus:ring-2 focus:ring-orange-500 outline-none transition-all ${getEmptyFieldToneClass(reportFields.attendingVet)}`}
                    aria-haspopup="listbox"
                    aria-expanded={isAttendingVetDropdownOpen}
                    onClick={() => setIsAttendingVetDropdownOpen(v => !v)}
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
                      {['', '町田健吾', '江成翔馬', '神田珠希', '小林嵩', '金田七海'].map((name) => {
                        const label = name || '選択してください';
                        const isSelected = reportFields.attendingVet === name;
                        return (
                          <li key={label}>
                            <button
                              type="button"
                              className={`w-full px-3 py-2 text-left text-base transition-colors ${isSelected ? 'bg-orange-50 text-orange-700' : 'text-slate-800 hover:bg-slate-50'}`}
                              onClick={() => {
                                setReportFields(v => ({ ...v, attendingVet: name }));
                                setIsAttendingVetDropdownOpen(false);
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
              {/* 主訴（新規：テキスト入力） */}
              <div className="space-y-1">
                <label className="text-[10px] font-semibold text-slate-700 uppercase tracking-widest">主訴</label>
                <input className={`w-full border rounded-xl px-3 py-2 text-sm focus:ring-2 focus:ring-orange-500 outline-none transition-all ${getEmptyFieldToneClass(reportFields.chiefComplaint)}`}
                  id="chief-complaint-input"
                  placeholder="主な症状や主訴"
                  value={reportFields.chiefComplaint}
                  onChange={e => setReportFields(v => ({ ...v, chiefComplaint: e.target.value }))}
                />
              </div>
              <div className="space-y-1">
                <label className="text-[10px] font-semibold text-slate-700 uppercase tracking-widest">PAGE2写真区分ラベル</label>
                <div className="relative" ref={page2PhotoCategoryDropdownRef}>
                  <button
                    type="button"
                    className="w-full border border-slate-200 rounded-xl px-3 py-2 text-base text-left focus:ring-2 focus:ring-orange-500 outline-none transition-all bg-white text-slate-900"
                    aria-haspopup="listbox"
                    aria-expanded={isPage2PhotoCategoryDropdownOpen}
                    onClick={() => setIsPage2PhotoCategoryDropdownOpen(v => !v)}
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
                      ].map((item) => {
                        const isSelected = reportFields.page2PhotoCategory === item.value;
                        return (
                          <li key={`${item.value || 'empty'}-${item.label}`}>
                            <button
                              type="button"
                              className={`w-full px-3 py-2 text-left text-base transition-colors ${isSelected ? 'bg-orange-50 text-orange-700' : 'text-slate-800 hover:bg-slate-50'}`}
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
            </div>

            <div className={`h-px bg-slate-200 transition-all duration-200 ${dateDividerOffsetClass}`} />

            {/* PAGE3設定 */}
            <div className="grid grid-cols-1 sm:grid-cols-3 gap-4 bg-white border border-slate-200 rounded-2xl shadow-sm p-4 md:p-5">
              <label className="flex items-center gap-2 text-sm text-slate-700 font-semibold">
                <input
                  type="checkbox"
                  checked={showPage3}
                  onChange={e => setShowPage3(e.target.checked)}
                />
                PAGE3を追加する
              </label>

              {showPage3 && (
                <div className="space-y-1">
                  <label className="text-[10px] font-semibold text-slate-700 uppercase tracking-widest">出力順（PAGE2/PAGE3）</label>
                  <div className="relative" ref={pageOrderDropdownRef}>
                    <button
                      type="button"
                      className="w-full border border-slate-200 rounded-xl px-3 py-2 text-base text-left focus:ring-2 focus:ring-orange-500 outline-none transition-all bg-white text-slate-900"
                      aria-haspopup="listbox"
                      aria-expanded={isPageOrderDropdownOpen}
                      onClick={() => setIsPageOrderDropdownOpen(v => !v)}
                    >
                      {pageOrderMode === 'page3-page2' ? 'PAGE3 → PAGE2' : 'PAGE2 → PAGE3'}
                    </button>

                    {isPageOrderDropdownOpen && (
                      <ul
                        role="listbox"
                        className="absolute z-40 mt-1 max-h-56 w-full overflow-auto rounded-xl border border-slate-200 bg-white py-1 shadow-lg"
                      >
                        {[
                          { value: 'page2-page3', label: 'PAGE2 → PAGE3' },
                          { value: 'page3-page2', label: 'PAGE3 → PAGE2' },
                        ].map((item) => {
                          const isSelected = pageOrderMode === item.value;
                          return (
                            <li key={`${item.value}-${item.label}`}>
                              <button
                                type="button"
                                className={`w-full px-3 py-2 text-left text-base transition-colors ${isSelected ? 'bg-orange-50 text-orange-700' : 'text-slate-800 hover:bg-slate-50'}`}
                                onClick={() => {
                                  setPageOrderMode(item.value as 'page2-page3' | 'page3-page2');
                                  setIsPageOrderDropdownOpen(false);
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
                <div className="space-y-1">
                  <label className="text-[10px] font-semibold text-slate-700 uppercase tracking-widest">術後経過の配置</label>
                  <div className="relative" ref={postPlacementDropdownRef}>
                    <button
                      type="button"
                      className="w-full border border-slate-200 rounded-xl px-3 py-2 text-base text-left focus:ring-2 focus:ring-orange-500 outline-none transition-all bg-white text-slate-900"
                      aria-haspopup="listbox"
                      aria-expanded={isPostPlacementDropdownOpen}
                      onClick={() => setIsPostPlacementDropdownOpen(v => !v)}
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
                        ].map((item) => {
                          const isSelected = postPlacement === item.value;
                          return (
                            <li key={`${item.value}-${item.label}`}>
                              <button
                                type="button"
                                className={`w-full px-3 py-2 text-left text-base transition-colors ${isSelected ? 'bg-orange-50 text-orange-700' : 'text-slate-800 hover:bg-slate-50'}`}
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
            <div className="grid grid-cols-1 gap-4 bg-white border border-slate-200 rounded-2xl shadow-sm p-4 md:p-5">
              <div className="space-y-1">
                <label className="text-[10px] font-semibold text-slate-700 uppercase tracking-widest">【初診時】本文 (Page 1)</label>
                <textarea className="w-full border border-slate-200 rounded-xl px-3 py-2 text-sm min-h-[80px] focus:ring-2 focus:ring-orange-500 outline-none transition-all"
                  id="initial-textarea"
                  placeholder="初診時の所見など..."
                  value={reportFields.initialText}
                  onChange={e => setReportFields(v => ({ ...v, initialText: e.target.value }))}
                />
              </div>
              <div className="space-y-1">
                <label className="text-[10px] font-semibold text-slate-700 uppercase tracking-widest">【検査・処置内容】本文 (Page 2)</label>
                <textarea className="w-full border border-slate-200 rounded-xl px-3 py-2 text-sm min-h-[80px] focus:ring-2 focus:ring-orange-500 outline-none transition-all"
                  placeholder="実施した検査や処置の詳細..."
                  value={reportFields.procedureText}
                  onChange={e => setReportFields(v => ({ ...v, procedureText: e.target.value }))}
                />
              </div>
              <div className="space-y-1">
                <label className="text-[10px] font-semibold text-slate-700 uppercase tracking-widest">【術後経過】本文 ({showPage3 && postPlacement === 'page3' ? 'Page 3' : 'Page 2'})</label>
                <textarea className="w-full border border-slate-200 rounded-xl px-3 py-2 text-sm min-h-[80px] focus:ring-2 focus:ring-orange-500 outline-none transition-all"
                  placeholder="術後の状態や今後の予定..."
                  value={reportFields.postText}
                  onChange={e => setReportFields(v => ({ ...v, postText: e.target.value }))}
                />
              </div>
              {showPage3 && (
                <div className="space-y-1">
                  <label className="text-[10px] font-semibold text-slate-700 uppercase tracking-widest">【PAGE3】自由入力</label>
                  <textarea className="w-full border border-slate-200 rounded-xl px-3 py-2 text-sm min-h-[80px] focus:ring-2 focus:ring-orange-500 outline-none transition-all"
                    placeholder="PAGE3に出す補足テキスト..."
                    value={reportFields.page3Text || ''}
                    onChange={e => setReportFields(v => ({ ...v, page3Text: e.target.value }))}
                  />
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
        <div className="lg:col-span-12 relative w-full my-4 h-[48px]">
          <div className="absolute right-0 top-0 z-10">
            <button
              onClick={() => {
                console.log('image upload button clicked');
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
          <div className="lg:col-span-12 space-y-8">
            <LayoutControls options={options} setOptions={setOptions} />

            <div className="w-full max-w-5xl mx-auto bg-white p-7 rounded-[2.5rem] shadow-sm border border-slate-200 overflow-hidden space-y-8 relative">
              <section>
                <div className="mb-4 flex items-center gap-3 flex-wrap justify-between">
                  <div>
                    <h3 className="inline-flex items-center gap-2 px-4 py-2 rounded-full border text-sm font-semibold shadow-sm bg-orange-50 border-orange-200 text-orange-700">Page {currentPage} - STEP1:画像編集・段落選択</h3>
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
          <div className="flex items-center gap-3 mb-4 flex-wrap">
            <h3 className="inline-flex items-center gap-2 px-4 py-2 rounded-full border text-sm font-semibold shadow-sm bg-sky-50 border-sky-200 text-sky-700">
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
        <h3 className="inline-flex items-center gap-2 px-4 py-2 rounded-full border text-sm font-semibold shadow-sm bg-emerald-50 border-emerald-200 text-emerald-700 mb-4">STEP3: プレビュー</h3>
        <p className="text-[10px] text-slate-500 font-bold uppercase tracking-wider">確定済みの画像が反映されます</p>
        {pptxStatus && <p className="text-[11px] text-slate-600 font-bold mt-1">{pptxStatus}</p>}
      </div>
      <div className="flex flex-wrap gap-3">
        <button onClick={openGmailDraft}
          disabled={isCreatingDraft || !(reportFields.refHospitalEmail || '').trim()}
          title="Gmail下書きを作成し、PDFを添付します（送信はしません）"
          className="bg-slate-100 text-slate-700 border border-slate-200 px-4 py-2 rounded-xl text-[10px] font-semibold hover:bg-slate-200 transition-all flex items-center gap-2 disabled:opacity-50 disabled:cursor-not-allowed">
          {isCreatingDraft ? "作成中…" : "PDF / Gmail"}
        </button>
        <button onClick={printPdf}
          title="印刷ダイアログを開きます（PDF保存/プリンタ印刷）"
          className="bg-orange-500 hover:bg-orange-600 text-white font-semibold shadow-md rounded-xl px-4 py-2 transition-all focus:outline-none focus:ring-2 focus:ring-orange-400 text-[10px] flex items-center gap-2">
          PDF 印刷
        </button>
        <button onClick={downloadPptx} disabled={isSavingPptx}
          className="bg-slate-100 text-slate-700 border border-slate-200 px-4 py-2 rounded-xl text-[10px] font-semibold hover:bg-slate-200 flex items-center gap-2 transition-all disabled:opacity-50 disabled:cursor-not-allowed">
          {isSavingPptx ? '保存中…' : 'PPTX出力/編集'}
        </button>
      </div>
      </div>
      </div>
    </div>
  );
};


export default App;
