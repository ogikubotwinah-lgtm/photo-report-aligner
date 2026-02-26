import React, { useState, useRef, useCallback, useMemo, useEffect } from 'react';
import type { ImageData, LayoutOptions } from "./types";
import LayoutControls from './components/LayoutControls';
import pptxgen from 'pptxgenjs';
import { LAYOUT } from './layout';
import RowBoard from "./components/RowBoard";
import TemplatePicker from './components/TemplatePicker';
import TEMPLATES from './data/templates';


// Minimal date normalizer used by onBlur handlers in this file.
function normalizeJapaneseDate(input: string): string {
  if (!input) return input;
  const s = String(input).trim();

  // If already in Japanese year/month/day form, leave as-is
  if (s.includes('年')) return s;

  // Keep only digits for compact checks
  const onlyDigits = s.replace(/\D+/g, '');

  if (/^\d{8}$/.test(onlyDigits)) {
    const y = parseInt(onlyDigits.slice(0, 4), 10);
    const m = parseInt(onlyDigits.slice(4, 6), 10);
    const d = parseInt(onlyDigits.slice(6, 8), 10);
    const date = new Date(y, m - 1, d);
    if (date.getFullYear() === y && date.getMonth() === m - 1 && date.getDate() === d) {
      return `${y}年${m}月${d}日`;
    }
    return input;
  }

  if (/^\d{6}$/.test(onlyDigits)) {
    const y = parseInt(onlyDigits.slice(0, 4), 10);
    const m = parseInt(onlyDigits.slice(4, 6), 10);
    if (m >= 1 && m <= 12) {
      return `${y}年${m}月1日`;
    }
    return input;
  }

  if (/^\d{4}$/.test(onlyDigits)) {
    const y = parseInt(onlyDigits.slice(0, 4), 10);
    return `${y}年1月1日`;
  }

  return input;
}


const App: React.FC = () => {
  // 今どの段落にいるか

  const rowBoardRef = useRef<HTMLDivElement | null>(null);

  // ページ管理用の状態
  const [currentPage, setCurrentPage] = useState<number>(1);
  const [allPagesImages, setAllPagesImages] = useState<Record<number, ImageData[]>>({ 1: [], 2: [] });
  const [allPagesHistory, setAllPagesHistory] = useState<Record<number, ImageData[][]>>({ 1: [], 2: [] });
  const [previewId, setPreviewId] = useState<string | null>(null);




  // 報告書テキスト入力ステート
  const [reportFields, setReportFields] = useState({
    reportDate: new Date().toLocaleDateString('ja-JP', {
  year: 'numeric',
  month: 'long',
  day: 'numeric'
}),

    refHospital: '',
    refDoctor: '',
    ownerLastName: '',
    petName: '',
    firstVisitDate: '',
    anesthesiaDate: '',
    initialText: '',
    procedureText: '',
    postText: ''
  });

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

  /**
   * テキストが枠内に収まるようにフォントサイズを計算するユーティリティ
   */
  const fitTextToBox = (text: string, box: { w: number, h: number }, basePt: number, minPt: number): number => {
    if (!text) return basePt;
    const chars = text.length;
    let currentSize = basePt;
    const boxArea = box.w * box.h;
    while (currentSize > minPt) {
      const charSizeCm = (currentSize / 72) * 2.54;
      const estimatedRequiredArea = chars * (charSizeCm * charSizeCm) * 1.2;
      if (estimatedRequiredArea <= boxArea) {
        return currentSize;
      }
      currentSize -= 0.5;
    }
    return minPt;
  };

  const [options, setOptions] = useState<LayoutOptions>({
    spacing: 1,
    padding: 0,
    targetHeight: 250,
    containerWidth: 650,
    backgroundColor: '#ffffff'
  });

  const [copyStatus, setCopyStatus] = useState<string>('SVGコピー');

  // テンプレ挿入の undo 用（直前の挿入を1回だけ戻す）
  const [lastInsert, setLastInsert] = useState<
    | { field: 'initialText' | 'procedureText' | 'postText'; prevValue: string }
    | null
  >(null);
  const [page1Confirmed, setPage1Confirmed] = useState(false);
  const [page2Confirmed, setPage2Confirmed] = useState(false);
  const isCurrentPageConfirmed = (currentPage === 1 && page1Confirmed) || (currentPage === 2 && page2Confirmed);
  const fileInputRef = useRef<HTMLInputElement>(null);

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
  if (page1Confirmed || page2Confirmed) {
    rowBoardRef.current?.scrollIntoView({ behavior: 'smooth', block: 'start' });
  }
}, [page1Confirmed, page2Confirmed]);

  const unassignedImages = useMemo(() => images.filter(img => img.row === 0), [images]);



  // ✅ 「確定」不要：段落(row>0)に入っている画像は常にプレビュー/出力対象にする
const confirmedImages = useMemo(() => images.filter(img => img.row > 0), [images]);
  const displayRows = useMemo(() => [1, 2, 3, 4].map(r => confirmedImages.filter(img => img.row === r)), [confirmedImages]);

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

  const getCalculatedLayout = useCallback(() => calculateLayoutForAnyPage(displayRows, currentPage), [calculateLayoutForAnyPage, displayRows, currentPage]);

const calculateSvgData = useCallback(() => {
  const { rowResults } = getCalculatedLayout();
  const dims = getPageDimensions(currentPage);

  // scale and pxPerCm are based on the full slide width so layout.ts cm coords map directly
  const scale = options.containerWidth / LAYOUT.SLIDE.WIDTH_CM;
  const pxPerCm = scale;
  const fullSlideW = LAYOUT.SLIDE.WIDTH_CM * scale; // === options.containerWidth
  const fullSlideH = LAYOUT.SLIDE.HEIGHT_CM * scale;

  // slideOffset positions the whole slide within the preview (usually 0)
  const slideOffsetX = 0;
  const slideOffsetY = 0;

  // image area offset: only images use the page's IMAGES.LEFT / START_Y
  const imgCfg = currentPage === 1 ? LAYOUT.PAGE1.IMAGES : LAYOUT.PAGE2.IMAGES;
  const imageOffsetX = slideOffsetX + imgCfg.LEFT * pxPerCm;
  const imageOffsetY = slideOffsetY + imgCfg.START_Y * pxPerCm;

  const svgParts: string[] = [];

  rowResults.forEach((row) => {
    row.images.forEach((item: any) => {
      const absX = imageOffsetX + item.x;
      const absY = imageOffsetY + item.y;

      const cx = absX + item.w / 2;
      const cy = absY + item.h / 2;

      const isPortrait = item.img.rotation === 90 || item.img.rotation === 270;
      let drawW = item.w, drawH = item.h;
      if (isPortrait) { drawW = item.h; drawH = item.w; }

      const transform =
        item.img.rotation !== 0
          ? `transform="rotate(${item.img.rotation} ${cx} ${cy})"`
          : '';

      svgParts.push(
        `  <image x="${cx - drawW / 2}" y="${cy - drawH / 2}" width="${drawW}" height="${drawH}" href="${item.img.dataUrl}" ${transform} />`
      );
    });
  });

  const escapeXml = (s: string) =>
    String(s ?? '')
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');

  // Helper: convert pt -> px for SVG
  const ptToPx = (pt: number) => pt * (96 / 72);
  const svgFontFamily = "Meiryo, 'MS PGothic', 'Noto Sans JP', sans-serif";

  // Canvas-based text measurement helpers (for more accurate wrap)
  const getMeasureContext = (fontPx: number, fontFamily: string) => {
    const canvas = document.createElement('canvas');
    const ctx = canvas.getContext('2d')!;
    ctx.font = `${fontPx}px ${fontFamily}`;
    return ctx;
  };

  const isJapaneseChar = (ch: string) => /[\u3000-\u30FF\u4E00-\u9FFF\uFF01-\uFF60]/.test(ch);

  const wrapTextByMeasure = (text: string, maxWidthPx: number, ctx: CanvasRenderingContext2D) => {
    const paragraphs = text.split('\n');
    const out: string[] = [];
    paragraphs.forEach(par => {
      if (par === '') { out.push(''); return; }
      let i = 0;
      const L = par.length;
      while (i < L) {
        // If japanese char at i, do char-by-char
        if (isJapaneseChar(par[i])) {
          let line = '';
          while (i < L) {
            const next = line + par[i];
            const w = ctx.measureText(next).width;
            if (w > maxWidthPx && line.length > 0) break;
            line = next;
            i += 1;
          }
          out.push(line);
        } else {
          // ascii/latin: try to take words
          const rest = par.slice(i);
          const words = rest.split(/(\s+)/);
          let line = '';
          let consumed = 0;
          for (let wi = 0; wi < words.length; wi++) {
            const wToken = words[wi];
            const candidate = line + wToken;
            const width = ctx.measureText(candidate).width;
            if (width > maxWidthPx && line.length > 0) break;
            line = candidate;
            consumed += wToken.length;
          }
          if (line.length === 0) {
            // fallback: single char
            line = par[i];
            i += 1;
          } else {
            i += consumed;
          }
          out.push(line);
        }
      }
    });
    return out;
  };

  const clampLines = (lines: string[], maxLines: number) => {
    if (lines.length <= maxLines) return lines;
    const visible = lines.slice(0, maxLines);
    const last = visible[visible.length - 1] || '';
    visible[visible.length - 1] = last.length > 0 ? (last.replace(/\s+$/, '') + '…') : '…';
    return visible;
  };

  // Render fixed header (logo + hospital info) and text fields using LAYOUT constants
  if (currentPage === 1) {
  const textCfg = LAYOUT.PAGE1.TEXT as any;

    // Logo placeholder
    if (textCfg.LOGO) {
      const lx = slideOffsetX + textCfg.LOGO.x * pxPerCm;
      const ly = slideOffsetY + textCfg.LOGO.y * pxPerCm;
      const lw = textCfg.LOGO.w * pxPerCm;
      const lh = textCfg.LOGO.h * pxPerCm;
      svgParts.push(`  <rect x="${lx}" y="${ly}" width="${lw}" height="${lh}" fill="#f3f4f6" stroke="#e5e7eb" />`);
    svgParts.push(`  <text x="${lx + lw / 2}" y="${ly + lh / 2}" font-size="${ptToPx(10)}" fill="#666" text-anchor="middle" dominant-baseline="middle">ロゴ</text>`);
    }

    // Hospital info (fixed)
    if (textCfg.HOSPITAL_INFO) {
      const hx = slideOffsetX + textCfg.HOSPITAL_INFO.x * pxPerCm;
      const hy = slideOffsetY + (textCfg.HOSPITAL_INFO.y + textCfg.HOSPITAL_INFO.h / 2) * pxPerCm;
      svgParts.push(`  <text x="${hx}" y="${hy}" font-size="${ptToPx(LAYOUT.FONTS.INFO_NAME)}" fill="#111" font-family="${svgFontFamily}" dominant-baseline="hanging">荻窪ツイン動物病院</text>`);
    }

    // Header fields (report date, title, recipient, doctor, owner, pet)
    const reportDateX = slideOffsetX + textCfg.REPORT_DATE.x * pxPerCm;
    const reportDateY = slideOffsetY + textCfg.REPORT_DATE.y * pxPerCm;
    svgParts.push(`  <text x="${reportDateX}" y="${reportDateY}" font-size="${ptToPx(LAYOUT.FONTS.BODY_BASE)}" fill="#000" font-family="${svgFontFamily}" text-anchor="end" dominant-baseline="hanging">${escapeXml('報告日：' + (reportFields.reportDate || ''))}</text>`);

    const titleX = slideOffsetX + textCfg.TITLE.x * pxPerCm;
    const titleY = slideOffsetY + textCfg.TITLE.y * pxPerCm;
    svgParts.push(`  <text x="${titleX}" y="${titleY}" font-size="${ptToPx(LAYOUT.FONTS.MAIN_TITLE)}" fill="#000" font-family="${svgFontFamily}" dominant-baseline="hanging">${escapeXml('ご紹介患者についてのご報告')}</text>`);

    const refHospX = slideOffsetX + textCfg.REF_HOSPITAL.x * pxPerCm;
    const refHospY = slideOffsetY + textCfg.REF_HOSPITAL.y * pxPerCm;
    svgParts.push(`  <text x="${refHospX}" y="${refHospY}" font-size="${ptToPx(LAYOUT.FONTS.INFO_NAME)}" fill="#000" font-family="${svgFontFamily}" dominant-baseline="hanging">${escapeXml((reportFields.refHospital || '○○動物病院') + '　御中')}</text>`);

    const refDocX = slideOffsetX + textCfg.REF_DOCTOR.x * pxPerCm;
    const refDocY = slideOffsetY + textCfg.REF_DOCTOR.y * pxPerCm;
    svgParts.push(`  <text x="${refDocX}" y="${refDocY}" font-size="${ptToPx(LAYOUT.FONTS.INFO_NAME)}" fill="#000" font-family="${svgFontFamily}" dominant-baseline="hanging">${escapeXml((reportFields.refDoctor || '△△') + ' 先生')}</text>`);

    const ownerX = slideOffsetX + textCfg.OWNER_LASTNAME.x * pxPerCm;
    const ownerY = slideOffsetY + textCfg.OWNER_LASTNAME.y * pxPerCm;
    svgParts.push(`  <text x="${ownerX}" y="${ownerY}" font-size="${ptToPx(LAYOUT.FONTS.INFO_NAME)}" fill="#000" font-family="${svgFontFamily}" dominant-baseline="hanging">${escapeXml('飼い主様：' + (reportFields.ownerLastName || '姓') + ' 様')}</text>`);

    const petX = slideOffsetX + textCfg.PET_NAME.x * pxPerCm;
    const petY = slideOffsetY + textCfg.PET_NAME.y * pxPerCm;
    svgParts.push(`  <text x="${petX}" y="${petY}" font-size="${ptToPx(LAYOUT.FONTS.INFO_NAME)}" fill="#000" font-family="${svgFontFamily}" dominant-baseline="hanging">${escapeXml('ペット名：' + (reportFields.petName || '名前'))}</text>`);

    // Combined body block into FREE_TEXT_INITIAL box
    const bodyCfg = textCfg.FREE_TEXT_INITIAL;
    if (bodyCfg) {
      const bx = slideOffsetX + bodyCfg.x * pxPerCm;
      const by = slideOffsetY + bodyCfg.y * pxPerCm;
      const bw = bodyCfg.w * pxPerCm;
      const bh = bodyCfg.h * pxPerCm;

      const fontPt = LAYOUT.FONTS.BODY_BASE;
      const fontPx = ptToPx(fontPt);
      const fontFamily = "Meiryo, 'MS PGothic', 'Noto Sans JP', sans-serif";
      const lineHeight = fontPx * 1.15;
      const maxLines = Math.max(1, Math.floor(bh / lineHeight));

      const combined = ['【初診時】', reportFields.initialText || '', '【検査・処置内容】', reportFields.procedureText || '', '【術後経過】', reportFields.postText || ''].join('\n');
      const ctx = getMeasureContext(fontPx, fontFamily);
      const wrapped = wrapTextByMeasure(combined, bw, ctx);
      const visible = clampLines(wrapped, maxLines);

      // Build tspans; first line bold label if present
      const parts: string[] = [];
      visible.forEach((ln, idx) => {
        if (idx === 0) parts.push(`<tspan x="${bx}" dy="0" font-weight="700">${escapeXml(ln)}</tspan>`);
        else parts.push(`<tspan x="${bx}" dy="${lineHeight}">${escapeXml(ln)}</tspan>`);
      });

      svgParts.push(`  <text x="${bx}" y="${by}" font-size="${fontPx}" fill="#000" font-family="${fontFamily}" dominant-baseline="hanging">${parts.join('')}</text>`);
    }
  } else if (currentPage === 2) {
    // Page 2: use PAGE2 text cfg boxes
    const textCfg = LAYOUT.PAGE2.TEXT as any;
    const bodyProc = textCfg.FREE_TEXT_PROCEDURE;
    const bodyPost = textCfg.FREE_TEXT_POSTOP;

    const renderBlock = (cfg: any, content: string) => {
      if (!cfg) return;
      const bx = slideOffsetX + cfg.x * pxPerCm;
      const by = slideOffsetY + cfg.y * pxPerCm;
      const bw = cfg.w * pxPerCm;
      const bh = cfg.h * pxPerCm;
      const fontPt = LAYOUT.FONTS.BODY_BASE;
      const fontPx = ptToPx(fontPt);
      const fontFamily = "Meiryo, 'MS PGothic', 'Noto Sans JP', sans-serif";
      const lineHeight = fontPx * 1.15;
      const maxLines = Math.max(1, Math.floor(bh / lineHeight));
      const raw = content || '';
      const ctx = getMeasureContext(fontPx, fontFamily);
      const wrapped = wrapTextByMeasure(raw, bw, ctx);
      const visible = clampLines(wrapped, maxLines);
      const parts: string[] = [];
      visible.forEach((ln, idx) => {
        if (idx === 0) parts.push(`<tspan x="${bx}" dy="0">${escapeXml(ln)}</tspan>`);
        else parts.push(`<tspan x="${bx}" dy="${lineHeight}">${escapeXml(ln)}</tspan>`);
      });
      svgParts.push(`  <text x="${bx}" y="${by}" font-size="${fontPx}" fill="#000" font-family="${fontFamily}" dominant-baseline="hanging">${parts.join('')}</text>`);
    };

    renderBlock(bodyProc, reportFields.procedureText || '');
    renderBlock(bodyPost, reportFields.postText || '');
  }

  const svgCode = `
<svg xmlns="http://www.w3.org/2000/svg" width="100%" height="100%" viewBox="0 0 ${fullSlideW} ${fullSlideH}">
  <rect width="100%" height="100%" fill="${options.backgroundColor}" />
${svgParts.join('\n')}
</svg>`.trim();

  return { svgCode, fullSlideW, fullSlideH };
}, [
  getCalculatedLayout,
  options.backgroundColor,
  options.containerWidth,
  getPageDimensions,
  currentPage,
  reportFields
]);


  const copyToClipboard = () => {
    const { svgCode } = calculateSvgData();
    navigator.clipboard.writeText(svgCode).then(() => {
      setCopyStatus('コピー完了!');
      setTimeout(() => setCopyStatus('SVGコピー'), 2000);
    });
  };

  const downloadPptx = async (e: React.MouseEvent) => {
    e.preventDefault();
    const pptx = new pptxgen();
    pptx.defineLayout({
      name: LAYOUT.SLIDE.NAME,
      width: LAYOUT.SLIDE.WIDTH_CM / 2.54,
      height: LAYOUT.SLIDE.HEIGHT_CM / 2.54
    });
    pptx.layout = LAYOUT.SLIDE.NAME;

    const buildSlide = (slide: pptxgen.Slide, pageNum: number, imagesData: ImageData[]) => {
      slide.background = { fill: options.backgroundColor.replace('#', '') };
      const pageDims = getPageDimensions(pageNum);
      const pxToInch = (pageDims.maxW / 2.54) / options.containerWidth;
      const startXInch = pageDims.left / 2.54;
      const startYInch = pageDims.startY / 2.54;
      const pageDisplayRows = [1, 2, 3, 4].map(r => imagesData.filter(img => img.row === r));
      const { rowResults } = calculateLayoutForAnyPage(pageDisplayRows, pageNum);
      rowResults.forEach(row => {
        row.images.forEach((item: any) => {
          slide.addImage({
            data: item.img.dataUrl,
            x: startXInch + (item.x * pxToInch),
            y: startYInch + (item.y * pxToInch),
            w: item.w * pxToInch,
            h: item.h * pxToInch,
            rotate: item.img.rotation
          });
        });
      });

      if (pageNum === 1) {
        const textCfg = LAYOUT.PAGE1.TEXT;
        // Fixed logo placeholder
        if (textCfg.LOGO) {
          slide.addText('ロゴ', {
            x: textCfg.LOGO.x / 2.54, y: textCfg.LOGO.y / 2.54, w: textCfg.LOGO.w / 2.54, h: textCfg.LOGO.h / 2.54,
            align: 'center', valign: 'middle', color: '666666', fill: { color: 'F3F4F6' }, fontSize: 10
          });
        }
        // Fixed hospital info
        if (textCfg.HOSPITAL_INFO) {
          slide.addText('荻窪ツイン動物病院', {
            x: textCfg.HOSPITAL_INFO.x / 2.54, y: textCfg.HOSPITAL_INFO.y / 2.54, w: textCfg.HOSPITAL_INFO.w / 2.54, h: textCfg.HOSPITAL_INFO.h / 2.54,
            fontSize: LAYOUT.FONTS.INFO_NAME
          });
        }
        slide.addText(`報告日：${reportFields.reportDate}`, {
          x: textCfg.REPORT_DATE.x / 2.54, y: textCfg.REPORT_DATE.y / 2.54, w: textCfg.REPORT_DATE.w / 2.54, h: textCfg.REPORT_DATE.h / 2.54,
          fontSize: LAYOUT.FONTS.BODY_BASE, align: 'right'
        });
        slide.addText('ご紹介患者についてのご報告', {
          x: textCfg.TITLE.x / 2.54, y: textCfg.TITLE.y / 2.54, w: textCfg.TITLE.w / 2.54, h: textCfg.TITLE.h / 2.54,
          fontSize: LAYOUT.FONTS.MAIN_TITLE, align: 'center', bold: true
        });
        slide.addText(([
          { text: reportFields.refHospital || '○○動物病院', options: { underline: true } },
          { text: ' 御中' }
        ] as any), {
          x: textCfg.REF_HOSPITAL.x / 2.54, y: textCfg.REF_HOSPITAL.y / 2.54, w: textCfg.REF_HOSPITAL.w / 2.54, h: textCfg.REF_HOSPITAL.h / 2.54,
          fontSize: LAYOUT.FONTS.INFO_NAME
        });
        slide.addText(([
          { text: reportFields.refDoctor || '△△', options: { underline: true } },
          { text: ' 先生' }
        ] as any), {
          x: textCfg.REF_DOCTOR.x / 2.54, y: textCfg.REF_DOCTOR.y / 2.54, w: textCfg.REF_DOCTOR.w / 2.54, h: textCfg.REF_DOCTOR.h / 2.54,
          fontSize: LAYOUT.FONTS.INFO_NAME
        });
        slide.addText(([
          { text: '飼い主様：' },
          { text: reportFields.ownerLastName || '姓', options: { underline: true } },
          { text: ' 様' }
        ] as any), {
          x: textCfg.OWNER_LASTNAME.x / 2.54, y: textCfg.OWNER_LASTNAME.y / 2.54, w: textCfg.OWNER_LASTNAME.w / 2.54, h: textCfg.OWNER_LASTNAME.h / 2.54,
          fontSize: LAYOUT.FONTS.INFO_NAME
        });
        slide.addText(([
          { text: 'ペット名：' },
          { text: reportFields.petName || '名前', options: { underline: true } }
        ] as any), {
          x: textCfg.PET_NAME.x / 2.54, y: textCfg.PET_NAME.y / 2.54, w: textCfg.PET_NAME.w / 2.54, h: textCfg.PET_NAME.h / 2.54,
          fontSize: LAYOUT.FONTS.INFO_NAME
        });
        slide.addText('【初診時】', {
          x: textCfg.SECTION_HEADER.x / 2.54, y: textCfg.SECTION_HEADER.y / 2.54, w: textCfg.SECTION_HEADER.w / 2.54, h: textCfg.SECTION_HEADER.h / 2.54,
          fontSize: LAYOUT.FONTS.SECTION_HEADER, bold: true
        });
        slide.addText(`初診日：${reportFields.firstVisitDate || '202X年XX月XX日'}`, {
          x: textCfg.FIRST_VISIT_DATE.x / 2.54, y: textCfg.FIRST_VISIT_DATE.y / 2.54, w: textCfg.FIRST_VISIT_DATE.w / 2.54, h: textCfg.FIRST_VISIT_DATE.h / 2.54,
          fontSize: LAYOUT.FONTS.BODY_BASE
        });
        slide.addText(`全身麻酔日：${reportFields.anesthesiaDate || '202X年XX月XX日'}`, {
          x: textCfg.ANESTHESIA_DATE.x / 2.54, y: textCfg.ANESTHESIA_DATE.y / 2.54, w: textCfg.ANESTHESIA_DATE.w / 2.54, h: textCfg.ANESTHESIA_DATE.h / 2.54,
          fontSize: LAYOUT.FONTS.BODY_BASE
        });
        const initialText = reportFields.initialText || "ここに初診時の内容が入ります...";
        const fitSize = fitTextToBox(initialText, textCfg.FREE_TEXT_INITIAL, LAYOUT.FONTS.BODY_BASE, LAYOUT.FONTS.MIN_SIZE);
        slide.addText(initialText, {
          x: textCfg.FREE_TEXT_INITIAL.x / 2.54, y: textCfg.FREE_TEXT_INITIAL.y / 2.54, w: textCfg.FREE_TEXT_INITIAL.w / 2.54, h: textCfg.FREE_TEXT_INITIAL.h / 2.54,
          fontSize: fitSize, align: 'left', valign: 'top', wrap: true
        });
      } else if (pageNum === 2) {
        const textCfg = LAYOUT.PAGE2.TEXT;
        slide.addText('【検査・処置内容】', {
          x: textCfg.SECTION_HEADER_PROCEDURE.x / 2.54, y: textCfg.SECTION_HEADER_PROCEDURE.y / 2.54, w: textCfg.SECTION_HEADER_PROCEDURE.w / 2.54, h: textCfg.SECTION_HEADER_PROCEDURE.h / 2.54,
          fontSize: LAYOUT.FONTS.SECTION_HEADER, bold: true
        });
        const procText = reportFields.procedureText || "ここに検査・処置内容が入ります...";
        const fitSizeProc = fitTextToBox(procText, textCfg.FREE_TEXT_PROCEDURE, LAYOUT.FONTS.BODY_BASE, LAYOUT.FONTS.MIN_SIZE);
        slide.addText(procText, {
          x: textCfg.FREE_TEXT_PROCEDURE.x / 2.54, y: textCfg.FREE_TEXT_PROCEDURE.y / 2.54, w: textCfg.FREE_TEXT_PROCEDURE.w / 2.54, h: textCfg.FREE_TEXT_PROCEDURE.h / 2.54,
          fontSize: fitSizeProc, align: 'left', valign: 'top', wrap: true
        });
        slide.addText('【術後経過】', {
          x: textCfg.SECTION_HEADER_POSTOP.x / 2.54, y: textCfg.SECTION_HEADER_POSTOP.y / 2.54, w: textCfg.SECTION_HEADER_POSTOP.w / 2.54, h: textCfg.SECTION_HEADER_POSTOP.h / 2.54,
          fontSize: LAYOUT.FONTS.SECTION_HEADER, bold: true
        });
        const postText = reportFields.postText || "ここに術後経過が入ります...";
        const fitSizePost = fitTextToBox(postText, textCfg.FREE_TEXT_POSTOP, LAYOUT.FONTS.BODY_BASE, LAYOUT.FONTS.MIN_SIZE);
        slide.addText(postText, {
          x: textCfg.FREE_TEXT_POSTOP.x / 2.54, y: textCfg.FREE_TEXT_POSTOP.y / 2.54, w: textCfg.FREE_TEXT_POSTOP.w / 2.54, h: textCfg.FREE_TEXT_POSTOP.h / 2.54,
          fontSize: fitSizePost, align: 'left', valign: 'top', wrap: true
        });
      }
    };

    [1, 2].forEach(pageNum => {
      const pageImages = allPagesImages[pageNum];
      // ✅ 「確定」不要：row>0 をそのまま採用
const pageConfirmed = pageImages.filter(img => img.row > 0);
      if (pageConfirmed.length === 0 && pageNum === 2) return;
      const slide = pptx.addSlide();
      buildSlide(slide, pageNum, pageConfirmed);
    });

    pptx.writeFile({ fileName: `Photo_Report_A4.pptx` });
  };

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
              <svg className="w-10 h-10" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2" d="M6 18L18 6M6 6l12 12"/></svg>
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
          <div>
            <h3 className="font-black text-slate-800 text-base mb-1">報告書データ入力</h3>
            <p className="text-[10px] text-slate-400 font-bold uppercase tracking-wider">現在の段落配置が自動で反映されます</p>
          </div>

          <div className="space-y-6">
            {/* 基本情報グリッド */}
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-4">
              <div className="space-y-1">
                <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest">報告日</label>
                <input className="w-full border border-slate-200 rounded-xl px-3 py-2 text-sm focus:ring-2 focus:ring-orange-500 outline-none transition-all"
                  placeholder="2026年2月16日"
                  value={reportFields.reportDate}
                  onChange={e => setReportFields(v => ({ ...v, reportDate: e.target.value }))}
                  onBlur={e => setReportFields(v => ({ ...v, reportDate: normalizeJapaneseDate(e.target.value) }))}
                />
              </div>
              <div className="space-y-1">
                <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest">紹介病院名</label>
                <input className="w-full border border-slate-200 rounded-xl px-3 py-2 text-sm focus:ring-2 focus:ring-orange-500 outline-none transition-all"
                  placeholder="○○動物病院"
                  value={reportFields.refHospital}
                  onChange={e => setReportFields(v => ({ ...v, refHospital: e.target.value }))}
                />
              </div>
              <div className="space-y-1">
                <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest">先生名</label>
                <input className="w-full border border-slate-200 rounded-xl px-3 py-2 text-sm focus:ring-2 focus:ring-orange-500 outline-none transition-all"
                  placeholder="△△ 先生"
                  value={reportFields.refDoctor}
                  onChange={e => setReportFields(v => ({ ...v, refDoctor: e.target.value }))}
                />
              </div>
              <div className="space-y-1">
                <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest">飼い主姓</label>
                <input className="w-full border border-slate-200 rounded-xl px-3 py-2 text-sm focus:ring-2 focus:ring-orange-500 outline-none transition-all"
                  placeholder="山田"
                  value={reportFields.ownerLastName}
                  onChange={e => setReportFields(v => ({ ...v, ownerLastName: e.target.value }))}
                />
              </div>
              <div className="space-y-1">
                <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest">ペット名</label>
                <input className="w-full border border-slate-200 rounded-xl px-3 py-2 text-sm focus:ring-2 focus:ring-orange-500 outline-none transition-all"
                  placeholder="タロウ"
                  value={reportFields.petName}
                  onChange={e => setReportFields(v => ({ ...v, petName: e.target.value }))}
                />
              </div>
              <div className="space-y-1">
                <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest">初診日</label>
                <input className="w-full border border-slate-200 rounded-xl px-3 py-2 text-sm focus:ring-2 focus:ring-orange-500 outline-none transition-all"
                  placeholder="202X年XX月XX日"
                  value={reportFields.firstVisitDate}
                  onChange={e => setReportFields(v => ({ ...v, firstVisitDate: e.target.value }))}
                  onBlur={e => setReportFields(v => ({ ...v, firstVisitDate: normalizeJapaneseDate(e.target.value) }))}
                />
              </div>
              <div className="space-y-1">
                <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest">全身麻酔日</label>
                <input className="w-full border border-slate-200 rounded-xl px-3 py-2 text-sm focus:ring-2 focus:ring-orange-500 outline-none transition-all"
                  placeholder="202X年XX月XX日"
                  value={reportFields.anesthesiaDate}
                  onChange={e => setReportFields(v => ({ ...v, anesthesiaDate: e.target.value }))}
                  onBlur={e => setReportFields(v => ({ ...v, anesthesiaDate: normalizeJapaneseDate(e.target.value) }))}
                />
              </div>
            </div>

            {/* 自由記載エリア */}
            <div className="grid grid-cols-1 gap-4">
              <div className="space-y-1">
                <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest">【初診時】本文 (Page 1)</label>
                <textarea className="w-full border border-slate-200 rounded-xl px-3 py-2 text-sm min-h-[80px] focus:ring-2 focus:ring-orange-500 outline-none transition-all"
                  placeholder="初診時の所見など..."
                  value={reportFields.initialText}
                  onChange={e => setReportFields(v => ({ ...v, initialText: e.target.value }))}
                />
              </div>
              <div className="space-y-1">
                <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest">【検査・処置内容】本文 (Page 2)</label>
                <textarea className="w-full border border-slate-200 rounded-xl px-3 py-2 text-sm min-h-[80px] focus:ring-2 focus:ring-orange-500 outline-none transition-all"
                  placeholder="実施した検査や処置の詳細..."
                  value={reportFields.procedureText}
                  onChange={e => setReportFields(v => ({ ...v, procedureText: e.target.value }))}
                />
              </div>
              <div className="space-y-1">
                <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest">【術後経過】本文 (Page 2)</label>
                <textarea className="w-full border border-slate-200 rounded-xl px-3 py-2 text-sm min-h-[80px] focus:ring-2 focus:ring-orange-500 outline-none transition-all"
                  placeholder="術後の状態や今後の予定..."
                  value={reportFields.postText}
                  onChange={e => setReportFields(v => ({ ...v, postText: e.target.value }))}
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
        <div className="lg:col-span-12 flex justify-between items-center">
          <div className="flex bg-slate-100 p-1.5 rounded-2xl border border-slate-200 gap-1.5">
            {[1, 2].map(p => (
              <button
                key={p}
                onClick={() => setCurrentPage(p)}
                className={`flex items-center gap-2.5 px-6 py-2 rounded-xl text-xs font-black transition-all border ${
                  currentPage === p
                  ? 'bg-white text-violet-600 border-violet-100 shadow-md scale-105'
                  : 'text-slate-400 border-transparent hover:text-slate-500 hover:bg-slate-50'
                }`}
              >
                <svg className={`w-4 h-4 ${currentPage === p ? 'text-violet-500' : 'text-slate-300'}`} fill="none" stroke="currentColor" viewBox="0 0 24 24">
                  <path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2.5" d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z" />
                </svg>
                <span className="uppercase tracking-[0.2em]">PAGE {p}</span>
              </button>
            ))}
          </div>
          <div className="flex gap-3">
            <button onClick={() => fileInputRef.current?.click()} className="bg-indigo-600 hover:bg-indigo-700 text-white px-6 py-3 rounded-2xl font-bold shadow-lg active:scale-95 flex items-center gap-2 transition-all">
              <svg className="w-6 h-6" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2.5" d="M12 4v16m8-8H4" /></svg>
              画像を追加
            </button>
          </div>
          <input type="file" ref={fileInputRef} onChange={handleFileUpload} multiple accept="image/*" className="hidden" />
        </div>

        {/* 左カラム - 確定前のみ表示 */}
        {!isCurrentPageConfirmed && (
        <div className="lg:col-span-5 space-y-8">
          <LayoutControls options={options} setOptions={setOptions} />

          <div className="bg-white p-7 rounded-[2.5rem] shadow-sm border border-slate-200 overflow-hidden space-y-8 relative">
            <section>
              <div className="mb-6 flex justify-between items-start">
                <div>
                  <h3 className="font-black text-slate-800 text-base mb-1">Page {currentPage} - STEP 1.1: 段落の選択</h3>
                  <p className="text-[10px] text-slate-400 font-bold uppercase tracking-wider">画像の下にある段落番号を選んでください</p>
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
        }
      }}
      className={`h-12 rounded-lg text-lg font-semibold transition-all shadow-sm active:scale-95 border ${
        img.row === num
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

        {/* 右カラム - 確定前のみ表示 */}
        {!isCurrentPageConfirmed && (
        <div className="lg:col-span-7 space-y-6">
          {/* プレビューセクション */}
          <div className="bg-white p-6 rounded-[2.5rem] shadow-sm border border-slate-200 overflow-hidden">
            <div className="p-4 bg-slate-100/50 flex justify-center items-center overflow-hidden">
              {confirmedImages.length > 0 ? (
                <div className="shadow-[0_15px_30px_rgba(0,0,0,0.1)] bg-white border border-slate-300 relative"
                  style={{ height: '480px', aspectRatio: '19.05 / 27.517' }}>
                  <div className="w-full h-full"
                    dangerouslySetInnerHTML={{ __html: calculateSvgData().svgCode }} />
                </div>
              ) : (
                <div className="text-center py-32 opacity-40 text-slate-600 font-black tracking-widest text-[10px]">
                  画像を段落に割り当てると自動でプレビューに反映されます
                </div>
              )}
            </div>
          </div>
        </div>
        )}
        {/* 段落ドラッグ移動 */}
        <div ref={rowBoardRef} className="lg:col-span-12 bg-white p-7 rounded-[2.5rem] shadow-sm border border-slate-200 space-y-4">
          <div>
            <h3 className="font-black text-slate-800 text-base mb-1">
              {isCurrentPageConfirmed ? `Page ${currentPage} - STEP 1.1: 段落` : '段落ドラッグ移動'}
            </h3>
            <p className="text-[10px] text-slate-400 font-bold uppercase tracking-wider">「段落移動」ハンドルをドラッグして段落間を移動できます</p>
          </div>
          <RowBoard images={images} setImages={setImages} rows={4} />
        </div>
        {/* 確定後のプレビュー（RowBoard の下） */}
        {isCurrentPageConfirmed && (
          <div className="lg:col-span-12 bg-white p-6 rounded-[2.5rem] shadow-sm border border-slate-200 overflow-hidden">
            <div className="p-4 bg-slate-100/50 flex justify-center items-center overflow-hidden">
              {confirmedImages.length > 0 ? (
                <div className="shadow-[0_15px_30px_rgba(0,0,0,0.1)] bg-white border border-slate-300 relative"
                  style={{ height: '480px', aspectRatio: '19.05 / 27.517' }}>
                  <div className="w-full h-full"
                    dangerouslySetInnerHTML={{ __html: calculateSvgData().svgCode }} />
                </div>
              ) : (
                <div className="text-center py-32 opacity-40 text-slate-600 font-black tracking-widest text-[10px]">
                  画像を段落に割り当てて順序確定するとプレビューできます
                </div>
              )}
            </div>
          </div>
        )}
      </main>

      {/* Sticky bottom bar */}
      <div className="sticky bottom-0 z-50 bg-white/90 backdrop-blur border-t border-slate-200 p-3">
        <div className="max-w-7xl mx-auto flex items-center justify-between px-6">
          <div>
            <h3 className="font-black text-slate-800 text-base">プレビュー</h3>
            <p className="text-[10px] text-slate-400 font-bold uppercase tracking-wider">確定済みの画像が反映されます</p>
          </div>
          <div className="flex gap-3">
            <button onClick={copyToClipboard} disabled={confirmedImages.length === 0}
              className="bg-slate-800 text-white border border-slate-700 px-5 py-2 rounded-xl text-[10px] font-black hover:bg-slate-700 disabled:opacity-30 transition-all flex items-center gap-2">
              {copyStatus}
            </button>
            <button onClick={downloadPptx} disabled={confirmedImages.length === 0}
              className="bg-orange-600 text-white px-5 py-2 rounded-xl text-[10px] font-black hover:bg-orange-700 shadow-xl shadow-orange-900/20 flex items-center gap-2 transition-all disabled:opacity-30">
              PPTX保存
            </button>
          </div>
        </div>
      </div>
    </div>
  );
};


export default App;
