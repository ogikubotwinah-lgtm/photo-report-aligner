import React, { useState, useRef, useCallback, useMemo, useEffect } from 'react';
import type { ImageData, LayoutOptions } from "./types";
import LayoutControls from './components/LayoutControls';
import pptxgen from 'pptxgenjs';
import { LAYOUT } from './layout';
import { Reorder } from 'framer-motion';
import RowBoard from "./components/RowBoard";


const App: React.FC = () => {
  // STEP1.2 の各段落DOMを保存する箱
  const step12RowRefs = useRef<Record<number, HTMLDivElement | null>>({});

  // 今どの段落にいるか

  // STEP1.2の全体位置（任意）
  const step12Ref = useRef<HTMLDivElement | null>(null);
  const scrollStep12NextRef = useRef(false);

  // Reorder中に履歴を1回だけ記録する用
  const recordedOnceRef = useRef(false);

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
  const [step12DraggingId, setStep12DraggingId] = useState<string | null>(null);
  const [step12OverRow, setStep12OverRow] = useState<number | null>(null);
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

const reorderPendingRowByIds = useCallback((rowNum: number, orderedIds: string[]) => {


  setImages((prev: ImageData[]) => {
    const rowItems = prev.filter(i => i.row === rowNum && !i.orderConfirmed);
    const others   = prev.filter(i => !(i.row === rowNum && !i.orderConfirmed));

    const byId = new Map(rowItems.map(i => [i.id, i]));
    const nextRow: ImageData[] = [];

    // orderedIds の順に並べる
    for (const id of orderedIds) {
      const item = byId.get(id);
      if (item) nextRow.push(item);
    }

    // 万一 orderedIds に漏れがあったら末尾に追加
    for (const item of rowItems) {
      if (!orderedIds.includes(item.id)) nextRow.push(item);
    }

    return [...others, ...nextRow];
  });
}, [recordHistory, setImages]);






const confirmAllPendingRows = useCallback(() => {
  recordHistory();
  setImages((prev: ImageData[]) =>
    prev.map(img => (img.row > 0 && !img.orderConfirmed ? { ...img, orderConfirmed: true } : img))
  );
}, [recordHistory, setImages]);

  const hasUnconfirmedImages = useMemo(() => images.some(img => img.row > 0 && !img.orderConfirmed), [images]);

useEffect(() => {
  const handleEnter = (e: KeyboardEvent) => {
    // 入力中にEnterを奪わない（安全）
    const tag = (e.target as HTMLElement | null)?.tagName;
    const isTyping = tag === 'INPUT' || tag === 'TEXTAREA';
    if (isTyping) return;

    if (e.key === 'Enter' && hasUnconfirmedImages) {
      confirmAllPendingRows();
    }
  };

  window.addEventListener('keydown', handleEnter);
  return () => window.removeEventListener('keydown', handleEnter);
}, [hasUnconfirmedImages, confirmAllPendingRows]);



  const unassignedImages = useMemo(() => images.filter(img => img.row === 0), [images]);
  


  const confirmedImages = useMemo(() => images.filter(img => img.row > 0 && img.orderConfirmed), [images]);
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
      const isFew = (pageNum === 2 && count === 3) ? false : count <= 3;
      let rowH;
      rowH = (containerW - (count - 1) * baseSpacing) / totalAR;

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

  const scale = options.containerWidth / dims.maxW;
  const fullSlideW = LAYOUT.SLIDE.WIDTH_CM * scale;
  const fullSlideH = LAYOUT.SLIDE.HEIGHT_CM * scale;

  const offsetX = dims.left * scale;
  const offsetY = dims.startY * scale;

  const svgParts: string[] = [];

  rowResults.forEach((row) => {
    row.images.forEach((item: any) => {
      const absX = offsetX + item.x;
      const absY = offsetY + item.y;

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
  currentPage
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
      const pageConfirmed = pageImages.filter(img => img.row > 0 && img.orderConfirmed);
      if (pageConfirmed.length === 0 && pageNum === 2) return;
      const slide = pptx.addSlide();
      buildSlide(slide, pageNum, pageConfirmed);
    });

    pptx.writeFile({ fileName: `Photo_Report_A4.pptx` });
  };

  return (
    <div className="min-h-screen bg-slate-50 pb-20 font-sans">
      <RowBoard images={images} setImages={setImages} />
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
        <div className="max-w-7xl mx-auto flex justify-between items-center">
          <div className="flex items-center gap-3">
            <div className="bg-orange-600 p-2.5 rounded-2xl text-white shadow-lg shadow-orange-100">
              <svg className="w-6 h-6" fill="currentColor" viewBox="0 0 24 24"><path d="M19 2H5C3.9 2 3 2.9 3 4V20C3 21.1 3.9 22 5 22H19C20.1 22 21 21.1 21 20V4C21 2.9 20.1 2 19 2M15.5 14H14V17.5H10V14H8.5L12 10.5L15.5 14M16.5 9H7.5V7H16.5V9Z"/></svg>
            </div>
            <div>
              <h1 className="text-xl font-black text-slate-900 tracking-tight leading-none">Photo Report Aligner</h1>
              <p className="text-[10px] text-orange-600 font-black uppercase tracking-[0.2em] mt-1">Smart Sequential Workflow</p>
            </div>
          </div>
          
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
              <svg className="w-5 h-5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="2.5" d="M12 4v16m8-8H4" /></svg>
              画像を追加
            </button>
          </div>
          <input type="file" ref={fileInputRef} onChange={handleFileUpload} multiple accept="image/*" className="hidden" />
        </div>
      </nav>

      <main className="max-w-7xl mx-auto px-6 mt-10 grid grid-cols-1 lg:grid-cols-12 gap-10">
        {/* 左カラム */}
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
                  <div key={img.id} className="p-4 bg-slate-50/50 border border-slate-100 rounded-[2rem] hover:border-orange-200 transition-all shadow-sm flex items-center gap-4 group">
                    <div className="relative flex-shrink-0 cursor-pointer" onClick={() => setPreviewId(img.id)}>
                      <img src={img.dataUrl} style={{ transform: `rotate(${img.rotation}deg)` }} className="w-14 h-14 object-contain rounded-xl border border-slate-100 bg-white shadow-inner" />
                    </div>
                    <div className="flex-1 min-w-0">
                      <div className="flex justify-between items-center mb-2">
                        <p className="text-[9px] font-bold text-slate-400 truncate pr-2 uppercase tracking-tight leading-none">{img.name}</p>
                        <div className="flex gap-1.5 items-center">
                          <button onClick={() => rotateImage(img.id, 'left')} className="w-7 h-7 flex items-center justify-center rounded-lg bg-indigo-600 text-white hover:bg-indigo-700 border border-indigo-500 transition-all shadow-sm">
                            <svg className="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="3" d="M4 4v5h.582m15.356 2A8.001 8.001 0 004.582 9m0 0H9" /></svg>
                          </button>
                          <button onClick={() => rotateImage(img.id, 'right')} className="w-7 h-7 flex items-center justify-center rounded-lg bg-indigo-600 text-white hover:bg-indigo-700 border border-indigo-500 transition-all shadow-sm">
                            <svg className="w-3.5 h-3.5" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="3" d="M20 4v5h-.582m-15.356 2A8.001 8.001 0 0119.418 9m0 0H15" /></svg>
                          </button>
                          <button onClick={() => removeImage(img.id)} className="h-7 rounded-lg bg-red-50 text-red-500 hover:bg-red-600 hover:text-white flex items-center justify-center transition-all border border-red-100 shadow-sm"><svg className="w-4 h-4" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path strokeLinecap="round" strokeLinejoin="round" strokeWidth="3" d="M19 7l-.867 12.142A2 2 0 0116.138 21H7.862a2 2 0 01-1.995-1.858L5 7m5 4v6m4-6v6m1-10V4a1 1 0 00-1-1h-4a1 1 0 00-1 1v3M4 7h16" /></svg></button>
                        </div>
                      </div>
                      <div className="grid grid-cols-4 gap-1">
  {[1, 2, 3, 4].map(num => (
    <button
      key={num}
      onClick={() => {
        scrollStep12NextRef.current = true;
        updateImageRow(img.id, num);
      }}
      className="h-8 rounded-lg bg-white text-orange-600 text-[10px] font-black border border-slate-200 hover:bg-orange-600 hover:text-white transition-all shadow-sm active:scale-95"
    >
      {num}
    </button>
  ))}
</div>

                    </div>
                  </div>
                ))}

                {hasUnconfirmedImages && (
  <div className="mt-6 space-y-6">
    <div ref={step12Ref} />
    <div className="font-black text-slate-800 text-sm">
      STEP 1.2: 挿入位置の決定
    </div>

    <div className="space-y-8">
      {[1, 2, 3, 4].map((rowNum) => {
        const rowImages = images.filter(i => i.row === rowNum && !i.orderConfirmed);
        const orderedIds = rowImages.map(i => i.id);
        const isDropTarget = step12OverRow === rowNum && !!step12DraggingId;
        if (rowImages.length === 0 && !step12DraggingId) return null;
        return (
          <div
            key={rowNum}
            ref={(el) => { step12RowRefs.current[rowNum] = el; }}
            className="space-y-3"
            onDragOver={(e) => { if (!step12DraggingId) return; e.preventDefault(); setStep12OverRow(rowNum); }}
            onDragEnter={() => { if (step12DraggingId) setStep12OverRow(rowNum); }}
            onDragLeave={(e) => { if (!e.currentTarget.contains(e.relatedTarget as Node)) setStep12OverRow(null); }}
            onDrop={(e) => {
              e.preventDefault();
              const id = e.dataTransfer.getData("text/plain");
              if (id) updateImageRow(id, rowNum);
              setStep12DraggingId(null);
              setStep12OverRow(null);
            }}
            style={{
              outline: isDropTarget ? '2px solid #4f46e5' : undefined,
              background: isDropTarget ? '#eef2ff' : undefined,
              borderRadius: 16,
              padding: isDropTarget ? 6 : undefined,
              transition: 'all 0.12s',
            }}
          >
            <div className="flex items-center gap-2">
              <div className="bg-indigo-100 text-indigo-700 px-3 py-1 rounded-lg text-[10px] font-black uppercase tracking-wider">
                段落 {rowNum}
              </div>
              <div className="h-px bg-slate-100 flex-1" />
            </div>
            <Reorder.Group
              key={`row-${rowNum}`}
              axis="x"
              values={orderedIds}
              onReorder={(newOrder) => reorderPendingRowByIds(rowNum, newOrder)}
              className="flex flex-wrap gap-3 p-3 bg-slate-50/50 rounded-2xl border border-slate-100/50"
              style={{ minHeight: isDropTarget && rowImages.length === 0 ? 72 : undefined }}
            >
              {rowImages.map((img, idx) => (
                <Reorder.Item
                  key={img.id}
                  value={img.id}
                  className="bg-white border rounded-2xl shadow-sm select-none"
                  whileDrag={{
                    scale: 1.05,
                    boxShadow: "0px 16px 40px rgba(0,0,0,0.20)",
                    zIndex: 50,
                  }}
                  onClick={() => setPreviewId(img.id)}
                  onPointerDown={() => {
                    if (!recordedOnceRef.current) {
                      recordHistory();
                      recordedOnceRef.current = true;
                    }
                  }}
                  onPointerUp={() => { recordedOnceRef.current = false; }}
                >
                  {/* 段落間移動ハンドル */}
                  <div
                    draggable
                    onPointerDownCapture={(e) => e.stopPropagation()}
                    onMouseDownCapture={(e) => e.stopPropagation()}
                    onClick={(e) => e.stopPropagation()}
                    onDragStart={(e) => {
                      e.stopPropagation();
                      setStep12DraggingId(img.id);
                      e.dataTransfer.setData("text/plain", img.id);
                      e.dataTransfer.effectAllowed = "move";
                    }}
                    onDragEnd={() => {
                      setStep12DraggingId(null);
                      setStep12OverRow(null);
                    }}
                    className="py-1 px-2 rounded-t-2xl bg-slate-100 hover:bg-indigo-100 text-[10px] text-slate-400 hover:text-indigo-500 font-bold text-center transition-colors"
                    style={{ cursor: 'grab' }}
                    title="ドラッグして別段落へ移動"
                  >
                    ⠿ 移動
                  </div>
                  <div className="relative p-2" style={{ cursor: step12DraggingId ? 'default' : 'grab' }}>
                    <img
                      src={img.dataUrl}
                      style={{ transform: `rotate(${img.rotation}deg)` }}
                      className="w-16 h-16 object-contain rounded-xl bg-white pointer-events-none"
                    />
                    <div className="absolute -top-2 -left-2 w-6 h-6 rounded-lg bg-indigo-600 text-white text-[10px] font-black flex items-center justify-center border-2 border-white shadow-sm">
                      {idx + 1}
                    </div>
                  </div>
                </Reorder.Item>
              ))}
            </Reorder.Group>
          </div>
        );
      })}
    </div>

    <div className="pt-4 border-t border-slate-100 flex justify-end">
      <button
        onClick={confirmAllPendingRows}
        className="px-6 py-2.5 rounded-xl bg-orange-600 text-white text-[10px] font-black hover:bg-orange-700 shadow-lg shadow-orange-900/10 active:scale-95 transition-all flex items-center gap-2"
      >
        <svg
          className="w-4 h-4"
          fill="none"
          stroke="currentColor"
          viewBox="0 0 24 24"
        >
          <path
            strokeLinecap="round"
            strokeLinejoin="round"
            strokeWidth="3"
            d="M5 13l4 4L19 7"
          />
        </svg>
        並び変え確定
      </button>
    </div>
  </div>
)}

              </div>
            </section>
          </div>
        </div>

        {/* 右カラム */}
        <div className="lg:col-span-7 space-y-6">
          {/* 報告書データ入力フォーム */}
          <div className="bg-white p-7 rounded-[2.5rem] shadow-sm border border-slate-200 space-y-6">
            <div>
              <h3 className="font-black text-slate-800 text-base mb-1">報告書データ入力</h3>
              <p className="text-[10px] text-slate-400 font-bold uppercase tracking-wider">PPTXの各項目に反映される内容を編集します</p>
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
                  />
                </div>
                <div className="space-y-1">
                  <label className="text-[10px] font-black text-slate-500 uppercase tracking-widest">全身麻酔日</label>
                  <input className="w-full border border-slate-200 rounded-xl px-3 py-2 text-sm focus:ring-2 focus:ring-orange-500 outline-none transition-all"
                    placeholder="202X年XX月XX日"
                    value={reportFields.anesthesiaDate}
                    onChange={e => setReportFields(v => ({ ...v, anesthesiaDate: e.target.value }))}
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
            </div>
          </div>

          {/* プレビューセクション */}
          <div className="bg-white p-6 rounded-[2.5rem] shadow-sm border border-slate-200 overflow-hidden">
            <div className="flex items-center justify-between mb-4">
              <div>
                <h3 className="font-black text-slate-800 text-base mb-1">プレビュー</h3>
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
        </div>
      <h2 style={{ marginTop: 20 }}>▼ 段落ドラッグ移動（RowBoard）</h2>
<div style={{ padding: 8 }}>images: {images.length} 枚</div>
<RowBoard images={images} setImages={setImages} rows={4} />
      </main>
    </div>
  );
};


export default App;
