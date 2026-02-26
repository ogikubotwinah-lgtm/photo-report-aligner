import pptxgen from 'pptxgenjs';
import { LAYOUT } from './layout';
import type { ReportFields } from './types';

// --- general helpers --------------------------------------------------------

export function formatSection(title: string, body: string) {
  const t = (body ?? '').trim();
  if (!t) return '';
  return `【${title}】\n${t}\n`;
}

export function fitTextToBox(
  text: string,
  box: { w: number; h: number },
  basePt: number,
  minPt: number
): number {
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
}

const escapeXml = (s: string) =>
  String(s ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');

const ptToPx = (pt: number) => pt * (96 / 72);
const isJapaneseChar = (ch: string) =>
  /[\u3000-\u30FF\u4E00-\u9FFF\uFF01-\uFF60]/.test(ch);

function getMeasureContext(fontPx: number, fontFamily: string) {
  const canvas = document.createElement('canvas');
  const ctx = canvas.getContext('2d')!;
  ctx.font = `${fontPx}px ${fontFamily}`;
  return ctx;
}

function wrapTextByMeasure(
  text: string,
  maxWidthPx: number,
  ctx: CanvasRenderingContext2D
) {
  const paragraphs = text.split('\n');
  const out: string[] = [];
  paragraphs.forEach(par => {
    if (par === '') {
      out.push('');
      return;
    }
    let i = 0;
    const L = par.length;
    while (i < L) {
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
}

function clampLines(lines: string[], maxLines: number) {
  if (lines.length <= maxLines) return lines;
  const visible = lines.slice(0, maxLines);
  const last = visible[visible.length - 1] || '';
  visible[visible.length - 1] =
    last.length > 0 ? (last.replace(/\s+$/, '') + '…') : '…';
  return visible;
}

// --- SVG renderer -----------------------------------------------------------

export function buildSvgTextParts(
  pageNum: number,
  reportFields: ReportFields,
  pxPerCm: number,
  slideOffsetX: number,
  slideOffsetY: number
): string[] {
  const svgParts: string[] = [];
  const svgFontFamily = "Meiryo, 'MS PGothic', 'Noto Sans JP', sans-serif";

  if (pageNum === 1) {
    const textCfg = LAYOUT.PAGE1.TEXT as any;

    // logo box placeholder
    if (textCfg.LOGO) {
      const lx = slideOffsetX + textCfg.LOGO.x * pxPerCm;
      const ly = slideOffsetY + textCfg.LOGO.y * pxPerCm;
      const lw = textCfg.LOGO.w * pxPerCm;
      const lh = textCfg.LOGO.h * pxPerCm;
      svgParts.push(
        `  <rect x="${lx}" y="${ly}" width="${lw}" height="${lh}" fill="#f3f4f6" stroke="#e5e7eb" />`
      );
      svgParts.push(
        `  <text x="${lx + lw / 2}" y="${ly + lh / 2}" font-size="${ptToPx(
          10
        )}" fill="#666" text-anchor="middle" dominant-baseline="middle">ロゴ</text>`
      );
    }

    if (textCfg.HOSPITAL_INFO) {
      const hx = slideOffsetX + textCfg.HOSPITAL_INFO.x * pxPerCm;
      const hy =
        slideOffsetY + (textCfg.HOSPITAL_INFO.y + textCfg.HOSPITAL_INFO.h / 2) *
        pxPerCm;
      svgParts.push(
        `  <text x="${hx}" y="${hy}" font-size="${ptToPx(
          LAYOUT.FONTS.INFO_NAME
        )}" fill="#111" font-family="${svgFontFamily}" dominant-baseline="hanging">荻窪ツイン動物病院</text>`
      );
    }

    // 1) Title
    const titleX = slideOffsetX + textCfg.TITLE.x * pxPerCm;
    const titleY = slideOffsetY + textCfg.TITLE.y * pxPerCm;
    svgParts.push(
      `  <text x="${titleX}" y="${titleY}" font-size="${ptToPx(
        LAYOUT.FONTS.MAIN_TITLE
      )}" fill="#000" font-family="${svgFontFamily}" dominant-baseline="hanging">${escapeXml(
        'ご紹介患者についてのご報告'
      )}</text>`
    );

    // 2) 紹介病院 + 先生名を同じ行（横並び）
    const refHospX = slideOffsetX + textCfg.REF_HOSPITAL.x * pxPerCm;
    const refHospY = slideOffsetY + textCfg.REF_HOSPITAL.y * pxPerCm;
    const hospitality = (reportFields.refHospital || '○○動物病院') + '　' + (reportFields.refDoctor || '△△') + ' 先生';
    svgParts.push(
      `  <text x="${refHospX}" y="${refHospY}" font-size="${ptToPx(
        LAYOUT.FONTS.INFO_NAME
      )}" fill="#000" font-family="${svgFontFamily}" dominant-baseline="hanging">${escapeXml(
        hospitality
      )}</text>`
    );

    // 3) Owner & Pet names
    const ownerX = slideOffsetX + textCfg.OWNER_LASTNAME.x * pxPerCm;
    const ownerY = slideOffsetY + textCfg.OWNER_LASTNAME.y * pxPerCm;
    svgParts.push(
      `  <text x="${ownerX}" y="${ownerY}" font-size="${ptToPx(
        LAYOUT.FONTS.INFO_NAME
      )}" fill="#000" font-family="${svgFontFamily}" dominant-baseline="hanging">${escapeXml(
        '飼い主様：' + (reportFields.ownerLastName || '') + ' 様'
      )}</text>`
    );

    const petX = slideOffsetX + textCfg.PET_NAME.x * pxPerCm;
    const petY = slideOffsetY + textCfg.PET_NAME.y * pxPerCm;
    svgParts.push(
      `  <text x="${petX}" y="${petY}" font-size="${ptToPx(
        LAYOUT.FONTS.INFO_NAME
      )}" fill="#000" font-family="${svgFontFamily}" dominant-baseline="hanging">${escapeXml(
        'ペット名：' + (reportFields.petName || '')
      )}</text>`
    );

    // 4) 担当獣医師
    if (textCfg.ATTENDING_VET && reportFields.attendingVet) {
      const vetX = slideOffsetX + textCfg.ATTENDING_VET.x * pxPerCm;
      const vetY = slideOffsetY + textCfg.ATTENDING_VET.y * pxPerCm;
      svgParts.push(
        `  <text x="${vetX}" y="${vetY}" font-size="${ptToPx(
          LAYOUT.FONTS.INFO_NAME
        )}" fill="#000" font-family="${svgFontFamily}" dominant-baseline="hanging">${escapeXml(
          '担当獣医師：' + reportFields.attendingVet
        )}</text>`
      );
    }

    // 5) 定型文①
    if (textCfg.FIXED_INTRO_TEXT) {
      const introText = `この度 ${reportFields.ownerLastName || '[ 飼い主姓 ]'} 様の ${reportFields.petName || '[ ペット名 ]'} ちゃんをご紹介いただきまして ありがとうございました。拝見させていただいた結果につきまして下記の通りご報告申し上げます。`;
      const bx = slideOffsetX + textCfg.FIXED_INTRO_TEXT.x * pxPerCm;
      const by = slideOffsetY + textCfg.FIXED_INTRO_TEXT.y * pxPerCm;
      const bw = textCfg.FIXED_INTRO_TEXT.w * pxPerCm;
      const bh = textCfg.FIXED_INTRO_TEXT.h * pxPerCm;

      const fontPt = LAYOUT.FONTS.BODY_BASE;
      const fontPx = ptToPx(fontPt);
      const lineHeight = fontPx * 1.15;
      const maxLines = Math.max(1, Math.floor(bh / lineHeight));

      const ctx = getMeasureContext(fontPx, svgFontFamily);
      const wrapped = wrapTextByMeasure(introText, bw, ctx);
      const visible = clampLines(wrapped, maxLines);

      const parts: string[] = [];
      visible.forEach((ln, idx) => {
        if (idx === 0) parts.push(`<tspan x="${bx}" dy="0">${escapeXml(ln)}</tspan>`);
        else parts.push(`<tspan x="${bx}" dy="${lineHeight}">${escapeXml(ln)}</tspan>`);
      });

      svgParts.push(`  <text x="${bx}" y="${by}" font-size="${fontPx}" fill="#000" font-family="${svgFontFamily}" dominant-baseline="hanging">${parts.join('')}</text>`);
    }

    // 6) 初診日 / 全身麻酔日
    const firstVisitX = slideOffsetX + textCfg.FIRST_VISIT_DATE.x * pxPerCm;
    const firstVisitY = slideOffsetY + textCfg.FIRST_VISIT_DATE.y * pxPerCm;
    svgParts.push(
      `  <text x="${firstVisitX}" y="${firstVisitY}" font-size="${ptToPx(
        LAYOUT.FONTS.BODY_BASE
      )}" fill="#000" font-family="${svgFontFamily}" dominant-baseline="hanging">${escapeXml(
        '初診日：' + (reportFields.firstVisitDate || '')
      )}</text>`
    );

    const anesthesiaX = slideOffsetX + textCfg.ANESTHESIA_DATE.x * pxPerCm;
    const anesthesiaY = slideOffsetY + textCfg.ANESTHESIA_DATE.y * pxPerCm;
    svgParts.push(
      `  <text x="${anesthesiaX}" y="${anesthesiaY}" font-size="${ptToPx(
        LAYOUT.FONTS.BODY_BASE
      )}" fill="#000" font-family="${svgFontFamily}" dominant-baseline="hanging">${escapeXml(
        '全身麻酔日：' + (reportFields.anesthesiaDate || '')
      )}</text>`
    );

    // 7) 定型文②
    if (textCfg.FIXED_CLOSING_TEXT) {
      const closingText = `「${reportFields.chiefComplaint || '[ 主訴 ]'}」という主訴の為、拝見いたしました。`;
      const cx = slideOffsetX + textCfg.FIXED_CLOSING_TEXT.x * pxPerCm;
      const cy = slideOffsetY + textCfg.FIXED_CLOSING_TEXT.y * pxPerCm;
      const cw = textCfg.FIXED_CLOSING_TEXT.w * pxPerCm;
      const ch = textCfg.FIXED_CLOSING_TEXT.h * pxPerCm;

      const fontPt = LAYOUT.FONTS.BODY_BASE;
      const fontPx = ptToPx(fontPt);
      const lineHeight = fontPx * 1.15;
      const maxLines = Math.max(1, Math.floor(ch / lineHeight));

      const ctx = getMeasureContext(fontPx, svgFontFamily);
      const wrapped = wrapTextByMeasure(closingText, cw, ctx);
      const visible = clampLines(wrapped, maxLines);

      const parts: string[] = [];
      visible.forEach((ln, idx) => {
        if (idx === 0) parts.push(`<tspan x="${cx}" dy="0">${escapeXml(ln)}</tspan>`);
        else parts.push(`<tspan x="${cx}" dy="${lineHeight}">${escapeXml(ln)}</tspan>`);
      });

      svgParts.push(`  <text x="${cx}" y="${cy}" font-size="${fontPx}" fill="#000" font-family="${svgFontFamily}" dominant-baseline="hanging">${parts.join('')}</text>`);
    }

    // 8) Section header 【初診時】
    if (textCfg.SECTION_HEADER) {
      const headerX = slideOffsetX + textCfg.SECTION_HEADER.x * pxPerCm;
      const headerY = slideOffsetY + textCfg.SECTION_HEADER.y * pxPerCm;
      svgParts.push(
        `  <text x="${headerX}" y="${headerY}" font-size="${ptToPx(
          LAYOUT.FONTS.SECTION_HEADER
        )}" fill="#000" font-family="${svgFontFamily}" dominant-baseline="hanging" font-weight="bold">【初診時】</text>`
      );
    }

    // 9) Images header 【初診時の肉眼写真等】
    if (textCfg.IMAGES_HEADER) {
      const imgHeaderX = slideOffsetX + textCfg.IMAGES_HEADER.x * pxPerCm;
      const imgHeaderY = slideOffsetY + textCfg.IMAGES_HEADER.y * pxPerCm;
      svgParts.push(
        `  <text x="${imgHeaderX}" y="${imgHeaderY}" font-size="${ptToPx(
          LAYOUT.FONTS.SECTION_HEADER
        )}" fill="#000" font-family="${svgFontFamily}" dominant-baseline="hanging" font-weight="bold">【初診時の肉眼写真等】</text>`
      );
    }

    // 初診時本文
    const bodyCfg = textCfg.FREE_TEXT_INITIAL;
    if (bodyCfg && reportFields.initialText) {
      const bx = slideOffsetX + bodyCfg.x * pxPerCm;
      const by = slideOffsetY + bodyCfg.y * pxPerCm;
      const bw = bodyCfg.w * pxPerCm;
      const bh = bodyCfg.h * pxPerCm;

      const fontPt = LAYOUT.FONTS.BODY_BASE;
      const fontPx = ptToPx(fontPt);
      const lineHeight = fontPx * 1.15;
      const maxLines = Math.max(1, Math.floor(bh / lineHeight));

      const ctx = getMeasureContext(fontPx, svgFontFamily);
      const wrapped = wrapTextByMeasure(reportFields.initialText, bw, ctx);
      const visible = clampLines(wrapped, maxLines);

      const parts: string[] = [];
      visible.forEach((ln, idx) => {
        if (idx === 0) parts.push(`<tspan x="${bx}" dy="0">${escapeXml(ln)}</tspan>`);
        else parts.push(`<tspan x="${bx}" dy="${lineHeight}">${escapeXml(ln)}</tspan>`);
      });

      svgParts.push(`  <text x="${bx}" y="${by}" font-size="${fontPx}" fill="#000" font-family="${svgFontFamily}" dominant-baseline="hanging">${parts.join('')}</text>`);
    }
  } else if (pageNum === 2) {
    const textCfg = LAYOUT.PAGE2.TEXT as any;
    const renderBlock = (cfg: any, content: string) => {
      if (!cfg) return;
      const bx = slideOffsetX + cfg.x * pxPerCm;
      const by = slideOffsetY + cfg.y * pxPerCm;
      const bw = cfg.w * pxPerCm;
      const bh = cfg.h * pxPerCm;
      const fontPt = LAYOUT.FONTS.BODY_BASE;
      const fontPx = ptToPx(fontPt);
      const fontFamily = svgFontFamily;
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

    renderBlock(textCfg.FREE_TEXT_PROCEDURE, formatSection('検査・処置内容', reportFields.procedureText || ''));
    renderBlock(textCfg.FREE_TEXT_POSTOP, formatSection('術後経過', reportFields.postText || ''));
  }

  return svgParts;
}

// --- PPTX text -------------------------------------------------------------

export function addPptxText(
  slide: pptxgen.Slide,
  pageNum: number,
  reportFields: ReportFields
) {
  if (pageNum === 1) {
    const textCfg = LAYOUT.PAGE1.TEXT as any;

    // logo placeholder
    if (textCfg.LOGO) {
      slide.addText('ロゴ', {
        x: textCfg.LOGO.x / 2.54,
        y: textCfg.LOGO.y / 2.54,
        w: textCfg.LOGO.w / 2.54,
        h: textCfg.LOGO.h / 2.54,
        align: 'center',
        valign: 'middle',
        color: '666666',
        fill: { color: 'F3F4F6' },
        fontSize: 10,
      });
    }

    // hospital info
    if (textCfg.HOSPITAL_INFO) {
      slide.addText('荻窪ツイン動物病院', {
        x: textCfg.HOSPITAL_INFO.x / 2.54,
        y: textCfg.HOSPITAL_INFO.y / 2.54,
        w: textCfg.HOSPITAL_INFO.w / 2.54,
        h: textCfg.HOSPITAL_INFO.h / 2.54,
        fontSize: LAYOUT.FONTS.INFO_NAME,
      });
    }

    // 1) Title
    slide.addText('ご紹介患者についてのご報告', {
      x: textCfg.TITLE.x / 2.54,
      y: textCfg.TITLE.y / 2.54,
      w: textCfg.TITLE.w / 2.54,
      h: textCfg.TITLE.h / 2.54,
      fontSize: LAYOUT.FONTS.MAIN_TITLE,
      bold: true,
    });

    // 2) 紹介病院 + 先生名を同じ行（横並び）
    const hospitality = (reportFields.refHospital || '○○動物病院') + '　' + (reportFields.refDoctor || '△△') + ' 先生';
    slide.addText(hospitality, {
      x: textCfg.REF_HOSPITAL.x / 2.54,
      y: textCfg.REF_HOSPITAL.y / 2.54,
      w: textCfg.REF_HOSPITAL.w / 2.54,
      h: textCfg.REF_HOSPITAL.h / 2.54,
      fontSize: LAYOUT.FONTS.INFO_NAME,
    });

    // 3) Owner & Pet names
    slide.addText([
      { text: '飼い主様：' },
      { text: reportFields.ownerLastName || '姓', options: { underline: true } },
      { text: ' 様' }
    ] as any, {
      x: textCfg.OWNER_LASTNAME.x / 2.54,
      y: textCfg.OWNER_LASTNAME.y / 2.54,
      w: textCfg.OWNER_LASTNAME.w / 2.54,
      h: textCfg.OWNER_LASTNAME.h / 2.54,
      fontSize: LAYOUT.FONTS.INFO_NAME,
    });
    slide.addText([
      { text: 'ペット名：' },
      { text: reportFields.petName || '名前', options: { underline: true } }
    ] as any, {
      x: textCfg.PET_NAME.x / 2.54,
      y: textCfg.PET_NAME.y / 2.54,
      w: textCfg.PET_NAME.w / 2.54,
      h: textCfg.PET_NAME.h / 2.54,
      fontSize: LAYOUT.FONTS.INFO_NAME,
    });

    // 4) 担当獣医師
    if (textCfg.ATTENDING_VET && reportFields.attendingVet) {
      slide.addText('担当獣医師：' + reportFields.attendingVet, {
        x: textCfg.ATTENDING_VET.x / 2.54,
        y: textCfg.ATTENDING_VET.y / 2.54,
        w: textCfg.ATTENDING_VET.w / 2.54,
        h: textCfg.ATTENDING_VET.h / 2.54,
        fontSize: LAYOUT.FONTS.INFO_NAME,
      });
    }

    // 5) 定型文①
    if (textCfg.FIXED_INTRO_TEXT) {
      const introText = `この度 ${reportFields.ownerLastName || '[ 飼い主姓 ]'} 様の ${reportFields.petName || '[ ペット名 ]'} ちゃんをご紹介いただきまして ありがとうございました。拝見させていただいた結果につきまして下記の通りご報告申し上げます。`;
      const fitSize = fitTextToBox(introText, textCfg.FIXED_INTRO_TEXT, LAYOUT.FONTS.BODY_BASE, LAYOUT.FONTS.MIN_SIZE);
      slide.addText(introText, {
        x: textCfg.FIXED_INTRO_TEXT.x / 2.54,
        y: textCfg.FIXED_INTRO_TEXT.y / 2.54,
        w: textCfg.FIXED_INTRO_TEXT.w / 2.54,
        h: textCfg.FIXED_INTRO_TEXT.h / 2.54,
        fontSize: fitSize,
        valign: 'top',
        wrap: true,
      });
    }

    // 6) 初診日 / 全身麻酔日
    slide.addText(`初診日：${reportFields.firstVisitDate || ''}`, {
      x: textCfg.FIRST_VISIT_DATE.x / 2.54,
      y: textCfg.FIRST_VISIT_DATE.y / 2.54,
      w: textCfg.FIRST_VISIT_DATE.w / 2.54,
      h: textCfg.FIRST_VISIT_DATE.h / 2.54,
      fontSize: LAYOUT.FONTS.BODY_BASE,
    });

    slide.addText(`全身麻酔日：${reportFields.anesthesiaDate || ''}`, {
      x: textCfg.ANESTHESIA_DATE.x / 2.54,
      y: textCfg.ANESTHESIA_DATE.y / 2.54,
      w: textCfg.ANESTHESIA_DATE.w / 2.54,
      h: textCfg.ANESTHESIA_DATE.h / 2.54,
      fontSize: LAYOUT.FONTS.BODY_BASE,
    });

    // 7) 定型文②
    if (textCfg.FIXED_CLOSING_TEXT) {
      const closingText = `「${reportFields.chiefComplaint || '[ 主訴 ]'}」という主訴の為、拝見いたしました。`;
      const fitSize = fitTextToBox(closingText, textCfg.FIXED_CLOSING_TEXT, LAYOUT.FONTS.BODY_BASE, LAYOUT.FONTS.MIN_SIZE);
      slide.addText(closingText, {
        x: textCfg.FIXED_CLOSING_TEXT.x / 2.54,
        y: textCfg.FIXED_CLOSING_TEXT.y / 2.54,
        w: textCfg.FIXED_CLOSING_TEXT.w / 2.54,
        h: textCfg.FIXED_CLOSING_TEXT.h / 2.54,
        fontSize: fitSize,
        valign: 'top',
        wrap: true,
      });
    }

    // 8) Section header 【初診時】
    if (textCfg.SECTION_HEADER) {
      slide.addText('【初診時】', {
        x: textCfg.SECTION_HEADER.x / 2.54,
        y: textCfg.SECTION_HEADER.y / 2.54,
        w: textCfg.SECTION_HEADER.w / 2.54,
        h: textCfg.SECTION_HEADER.h / 2.54,
        fontSize: LAYOUT.FONTS.SECTION_HEADER,
        bold: true,
      });
    }

    // 9) Images header 【初診時の肉眼写真等】
    if (textCfg.IMAGES_HEADER) {
      slide.addText('【初診時の肉眼写真等】', {
        x: textCfg.IMAGES_HEADER.x / 2.54,
        y: textCfg.IMAGES_HEADER.y / 2.54,
        w: textCfg.IMAGES_HEADER.w / 2.54,
        h: textCfg.IMAGES_HEADER.h / 2.54,
        fontSize: LAYOUT.FONTS.SECTION_HEADER,
        bold: true,
      });
    }

    // 初診時本文
    if (textCfg.FREE_TEXT_INITIAL && reportFields.initialText) {
      const fitSize = fitTextToBox(reportFields.initialText, textCfg.FREE_TEXT_INITIAL, LAYOUT.FONTS.BODY_BASE, LAYOUT.FONTS.MIN_SIZE);
      slide.addText(reportFields.initialText, {
        x: textCfg.FREE_TEXT_INITIAL.x / 2.54,
        y: textCfg.FREE_TEXT_INITIAL.y / 2.54,
        w: textCfg.FREE_TEXT_INITIAL.w / 2.54,
        h: textCfg.FREE_TEXT_INITIAL.h / 2.54,
        fontSize: fitSize,
        valign: 'top',
        wrap: true,
      });
    }
  } else if (pageNum === 2) {
    const textCfg = LAYOUT.PAGE2.TEXT as any;
    if (textCfg.SECTION_HEADER_PROCEDURE) {
      slide.addText('【検査・処置内容】', {
        x: textCfg.SECTION_HEADER_PROCEDURE.x / 2.54,
        y: textCfg.SECTION_HEADER_PROCEDURE.y / 2.54,
        w: textCfg.SECTION_HEADER_PROCEDURE.w / 2.54,
        h: textCfg.SECTION_HEADER_PROCEDURE.h / 2.54,
        fontSize: LAYOUT.FONTS.SECTION_HEADER,
        bold: true,
      });
    }
    const procText = reportFields.procedureText || "ここに検査・処置内容が入ります...";
    const fitSizeProc = fitTextToBox(procText, textCfg.FREE_TEXT_PROCEDURE, LAYOUT.FONTS.BODY_BASE, LAYOUT.FONTS.MIN_SIZE);
    slide.addText(procText, {
      x: textCfg.FREE_TEXT_PROCEDURE.x / 2.54,
      y: textCfg.FREE_TEXT_PROCEDURE.y / 2.54,
      w: textCfg.FREE_TEXT_PROCEDURE.w / 2.54,
      h: textCfg.FREE_TEXT_PROCEDURE.h / 2.54,
      fontSize: fitSizeProc,
      align: 'left',
      valign: 'top',
      wrap: true,
    });
    if (textCfg.SECTION_HEADER_POSTOP) {
      slide.addText('【術後経過】', {
        x: textCfg.SECTION_HEADER_POSTOP.x / 2.54,
        y: textCfg.SECTION_HEADER_POSTOP.y / 2.54,
        w: textCfg.SECTION_HEADER_POSTOP.w / 2.54,
        h: textCfg.SECTION_HEADER_POSTOP.h / 2.54,
        fontSize: LAYOUT.FONTS.SECTION_HEADER,
        bold: true,
      });
    }
    const postText = reportFields.postText || "ここに術後経過が入ります...";
    const fitSizePost = fitTextToBox(postText, textCfg.FREE_TEXT_POSTOP, LAYOUT.FONTS.BODY_BASE, LAYOUT.FONTS.MIN_SIZE);
    slide.addText(postText, {
      x: textCfg.FREE_TEXT_POSTOP.x / 2.54,
      y: textCfg.FREE_TEXT_POSTOP.y / 2.54,
      w: textCfg.FREE_TEXT_POSTOP.w / 2.54,
      h: textCfg.FREE_TEXT_POSTOP.h / 2.54,
      fontSize: fitSizePost,
      align: 'left',
      valign: 'top',
      wrap: true,
    });
  }
}
