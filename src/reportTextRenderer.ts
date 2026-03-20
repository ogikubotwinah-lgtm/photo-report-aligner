import pptxgen from 'pptxgenjs';
import { LAYOUT } from './layout';
import type { ReportFields } from './types';
import logoDataUrl from './assets/logo.png?inline';
import { cmToInch } from './utils/units';
import { getStampUrlByVetName } from './stamps';

const INTRO_LINE_SPACING_MULTIPLE = 1.25;
const PAGE1_INITIAL_SECTION_HEADER_Y_CM = 11.07;
const PAGE1_INITIAL_BODY_Y_CM = 11.8;
const PAGE1_IMAGES_HEADER_AFTER_BODY_LINES = 1;
const PAGE1_IMAGE_START_EXTRA_CM = 0.0;

function getPage1InitialBodyMetricsCm(
  textCfg: any,
  reportFields: ReportFields,
  fontFamily: string
) {
  const bodyCfg = textCfg.FREE_TEXT_INITIAL;

  const fontPt = LAYOUT.FONTS.BODY_BASE;
  const fontPx = ptToPx(fontPt);
  const lineHeightPx = fontPx * 1.15;
  const lineHeightCm = (lineHeightPx / 96) * 2.54;

  if (!bodyCfg) {
    return { lineHeightCm, bodyLineCount: 1 };
  }

  const bodyWidthPx = (bodyCfg.w / 2.54) * 96;
  const bodyHeightPx = (bodyCfg.h / 2.54) * 96;
  const maxLines = Math.max(1, Math.floor(bodyHeightPx / lineHeightPx));

  const initialText = String(reportFields.initialText || '');
  if (initialText.trim() === '') {
    return { lineHeightCm, bodyLineCount: 1 };
  }

  const ctx = getMeasureContext(fontPx, fontFamily);
  const wrapped = wrapTextByMeasure(initialText, bodyWidthPx, ctx);
  const visible = clampLines(wrapped, maxLines);

  return {
    lineHeightCm,
    bodyLineCount: Math.max(1, visible.length),
  };
}

function getPage1ClosingTextLineCount(
  textCfg: any,
  reportFields: ReportFields,
  fontFamily: string
): number {
  const closingCfg = textCfg.FIXED_CLOSING_TEXT;
  if (!closingCfg) return 1;
  const closingText = `「${reportFields.chiefComplaint || '[ 主訴 ]'}」という主訴の為、拝見いたしました。`;
  const fontPx = ptToPx(LAYOUT.FONTS.BODY_BASE);
  const lineHeight = fontPx * 1.15;
  const cwPx = (closingCfg.w / 2.54) * 96;
  const chPx = (closingCfg.h / 2.54) * 96;
  const maxLines = Math.max(1, Math.floor(chPx / lineHeight));
  const ctx = getMeasureContext(fontPx, fontFamily);
  const wrapped = wrapTextByMeasure(closingText, cwPx, ctx);
  const visible = clampLines(wrapped, maxLines);
  return Math.max(1, visible.length);
}

function getPage1InitialBlockLayout(
  textCfg: any,
  reportFields: ReportFields,
  fontFamily: string
) {
  const bodyCfg = textCfg.FREE_TEXT_INITIAL;
  const sectionHeaderCfg = textCfg.SECTION_HEADER;
  const imagesHeaderCfg = textCfg.IMAGES_HEADER;

  const baseSectionHeaderY =
    typeof sectionHeaderCfg?.y === 'number'
      ? PAGE1_INITIAL_SECTION_HEADER_Y_CM
      : sectionHeaderCfg?.y ?? 0;
  const baseBodyY =
    typeof bodyCfg?.y === 'number' ? PAGE1_INITIAL_BODY_Y_CM : bodyCfg?.y ?? 0;

  const { lineHeightCm, bodyLineCount } = getPage1InitialBodyMetricsCm(
    textCfg,
    reportFields,
    fontFamily
  );

  // 主訴（定型文②）の実描画行数が 2行以上のとき、その分だけ初診セクションを下げる
  const closingLineCount = getPage1ClosingTextLineCount(textCfg, reportFields, fontFamily);
  const extraClosingLineCm = Math.max(0, closingLineCount - 1) * lineHeightCm;

  const sectionHeaderY = baseSectionHeaderY + extraClosingLineCm;
  const bodyY = baseBodyY + extraClosingLineCm;

  const imagesHeaderY =
    bodyY + lineHeightCm * (bodyLineCount + PAGE1_IMAGES_HEADER_AFTER_BODY_LINES);

  const imagesHeaderHeightCm = imagesHeaderCfg?.h ?? 0;
  const imageStartY =
    imagesHeaderY + imagesHeaderHeightCm + PAGE1_IMAGE_START_EXTRA_CM;

  return { sectionHeaderY, bodyY, imagesHeaderY, imageStartY };
}

export function getPage1ImageStartYcm(reportFields: ReportFields) {
  const textCfg = LAYOUT.PAGE1.TEXT as any;
  const fontFamily = "Meiryo, 'MS PGothic', 'Noto Sans JP', sans-serif";
  return getPage1InitialBlockLayout(textCfg, reportFields, fontFamily).imageStartY;
}


// 医療情報ブロック左端はレイアウト設定内の HOSPITAL_INFO.x を使うため
// 直接定義は不要。




// --- general helpers --------------------------------------------------------

export function formatSection(title: string, body: string) {
  const t = (body ?? '').trim();
  if (!t) return '';
  return `【${title}】\n${t}\n`;
}

function indentEachLine(text: string) {
  return String(text ?? '')
    .split('\n')
    .map((line) => (line === '' ? line : `　${line}`))
    .join('\n');
}

function buildPage3Body(reportFields: ReportFields, options?: RenderOptions) {
  const page3FreeText = String(reportFields.page3Text || '').trim();
  const shouldIndentPostOnPage3 =
    options?.showPage3 &&
    options?.postPlacement === 'page3' &&
    (options?.indentPostOnPage3 ?? true);
  const postBodyForPage3 = shouldIndentPostOnPage3
    ? indentEachLine(reportFields.postText || '')
    : (reportFields.postText || '');
  const postSection =
    options?.showPage3 && options?.postPlacement === 'page3' && postBodyForPage3.trim() !== ''
      ? `【術後経過】\n${postBodyForPage3}\n`
      : '';

  const parts: string[] = [];
  if (page3FreeText) parts.push(page3FreeText);
  if (postSection) parts.push(postSection);

  // お礼文を最終ページ（PAGE3）に追加（5行空けはテキスト内の空行で表現）
  const thankYou = getThankYouBody(reportFields);
  if (thankYou) {
    parts.push('\n\n\n\n' + thankYou);
  }

  return parts.join('\n\n');
}

function getThankYouBody(reportFields: ReportFields) {
  const thankYouTextType = String((reportFields as any)?.thankYouTextType || 'first-time');
  if (thankYouTextType === 'existing') {
    return '平素よりご紹介いただき、誠にありがとうございます。\nご不明な点などございましたら、ご遠慮なくお知らせください。\n今後とも何卒よろしくお願い申し上げます。';
  }
  return 'この度は初めてご紹介のご縁を賜り、誠にありがとうございます。\nご不明な点などございましたら、ご遠慮なくお知らせください。\n今後とも何卒よろしくお願い申し上げます。';
}

function buildPage1DateLines(reportFields: ReportFields) {
  const ordered = [
    { label: '初診日', value: String(reportFields?.firstVisitDate || '') },
    { label: '鎮静日', value: String((reportFields as any)?.sedationDate || '') },
    { label: '全身麻酔日', value: String(reportFields?.anesthesiaDate || '') },
  ];

  return ordered.filter(item => String(item?.value || '').trim() !== '');
}

function getPage1DividerYcmByDateCount(dateCount: number, fallbackYcm: number) {
  if (dateCount === 1) return 9.35;
  if (dateCount === 2) return 9.79;
  if (dateCount >= 3) return 10.2;
  return fallbackYcm;
}

function getPage1ClosingTextYcmByDateCount(dateCount: number, fallbackYcm: number) {
  return dateCount >= 3 ? 10.3 : fallbackYcm;
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

const getVetSealChar = (fullName: string) => {
  if (!fullName) return '印';
  const surnameMap: Record<string, string> = {
    '町田健吾': '田',
    '江成翔馬': '成',
    '神田珠希': '田',
    '小林嵩': '林',
    '金田七海': '田'
  };
  if (surnameMap[fullName]) return surnameMap[fullName];
  return fullName.slice(-1);
};

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
        let jpCharCount = 0;
        const fontPx = parseFloat(ctx.font) || 16;
        while (i < L) {
          const next = line + par[i];
          const measured = ctx.measureText(next).width;
          const nextJpCount = jpCharCount + (isJapaneseChar(par[i]) ? 1 : 0);
          const w = Math.max(measured, nextJpCount * fontPx);
          if (w > maxWidthPx && line.length > 0) break;
          line = next;
          jpCharCount = nextJpCount;
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

const normalizeJapaneseSentence = (text: string) =>
  String(text ?? '')
    .replace(/[ \t]+/g, '')
    .replace(/　+/g, '')
    .trim();

// --- SVG renderer -----------------------------------------------------------

type RenderOptions = {
  showPage3?: boolean;
  postPlacement?: 'page2' | 'page3';
  indentPostOnPage3?: boolean;
  page3ImagesBottomYcm?: number;
  page2ImagesBottomYcm?: number;
  previewYOffsets?: Partial<import('./App').PreviewYOffsetMap>;
};

export function buildSvgTextParts(
  pageNum: number,
  reportFields: ReportFields,
  pxPerCm: number,
  slideOffsetX: number,
  slideOffsetY: number,
  options?: RenderOptions
): string[] {
  const svgParts: string[] = [];
  const svgFontFamily = "Meiryo, 'MS PGothic', 'Noto Sans JP', sans-serif";

  if (pageNum === 1) {
    const textCfg = LAYOUT.PAGE1.TEXT as any;
    const initialBlockLayout = getPage1InitialBlockLayout(
      textCfg,
      reportFields,
      svgFontFamily
    );

    // ロゴ画像（固定サイズ：layout.ts の w/h をそのまま使う）
// ロゴ画像（固定座標＆固定サイズ：layout.ts の x/y/w/h をそのまま使う）
// ※ PPTX と同じ保険（w/h が無い・0 の時は実測値にフォールバック）
if (textCfg.LOGO) {
  const wCm =
    typeof textCfg.LOGO.w === 'number' && textCfg.LOGO.w > 0 ? textCfg.LOGO.w : 1.45;
  const hCm =
    typeof textCfg.LOGO.h === 'number' && textCfg.LOGO.h > 0 ? textCfg.LOGO.h : 1.26;

  const lx = slideOffsetX + textCfg.LOGO.x * pxPerCm;
  const ly = slideOffsetY + textCfg.LOGO.y * pxPerCm;
  const lw = wCm * pxPerCm;
  const lh = hCm * pxPerCm;

  svgParts.push(
    `  <image x="${lx}" y="${ly}" width="${lw}" height="${lh}" href="${logoDataUrl}" xlink:href="${logoDataUrl}" />`
  );
}

    // 報告日（右上）
    if (textCfg.REPORT_DATE) {
      const dx = slideOffsetX + textCfg.REPORT_DATE.x * pxPerCm;
      const dy = slideOffsetY + textCfg.REPORT_DATE.y * pxPerCm;
      svgParts.push(
        `  <text x="${dx + textCfg.REPORT_DATE.w * pxPerCm}" y="${dy}" font-size="${ptToPx(
          LAYOUT.FONTS.REPORT_DATE
        )}" fill="#000" font-family="${svgFontFamily}" dominant-baseline="hanging" text-anchor="end">${escapeXml(
          reportFields.reportDate || ''
        )}</text>`
      );
    }
// 病院名＋住所メール（左端は HOSPITAL_INFO.x プロパティに統一）
if (textCfg.HOSPITAL_INFO) {
  const x = slideOffsetX + textCfg.HOSPITAL_INFO.x * pxPerCm;
  const y = slideOffsetY + textCfg.HOSPITAL_INFO.y * pxPerCm;
  svgParts.push(
    `  <text x="${x}" y="${y}" font-size="${ptToPx(LAYOUT.FONTS.INFO_NAME)}" fill="#111" font-family="${svgFontFamily}" dominant-baseline="hanging">荻窪ツイン動物病院</text>`
  );
}

if (textCfg.HOSPITAL_ADDR) {
  const ax = slideOffsetX + textCfg.HOSPITAL_ADDR.x * pxPerCm;
  const ay = slideOffsetY + textCfg.HOSPITAL_ADDR.y * pxPerCm;

  svgParts.push(
    `  <text x="${ax}" y="${ay}" font-size="${ptToPx(
      LAYOUT.FONTS.INFO_DETAIL
    )}" fill="#111" font-family="${svgFontFamily}" dominant-baseline="hanging">東京都杉並区上荻1-23-18</text>`
  );
}

if (textCfg.HOSPITAL_EMAIL) {
  const ex = slideOffsetX + textCfg.HOSPITAL_EMAIL.x * pxPerCm;
  const ey = slideOffsetY + textCfg.HOSPITAL_EMAIL.y * pxPerCm;

  svgParts.push(
    `  <text x="${ex}" y="${ey}" font-size="${ptToPx(
      LAYOUT.FONTS.INFO_DETAIL
    )}" fill="#111" font-family="${svgFontFamily}" dominant-baseline="hanging">ogikubotwinah@gmail.com</text>`
  );
}

    // 1) Title（中央寄せ）
    const titleX = slideOffsetX + textCfg.TITLE.x * pxPerCm;
    const titleY = slideOffsetY + textCfg.TITLE.y * pxPerCm;
    const titleW = textCfg.TITLE.w * pxPerCm;
    svgParts.push(
      `  <text x="${titleX + titleW / 2}" y="${titleY}" font-size="${ptToPx(
        LAYOUT.FONTS.MAIN_TITLE
      )}" fill="#000" font-family="${svgFontFamily}" dominant-baseline="hanging" text-anchor="middle" font-weight="bold">${escapeXml(
        'ご紹介患者についてのご報告'
      )}</text>`
    );

    // 2) 紹介病院 + 先生名 (1行、13.5pt HGP明朝B) with individual underlines
    const refHospX = slideOffsetX + textCfg.REF_HOSPITAL.x * pxPerCm;
    const refHospY = slideOffsetY + textCfg.REF_HOSPITAL.y * pxPerCm;
    const hospName = reportFields.refHospital || '○○動物病院';
    const docName = (reportFields.refDoctor || '').trim();
    if (docName) {
      svgParts.push(
        `  <text x="${refHospX}" y="${refHospY}" font-size="${ptToPx(
          13.5
        )}" fill="#000" font-family="HGP明朝B" dominant-baseline="hanging" text-anchor="start" >` +
          `<tspan text-decoration="underline">${escapeXml(hospName)}</tspan>` +
          `<tspan>　</tspan>` +
          `<tspan text-decoration="underline">${escapeXml(docName)}</tspan>` +
          `<tspan> 先生</tspan>` +
        `</text>`
      );
    } else {
      svgParts.push(
        `  <text x="${refHospX}" y="${refHospY}" font-size="${ptToPx(
          13.5
        )}" fill="#000" font-family="HGP明朝B" dominant-baseline="hanging" text-anchor="start" >` +
          `<tspan text-decoration="underline">${escapeXml(hospName)}</tspan>` +
        `</text>`
      );
    }

    // 3) 担当獣医師（ラベル＋名前下線）
    if (reportFields.attendingVet) {
      if (textCfg.ATTENDING_VET_LABEL) {
        const vetLabelX = slideOffsetX + textCfg.ATTENDING_VET_LABEL.x * pxPerCm;
        const vetLabelY = slideOffsetY + textCfg.ATTENDING_VET_LABEL.y * pxPerCm;
        svgParts.push(
          `  <text x="${vetLabelX}" y="${vetLabelY}" font-size="${ptToPx(
            11
          )}" fill="#000" font-family="HGP明朝B" dominant-baseline="hanging">担当獣医師：</text>`
        );
      }
      if (textCfg.ATTENDING_VET_NAME) {
        const vetNameX = slideOffsetX + textCfg.ATTENDING_VET_NAME.x * pxPerCm;
        const vetNameY = slideOffsetY + textCfg.ATTENDING_VET_NAME.y * pxPerCm;
        svgParts.push(
          `  <text x="${vetNameX}" y="${vetNameY}" font-size="${ptToPx(
            11
          )}" fill="#000" font-family="HGP明朝B" dominant-baseline="hanging">${escapeXml(
            reportFields.attendingVet
          )}</text>`
        );
        if (textCfg.ATTENDING_VET_LINE) {
          const lineX = slideOffsetX + textCfg.ATTENDING_VET_LINE.x * pxPerCm;
          const lineY = slideOffsetY + textCfg.ATTENDING_VET_LINE.y * pxPerCm;
          const lineW = textCfg.ATTENDING_VET_LINE.w * pxPerCm;
          svgParts.push(
            `  <line x1="${lineX}" y1="${lineY}" x2="${lineX + lineW}" y2="${lineY}" stroke="#000" stroke-width="${ptToPx(0.5)}" />`
          );
        }
      }
      // 印鑑を1cm角枠付き文字で描く
      if (textCfg.SEAL) {
        const sealChar = getVetSealChar(reportFields.attendingVet || '');
        const stampUrl = getStampUrlByVetName(reportFields.attendingVet || '');
        const hasStampImage = !!stampUrl;
        const sx = slideOffsetX + textCfg.SEAL.x * pxPerCm;
        const sy = slideOffsetY + textCfg.SEAL.y * pxPerCm;
        const sw = textCfg.SEAL.w * pxPerCm;
        const sh = textCfg.SEAL.h * pxPerCm;

        if (stampUrl) {
          const stampHref = escapeXml(stampUrl);
          svgParts.push(
            `  <image x="${sx}" y="${sy}" width="${sw}" height="${sh}" href="${stampHref}" xlink:href="${stampHref}" preserveAspectRatio="xMidYMid meet" />`
          );
        }

        if (!hasStampImage) {
          svgParts.push(
            `  <rect x="${sx}" y="${sy}" width="${sw}" height="${sh}" fill="none" stroke="#000" stroke-width="${ptToPx(0.5)}"/>`
          );
        }

        if (!hasStampImage) {
          svgParts.push(
            `  <text x="${sx + sw / 2}" y="${sy + sh / 2}" font-size="${ptToPx(18)}" fill="#000" font-family="${svgFontFamily}" dominant-baseline="middle" text-anchor="middle" font-weight="bold">${escapeXml(
              sealChar
            )}</text>`
          );
        }
      }
    }

    // 4) 定型文①（飼い主姓・ペット名のみ下線）
    if (textCfg.FIXED_INTRO_TEXT) {
      const owner = (reportFields.ownerLastName || '[ 飼い主姓 ]').trim();
      const pet = (reportFields.petName || '[ ペット名 ]').trim();
      const paragraph1 = normalizeJapaneseSentence(
        `この度${owner}様の${pet}ちゃんをご紹介いただきましてありがとうございました。`
      );
      const paragraph2 = normalizeJapaneseSentence(
        '拝見させていただいた結果につきまして下記の通りご報告申し上げます。'
      );

      const bx = slideOffsetX + textCfg.FIXED_INTRO_TEXT.x * pxPerCm;
      const by = slideOffsetY + textCfg.FIXED_INTRO_TEXT.y * pxPerCm;
      const bw = textCfg.FIXED_INTRO_TEXT.w * pxPerCm;
      const bh = textCfg.FIXED_INTRO_TEXT.h * pxPerCm;

      const fontPt = LAYOUT.FONTS.BODY_BASE;
      const fontPx = ptToPx(fontPt);
      const lineHeight = fontPx * INTRO_LINE_SPACING_MULTIPLE;
      const maxLines = Math.max(1, Math.floor(bh / lineHeight));

      const ctx = getMeasureContext(fontPx, svgFontFamily);
      const paragraph1Lines = wrapTextByMeasure(paragraph1, bw, ctx);
      const paragraph2Lines = wrapTextByMeasure(paragraph2, bw, ctx);
      const allLines = [...paragraph1Lines, ...paragraph2Lines];
      const visibleLines = clampLines(allLines, maxLines);

      const parts: string[] = [];
      visibleLines.forEach((ln, idx) => {
        let lineEsc = escapeXml(ln);
        lineEsc = lineEsc
          .replace(escapeXml(owner), `<tspan text-decoration="underline">${escapeXml(owner)}</tspan>`)
          .replace(escapeXml(pet), `<tspan text-decoration="underline">${escapeXml(pet)}</tspan>`);
        if (idx === 0) parts.push(`<tspan x="${bx}" dy="0">${lineEsc}</tspan>`);
        else parts.push(
          `<tspan x="${bx}" dy="${lineHeight}">${lineEsc}</tspan>`
        );
      });

      svgParts.push(
        `  <text x="${bx}" y="${by}" font-size="${fontPx}" fill="#000" font-family="${svgFontFamily}" dominant-baseline="hanging">${parts.join(
          ''
        )}</text>`
      );
    }

    // 5) 日付行（初診日 / 鎮静日 / 全身麻酔日）
    const page1DateLines = buildPage1DateLines(reportFields);
    const safePage1DateLines = Array.isArray(page1DateLines) ? page1DateLines : [];
    const firstVisitCfg = textCfg.FIRST_VISIT_DATE;
    const anesthesiaCfg = textCfg.ANESTHESIA_DATE;
    const dateCount = safePage1DateLines.length;
    const dateLineGapCm = firstVisitCfg && anesthesiaCfg
      ? Math.max((anesthesiaCfg.y ?? 0) - (firstVisitCfg.y ?? 0), firstVisitCfg.h ?? 0.45)
      : (firstVisitCfg?.h ?? 0.45);

    if (firstVisitCfg && safePage1DateLines.length > 0) {
      const baseDateX = slideOffsetX + firstVisitCfg.x * pxPerCm;
      const baseDateYcm = firstVisitCfg.y;
      safePage1DateLines.forEach((line, idx) => {
        const y = slideOffsetY + (baseDateYcm + idx * dateLineGapCm) * pxPerCm;
        svgParts.push(
          `  <text x="${baseDateX}" y="${y}" font-size="${ptToPx(
            LAYOUT.FONTS.BODY_BASE
          )}" fill="#000" font-family="${svgFontFamily}" dominant-baseline="hanging">${escapeXml(
            `${line.label}：${String(line.value || '').trim()}`
          )}</text>`
        );
      });
    }

    // 7) 定型文②
    if (textCfg.FIXED_CLOSING_TEXT) {
      const closingText = `「${
        reportFields.chiefComplaint || '[ 主訴 ]'
      }」という主訴の為、拝見いたしました。`;

      const cx = slideOffsetX + textCfg.FIXED_CLOSING_TEXT.x * pxPerCm;
      const cy = slideOffsetY + getPage1ClosingTextYcmByDateCount(dateCount, textCfg.FIXED_CLOSING_TEXT.y) * pxPerCm;
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
        if (idx === 0)
          parts.push(`<tspan x="${cx}" dy="0">${escapeXml(ln)}</tspan>`);
        else
          parts.push(
            `<tspan x="${cx}" dy="${lineHeight}">${escapeXml(ln)}</tspan>`
          );
      });

      svgParts.push(
        `  <text x="${cx}" y="${cy}" font-size="${fontPx}" fill="#000" font-family="${svgFontFamily}" dominant-baseline="hanging">${parts.join(
          ''
        )}</text>`
      );
    }

    // 切り取り線
    ['CUT_LINE_TOP', 'CUT_LINE_BOTTOM', 'IMAGES_BOTTOM_LINE'].forEach(key => {
      const cfg = (textCfg as any)[key];
      if (cfg) {
        const lx = slideOffsetX + cfg.x * pxPerCm;
        const defaultLineYcm = cfg.y + cfg.h / 2;
        const lineYcm = key === 'CUT_LINE_BOTTOM'
          ? getPage1DividerYcmByDateCount(dateCount, defaultLineYcm)
          : defaultLineYcm;
        const ly = slideOffsetY + lineYcm * pxPerCm;
        const lw = cfg.w * pxPerCm;
        svgParts.push(`  <line x1="${lx}" y1="${ly}" x2="${lx + lw}" y2="${ly}" stroke="#000" stroke-width="${ptToPx(0.5)}" />`);
      }
    });

    // 「記」
    if (textCfg.KI) {
      const kx = slideOffsetX + textCfg.KI.x * pxPerCm;
      const ky = slideOffsetY + textCfg.KI.y * pxPerCm;
      const kw = textCfg.KI.w * pxPerCm;
      svgParts.push(
        `  <text x="${kx + kw / 2}" y="${ky}" font-size="${ptToPx(
          LAYOUT.FONTS.KI
        )}" fill="#000" font-family="HG明朝B" dominant-baseline="hanging" text-anchor="middle" font-weight="bold">記</text>`
      );
    }

    // 8) Section header 【初診時】
    if (textCfg.SECTION_HEADER) {
      const headerX = slideOffsetX + textCfg.SECTION_HEADER.x * pxPerCm;
      const headerY = slideOffsetY + initialBlockLayout.sectionHeaderY * pxPerCm;
      svgParts.push(
        `  <text x="${headerX}" y="${headerY}" font-size="${ptToPx(
          LAYOUT.FONTS.SECTION_HEADER
        )}" fill="#000" font-family="${svgFontFamily}" dominant-baseline="hanging" font-weight="bold">【初診時】</text>`
      );
    }

    // 初診時本文
    const bodyCfg2 = textCfg.FREE_TEXT_INITIAL;
    if (bodyCfg2 && reportFields.initialText) {
      const bx2 = slideOffsetX + bodyCfg2.x * pxPerCm;
      const by2 = slideOffsetY + initialBlockLayout.bodyY * pxPerCm;
      const bw2 = bodyCfg2.w * pxPerCm;
      const bh2 = bodyCfg2.h * pxPerCm;

      const fontPt2 = LAYOUT.FONTS.BODY_BASE;
      const fontPx2 = ptToPx(fontPt2);
      const lineHeight2 = fontPx2 * 1.15;
      const maxLines2 = Math.max(1, Math.floor(bh2 / lineHeight2));

      const ctx2 = getMeasureContext(fontPx2, svgFontFamily);
      const wrapped2 = wrapTextByMeasure(reportFields.initialText, bw2, ctx2);
      const visible2 = clampLines(wrapped2, maxLines2);

      const parts2: string[] = [];
      visible2.forEach((ln, idx) => {
        if (idx === 0)
          parts2.push(`<tspan x="${bx2}" dy="0">${escapeXml(ln)}</tspan>`);
        else
          parts2.push(
            `<tspan x="${bx2}" dy="${lineHeight2}">${escapeXml(ln)}</tspan>`
          );
      });

      svgParts.push(
        `  <text x="${bx2}" y="${by2}" font-size="${fontPx2}" fill="#000" font-family="${svgFontFamily}" dominant-baseline="hanging">${parts2.join(
          ''
        )}</text>`
      );
    }

    // 9) Images header 【処置前の肉眼写真等】
    if (textCfg.IMAGES_HEADER) {
      const imgHeaderX = slideOffsetX + textCfg.IMAGES_HEADER.x * pxPerCm;
      const imgHeaderY = slideOffsetY + initialBlockLayout.imagesHeaderY * pxPerCm;
      svgParts.push(
        `  <text x="${imgHeaderX}" y="${imgHeaderY}" font-size="${ptToPx(
          LAYOUT.FONTS.SECTION_HEADER
        )}" fill="#000" font-family="${svgFontFamily}" dominant-baseline="hanging" font-weight="bold">【処置前の肉眼写真等】</text>`
      );
    }
  } else if (pageNum === 2) {
    const lineCfg = (LAYOUT.PAGE2 as any).LINES;
    const textCfg = LAYOUT.PAGE2.TEXT as any;
    const placePostOnPage3 = options?.showPage3 && options?.postPlacement === 'page3';

    if (lineCfg?.SEP_TOP) {
      const lx = slideOffsetX + lineCfg.SEP_TOP.x * pxPerCm;
      const ly = slideOffsetY + lineCfg.SEP_TOP.y * pxPerCm;
      const lw = lineCfg.SEP_TOP.w * pxPerCm;
      svgParts.push(`  <line x1="${lx}" y1="${ly}" x2="${lx + lw}" y2="${ly}" stroke="#000" stroke-width="${ptToPx(0.5)}" />`);
    }

    if (lineCfg?.SEP_BOTTOM) {
      const lx = slideOffsetX + lineCfg.SEP_BOTTOM.x * pxPerCm;
      const ly = slideOffsetY + lineCfg.SEP_BOTTOM.y * pxPerCm;
      const lw = lineCfg.SEP_BOTTOM.w * pxPerCm;
      svgParts.push(`  <line x1="${lx}" y1="${ly}" x2="${lx + lw}" y2="${ly}" stroke="#000" stroke-width="${ptToPx(0.5)}" />`);
    }

    const headerProcCfg = textCfg.SECTION_HEADER_PROCEDURE;
    const bodyProcCfg = textCfg.FREE_TEXT_PROCEDURE;
    const headerPostCfg = textCfg.SECTION_HEADER_POSTOP;
    const bodyPostCfg = textCfg.FREE_TEXT_POSTOP;

    const fontPt = LAYOUT.FONTS.BODY_BASE;
    const fontPx = ptToPx(fontPt);
    const lineHeightPx = fontPx * 1.15;
    const lineHeightCm = lineHeightPx / pxPerCm;
    const fontFamily = svgFontFamily;

    const pageBottomCm = Math.max(
      bodyProcCfg ? bodyProcCfg.y + bodyProcCfg.h : 0,
      !placePostOnPage3 && bodyPostCfg ? bodyPostCfg.y + bodyPostCfg.h : 0
    );

    const bodyCtx = getMeasureContext(fontPx, fontFamily);

    const drawBody = (cfg: any, yCm: number, text: string, maxLines: number) => {
      if (!cfg || maxLines <= 0) return 0;
      const bx = slideOffsetX + cfg.x * pxPerCm;
      const by = slideOffsetY + yCm * pxPerCm;
      const bw = cfg.w * pxPerCm;

      const wrapped = wrapTextByMeasure(text || '', bw, bodyCtx);
      const visible = clampLines(wrapped, maxLines);
      if (visible.length === 0) return 0;

      const parts: string[] = [];
      visible.forEach((ln, idx) => {
        const lineText = ln === '' ? '　' : ln;
        if (idx === 0) parts.push(`<tspan x="${bx}" dy="0">${escapeXml(lineText)}</tspan>`);
        else parts.push(`<tspan x="${bx}" dy="${lineHeightPx}">${escapeXml(lineText)}</tspan>`);
      });

      svgParts.push(
        `  <text x="${bx}" y="${by}" font-size="${fontPx}" fill="#000" font-family="${fontFamily}" dominant-baseline="hanging">${parts.join('')}</text>`
      );

      return visible.length;
    };

    // PAGE2 auto layout base Y (offsets are applied later in preview pipeline)
    let layoutCursorYcm = headerProcCfg?.y ?? bodyProcCfg?.y ?? 0;
    if (typeof options?.page2ImagesBottomYcm === 'number') {
      layoutCursorYcm = Math.max(layoutCursorYcm, options.page2ImagesBottomYcm + lineHeightCm);
    }

    const examHeadingBaseYcm = layoutCursorYcm;
    if (headerProcCfg) {
      const hx = slideOffsetX + headerProcCfg.x * pxPerCm;
      const hy = slideOffsetY + examHeadingBaseYcm * pxPerCm;
      svgParts.push(
        `  <text x="${hx}" y="${hy}" font-size="${ptToPx(LAYOUT.FONTS.SECTION_HEADER)}" fill="#000" font-family="${svgFontFamily}" dominant-baseline="hanging" font-weight="bold">【検査・処置内容】</text>`
      );
      layoutCursorYcm = examHeadingBaseYcm + headerProcCfg.h;
    }

    const reservePostCm = !placePostOnPage3
      ? (headerPostCfg?.h ?? 0) + lineHeightCm * 2
      : 0;
    const procMaxLines = Math.max(
      0,
      Math.floor((pageBottomCm - layoutCursorYcm - reservePostCm) / lineHeightCm)
    );
    const examBodyBaseYcm = layoutCursorYcm;
    const procLines = drawBody(bodyProcCfg, examBodyBaseYcm, reportFields.procedureText || '', procMaxLines);
    layoutCursorYcm = examBodyBaseYcm + procLines * lineHeightCm;

    if (!placePostOnPage3) {
      layoutCursorYcm += lineHeightCm;

      const postHeadingBaseYcm = layoutCursorYcm;
      if (headerPostCfg) {
        const hx = slideOffsetX + headerPostCfg.x * pxPerCm;
        const hy = slideOffsetY + postHeadingBaseYcm * pxPerCm;
        svgParts.push(
          `  <text x="${hx}" y="${hy}" font-size="${ptToPx(LAYOUT.FONTS.SECTION_HEADER)}" fill="#000" font-family="${svgFontFamily}" dominant-baseline="hanging" font-weight="bold">【術後経過】</text>`
        );
        layoutCursorYcm = postHeadingBaseYcm + headerPostCfg.h;
      }

      const postMaxLines = Math.max(0, Math.floor((pageBottomCm - layoutCursorYcm) / lineHeightCm));
      const postBodyBaseYcm = layoutCursorYcm;
      const postLines = drawBody(bodyPostCfg, postBodyBaseYcm, reportFields.postText || '', postMaxLines);
      layoutCursorYcm = postBodyBaseYcm + postLines * lineHeightCm;

      // PAGE3が有効ならお礼文はPAGE3側に出すためスキップ
      if (!options?.showPage3) {
        const thankYouBody = getThankYouBody(reportFields);
        if (thankYouBody) {
          // 5行空けを試み、スペース不足なら最低2行まで縮める
          const headingH = headerPostCfg?.h ?? 0;
          let gapLines = 5;
          while (gapLines > 2 && layoutCursorYcm + gapLines * lineHeightCm + headingH > pageBottomCm) {
            gapLines--;
          }
          layoutCursorYcm += gapLines * lineHeightCm;

          const thanksMaxLines = Math.max(0, Math.floor((pageBottomCm - layoutCursorYcm) / lineHeightCm));
          const closingBodyBaseYcm = layoutCursorYcm;
          drawBody(bodyPostCfg, closingBodyBaseYcm, thankYouBody, thanksMaxLines);
        }
      }
    }
  } else if (pageNum === 3) {
    const lineCfg = (LAYOUT.PAGE3 as any).LINES;
    const textCfg = LAYOUT.PAGE3.TEXT as any;

    if (lineCfg?.SEP_TOP) {
      const lx = slideOffsetX + lineCfg.SEP_TOP.x * pxPerCm;
      const ly = slideOffsetY + lineCfg.SEP_TOP.y * pxPerCm;
      const lw = lineCfg.SEP_TOP.w * pxPerCm;
      svgParts.push(`  <line x1="${lx}" y1="${ly}" x2="${lx + lw}" y2="${ly}" stroke="#000" stroke-width="${ptToPx(0.5)}" />`);
    }

    if (lineCfg?.SEP_BOTTOM) {
      const lx = slideOffsetX + lineCfg.SEP_BOTTOM.x * pxPerCm;
      const ly = slideOffsetY + lineCfg.SEP_BOTTOM.y * pxPerCm;
      const lw = lineCfg.SEP_BOTTOM.w * pxPerCm;
      svgParts.push(`  <line x1="${lx}" y1="${ly}" x2="${lx + lw}" y2="${ly}" stroke="#000" stroke-width="${ptToPx(0.5)}" />`);
    }

    // --- PAGE3: 縦カーソル方式で安定レイアウト ---
    // 設定
    const page3StartY = 1.73;
    const page3StartX = 1.0;
    let cursorY = page3StartY;
    const fontFamily = svgFontFamily;
    const fontPt = LAYOUT.FONTS.BODY_BASE;
    const fontPx = ptToPx(fontPt);
    const lineHeightPx = fontPx * 1.15;
    const lineHeightCm = lineHeightPx / pxPerCm;

    // 1. 画像（未実装: 画像描画ロジックが必要な場合はここに追加）
    // 画像の最下端まで cursorY を進める（仮実装: 画像がなければスキップ）
    // TODO: 画像描画が必要な場合はここで cursorY を調整

    // 2. 【PAGE3】自由入力
    let freeTextHeight = 0;
    if (textCfg.FREE_TEXT_PAGE3 && reportFields.page3Text) {
      const box = textCfg.FREE_TEXT_PAGE3;
      const bw = box.w * pxPerCm;
      const bh = box.h * pxPerCm;
      const bx = slideOffsetX + page3StartX * pxPerCm;
      // 画像がある場合は2行空け、なければそのまま
      // 今回は画像なし前提なので2行空けは省略
      const ctx = getMeasureContext(fontPx, fontFamily);
      const wrapped = wrapTextByMeasure(reportFields.page3Text, bw, ctx);
      const maxLines = Math.max(1, Math.floor(bh / lineHeightPx));
      const visible = clampLines(wrapped, maxLines);
      freeTextHeight = visible.length * lineHeightCm;
      const baseY = cursorY;
      const renderedY = baseY + (options?.previewYOffsets?.page3FreeText ?? 0);
      const by = slideOffsetY + renderedY * pxPerCm;
      const parts: string[] = [];
      visible.forEach((ln, idx) => {
        const lineText = ln === '' ? '　' : ln;
        if (idx === 0)
          parts.push(`<tspan x="${bx}" dy="0">${escapeXml(lineText)}</tspan>`);
        else
          parts.push(`<tspan x="${bx}" dy="${lineHeightPx}">${escapeXml(lineText)}</tspan>`);
      });
      svgParts.push(
        `  <text x="${bx}" y="${by}" font-size="${fontPx}" fill="#000" font-family="${fontFamily}" dominant-baseline="hanging">${parts.join('')}</text>`
      );
      cursorY += freeTextHeight;
    }

    // 3. 【術後経過】タイトル
    let postopTitleHeight = 0;
    if (textCfg.SECTION_HEADER_POSTOP_PAGE3) {
      // 自由入力があれば1行空け
      if (freeTextHeight > 0) cursorY += lineHeightCm;
      const box = textCfg.SECTION_HEADER_POSTOP_PAGE3;
      const fontPxHeader = ptToPx(LAYOUT.FONTS.SECTION_HEADER);
      const bx = slideOffsetX + page3StartX * pxPerCm;
      const baseY = cursorY;
      const renderedY = baseY + (options?.previewYOffsets?.page3PostopTitle ?? 0);
      const by = slideOffsetY + renderedY * pxPerCm;
      svgParts.push(
        `  <text x="${bx}" y="${by}" font-size="${fontPxHeader}" fill="#000" font-family="${svgFontFamily}" dominant-baseline="hanging" font-weight="bold">【術後経過】</text>`
      );
      postopTitleHeight = lineHeightCm;
      cursorY += postopTitleHeight;
    }

    // 4. 【術後経過】本文
    let postopBodyHeight = 0;
    if (textCfg.FREE_TEXT_POSTOP_PAGE3 && reportFields.postText) {
      const box = textCfg.FREE_TEXT_POSTOP_PAGE3;
      const bw = box.w * pxPerCm;
      const bh = box.h * pxPerCm;
      const bx = slideOffsetX + page3StartX * pxPerCm;
      const ctx = getMeasureContext(fontPx, fontFamily);
      const wrapped = wrapTextByMeasure(reportFields.postText, bw, ctx);
      const maxLines = Math.max(1, Math.floor(bh / lineHeightPx));
      const visible = clampLines(wrapped, maxLines);
      postopBodyHeight = visible.length * lineHeightCm;
      const baseY = cursorY;
      const renderedY = baseY + (options?.previewYOffsets?.page3PostopBody ?? 0);
      const by = slideOffsetY + renderedY * pxPerCm;
      const parts: string[] = [];
      visible.forEach((ln, idx) => {
        const lineText = ln === '' ? '　' : ln;
        if (idx === 0)
          parts.push(`<tspan x="${bx}" dy="0">${escapeXml(lineText)}</tspan>`);
        else
          parts.push(`<tspan x="${bx}" dy="${lineHeightPx}">${escapeXml(lineText)}</tspan>`);
      });
      svgParts.push(
        `  <text x="${bx}" y="${by}" font-size="${fontPx}" fill="#000" font-family="${fontFamily}" dominant-baseline="hanging">${parts.join('')}</text>`
      );
      cursorY += postopBodyHeight;
    }

    // 5. 【お礼文】本文（PAGE3に出る場合）
    if (textCfg.FREE_TEXT_THANKS_PAGE3) {
      const thankYouBody = getThankYouBody(reportFields);
      if (thankYouBody) {
        // 術後経過本文があれば1行空け
        if (postopBodyHeight > 0) cursorY += lineHeightCm;
        const box = textCfg.FREE_TEXT_THANKS_PAGE3;
        const bw = box.w * pxPerCm;
        const bh = box.h * pxPerCm;
        const bx = slideOffsetX + page3StartX * pxPerCm;
        const ctx = getMeasureContext(fontPx, fontFamily);
        const wrapped = wrapTextByMeasure(thankYouBody, bw, ctx);
        const maxLines = Math.max(1, Math.floor(bh / lineHeightPx));
        const visible = clampLines(wrapped, maxLines);
        const baseY = cursorY;
        const renderedY = baseY + (options?.previewYOffsets?.page3ThanksBody ?? 0);
        const by = slideOffsetY + renderedY * pxPerCm;
        const parts: string[] = [];
        visible.forEach((ln, idx) => {
          const lineText = ln === '' ? '　' : ln;
          if (idx === 0)
            parts.push(`<tspan x="${bx}" dy="0">${escapeXml(lineText)}</tspan>`);
          else
            parts.push(`<tspan x="${bx}" dy="${lineHeightPx}">${escapeXml(lineText)}</tspan>`);
        });
        svgParts.push(
          `  <text x="${bx}" y="${by}" font-size="${fontPx}" fill="#000" font-family="${fontFamily}" dominant-baseline="hanging">${parts.join('')}</text>`
        );
        cursorY += visible.length * lineHeightCm;
      }
    }
  }

  return svgParts;
}
// --- PPTX text -------------------------------------------------------------

export function addPptxText(
  slide: pptxgen.Slide,
  pageNum: number,
  reportFields: ReportFields,
  options?: RenderOptions
) {
  if (pageNum === 1) {
    const textCfg = LAYOUT.PAGE1.TEXT as any;
    const fontFamily = "Meiryo, 'MS PGothic', 'Noto Sans JP', sans-serif";
    const initialBlockLayout = getPage1InitialBlockLayout(
      textCfg,
      reportFields,
      fontFamily
    );

    // logo image（固定座標＆固定サイズ：layout.ts の x/y/w/h をそのまま使う）
// ロゴ画像（固定座標＆固定サイズ）
if (textCfg.LOGO) {
  slide.addImage({
    data: logoDataUrl,
    x: cmToInch(textCfg.LOGO.x),
    y: cmToInch(textCfg.LOGO.y),
    w: cmToInch(textCfg.LOGO.w),
    h: cmToInch(textCfg.LOGO.h),
  });
}

    // 報告日（右上）
    if (textCfg.REPORT_DATE) {
      slide.addText(reportFields.reportDate || '', {
        x: cmToInch(textCfg.REPORT_DATE.x),
        y: cmToInch(textCfg.REPORT_DATE.y),
        w: cmToInch(textCfg.REPORT_DATE.w),
        h: cmToInch(textCfg.REPORT_DATE.h),
        fontSize: LAYOUT.FONTS.REPORT_DATE,
        align: 'right'
      });
    }
    
    // hospital information block left edge already computed above
    // (shared constant used instead of redeclaring)

    // hospital info（病院名・住所・メール）
// 病院名
if (textCfg.HOSPITAL_INFO) {
  slide.addText('荻窪ツイン動物病院', {
    x: cmToInch(textCfg.HOSPITAL_INFO.x),
    y: cmToInch(textCfg.HOSPITAL_INFO.y),
    w: cmToInch(textCfg.HOSPITAL_INFO.w),
    h: cmToInch(textCfg.HOSPITAL_INFO.h),
    fontSize: LAYOUT.FONTS.INFO_NAME
  });
}
// 住所
if (textCfg.HOSPITAL_ADDR) {
  slide.addText('東京都杉並区上荻1-23-18', {
    x: cmToInch(textCfg.HOSPITAL_ADDR.x),
    y: cmToInch(textCfg.HOSPITAL_ADDR.y),
    w: cmToInch(textCfg.HOSPITAL_ADDR.w),
    h: cmToInch(textCfg.HOSPITAL_ADDR.h),
    fontSize: LAYOUT.FONTS.INFO_DETAIL
  });
}
// メール
if (textCfg.HOSPITAL_EMAIL) {
  slide.addText('ogikubotwinah@gmail.com', {
    x: cmToInch(textCfg.HOSPITAL_EMAIL.x),
    y: cmToInch(textCfg.HOSPITAL_EMAIL.y),
    w: cmToInch(textCfg.HOSPITAL_EMAIL.w),
    h: cmToInch(textCfg.HOSPITAL_EMAIL.h),
    fontSize: LAYOUT.FONTS.INFO_DETAIL
  });
}

    // 1) Title (中央揃え)
    slide.addText('ご紹介患者についてのご報告', {
      x: cmToInch(textCfg.TITLE.x),
      y: cmToInch(textCfg.TITLE.y),
      w: cmToInch(textCfg.TITLE.w),
      h: cmToInch(textCfg.TITLE.h),
      fontSize: LAYOUT.FONTS.MAIN_TITLE,
      bold: true,
      align: 'center',
      fontFace: 'HGP創英プレゼンスEB'
    });

    // 2) 紹介病院 + 先生名を同じ行（1行固定、13.5pt HGP明朝B）下線はそれぞれ
    const hospName2 = reportFields.refHospital || '○○動物病院';
    const docName2 = (reportFields.refDoctor || '').trim();
    if (docName2) {
      slide.addText([
        { text: hospName2, options: { underline: true } },
        { text: '　' },
        { text: docName2, options: { underline: true } },
        { text: ' 先生' }
      ] as any, {
        x: cmToInch(textCfg.REF_HOSPITAL.x),
        y: cmToInch(textCfg.REF_HOSPITAL.y),
        w: cmToInch(textCfg.REF_HOSPITAL.w),
        h: cmToInch(textCfg.REF_HOSPITAL.h),
        fontSize: 13.5,
        fontFace: 'HGP明朝B',
        wrap: false
      });
    } else {
      slide.addText([
        { text: hospName2, options: { underline: true } }
      ] as any, {
        x: cmToInch(textCfg.REF_HOSPITAL.x),
        y: cmToInch(textCfg.REF_HOSPITAL.y),
        w: cmToInch(textCfg.REF_HOSPITAL.w),
        h: cmToInch(textCfg.REF_HOSPITAL.h),
        fontSize: 13.5,
        fontFace: 'HGP明朝B',
        wrap: false
      });
    }

    // 4) 担当獣医師（ラベル＋名前下線＋印鑑）
// 担当獣医師（ラベル）
if (textCfg.ATTENDING_VET_LABEL) {
  slide.addText('担当獣医師：', {
    x: cmToInch(textCfg.ATTENDING_VET_LABEL.x),
    y: cmToInch(textCfg.ATTENDING_VET_LABEL.y),
    w: cmToInch(textCfg.ATTENDING_VET_LABEL.w),
    h: cmToInch(textCfg.ATTENDING_VET_LABEL.h),
    fontSize: 11,
    fontFace: 'HGP明朝B',
    valign: 'top'
  });
}

// 担当獣医師（名前）＋「ラベル終端〜印鑑手前」下線（rectで描画）
if (textCfg.ATTENDING_VET_NAME) {
  // 名前テキスト（位置は変えない）
  slide.addText(
    [{ text: reportFields.attendingVet || '' }] as any,
    {
      x: cmToInch(textCfg.ATTENDING_VET_NAME.x),
      y: cmToInch(textCfg.ATTENDING_VET_NAME.y),
      w: cmToInch(textCfg.ATTENDING_VET_NAME.w),
      h: cmToInch(textCfg.ATTENDING_VET_NAME.h),
      fontSize: 11,
      fontFace: 'HGP明朝B',
      valign: 'top'
    } as any
  );

  // 下線（rect：0.5pt相当）
  if (textCfg.ATTENDING_VET_LINE) {
    slide.addShape('line' as any, {
      x: cmToInch(textCfg.ATTENDING_VET_LINE.x),
      y: cmToInch(textCfg.ATTENDING_VET_LINE.y),
      w: cmToInch(textCfg.ATTENDING_VET_LINE.w),
      h: 0,
      line: { color: '000000', pt: 0.5 }
    });
  }
}
// 印鑑（1cm角・固定位置）
if (textCfg.SEAL) {
  const sealChar = getVetSealChar(reportFields.attendingVet || '');
  const stampUrl = getStampUrlByVetName(reportFields.attendingVet || '');
  const hasStampImage = !!stampUrl;
  const sealX = cmToInch(textCfg.SEAL.x);
  const sealY = cmToInch(textCfg.SEAL.y);
  const sealW = cmToInch(textCfg.SEAL.w);
  const sealH = cmToInch(textCfg.SEAL.h);

  if (hasStampImage) {
    slide.addImage({
      data: stampUrl,
      x: sealX,
      y: sealY,
      w: sealW,
      h: sealH,
      sizing: {
        type: 'contain',
        x: sealX,
        y: sealY,
        w: sealW,
        h: sealH,
      } as any,
    } as any);
  } else {
    slide.addText(sealChar, {
      x: sealX,
      y: sealY,
      w: sealW,
      h: sealH,
      fontSize: 18,
      align: 'center',
      valign: 'middle',
      bold: true,
    });
  }

  if (!hasStampImage) {
    slide.addShape('rect' as any, {
      x: sealX,
      y: sealY,
      w: sealW,
      h: sealH,
      line: { color: '000000', pt: 0.5 },
      fill: { color: 'FFFFFF', transparency: 100 } as any,
    });
  }
}

    if (textCfg.FIXED_INTRO_TEXT) {
      const owner = (reportFields.ownerLastName || '[ 飼い主姓 ]').trim();
      const pet = (reportFields.petName || '[ ペット名 ]').trim();
      const introLine2 = normalizeJapaneseSentence('拝見させていただいた結果につきまして下記の通りご報告申し上げます。');
      slide.addText(
        [
          { text: 'この度' },
          { text: owner, options: { underline: true } },
          { text: '様の' },
          { text: pet, options: { underline: true } },
          { text: `ちゃんをご紹介いただきましてありがとうございました。\n` },
          { text: introLine2 }
        ] as any,
        {
          x: cmToInch(textCfg.FIXED_INTRO_TEXT.x),
          y: cmToInch(textCfg.FIXED_INTRO_TEXT.y),
          w: cmToInch(textCfg.FIXED_INTRO_TEXT.w),
          h: cmToInch(textCfg.FIXED_INTRO_TEXT.h),
          fontSize: LAYOUT.FONTS.BODY_BASE,
          valign: 'top',
          fontFace: 'HG明朝B',
          lineSpacingMultiple: INTRO_LINE_SPACING_MULTIPLE,
          breakLine: true,
          wrap: true,
        } as any
      );
    }

    const page1DateLines = buildPage1DateLines(reportFields);
    const safePage1DateLines = Array.isArray(page1DateLines) ? page1DateLines : [];
    const firstVisitCfg = textCfg.FIRST_VISIT_DATE;
    const anesthesiaCfg = textCfg.ANESTHESIA_DATE;
    const dateCount = safePage1DateLines.length;
    const dateLineGapCm = firstVisitCfg && anesthesiaCfg
      ? Math.max((anesthesiaCfg.y ?? 0) - (firstVisitCfg.y ?? 0), firstVisitCfg.h ?? 0.45)
      : (firstVisitCfg?.h ?? 0.45);

    if (firstVisitCfg && safePage1DateLines.length > 0) {
      safePage1DateLines.forEach((line, idx) => {
        slide.addText(`${line.label}：${String(line.value || '').trim()}`, {
          x: cmToInch(firstVisitCfg.x),
          y: cmToInch(firstVisitCfg.y + idx * dateLineGapCm),
          w: cmToInch(firstVisitCfg.w),
          h: cmToInch(firstVisitCfg.h),
          fontSize: LAYOUT.FONTS.BODY_BASE
        });
      });
    }

    // 7) 定型文②
    if (textCfg.FIXED_CLOSING_TEXT) {
  const closingText = `「${reportFields.chiefComplaint || '[ 主訴 ]'}」という主訴の為、拝見いたしました。`;
  slide.addText(closingText, {
    x: cmToInch(textCfg.FIXED_CLOSING_TEXT.x),
    y: cmToInch(getPage1ClosingTextYcmByDateCount(dateCount, textCfg.FIXED_CLOSING_TEXT.y)),
    w: cmToInch(textCfg.FIXED_CLOSING_TEXT.w),
    h: cmToInch(textCfg.FIXED_CLOSING_TEXT.h),
    fontSize: LAYOUT.FONTS.BODY_BASE,
    valign: 'top'
  });
}

    // 切り取り線・画像下線
    ['CUT_LINE_TOP','CUT_LINE_BOTTOM','IMAGES_BOTTOM_LINE'].forEach(key=>{
      const cfg=(textCfg as any)[key];
      if(cfg){
        const defaultLineYcm = cfg.y + (cfg.h / 2);
        const lineYcm = key === 'CUT_LINE_BOTTOM'
          ? getPage1DividerYcmByDateCount(dateCount, defaultLineYcm)
          : defaultLineYcm;
        slide.addShape('line' as any,{
          x:cmToInch(cfg.x),
          y:cmToInch(lineYcm),
          w:cmToInch(cfg.w),
          h:0,
          line:{ color: '000000', pt: 0.5 }
        });
      }
    });

    // 「記」
    if(textCfg.KI){
      slide.addText('記',{
        x:cmToInch(textCfg.KI.x),y:cmToInch(textCfg.KI.y),w:cmToInch(textCfg.KI.w),h:cmToInch(textCfg.KI.h),
        fontSize:LAYOUT.FONTS.KI,fontFace:'HG明朝B',align:'center'
      });
    }

    // 8) Section header 【初診時】
    if (textCfg.SECTION_HEADER) {
      slide.addText('【初診時】', {
        x: cmToInch(textCfg.SECTION_HEADER.x),
        y: cmToInch(initialBlockLayout.sectionHeaderY),
        w: cmToInch(textCfg.SECTION_HEADER.w),
        h: cmToInch(textCfg.SECTION_HEADER.h),
        fontSize: LAYOUT.FONTS.SECTION_HEADER,
        bold: true
      });
    }

    // 初診時本文
    if (textCfg.FREE_TEXT_INITIAL && reportFields.initialText) {
      const fitSize = fitTextToBox(reportFields.initialText, textCfg.FREE_TEXT_INITIAL, LAYOUT.FONTS.BODY_BASE, LAYOUT.FONTS.MIN_SIZE);
      slide.addText(reportFields.initialText, {
        x: cmToInch(textCfg.FREE_TEXT_INITIAL.x),
        y: cmToInch(initialBlockLayout.bodyY),
        w: cmToInch(textCfg.FREE_TEXT_INITIAL.w),
        h: cmToInch(textCfg.FREE_TEXT_INITIAL.h),
        fontSize: fitSize,
        valign: 'top',
        wrap: true
      });
    }

    // 9) Images header 【処置前の肉眼写真等】
    if (textCfg.IMAGES_HEADER) {
      slide.addText('【処置前の肉眼写真等】', {
        x: cmToInch(textCfg.IMAGES_HEADER.x),
        y: cmToInch(initialBlockLayout.imagesHeaderY),
        w: cmToInch(textCfg.IMAGES_HEADER.w),
        h: cmToInch(textCfg.IMAGES_HEADER.h),
        fontSize: LAYOUT.FONTS.SECTION_HEADER,
        bold: true
      });
    }
  } else if (pageNum === 2) {
    const lineCfg = (LAYOUT.PAGE2 as any).LINES;
    const textCfg = LAYOUT.PAGE2.TEXT as any;
    const placePostOnPage3 = options?.showPage3 && options?.postPlacement === 'page3';

    if (lineCfg?.SEP_TOP) {
      slide.addShape('line' as any, {
        x: cmToInch(lineCfg.SEP_TOP.x),
        y: cmToInch(lineCfg.SEP_TOP.y),
        w: cmToInch(lineCfg.SEP_TOP.w),
        h: 0,
        line: { color: '000000', pt: 0.5 }
      });
    }

    if (lineCfg?.SEP_BOTTOM) {
      slide.addShape('line' as any, {
        x: cmToInch(lineCfg.SEP_BOTTOM.x),
        y: cmToInch(lineCfg.SEP_BOTTOM.y),
        w: cmToInch(lineCfg.SEP_BOTTOM.w),
        h: 0,
        line: { color: '000000', pt: 0.5 }
      });
    }

    const headerProcCfg = textCfg.SECTION_HEADER_PROCEDURE;
    const bodyProcCfg = textCfg.FREE_TEXT_PROCEDURE;
    const headerPostCfg = textCfg.SECTION_HEADER_POSTOP;
    const bodyPostCfg = textCfg.FREE_TEXT_POSTOP;

    const fontPt = LAYOUT.FONTS.BODY_BASE;
    const fontPx = ptToPx(fontPt);
    const lineHeightPx = fontPx * 1.15;
    const lineHeightCm = (lineHeightPx / 96) * 2.54;
    const fontFamily = "Meiryo, 'MS PGothic', 'Noto Sans JP', sans-serif";
    const bodyCtx = getMeasureContext(fontPx, fontFamily);

    const pageBottomCm = Math.max(
      bodyProcCfg ? bodyProcCfg.y + bodyProcCfg.h : 0,
      !placePostOnPage3 && bodyPostCfg ? bodyPostCfg.y + bodyPostCfg.h : 0
    );

    const getVisibleBody = (cfg: any, text: string, maxLines: number) => {
      if (!cfg || maxLines <= 0) return [] as string[];
      const wrapped = wrapTextByMeasure(text || '', cfg.w * (96 / 2.54), bodyCtx);
      return clampLines(wrapped, maxLines);
    };

    // PAGE2 auto layout base Y (offsets are applied later in preview pipeline)
    let layoutCursorYcm = headerProcCfg?.y ?? bodyProcCfg?.y ?? 0;
    if (typeof options?.page2ImagesBottomYcm === 'number') {
      layoutCursorYcm = Math.max(layoutCursorYcm, options.page2ImagesBottomYcm + lineHeightCm);
    }

    const examHeadingBaseYcm = layoutCursorYcm;
    if (headerProcCfg) {
      slide.addText('【検査・処置内容】', {
        x: cmToInch(headerProcCfg.x),
        y: cmToInch(examHeadingBaseYcm),
        w: cmToInch(headerProcCfg.w),
        h: cmToInch(headerProcCfg.h),
        fontSize: LAYOUT.FONTS.SECTION_HEADER,
        bold: true
      });
      layoutCursorYcm = examHeadingBaseYcm + headerProcCfg.h;
    }

    const reservePostCm = !placePostOnPage3
      ? (headerPostCfg?.h ?? 0) + lineHeightCm * 2
      : 0;
    const procMaxLines = Math.max(
      0,
      Math.floor((pageBottomCm - layoutCursorYcm - reservePostCm) / lineHeightCm)
    );
    const examBodyBaseYcm = layoutCursorYcm;
    const procLines = getVisibleBody(bodyProcCfg, reportFields.procedureText || '', procMaxLines);
    if (bodyProcCfg && procLines.length > 0) {
      slide.addText(procLines.join('\n'), {
        x: cmToInch(bodyProcCfg.x),
        y: cmToInch(examBodyBaseYcm),
        w: cmToInch(bodyProcCfg.w),
        h: cmToInch(procLines.length * lineHeightCm),
        fontSize: LAYOUT.FONTS.BODY_BASE,
        align: 'left',
        valign: 'top',
        wrap: true
      });
    }
    layoutCursorYcm = examBodyBaseYcm + procLines.length * lineHeightCm;

    if (!placePostOnPage3) {
      layoutCursorYcm += lineHeightCm;

      const postHeadingBaseYcm = layoutCursorYcm;
      if (headerPostCfg) {
        slide.addText('【術後経過】', {
          x: cmToInch(headerPostCfg.x),
          y: cmToInch(postHeadingBaseYcm),
          w: cmToInch(headerPostCfg.w),
          h: cmToInch(headerPostCfg.h),
          fontSize: LAYOUT.FONTS.SECTION_HEADER,
          bold: true
        });
        layoutCursorYcm = postHeadingBaseYcm + headerPostCfg.h;
      }

      const postMaxLines = Math.max(0, Math.floor((pageBottomCm - layoutCursorYcm) / lineHeightCm));
      const postBodyBaseYcm = layoutCursorYcm;
      const postLines = getVisibleBody(bodyPostCfg, reportFields.postText || '', postMaxLines);
      if (bodyPostCfg && postLines.length > 0) {
        slide.addText(postLines.join('\n'), {
          x: cmToInch(bodyPostCfg.x),
          y: cmToInch(postBodyBaseYcm),
          w: cmToInch(bodyPostCfg.w),
          h: cmToInch(postLines.length * lineHeightCm),
          fontSize: LAYOUT.FONTS.BODY_BASE,
          align: 'left',
          valign: 'top',
          wrap: true
        });
      }
      layoutCursorYcm = postBodyBaseYcm + postLines.length * lineHeightCm;

      // PAGE3が有効ならお礼文はPAGE3側に出すためスキップ
      if (!options?.showPage3) {
        const thankYouBody = getThankYouBody(reportFields);
        if (thankYouBody) {
          // 5行空けを試み、スペース不足なら最低2行まで縮める
          const headingH = headerPostCfg?.h ?? 0;
          let gapLines = 5;
          while (gapLines > 2 && layoutCursorYcm + gapLines * lineHeightCm + headingH > pageBottomCm) {
            gapLines--;
          }
          layoutCursorYcm += gapLines * lineHeightCm;

          const thanksMaxLines = Math.max(0, Math.floor((pageBottomCm - layoutCursorYcm) / lineHeightCm));
          const closingBodyBaseYcm = layoutCursorYcm;
          const thankYouLines = getVisibleBody(bodyPostCfg, thankYouBody, thanksMaxLines);
          if (bodyPostCfg && thankYouLines.length > 0) {
            slide.addText(thankYouLines.join('\n'), {
              x: cmToInch(bodyPostCfg.x),
              y: cmToInch(closingBodyBaseYcm),
              w: cmToInch(bodyPostCfg.w),
              h: cmToInch(thankYouLines.length * lineHeightCm),
              fontSize: LAYOUT.FONTS.BODY_BASE,
              align: 'left',
              valign: 'top',
              wrap: true
            });
          }
        }
      }
    }
  } else if (pageNum === 3) {
    const lineCfg = (LAYOUT.PAGE3 as any).LINES;
    const textCfg = LAYOUT.PAGE3.TEXT as any;

    if (lineCfg?.SEP_TOP) {
      slide.addShape('line' as any, {
        x: cmToInch(lineCfg.SEP_TOP.x),
        y: cmToInch(lineCfg.SEP_TOP.y),
        w: cmToInch(lineCfg.SEP_TOP.w),
        h: 0,
        line: { color: '000000', pt: 0.5 }
      });
    }

    if (lineCfg?.SEP_BOTTOM) {
      slide.addShape('line' as any, {
        x: cmToInch(lineCfg.SEP_BOTTOM.x),
        y: cmToInch(lineCfg.SEP_BOTTOM.y),
        w: cmToInch(lineCfg.SEP_BOTTOM.w),
        h: 0,
        line: { color: '000000', pt: 0.5 }
      });
    }

    // 1. 【PAGE3】自由入力
    if (textCfg.FREE_TEXT_PAGE3 && reportFields.page3Text) {
      const box = textCfg.FREE_TEXT_PAGE3;
      const fitSize = fitTextToBox(reportFields.page3Text, box, LAYOUT.FONTS.BODY_BASE, LAYOUT.FONTS.MIN_SIZE);
      slide.addText(reportFields.page3Text, {
        x: cmToInch(box.x),
        y: cmToInch(textCfg.FREE_TEXT_PAGE3.y + (options?.previewYOffsets?.page3FreeText ?? 0)),
        w: cmToInch(box.w),
        h: cmToInch(box.h),
        fontSize: fitSize,
        align: 'left',
        valign: 'top',
        wrap: true
      });
    }

    // 2. 【術後経過】タイトル
    if (textCfg.SECTION_HEADER_POSTOP_PAGE3) {
      slide.addText('【術後経過】', {
        x: cmToInch(textCfg.SECTION_HEADER_POSTOP_PAGE3.x),
        y: cmToInch(textCfg.SECTION_HEADER_POSTOP_PAGE3.y + (options?.previewYOffsets?.page3PostopTitle ?? 0)),
        w: cmToInch(textCfg.SECTION_HEADER_POSTOP_PAGE3.w),
        h: cmToInch(textCfg.SECTION_HEADER_POSTOP_PAGE3.h),
        fontSize: LAYOUT.FONTS.SECTION_HEADER,
        bold: true
      });
    }

    // 3. 【術後経過】本文 (Page3)
    if (textCfg.FREE_TEXT_POSTOP_PAGE3 && reportFields.postText) {
      const box = textCfg.FREE_TEXT_POSTOP_PAGE3;
      const fitSize = fitTextToBox(reportFields.postText, box, LAYOUT.FONTS.BODY_BASE, LAYOUT.FONTS.MIN_SIZE);
      slide.addText(reportFields.postText, {
        x: cmToInch(box.x),
        y: cmToInch(textCfg.FREE_TEXT_POSTOP_PAGE3.y + (options?.previewYOffsets?.page3PostopBody ?? 0)),
        w: cmToInch(box.w),
        h: cmToInch(box.h),
        fontSize: fitSize,
        align: 'left',
        valign: 'top',
        wrap: true
      });
    }

    // 4. 【お礼文】本文（PAGE3）
    if (textCfg.FREE_TEXT_THANKS_PAGE3) {
      const thankYouBody = getThankYouBody(reportFields);
      if (thankYouBody) {
        const box = textCfg.FREE_TEXT_THANKS_PAGE3;
        const fitSize = fitTextToBox(thankYouBody, box, LAYOUT.FONTS.BODY_BASE, LAYOUT.FONTS.MIN_SIZE);
        slide.addText(thankYouBody, {
          x: cmToInch(box.x),
          y: cmToInch(textCfg.FREE_TEXT_THANKS_PAGE3.y + (options?.previewYOffsets?.page3ThanksBody ?? 0)),
          w: cmToInch(box.w),
          h: cmToInch(box.h),
          fontSize: fitSize,
          align: 'left',
          valign: 'top',
          wrap: true
        });
      }
    }
  }
}
