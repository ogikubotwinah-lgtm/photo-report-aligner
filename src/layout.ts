/**
 * PPTXレイアウト設定定数
 * 単位はすべて cm です。
 */
const TOP_BLOCK_OFFSET_CM = -0.8;
const withTopOffset = <T extends { y: number }>(box: T): T => ({
  ...box,
  y: box.y + TOP_BLOCK_OFFSET_CM,
});

const PAGE2_BLOCK_OFFSET_CM = -3.58;
const withPage2BlockOffset = <T extends { y: number }>(box: T): T => ({
  ...box,
  y: box.y + PAGE2_BLOCK_OFFSET_CM,
});

export const LAYOUT = {
  SLIDE: {
    WIDTH_CM: 19.05,
    HEIGHT_CM: 27.517,
    NAME: "A4_REPORT",
  },

  // 右ブロック（固定方式に変更）
  RIGHT_BLOCK_LEFT: 11.16, // 右ブロック左端
  RIGHT_BLOCK_RIGHT: 18.02,
  HOSPITAL_BLOCK_WIDTH: 6.86, // 18.02 - 11.16

  FONTS: {
    MAIN_TITLE: 14,
    INFO_NAME: 12,
    REPORT_DATE: 11,
    INFO_DETAIL: 8,
    KI: 12.5,
    SECTION_HEADER: 11,
    BODY_BASE: 10.5,
    MIN_SIZE: 8.0,
  },

  PAGE1: {
    IMAGES: {
      LEFT: 1.0,
      RIGHT: 1.0,
      START_Y: 15.82,
      MARGIN_BOTTOM: 1.4,
      ALIGN_LEFT: false,
    },
    TEXT: {
      TITLE: { x: 0.1260 * 2.54, y: 0.3937 * 2.54, w: 7.2441 * 2.54, h: 0.3937 * 2.54 },
      REPORT_DATE: { x: 5.5118 * 2.54, y: 0.1969 * 2.54, w: 1.7717 * 2.54, h: 0.2362 * 2.54 },

      // ロゴ（確定）
      LOGO: withTopOffset({ x: 10.42, y: 4.12, w: 1.89, h: 1.9 }),

      // 病院情報（ロゴ基準で縦位置を確定）
      // x を同一値に揃え、SVG/PPTX 両方で一致させる
      HOSPITAL_INFO: withTopOffset({ x: 12.42, y: 4.36, w: 7.0, h: 0.6 }),
      HOSPITAL_ADDR: withTopOffset({ x: 12.42, y: 4.36 + 0.55, w: 7.0, h: 0.45 }),
      HOSPITAL_EMAIL: withTopOffset({ x: 12.42, y: 4.36 + 1.0, w: 7.0, h: 0.45 }),

      // 紹介病院と先生名（横幅17cm）
      REF_HOSPITAL: { x: 0.3937 * 2.54, y: 1.1811 * 2.54, w: 17.0, h: 0.3150 * 2.54 },

      // 担当獣医師＋印鑑（確定座標）
      ATTENDING_VET_LABEL: withTopOffset({ x: 12.1, y: 5.86, w: 3.5, h: 0.6 }),
      ATTENDING_VET_NAME: withTopOffset({ x: 14.9, y: 5.86, w: 3.0, h: 0.6 }),
      ATTENDING_VET_LINE: withTopOffset({ x: 14.52, y: 6.45, w: 2.7, h: 0.02 }),
      SEAL: withTopOffset({ x: 17.38, y: 5.65, w: 1.0, h: 1.0 }),

      // 定型文
      FIXED_INTRO_TEXT: withTopOffset({ x: 1.0, y: 6.97, w: 17.0, h: 1.7 }),

      // 初診日・全身麻酔日
      FIRST_VISIT_DATE: withTopOffset({ x: 1.0, y: 9.15 + 0.1, w: 10.0, h: 0.55 }),
      ANESTHESIA_DATE: withTopOffset({ x: 1.0, y: 9.8, w: 10.0, h: 0.55 }),

      // 主訴
      FIXED_CLOSING_TEXT: withTopOffset({ x: 1.0, y: 10.84, w: 17.0, h: 1.0 }),

      // 左端 x=1.0cm
      SECTION_HEADER: { x: 1.0, y: 11.5, w: 3.0, h: 0.6 },
      IMAGES_HEADER: { x: 1.0, y: 12.1, w: 6.0, h: 0.6 },
      FREE_TEXT_INITIAL: { x: 1.23, y: 12.7, w: 17.0, h: 2.8 },

      CUT_LINE_TOP: withTopOffset({ x: 1.0, y: 9.02, w: 17.38, h: 0.07 }),
      CUT_LINE_BOTTOM: withTopOffset({ x: 1.0, y: 10.55, w: 17.38, h: 0.07 }),
      KI: withTopOffset({ x: 1.0, y: 8.52, w: 17.38, h: 0.5 }),
      IMAGES_BOTTOM_LINE: { x: 1.0, y: 26.26, w: 17.38, h: 0.07 },
    },
  },

  PAGE2: {
    IMAGES: {
      LEFT: 0.63,
      RIGHT: 0.63,
      START_Y: 2.76,
      END_Y: 11.6,
      ALIGN_LEFT: true,
    },
    LINES: {
      SEP_TOP: { x: 1.0, y: 1.27, w: 17.38 },
      SEP_BOTTOM: { x: 1.0, y: 26.3, w: 17.38 },
    },
    TEXT: {
      SECTION_HEADER_PROCEDURE: withPage2BlockOffset({ x: 1.0, y: 12.0, w: 5.0, h: 0.6 }),
      FREE_TEXT_PROCEDURE: withPage2BlockOffset({ x: 1.23, y: 12.7, w: 17.0, h: 6.0 }),
      SECTION_HEADER_POSTOP: withPage2BlockOffset({ x: 1.0, y: 19.0, w: 5.0, h: 0.6 }),
      FREE_TEXT_POSTOP: withPage2BlockOffset({ x: 1.23, y: 19.7, w: 17.0, h: 6.0 }),
    },
  },

  PAGE3: {
    IMAGES: {
      LEFT: 0.63,
      RIGHT: 0.63,
      START_Y: 2.76,
      END_Y: 11.6,
      ALIGN_LEFT: true,
    },
    LINES: {
      SEP_TOP: { x: 1.0, y: 1.27, w: 17.38 },
      SEP_BOTTOM: { x: 1.0, y: 26.3, w: 17.38 },
    },
    TEXT: {
      FREE_TEXT_PAGE3: withPage2BlockOffset({ x: 1.23, y: 12.7, w: 17.0, h: 3.0 }), // 【PAGE3】自由入力
      SECTION_HEADER_POSTOP_PAGE3: withPage2BlockOffset({ x: 1.0, y: 16.0, w: 5.0, h: 0.6 }), // 【術後経過】タイトル
      FREE_TEXT_POSTOP_PAGE3: withPage2BlockOffset({ x: 1.23, y: 16.7, w: 17.0, h: 3.0 }), // 【術後経過】本文
      FREE_TEXT_THANKS_PAGE3: withPage2BlockOffset({ x: 1.23, y: 20.0, w: 17.0, h: 2.0 }), // 【お礼文】本文
    },
  },
} as const;