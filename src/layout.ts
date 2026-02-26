
/**
 * PPTXレイアウト設定定数
 * 単位はすべて cm です。
 */
export const LAYOUT = {
  SLIDE: {
    WIDTH_CM: 19.05,
    HEIGHT_CM: 27.517,
    NAME: 'A4_REPORT'
  },
  FONTS: {
    MAIN_TITLE: 14,
    INFO_NAME: 12,
    SECTION_HEADER: 11,
    BODY_BASE: 10.5,
    MIN_SIZE: 8.0
  },
  PAGE1: {
    IMAGES: {
      LEFT: 0.32,
      RIGHT: 0.32,
      START_Y: 15.82,
      MARGIN_BOTTOM: 1.4,
      ALIGN_LEFT: false
    },
    TEXT: {
      // Coordinates converted from PPTX (inch -> cm = *2.54)
      TITLE: { x: 0.1260 * 2.54, y: 0.3937 * 2.54, w: 7.2441 * 2.54, h: 0.3937 * 2.54 },
      REPORT_DATE: { x: 5.5118 * 2.54, y: 0.1969 * 2.54, w: 1.7717 * 2.54, h: 0.2362 * 2.54 },
      // LOGO (fixed): x=3.620in, y=1.765in, w=0.569in, h=0.494in
      LOGO: { x: 3.620 * 2.54, y: 1.765 * 2.54, w: 0.569 * 2.54, h: 0.494 * 2.54 },
      // HOSPITAL_INFO (fixed): x=4.199in, y=1.718in, w=2.896in, h=0.683in
      HOSPITAL_INFO: { x: 4.199 * 2.54, y: 1.718 * 2.54, w: 2.896 * 2.54, h: 0.683 * 2.54 },
      // 紹介病院と先生名を同じ行（y値揃える）
      REF_HOSPITAL: { x: 0.3937 * 2.54, y: 1.1811 * 2.54, w: 3.1496 * 2.54, h: 0.3150 * 2.54 },
      REF_DOCTOR: { x: 0.3937 * 2.54, y: 1.1811 * 2.54, w: 3.1496 * 2.54, h: 0.3150 * 2.54 },
      OWNER_LASTNAME: { x: 3.9370 * 2.54, y: 1.1811 * 2.54, w: 1.5748 * 2.54, h: 0.3150 * 2.54 },
      PET_NAME: { x: 5.5118 * 2.54, y: 1.1811 * 2.54, w: 1.5748 * 2.54, h: 0.3150 * 2.54 },
      // 担当獣医師（新規）
      ATTENDING_VET: { x: 0.3937 * 2.54, y: 1.9685 * 2.54, w: 6.299 * 2.54, h: 0.3150 * 2.54 },
      // 定型文①（新規・固定）
      FIXED_INTRO_TEXT: { x: 0.2480 * 2.54, y: 2.2441 * 2.54, w: 7.0039 * 2.54, h: 0.4725 * 2.54 },
      SECTION_HEADER: { x: 0.2480 * 2.54, y: 2.7165 * 2.54, w: 1.1811 * 2.54, h: 0.2362 * 2.54 },
      FIRST_VISIT_DATE: { x: 1.4961 * 2.54, y: 2.7165 * 2.54, w: 1.9685 * 2.54, h: 0.2362 * 2.54 },
      ANESTHESIA_DATE: { x: 3.9370 * 2.54, y: 2.7165 * 2.54, w: 2.3622 * 2.54, h: 0.2362 * 2.54 },
      // 定型文②（新規・固定）
      FIXED_CLOSING_TEXT: { x: 0.2480 * 2.54, y: 3.0315 * 2.54, w: 7.0039 * 2.54, h: 0.3543 * 2.54 },
      // 【初診時の肉眼写真等】見出し（新規）
      IMAGES_HEADER: { x: 0.2480 * 2.54, y: 3.3858 * 2.54, w: 2.1260 * 2.54, h: 0.2362 * 2.54 },
      // FREE_TEXT_INITIAL (本文枠): 移動なし（後で参照用として保持）
      FREE_TEXT_INITIAL: { x: 0.2480 * 2.54, y: 3.6220 * 2.54, w: 7.0039 * 2.54, h: 2.8189 * 2.54 }
    }
  },
  PAGE2: {
    IMAGES: {
      LEFT: 0.63,
      RIGHT: 0.63,
      START_Y: 2.76,
      END_Y: 11.6,
      ALIGN_LEFT: true
    },
    TEXT: {
      SECTION_HEADER_PROCEDURE: { x: 0.63, y: 12.0, w: 5.0, h: 0.6 },
      FREE_TEXT_PROCEDURE: { x: 0.63, y: 12.7, w: 17.79, h: 6.0 },
      SECTION_HEADER_POSTOP: { x: 0.63, y: 19.0, w: 5.0, h: 0.6 },
      FREE_TEXT_POSTOP: { x: 0.63, y: 19.7, w: 17.79, h: 6.0 }
    }
  }
};
