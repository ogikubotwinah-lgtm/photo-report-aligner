
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
      TITLE: { x: 0.32, y: 1.0, w: 18.4, h: 1.0 },
      REPORT_DATE: { x: 14.0, y: 0.5, w: 4.5, h: 0.6 },
      REF_HOSPITAL: { x: 1.0, y: 3.0, w: 8.0, h: 0.8 },
      REF_DOCTOR: { x: 1.0, y: 4.0, w: 8.0, h: 0.8 },
      OWNER_LASTNAME: { x: 10.0, y: 3.0, w: 4.0, h: 0.8 },
      PET_NAME: { x: 14.0, y: 3.0, w: 4.0, h: 0.8 },
      SECTION_HEADER: { x: 0.63, y: 6.0, w: 3.0, h: 0.6 },
      FIRST_VISIT_DATE: { x: 3.8, y: 6.0, w: 5.0, h: 0.6 },
      ANESTHESIA_DATE: { x: 10.0, y: 6.0, w: 6.0, h: 0.6 },
      FREE_TEXT_INITIAL: { x: 0.63, y: 6.7, w: 17.79, h: 8.0 }
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
