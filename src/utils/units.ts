// Conversion utilities used across SVG preview and PPTX export
export const CM_PER_INCH = 2.54;

export function cmToInch(cm: number): number {
  return cm / CM_PER_INCH;
}

export function cmToPx(cm: number, pxPerCm: number): number {
  return cm * pxPerCm;
}

export function ptToPx(pt: number): number {
  // 1pt = 1/72 inch; 1in = 96px
  return pt * (96 / 72);
}
