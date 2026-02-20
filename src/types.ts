
export interface ImageData {
  id: string;
  name: string;
  dataUrl: string;
  width: number;
  height: number;
  mimeType: string;
  row: number; // 0: Unassigned, 1-4: Assigned rows
  orderConfirmed: boolean; // Whether the position within the row is confirmed
  rotation: number; // 0, 90, 180, 270 degrees
}

export interface LayoutOptions {
  spacing: number;
  padding: number;
  targetHeight: number;
  containerWidth: number;
  backgroundColor: string;
}
