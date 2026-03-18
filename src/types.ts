export type ImageCrop = {
  left: number;
  top: number;
  right: number;
  bottom: number;
};

export interface ImageData {
  id: string;
  name: string;
  dataUrl: string;
  width: number;
  height: number;
  mimeType: string;
  row: number; // 0: unassigned, 1-4: assigned rows
  orderConfirmed: boolean;
  rotation: number; // 0, 90, 180, 270

  // トリミング復元用
  originalDataUrl?: string;
  originalWidth?: number;
  originalHeight?: number;

  // 表示上のトリミング範囲
  crop?: ImageCrop;

  // 回転・反転状態
  flipX?: boolean;
  flipY?: boolean;
}

export interface LayoutOptions {
  spacing: number;
  padding: number;
  targetHeight: number;
  containerWidth: number;
  backgroundColor: string;
}

export interface ReportFields {
  reportDate: string;

  refHospitalName: string;
  refHospital: string;
  refHospitalEmail: string;
  refDoctor: string;

  ownerLastName: string;
  petName: string;

  firstVisitDate: string;
  sedationDate: string;
  anesthesiaDate: string;

  attendingVet: string;

  initialText: string;
  procedureText: string;
  postText: string;

  thankYouTextType: string;

  page3Text: string;
  chiefComplaint: string;

  page2PhotoCategory: string;
  page3PhotoLabel: string;
}

export type AppSuggestions = {
  refHospitals: string[];
  doctors: string[];
  refHospitalEmails: Record<string, string>;
};

export type DateFieldKey =
  | "reportDate"
  | "firstVisitDate"
  | "sedationDate"
  | "anesthesiaDate";