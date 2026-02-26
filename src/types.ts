
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

export interface ReportFields {
  reportDate: string;
  refHospital: string;
  refDoctor: string;
  ownerLastName: string;
  petName: string;
  firstVisitDate: string;
  anesthesiaDate: string;
  attendingVet: string; // 新規：担当獣医師
  initialText: string;
  procedureText: string;
  postText: string;
  chiefComplaint: string; // 新規：主訴
}
