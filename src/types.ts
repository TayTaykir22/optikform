export interface OpticField {
  startRow: number;
  endRow: number;
  startCol: number;
  endCol: number;
  type: string;
  direction: string;
  chars: string;
  label: string;
}

export interface StudentData {
  [key: string]: string;
}

export interface Calibration {
  offsetX: number;
  offsetY: number;
  bubbleSpacing: number; // in mm, default 1/6 inch
  bubbleSize: number; // in mm
}

export interface SavedProfile {
  id: string;
  name: string;
  fmtText: string;
  calibration: Calibration;
  fieldMappings: Record<string, number>;
  pageSize: { width: number; height: number };
  createdAt: number;
}
