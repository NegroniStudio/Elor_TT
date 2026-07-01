export interface SheetData {
  sheetName: string;
  grid: any[][];
  rowCount: number;
  columnCount: number;
}

export interface VisualContent {
  pages?: number;
  version?: string;
  info?: any;
  excerpt?: string;
  html?: string;
  sheets?: SheetData[];
}

export interface AIAnalysis {
  summary: string;
  documentType: string;
  keyTopics: string[];
  suggestions: string[];
}

export interface ParsedDocument {
  success: boolean;
  filename: string;
  size: number;
  extension: string;
  visualContent: VisualContent;
  extractedText: string;
  clipped: boolean;
  aiAnalysis?: AIAnalysis;
}

export interface ChatMessage {
  id: string;
  sender: "user" | "model";
  text: string;
  timestamp: Date;
}
