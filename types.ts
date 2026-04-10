export interface WordFrequency {
  word: string;
  count: number;
  rowCount: number;
}

export interface AnalysisResult {
  totalWords: number;
  uniqueWords: number;
  frequencies: WordFrequency[];
  summary: string;
}

export interface ParsedSheetData {
  sheetName: string;
  headers: string[];
  rows: Record<string, any>[];
}

export interface ParsedWorkbookData {
  sheetNames: string[];
  sheets: Record<string, ParsedSheetData>;
}
