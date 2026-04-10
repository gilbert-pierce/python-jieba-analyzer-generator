import * as XLSX from 'xlsx';
import { ParsedWorkbookData, ParsedSheetData } from '../types';

export const parseExcelFile = async (file: File): Promise<ParsedWorkbookData> => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    
    reader.onload = (e) => {
      try {
        const data = e.target?.result;
        // Handle XLSX import variations (namespace vs default export)
        const xlsxLib = (XLSX as any).default || XLSX;
        
        const workbook = xlsxLib.read(data, { type: 'binary' });
        
        const sheetNames = workbook.SheetNames;
        const sheets: Record<string, ParsedSheetData> = {};

        sheetNames.forEach((sheetName: string) => {
          const sheet = workbook.Sheets[sheetName];
          const jsonData = xlsxLib.utils.sheet_to_json(sheet, { header: 1 });
          
          if (jsonData.length > 0) {
            const headers = (jsonData[0] as any[]).map(h => String(h || ''));
            const rowsRaw = jsonData.slice(1) as any[][];

            const rows = rowsRaw.map((row) => {
              const rowData: Record<string, any> = {};
              headers.forEach((header, index) => {
                rowData[header] = row[index];
              });
              return rowData;
            });

            sheets[sheetName] = { sheetName, headers, rows };
          } else {
            sheets[sheetName] = { sheetName, headers: [], rows: [] };
          }
        });

        resolve({ sheetNames, sheets });
      } catch (error) {
        reject(error);
      }
    };

    reader.onerror = (error) => reject(error);
    reader.readAsBinaryString(file);
  });
};
