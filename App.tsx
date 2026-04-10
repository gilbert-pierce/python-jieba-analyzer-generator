import React, { useState, useRef, useEffect } from 'react';
import { isTauri } from '@tauri-apps/api/core';
import { Download, Play, X, FileText, ChartBar, Table, List, Activity, User, MapPin } from 'lucide-react';
import * as XLSX from 'xlsx';
import { parseExcelFile } from './utils/excelParser';
import { analyzeTextLocally, segmentText, parseStopWords, extractKeywordsLocally, calculateEntropy } from './utils/segmentation';
import { ParsedSheetData, ParsedWorkbookData } from './types';
import initJieba, { tag as jiebaTag } from 'jieba-wasm';
// Deep path: jieba-wasm package exports do not expose the .wasm file to the bundler.
import jiebaWasmUrl from './node_modules/jieba-wasm/pkg/web/jieba_rs_wasm_bg.wasm?url';

const DEFAULT_STOP_WORDS_DISPLAY = "的,了,是,我,你,在,和,就,都,而,及,与,这,那,有,个,之,为,但,我们,公司,有限公司,\n，,。,！,？,、,“, ”,：,；";

type PreviewMode = 'stats' | 'segmentation' | 'row_keywords' | 'tfidf' | 'entropy' | 'ner_person' | 'ner_location' | null;

const SETTINGS_STORAGE_KEY = 'jieba-analyzer:settings:v1';

interface PreviewRow {
  col1: string | number;
  col2: string;
  col3: string | number;
  col4?: string | number;
  col5?: string | number;
}

export interface PosExclusions {
  noun: boolean;
  verb: boolean;
  adj: boolean;
  adv: boolean;
  prep: boolean;
  particle: boolean;
}

function detectTauriRuntime(): boolean {
  try {
    if (isTauri()) return true;
  } catch {
    // ignore
  }

  // Fallbacks: depending on bundling and API init timing, isTauri() may be false.
  const w = window as unknown as Record<string, unknown>;
  if (w && (w.__TAURI_INTERNALS__ || w.__TAURI__)) return true;

  const proto = window.location?.protocol || '';
  if (proto.startsWith('tauri')) return true;

  // Last resort heuristic.
  const ua = navigator.userAgent || '';
  if (ua.toLowerCase().includes('tauri')) return true;

  return false;
}

export default function App() {
  const [file, setFile] = useState<File | null>(null);
  const [parsedWorkbook, setParsedWorkbook] = useState<ParsedWorkbookData | null>(null);
  const [sheetNames, setSheetNames] = useState<string[]>([]);
  const [selectedSheet, setSelectedSheet] = useState<string>("");
  const [headers, setHeaders] = useState<string[]>([]);
  const [selectedColumn, setSelectedColumn] = useState<string>("");
  
  const [isAnalyzing, setIsAnalyzing] = useState(false);
  const [parsedData, setParsedData] = useState<ParsedSheetData | null>(null);
  
  // Custom Stop Words State
  const [customStopWords, setCustomStopWords] = useState(DEFAULT_STOP_WORDS_DISPLAY);
  const [customFixedWords, setCustomFixedWords] = useState("");
  const [filterEnglish, setFilterEnglish] = useState(false);
  const [filterNumbers, setFilterNumbers] = useState(false);
  const [minWordCount, setMinWordCount] = useState<number>(2);
  const [maxWordCount, setMaxWordCount] = useState<number>(50);

  // POS Exclusion State (True means Exclude)
  const [posExclusions, setPosExclusions] = useState<PosExclusions>({
    noun: false,
    verb: false,
    adj: false,
    adv: false,
    prep: false,
    particle: false
  });

  // Preview State
  const [previewMode, setPreviewMode] = useState<PreviewMode>(null);
  const [previewData, setPreviewData] = useState<PreviewRow[]>([]);
  const [jiebaReady, setJiebaReady] = useState(false);

  // Load persisted settings once.
  useEffect(() => {
    try {
      const raw = localStorage.getItem(SETTINGS_STORAGE_KEY);
      if (!raw) return;
      const parsed = JSON.parse(raw) as Partial<{
        customStopWords: string;
        customFixedWords: string;
        filterEnglish: boolean;
        filterNumbers: boolean;
        minWordCount: number;
        maxWordCount: number;
        posExclusions: PosExclusions;
      }>;

      if (typeof parsed.customStopWords === 'string') setCustomStopWords(parsed.customStopWords);
      if (typeof parsed.customFixedWords === 'string') setCustomFixedWords(parsed.customFixedWords);
      if (typeof parsed.filterEnglish === 'boolean') setFilterEnglish(parsed.filterEnglish);
      if (typeof parsed.filterNumbers === 'boolean') setFilterNumbers(parsed.filterNumbers);
      if (typeof parsed.minWordCount === 'number') setMinWordCount(parsed.minWordCount);
      if (typeof parsed.maxWordCount === 'number') setMaxWordCount(parsed.maxWordCount);
      if (parsed.posExclusions && typeof parsed.posExclusions === 'object') {
        setPosExclusions({
          noun: !!parsed.posExclusions.noun,
          verb: !!parsed.posExclusions.verb,
          adj: !!parsed.posExclusions.adj,
          adv: !!parsed.posExclusions.adv,
          prep: !!parsed.posExclusions.prep,
          particle: !!parsed.posExclusions.particle,
        });
      }
    } catch (e) {
      console.warn('Failed to load persisted settings', e);
    }
  }, []);

  // Persist settings (debounced) whenever they change.
  useEffect(() => {
    const t = window.setTimeout(() => {
      try {
        const payload = {
          customStopWords,
          customFixedWords,
          filterEnglish,
          filterNumbers,
          minWordCount,
          maxWordCount,
          posExclusions,
        };
        localStorage.setItem(SETTINGS_STORAGE_KEY, JSON.stringify(payload));
      } catch (e) {
        console.warn('Failed to persist settings', e);
      }
    }, 250);
    return () => window.clearTimeout(t);
  }, [
    customStopWords,
    customFixedWords,
    filterEnglish,
    filterNumbers,
    minWordCount,
    maxWordCount,
    posExclusions,
  ]);

  useEffect(() => {
    initJieba(jiebaWasmUrl).then(() => {
      setJiebaReady(true);
    }).catch(err => {
      console.error("Failed to init jieba-wasm:", err);
    });
  }, []);

  const fileInputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    if (parsedWorkbook && selectedSheet) {
      const sheetData = parsedWorkbook.sheets[selectedSheet];
      if (sheetData) {
        setParsedData(sheetData);
        setHeaders(sheetData.headers);
      }
    }
  }, [selectedSheet, parsedWorkbook]);

  useEffect(() => {
    if (headers.length > 0) {
        setSelectedColumn(headers[0]);
    } else {
        setSelectedColumn("");
    }
  }, [headers]);

  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const uploadedFile = e.target.files?.[0];
    if (uploadedFile) {
      try {
        setFile(uploadedFile);
        const data = await parseExcelFile(uploadedFile);
        setParsedWorkbook(data);
        setSheetNames(data.sheetNames);
        if (data.sheetNames.length > 0) {
          const firstSheet = data.sheetNames[0];
          setSelectedSheet(firstSheet);
          setParsedData(data.sheets[firstSheet]);
          setHeaders(data.sheets[firstSheet].headers);
        }
        setPreviewMode(null);
        setPreviewData([]);
      } catch (error) {
        console.error("Parse Error:", error);
        alert("Excel 解析失败，请确认文件格式");
      }
    }
  };

  const exportWebExcel = async (data: unknown[], filename: string, sheetName: string) => {
    const ws = XLSX.utils.json_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, sheetName);

    try {
      if (isTauri()) {
        const { save } = await import('@tauri-apps/plugin-dialog');
        const { writeFile } = await import('@tauri-apps/plugin-fs');
        const path = await save({
          defaultPath: filename,
          filters: [{ name: 'Excel', extensions: ['xlsx'] }],
        });
        if (!path) return;
        const out = XLSX.write(wb, { bookType: 'xlsx', type: 'array' }) as ArrayBuffer | Uint8Array;
        const bytes = out instanceof Uint8Array ? out : new Uint8Array(out);
        await writeFile(path, bytes);
      } else {
        XLSX.writeFile(wb, filename);
      }
    } catch (e) {
      console.error("Export failed:", e);
      alert("文件生成失败，请重试");
    }
  };

  const handleAnalyze = async (mode: PreviewMode) => {
    if (!parsedData || !selectedColumn || !mode) return;
    setIsAnalyzing(true);
    setPreviewMode(null);
    
    setTimeout(() => {
        try {
            const textData = parsedData.rows.map(row => String(row[selectedColumn] || ""));
            const stopWords = parseStopWords(customStopWords);
            const fixedCollocations = parseStopWords(customFixedWords);

            if (mode === 'stats') {
               const analysis = analyzeTextLocally(textData, customStopWords, posExclusions, fixedCollocations, filterEnglish, filterNumbers, jiebaReady ? jiebaTag : undefined);
               const pData = analysis.frequencies.slice(0, 30).map((f, i) => ({
                 col1: i + 1,
                 col2: f.word,
                 col3: f.count,
                 col4: f.rowCount,
                 col5: ((f.rowCount / textData.length) * 100).toFixed(2) + "%"
               }));
               setPreviewData(pData as any);
               setPreviewMode('stats');

            } else if (mode === 'segmentation') {
               const pData = [];
               for(let i=0; i<Math.min(30, textData.length); i++) {
                 const original = textData[i];
                 const segments = segmentText(original, stopWords, posExclusions, fixedCollocations, filterEnglish, filterNumbers, jiebaReady ? jiebaTag : undefined);
                 pData.push({
                   col1: i + 1,
                   col2: original.length > 20 ? original.substring(0,20)+"..." : original,
                   col3: segments.length,
                   col4: segments.join(" ").substring(0, 30) + (segments.join(" ").length > 30 ? "..." : "")
                 });
               }
               setPreviewData(pData);
               setPreviewMode('segmentation');
            
            } else if (mode === 'row_keywords') {
               const pData = [];
               for(let i=0; i<Math.min(30, textData.length); i++) {
                   const original = textData[i];
                   const keywords = extractKeywordsLocally(original, stopWords, 5, posExclusions, fixedCollocations, filterEnglish, filterNumbers, jiebaReady ? jiebaTag : undefined);
                   pData.push({
                       col1: i + 1,
                       col2: original.length > 20 ? original.substring(0,20)+"..." : original,
                       col3: keywords.join(", ")
                   });
               }
               setPreviewData(pData);
               setPreviewMode('row_keywords');

            } else if (mode === 'tfidf') {
               const analysis = analyzeTextLocally(textData, customStopWords, posExclusions, fixedCollocations, filterEnglish, filterNumbers, jiebaReady ? jiebaTag : undefined);
               const pData = analysis.frequencies.slice(0, 30).map((f, i) => ({
                 col1: i + 1,
                 col2: f.word,
                 col3: (f.count / analysis.totalWords).toFixed(4),
                 col4: f.rowCount,
                 col5: ((f.rowCount / textData.length) * 100).toFixed(2) + "%"
               }));
               setPreviewData(pData as any);
               setPreviewMode('tfidf');
            } else if (mode === 'entropy') {
               const pData = [];
               const allAnomalies = [];
               for(let i=0; i<textData.length; i++) {
                   const original = textData[i];
                   const segments = segmentText(original, stopWords, posExclusions, fixedCollocations, filterEnglish, filterNumbers, jiebaReady ? jiebaTag : undefined);
                   const wordCount = segments.length;
                   const entropy = calculateEntropy(segments);
                   const isAnomaly = wordCount < minWordCount || wordCount > maxWordCount;
                   
                   if (isAnomaly) {
                       allAnomalies.push({
                           col1: i + 1,
                           col2: original.length > 20 ? original.substring(0,20)+"..." : original,
                           col3: wordCount,
                           col4: entropy.toFixed(4),
                           col5: "异常"
                       });
                   }
               }
               setPreviewData(allAnomalies.slice(0, 30));
               setPreviewMode('entropy');
            } else if (mode === 'ner_person' || mode === 'ner_location') {
               if (!jiebaReady) {
                   alert("分词引擎正在加载中，请稍后再试...");
                   return;
               }
               const allNerData = [];
               const targetTag = mode === 'ner_person' ? 'nr' : 'ns';
               const stopWords = parseStopWords(customStopWords);
               
               for(let i=0; i<textData.length; i++) {
                   const original = textData[i];
                   if (!original.trim()) continue;
                   const tags = jiebaTag(original);
                   
                   const entities = tags
                       .filter((t: any) => t.tag === targetTag && !stopWords.has(t.word))
                       .map((t: any) => `${t.word}(${t.tag})`);
                       
                   if (entities.length > 0) {
                       allNerData.push({
                           col1: i + 1,
                           col2: original.length > 20 ? original.substring(0,20)+"..." : original,
                           col3: entities.join(", ")
                       });
                   }
               }
               setPreviewData(allNerData.slice(0, 30));
               setPreviewMode(mode);
            }

        } catch (err) {
            console.error(err);
            alert("分析出错: " + err);
        } finally {
            setIsAnalyzing(false);
        }
    }, 500);
  };

  const handleExportResult = async () => {
    if (!previewMode || !parsedData || !selectedColumn) return;

    try {
      const textData = parsedData.rows.map(row => String(row[selectedColumn] || ""));
      const stopWords = parseStopWords(customStopWords);
      const fixedCollocations = parseStopWords(customFixedWords);
      const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
      const baseFilename = `${file?.name.split('.')[0]}_${selectedSheet}_${selectedColumn}_${timestamp}`;

      if (previewMode === 'stats') {
         const analysis = analyzeTextLocally(textData, customStopWords, posExclusions, fixedCollocations, filterEnglish, filterNumbers, jiebaReady ? jiebaTag : undefined);
         const exportData = analysis.frequencies.map((item, index) => ({
           "Rank": index + 1,
           "Word": item.word,
           "Count": item.count,
           "Row Count": item.rowCount,
           "Row Percent": ((item.rowCount / textData.length) * 100).toFixed(2) + "%"
         }));
         await exportWebExcel(exportData, `${baseFilename}_stats.xlsx`, "Stats");

      } else if (previewMode === 'segmentation') {
         const exportData = parsedData.rows.map(row => {
            const txt = String(row[selectedColumn] || "");
            const segments = segmentText(txt, stopWords, posExclusions, fixedCollocations, filterEnglish, filterNumbers, jiebaReady ? jiebaTag : undefined);
            return {
              ...row,
              "Word_Count": segments.length,
              "Segmented_Text": segments.join(" ")
            };
         });
         await exportWebExcel(exportData, `${baseFilename}_segmented.xlsx`, "Segmented");

      } else if (previewMode === 'row_keywords') {
         const exportData = parsedData.rows.map(row => {
            const txt = String(row[selectedColumn] || "");
            const keywords = extractKeywordsLocally(txt, stopWords, 5, posExclusions, fixedCollocations, filterEnglish, filterNumbers, jiebaReady ? jiebaTag : undefined);
            return {
              ...row,
              "Keywords_Top5": keywords.join(",")
            };
         });
         await exportWebExcel(exportData, `${baseFilename}_row_keywords.xlsx`, "Keywords");

      } else if (previewMode === 'tfidf') {
         const analysis = analyzeTextLocally(textData, customStopWords, posExclusions, fixedCollocations, filterEnglish, filterNumbers, jiebaReady ? jiebaTag : undefined);
         const exportData = analysis.frequencies.slice(0, 200).map((item, index) => ({
           "Rank": index + 1,
           "Keyword": item.word,
           "Weight": (item.count / analysis.totalWords).toFixed(5),
           "Row Count": item.rowCount,
           "Row Percent": ((item.rowCount / textData.length) * 100).toFixed(2) + "%"
         }));
         await exportWebExcel(exportData, `${baseFilename}_tfidf.xlsx`, "Keywords");
      } else if (previewMode === 'entropy') {
         const exportData: any[] = [];
         parsedData.rows.forEach(row => {
            const txt = String(row[selectedColumn] || "");
            const segments = segmentText(txt, stopWords, posExclusions, fixedCollocations, filterEnglish, filterNumbers, jiebaReady ? jiebaTag : undefined);
            const wordCount = segments.length;
            const entropy = calculateEntropy(segments);
            const isAnomaly = wordCount < minWordCount || wordCount > maxWordCount;
            
            if (isAnomaly) {
                exportData.push({
                  ...row,
                  "Word_Count": wordCount,
                  "Entropy": entropy.toFixed(4),
                  "Status": "异常"
                });
            }
         });
         await exportWebExcel(exportData, `${baseFilename}_entropy.xlsx`, "Entropy");
      } else if (previewMode === 'ner_person' || previewMode === 'ner_location') {
         if (!jiebaReady) {
             alert("分词引擎正在加载中，请稍后再试...");
             return;
         }
         const exportData: any[] = [];
         const targetTag = previewMode === 'ner_person' ? 'nr' : 'ns';
         const columnName = previewMode === 'ner_person' ? 'Entities_Person_nr' : 'Entities_Location_ns';
         const stopWords = parseStopWords(customStopWords);
         
         parsedData.rows.forEach(row => {
            const txt = String(row[selectedColumn] || "");
            if (!txt.trim()) return;
            const tags = jiebaTag(txt);
            const entities = tags
                .filter((t: any) => t.tag === targetTag && !stopWords.has(t.word))
                .map((t: any) => `${t.word}(${t.tag})`);
                
            if (entities.length > 0) {
                exportData.push({
                  ...row,
                  [columnName]: entities.join(", ")
                });
            }
         });
         await exportWebExcel(exportData, `${baseFilename}_${previewMode}.xlsx`, "Entities");
      }
    } catch (e) {
      console.error(e);
      alert("导出失败");
    }
  };

  const runningInTauri = detectTauriRuntime();

  return (
    <div
      className={
        runningInTauri
          ? 'h-screen w-screen flex flex-col overflow-hidden bg-[#f0f2f5] font-sans'
          : 'min-h-screen flex items-center justify-center p-4 font-sans bg-gray-800'
      }
    >
      {/* 历史原因：原网页在页面内模拟“桌面窗口标题栏”，打包成桌面应用后会造成“套框”观感。
          这里统一移除内层假标题栏，避免任何环境下出现双标题栏。 */}
      <div
        className={
          runningInTauri
            ? 'flex min-h-0 min-w-0 flex-1 flex-col overflow-hidden'
            : 'relative flex h-[850px] w-[1000px] flex-col overflow-hidden rounded-sm border border-gray-400 bg-[#f0f2f5] shadow-2xl'
        }
      >
        <div
          className={`flex min-h-0 flex-1 flex-col gap-4 overflow-y-auto ${runningInTauri ? 'p-3' : 'p-4'}`}
        >
          
          {/* Section 1: Data Source */}
          <div className="bg-white p-4 rounded border border-gray-200 shadow-sm">
            <div className="text-sm font-bold text-gray-700 border-b pb-2 mb-3 flex items-center gap-2">
              <span className="bg-blue-100 text-blue-800 px-2 py-0.5 rounded text-xs">第一步</span>
              数据源设置
            </div>
            
            <div className="flex flex-col gap-3">
              <div className="flex items-center gap-3">
                <button onClick={() => fileInputRef.current?.click()} className="px-4 py-1.5 bg-[#e3f2fd] border border-blue-200 text-gray-700 text-sm hover:bg-blue-50 rounded-sm min-w-[120px]">
                  📂 选择 Excel/CSV
                </button>
                <input type="file" ref={fileInputRef} className="hidden" accept=".xlsx,.xls,.csv" onChange={handleFileUpload} />
                <div className="flex-1 bg-gray-50 border border-gray-300 h-8 flex items-center px-2 text-sm text-gray-600 rounded-sm truncate">
                  {file ? file.name : ""}
                </div>
              </div>

              <div className="flex items-center gap-3">
                 <div className="min-w-[120px] text-center text-sm text-gray-600">分析目标 Sheet：</div>
                 <select 
                   value={selectedSheet} 
                   onChange={(e) => setSelectedSheet(e.target.value)} 
                   disabled={sheetNames.length === 0} 
                   className="flex-1 border border-gray-300 h-8 px-2 text-sm rounded-sm bg-white text-gray-900 focus:outline-none focus:border-blue-500"
                 >
                   {sheetNames.length === 0 && <option value="">请先上传文件...</option>}
                   {sheetNames.map(s => <option key={s} value={s}>{s}</option>)}
                 </select>
              </div>

              <div className="flex items-center gap-3">
                 <div className="min-w-[120px] text-center text-sm text-gray-600">分析目标列：</div>
                 <select 
                   value={selectedColumn} 
                   onChange={(e) => setSelectedColumn(e.target.value)} 
                   disabled={headers.length === 0} 
                   className="flex-1 border border-gray-300 h-8 px-2 text-sm rounded-sm bg-white text-gray-900 focus:outline-none focus:border-blue-500"
                 >
                   {headers.length === 0 && <option value="">请先上传文件...</option>}
                   {headers.map(h => <option key={h} value={h}>{h}</option>)}
                 </select>
              </div>
            </div>
          </div>

          {/* Section 2: Stop Words & POS Exclusion Settings */}
          <div className="bg-white p-4 rounded border border-gray-200 shadow-sm">
            <div className="text-sm font-bold text-gray-700 border-b pb-2 mb-2 flex items-center gap-2">
              <span className="bg-blue-100 text-blue-800 px-2 py-0.5 rounded text-xs">第二步</span>
              停用词与词性设置
            </div>
            
            {/* POS Filter - Exclusion Mode */}
            <div className="mb-4 bg-gray-50 p-2 rounded border border-gray-100">
               <div className="text-xs font-bold text-gray-600 mb-2 flex items-center gap-2">
                 <span className="text-red-500">【词性停用】</span> 
                 <span className="font-normal">勾选下方词性，将在统计中<strong className="text-gray-800">排除</strong>这些类型的词汇。</span>
               </div>
               <div className="grid grid-cols-6 gap-2">
                  <label className="flex items-center gap-1 text-sm text-gray-700 cursor-pointer select-none">
                     <input type="checkbox" checked={posExclusions.noun} onChange={e => setPosExclusions({...posExclusions, noun: e.target.checked})} className="rounded text-blue-600 focus:ring-blue-500"/>
                     <span>名词 (n*)</span>
                  </label>
                  <label className="flex items-center gap-1 text-sm text-gray-700 cursor-pointer select-none">
                     <input type="checkbox" checked={posExclusions.verb} onChange={e => setPosExclusions({...posExclusions, verb: e.target.checked})} className="rounded text-blue-600 focus:ring-blue-500"/>
                     <span>动词 (v*)</span>
                  </label>
                  <label className="flex items-center gap-1 text-sm text-gray-700 cursor-pointer select-none">
                     <input type="checkbox" checked={posExclusions.adj} onChange={e => setPosExclusions({...posExclusions, adj: e.target.checked})} className="rounded text-blue-600 focus:ring-blue-500"/>
                     <span>形容词 (a*)</span>
                  </label>
                  <label className="flex items-center gap-1 text-sm text-gray-700 cursor-pointer select-none">
                     <input type="checkbox" checked={posExclusions.adv} onChange={e => setPosExclusions({...posExclusions, adv: e.target.checked})} className="rounded text-blue-600 focus:ring-blue-500"/>
                     <span>副词 (d*)</span>
                  </label>
                   <label className="flex items-center gap-1 text-sm text-gray-700 cursor-pointer select-none">
                     <input type="checkbox" checked={posExclusions.prep} onChange={e => setPosExclusions({...posExclusions, prep: e.target.checked})} className="rounded text-blue-600 focus:ring-blue-500"/>
                     <span>介词 (p*)</span>
                  </label>
                   <label className="flex items-center gap-1 text-sm text-gray-700 cursor-pointer select-none">
                     <input type="checkbox" checked={posExclusions.particle} onChange={e => setPosExclusions({...posExclusions, particle: e.target.checked})} className="rounded text-blue-600 focus:ring-blue-500"/>
                     <span>助词 (u*)</span>
                  </label>
               </div>
            </div>

            {/* Specific Stop Words */}
            <div className="mb-4">
               <div className="text-xs font-bold text-gray-600 mb-1">
                 <span className="text-red-500">【具体词停用】</span> 每行一个，或用中文/英文逗号分隔
               </div>
               <textarea 
                  value={customStopWords}
                  onChange={(e) => setCustomStopWords(e.target.value)}
                  className="w-full h-16 border border-gray-300 rounded p-2 text-sm font-mono focus:border-blue-500 outline-none resize-none"
                  placeholder="输入停用词，例如：的,了,是"
                />
            </div>

            {/* Fixed Collocations */}
            <div className="mb-4">
               <div className="text-xs font-bold text-gray-600 mb-1">
                 <span className="text-blue-500">【固定搭配】</span> 不进行拆分的词汇，每行一个，或用中文/英文逗号分隔
               </div>
               <textarea 
                  value={customFixedWords}
                  onChange={(e) => setCustomFixedWords(e.target.value)}
                  className="w-full h-16 border border-gray-300 rounded p-2 text-sm font-mono focus:border-blue-500 outline-none resize-none"
                  placeholder="输入固定搭配，例如：人工智能,机器学习"
                />
            </div>

            {/* Other Filters */}
            <div className="bg-gray-50 p-2 rounded border border-gray-100">
               <div className="text-xs font-bold text-gray-600 mb-2 flex items-center gap-2">
                 <span className="text-purple-500">【其他过滤】</span> 
                 <span className="font-normal">勾选下方选项，将在统计中<strong className="text-gray-800">排除</strong>这些类型的内容。</span>
               </div>
               <div className="flex gap-4">
                  <label className="flex items-center gap-1 text-sm text-gray-700 cursor-pointer select-none">
                     <input type="checkbox" checked={filterEnglish} onChange={e => setFilterEnglish(e.target.checked)} className="rounded text-blue-600 focus:ring-blue-500"/>
                     <span>过滤英文</span>
                  </label>
                  <label className="flex items-center gap-1 text-sm text-gray-700 cursor-pointer select-none">
                     <input type="checkbox" checked={filterNumbers} onChange={e => setFilterNumbers(e.target.checked)} className="rounded text-blue-600 focus:ring-blue-500"/>
                     <span>过滤数字</span>
                  </label>
               </div>
            </div>

            {/* Length & Entropy Settings */}
            <div className="flex gap-4 mt-4">
              <div className="flex-1 bg-gray-50 p-2 rounded border border-gray-100">
                 <div className="text-xs font-bold text-gray-600 mb-2 flex items-center gap-2">
                   <span className="text-teal-500">【异常长度检测】</span>
                   <span className="font-normal">设置词数阈值，用于信息熵与长度分析标记异常。</span>
                 </div>
                 <div className="flex gap-4 items-center">
                    <label className="flex items-center gap-2 text-sm text-gray-700">
                       <span>最小词数:</span>
                       <input type="number" value={minWordCount} onChange={e => setMinWordCount(Number(e.target.value))} className="w-16 border border-gray-300 rounded px-1 text-center h-7 focus:outline-none focus:border-blue-500" />
                    </label>
                    <label className="flex items-center gap-2 text-sm text-gray-700">
                       <span>最大词数:</span>
                       <input type="number" value={maxWordCount} onChange={e => setMaxWordCount(Number(e.target.value))} className="w-16 border border-gray-300 rounded px-1 text-center h-7 focus:outline-none focus:border-blue-500" />
                    </label>
                 </div>
              </div>
            </div>
          </div>

          {/* Section 3: Results & Actions */}
          <div className="flex-1 flex gap-4 h-[520px] min-h-[520px]">
            
            {/* Left: Result Preview Panel */}
            <div className="flex-1 bg-white border border-gray-300 rounded flex flex-col shadow-sm">
               {/* Panel Header */}
               <div className="bg-gray-50 px-4 py-3 border-b border-gray-200 flex items-center justify-between">
                 <div className="font-bold text-sm text-gray-800 flex items-center gap-2">
                    <Table size={16} className="text-gray-500" />
                    <span>结果预览</span>
                    {previewMode && (
                        <span className="text-xs bg-blue-100 text-blue-800 px-2 py-0.5 rounded-full font-medium">
                            {previewMode === 'stats' && "全局统计 Top 30"}
                            {previewMode === 'segmentation' && "分词预览 Top 30"}
                            {previewMode === 'row_keywords' && "逐行关键词 Top 30"}
                            {previewMode === 'tfidf' && "全局关键词权重 Top 30"}
                            {previewMode === 'entropy' && "信息熵与长度 Top 30"}
                            {previewMode === 'ner_person' && "人名提取 Top 30"}
                            {previewMode === 'ner_location' && "地名提取 Top 30"}
                        </span>
                    )}
                 </div>
                 
                 <button 
                   onClick={() => void handleExportResult()}
                   disabled={!previewMode}
                   className={`px-3 py-1.5 text-xs font-bold rounded flex items-center gap-1 transition-colors ${previewMode ? 'bg-blue-600 text-white hover:bg-blue-700 shadow-sm' : 'bg-gray-200 text-gray-400 cursor-not-allowed'}`}
                 >
                   <Download size={14} /> 导出完整文件
                 </button>
               </div>

               {/* Panel Content (Table) */}
               <div className="flex-1 overflow-y-auto relative bg-white">
                 {isAnalyzing ? (
                    <div className="absolute inset-0 flex flex-col items-center justify-center text-gray-500">
                      <div className="w-8 h-8 border-4 border-gray-200 border-t-blue-500 rounded-full animate-spin mb-2"></div>
                      <span className="text-sm">分析中...</span>
                    </div>
                 ) : !previewMode ? (
                   <div className="absolute inset-0 flex items-center justify-center text-gray-400 text-sm flex-col gap-2">
                     <FileText size={40} className="text-gray-300"/>
                     <p>请在右侧选择一项分析功能</p>
                   </div>
                 ) : (
                   <table className="w-full text-sm text-left border-collapse">
                     <thead className="bg-gray-50 sticky top-0 text-gray-600 text-xs uppercase font-semibold">
                       <tr>
                         <th className="px-4 py-2 border-b w-16 text-center">
                            {(previewMode === 'segmentation' || previewMode === 'row_keywords' || previewMode === 'entropy' || previewMode === 'ner_person' || previewMode === 'ner_location') ? '行号' : 'Rank'}
                         </th>
                         <th className="px-4 py-2 border-b">
                            {previewMode === 'stats' && "词语 (Word)"}
                            {(previewMode === 'segmentation' || previewMode === 'row_keywords' || previewMode === 'entropy' || previewMode === 'ner_person' || previewMode === 'ner_location') && "原文 (Original)"}
                            {previewMode === 'tfidf' && "关键词 (Keyword)"}
                         </th>
                         <th className="px-4 py-2 border-b w-1/3">
                            {previewMode === 'stats' && <div className="text-center">频次 (Count)</div>}
                            {previewMode === 'segmentation' && <div className="text-center">分词数量 (Word Count)</div>}
                            {previewMode === 'row_keywords' && "提取关键词 (Keywords)"}
                            {previewMode === 'tfidf' && <div className="text-center">权重 (Weight)</div>}
                            {previewMode === 'entropy' && <div className="text-center">词数 (Word Count)</div>}
                            {previewMode === 'ner_person' && "人名 (Person)"}
                            {previewMode === 'ner_location' && "地名 (Location)"}
                         </th>
                         {previewMode === 'segmentation' && (
                           <th className="px-4 py-2 border-b w-1/3">分词结果 (Segmented)</th>
                         )}
                         {(previewMode === 'stats' || previewMode === 'tfidf') && (
                           <>
                             <th className="px-4 py-2 border-b text-center">行数 (Row Count)</th>
                             <th className="px-4 py-2 border-b text-center">行占比 (Row %)</th>
                           </>
                         )}
                         {previewMode === 'entropy' && (
                           <>
                             <th className="px-4 py-2 border-b text-center">信息熵 (Entropy)</th>
                             <th className="px-4 py-2 border-b text-center">状态 (Status)</th>
                           </>
                         )}
                       </tr>
                     </thead>
                     <tbody className="divide-y divide-gray-100">
                       {previewData.map((row, idx) => (
                         <tr key={idx} className="hover:bg-blue-50 transition-colors">
                           <td className="px-4 py-2 text-center text-gray-500 font-mono text-xs">{row.col1}</td>
                           <td className="px-4 py-2 text-gray-800 break-all">{row.col2}</td>
                           <td className={`px-4 py-2 font-medium break-all ${
                               (previewMode === 'row_keywords' || previewMode === 'ner_person' || previewMode === 'ner_location') 
                               ? 'text-xs text-gray-600' 
                               : 'text-center text-blue-600'
                           }`}>
                             {row.col3}
                           </td>
                           {previewMode === 'segmentation' && (
                             <td className="px-4 py-2 text-xs text-gray-600 break-all">{row.col4}</td>
                           )}
                           {(previewMode === 'stats' || previewMode === 'tfidf') && (
                             <>
                               <td className="px-4 py-2 text-center text-gray-600">{row.col4}</td>
                               <td className={`px-4 py-2 text-center text-gray-600`}>{row.col5}</td>
                             </>
                           )}
                           {previewMode === 'entropy' && (
                             <>
                               <td className="px-4 py-2 text-center text-gray-600">{row.col4}</td>
                               <td className={`px-4 py-2 text-center ${row.col5 === '异常' ? 'text-red-500 font-bold' : 'text-green-600'}`}>{row.col5}</td>
                             </>
                           )}
                         </tr>
                       ))}
                     </tbody>
                   </table>
                 )}
               </div>
            </div>

            {/* Right: Action Buttons */}
            <div className="w-[240px] flex flex-col gap-4">
               
               <div className="flex-1 flex flex-col gap-3">
                   <div className="font-bold text-sm text-gray-700 border-b border-gray-300 pb-2">分析功能区</div>

                   <button 
                     onClick={() => handleAnalyze('stats')}
                     disabled={isAnalyzing || !file}
                     className={`w-full py-3 text-white font-bold rounded shadow-sm flex items-center px-4 gap-3 transition-transform active:scale-95 ${isAnalyzing || !file ? 'bg-gray-400' : 'bg-[#4caf50] hover:bg-[#43a047]'}`}
                   >
                     <Play size={20} /> 
                     <div className="flex flex-col items-start">
                        <span className="text-sm">全局词频统计</span>
                        <span className="text-[10px] font-normal opacity-90">预览 Top 30</span>
                     </div>
                   </button>

                   <button 
                     onClick={() => handleAnalyze('segmentation')}
                     disabled={isAnalyzing || !file}
                     className={`w-full py-3 text-white font-bold rounded shadow-sm flex items-center px-4 gap-3 transition-transform active:scale-95 ${isAnalyzing || !file ? 'bg-gray-400' : 'bg-[#ff9800] hover:bg-[#fb8c00]'}`}
                   >
                     <List size={20} /> 
                     <div className="flex flex-col items-start">
                        <span className="text-sm">逐行分词预览</span>
                        <span className="text-[10px] font-normal opacity-90">预览前 30 行</span>
                     </div>
                   </button>

                   <button 
                     onClick={() => handleAnalyze('row_keywords')}
                     disabled={isAnalyzing || !file}
                     className={`w-full py-3 text-white font-bold rounded shadow-sm flex items-center px-4 gap-3 transition-transform active:scale-95 ${isAnalyzing || !file ? 'bg-gray-400' : 'bg-[#009688] hover:bg-[#00796b]'}`}
                   >
                     <FileText size={20} /> 
                     <div className="flex flex-col items-start">
                        <span className="text-sm">逐行关键词提取</span>
                        <span className="text-[10px] font-normal opacity-90">预览前 30 行</span>
                     </div>
                   </button>

                   <button 
                     onClick={() => handleAnalyze('tfidf')}
                     disabled={isAnalyzing || !file}
                     className={`w-full py-3 text-white font-bold rounded shadow-sm flex items-center px-4 gap-3 transition-transform active:scale-95 ${isAnalyzing || !file ? 'bg-gray-400' : 'bg-[#9c27b0] hover:bg-[#7b1fa2]'}`}
                   >
                     <ChartBar size={20} /> 
                     <div className="flex flex-col items-start">
                        <span className="text-sm">全局 TF-IDF 关键词</span>
                        <span className="text-[10px] font-normal opacity-90">预览 Top 30</span>
                     </div>
                   </button>

                   <button 
                     onClick={() => handleAnalyze('entropy')}
                     disabled={isAnalyzing || !file}
                     className={`w-full py-3 text-white font-bold rounded shadow-sm flex items-center px-4 gap-3 transition-transform active:scale-95 ${isAnalyzing || !file ? 'bg-gray-400' : 'bg-[#e91e63] hover:bg-[#c2185b]'}`}
                   >
                     <FileText size={20} /> 
                     <div className="flex flex-col items-start">
                        <span className="text-sm">信息熵与长度分析</span>
                        <span className="text-[10px] font-normal opacity-90">预览前 30 行</span>
                     </div>
                   </button>

                   <button 
                     onClick={() => handleAnalyze('ner_person')}
                     disabled={isAnalyzing || !file}
                     className={`w-full py-3 text-white font-bold rounded shadow-sm flex items-center px-4 gap-3 transition-transform active:scale-95 ${isAnalyzing || !file ? 'bg-gray-400' : 'bg-[#3f51b5] hover:bg-[#303f9f]'}`}
                   >
                     <User size={20} /> 
                     <div className="flex flex-col items-start">
                        <span className="text-sm">提取人名</span>
                        <span className="text-[10px] font-normal opacity-90">基于 jieba-wasm (nr)</span>
                     </div>
                   </button>

                   <button 
                     onClick={() => handleAnalyze('ner_location')}
                     disabled={isAnalyzing || !file}
                     className={`w-full py-3 text-white font-bold rounded shadow-sm flex items-center px-4 gap-3 transition-transform active:scale-95 ${isAnalyzing || !file ? 'bg-gray-400' : 'bg-[#00bcd4] hover:bg-[#0097a7]'}`}
                   >
                     <MapPin size={20} /> 
                     <div className="flex flex-col items-start">
                        <span className="text-sm">提取地名</span>
                        <span className="text-[10px] font-normal opacity-90">基于 jieba-wasm (ns)</span>
                     </div>
                   </button>
                   
               </div>

            </div>

          </div>
        </div>
      </div>
    </div>
  );
}