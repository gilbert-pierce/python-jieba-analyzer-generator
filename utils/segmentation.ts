import { WordFrequency, AnalysisResult } from '../types';
import { PosExclusions } from '../App';

export const parseStopWords = (input: string): Set<string> => {
  const raw = input.split(/[\n,，\s]+/);
  return new Set(raw.map(w => w.trim()).filter(w => w.length > 0));
};

export const segmentText = (
  text: string, 
  stopWords: Set<string>, 
  posExclusions?: PosExclusions,
  fixedCollocations?: Set<string>,
  filterEnglish?: boolean,
  filterNumbers?: boolean,
  tagger?: (text: string) => {word: string, tag: string}[]
): string[] => {
  let segments: string[] = [];
  
  // Extract fixed collocations first to prevent them from being split
  let remainingText = text;
  const extractedFixed: string[] = [];
  
  if (fixedCollocations && fixedCollocations.size > 0) {
    const sortedFixed = Array.from(fixedCollocations).sort((a, b) => b.length - a.length);
    for (const fixed of sortedFixed) {
      let idx = remainingText.indexOf(fixed);
      while (idx !== -1) {
        extractedFixed.push(fixed);
        remainingText = remainingText.substring(0, idx) + " " + remainingText.substring(idx + fixed.length);
        idx = remainingText.indexOf(fixed);
      }
    }
  }

  // Use tagger if available for POS filtering
  if (tagger && posExclusions) {
      const tags = tagger(remainingText);
      for (const t of tags) {
          let exclude = false;
          if (posExclusions.noun && t.tag.startsWith('n')) exclude = true;
          if (posExclusions.verb && t.tag.startsWith('v')) exclude = true;
          if (posExclusions.adj && t.tag.startsWith('a')) exclude = true;
          if (posExclusions.adv && t.tag.startsWith('d')) exclude = true;
          if (posExclusions.prep && t.tag.startsWith('p')) exclude = true;
          if (posExclusions.particle && t.tag.startsWith('u')) exclude = true;
          
          if (!exclude) {
              segments.push(t.word);
          }
      }
  } else if (typeof Intl !== 'undefined' && (Intl as any).Segmenter) {
       // Use Intl.Segmenter if available (modern browsers)
       const Segmenter = (Intl as any).Segmenter;
       const segmenter = new Segmenter('zh-CN', { granularity: 'word' });
       const rawSegments = segmenter.segment(remainingText);
       for(const s of rawSegments) {
           segments.push(s.segment);
       }
  } else {
       // Fallback
       segments = remainingText.split('');
  }
  
  // Add back the fixed collocations
  segments.push(...extractedFixed);

  return segments.filter(w => {
      const clean = w.trim();
      const isChinese = /[\u4e00-\u9fa5]/.test(clean);
      if (clean.length < 2 && !isChinese && (!fixedCollocations || !fixedCollocations.has(clean))) return false;
      if (stopWords.has(clean)) return false;
      
      if (filterEnglish && /^[a-zA-Z]+$/.test(clean)) return false;
      if (filterNumbers && /^[0-9]+(\.[0-9]+)?$/.test(clean)) return false;
      
      return true;
  });
};

export const calculateEntropy = (words: string[]): number => {
  if (words.length === 0) return 0;
  const counts: Record<string, number> = {};
  words.forEach(w => counts[w] = (counts[w] || 0) + 1);
  let entropy = 0;
  const total = words.length;
  for (const w in counts) {
      const p = counts[w] / total;
      entropy -= p * Math.log2(p);
  }
  return entropy;
};

export const extractKeywordsLocally = (
  text: string, 
  stopWords: Set<string>, 
  topK: number = 5, 
  posExclusions?: PosExclusions,
  fixedCollocations?: Set<string>,
  filterEnglish?: boolean,
  filterNumbers?: boolean,
  tagger?: (text: string) => {word: string, tag: string}[]
): string[] => {
    const segments = segmentText(text, stopWords, posExclusions, fixedCollocations, filterEnglish, filterNumbers, tagger);
    
    // De-duplicate
    const uniqueWords = Array.from(new Set(segments));
    
    // Sort by length descending, then alphabetical
    uniqueWords.sort((a, b) => b.length - a.length || a.localeCompare(b));
    
    return uniqueWords.slice(0, topK);
};

export const analyzeTextLocally = (
  textData: string[], 
  customStopWordsString: string, 
  posExclusions?: PosExclusions,
  fixedCollocations?: Set<string>,
  filterEnglish?: boolean,
  filterNumbers?: boolean,
  tagger?: (text: string) => {word: string, tag: string}[]
): AnalysisResult => {
  const STOP_WORDS = parseStopWords(customStopWordsString);
  const frequencies: Record<string, number> = {};
  const rowCounts: Record<string, Set<number>> = {};
  let totalWords = 0;

  // Process all text
  textData.forEach((text, rowIndex) => {
      const words = segmentText(text, STOP_WORDS, posExclusions, fixedCollocations, filterEnglish, filterNumbers, tagger);
      words.forEach(cleanWord => {
          frequencies[cleanWord] = (frequencies[cleanWord] || 0) + 1;
          if (!rowCounts[cleanWord]) rowCounts[cleanWord] = new Set();
          rowCounts[cleanWord].add(rowIndex);
          totalWords++;
      });
  });

  // Sort and format
  const sortedFrequencies: WordFrequency[] = Object.entries(frequencies)
    .map(([word, count]) => ({ 
      word, 
      count,
      rowCount: rowCounts[word] ? rowCounts[word].size : 0
    }))
    .sort((a, b) => b.count - a.count)
    .slice(0, 100); // Top 100

  return {
    totalWords,
    uniqueWords: Object.keys(frequencies).length,
    frequencies: sortedFrequencies,
    summary: "本地快速分析完成 (Web预览模式)。"
  };
};