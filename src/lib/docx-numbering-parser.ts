/**
 * Hybrid DOCX Quiz Parser
 * ------------------------------------------------------------
 * Strategy:
 *  1. Baca langsung word/document.xml + word/numbering.xml + word/styles.xml
 *  2. Classify setiap paragraph: soal vs opsi vs ignore
 *     - numPr ilvl 0 with numeric format = SOAL
 *     - numPr ilvl 1 = OPSI
 *     - numPr ilvl 0 with alpha format (a, A, etc.) = OPSI
 *     - Different numId following a soal = OPSI (contextual)
 *     - Regex fallback buat manual numbering (1. 2. / A. B.)
 *  3. Bold detection langsung dari XML <w:b/>
 *
 * Satu pass — dokumen campuran (auto-numbering + manual) tetep jalan.
 */

import JSZip from "jszip";
import { XMLParser } from "fast-xml-parser";

// ---------- Types ----------

export interface ParsedParagraph {
  text: string;
  isBold: boolean;
  numId: number | null;
  ilvl: number | null;
  hasStyleNumbering: boolean;
  numFormat: string | null; // "decimal", "lowerLetter", "upperLetter", etc.
}

export interface QuizItem {
  soal: string;
  opsi: string[];
  jawabanIndex?: number;
}

// ---------- Step 1: Unzip & parse XML ----------

async function loadDocxXml(buffer: Buffer) {
  const zip = await JSZip.loadAsync(buffer);

  const documentXmlFile = zip.file("word/document.xml");
  const numberingXmlFile = zip.file("word/numbering.xml");
  const stylesXmlFile = zip.file("word/styles.xml");

  if (!documentXmlFile) {
    throw new Error(
      "word/document.xml tidak ditemukan — file .docx corrupt atau bukan docx valid."
    );
  }

  const documentXml = await documentXmlFile.async("text");
  const numberingXml = numberingXmlFile
    ? await numberingXmlFile.async("text")
    : null;
  const stylesXml = stylesXmlFile ? await stylesXmlFile.async("text") : null;

  const parser = new XMLParser({
    ignoreAttributes: false,
    attributeNamePrefix: "@_",
    preserveOrder: false,
    trimValues: false,
    parseTagValue: false,
  });

  const documentObj = parser.parse(documentXml);
  const numberingObj = numberingXml ? parser.parse(numberingXml) : null;
  const stylesObj = stylesXml ? parser.parse(stylesXml) : null;

  return { documentObj, numberingObj, stylesObj };
}

function extractStyleNumberingMap(stylesObj: unknown): Set<string> {
  const result = new Set<string>();
  if (!stylesObj) return result;

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const styles = toArray((stylesObj as any)?.["w:styles"]?.["w:style"]);

  for (const style of styles) {
    const styleId = style?.["@_w:styleId"];
    if (!styleId) continue;

    const numPr = style?.["w:pPr"]?.["w:numPr"];
    if (numPr) {
      result.add(styleId);
    }
  }

  return result;
}

// ---------- Numbering Format Map ----------

/**
 * Builds a map: "numId:ilvl" -> numFmt (e.g. "decimal", "lowerLetter", "upperLetter", "lowerRoman", etc.)
 * This is used to distinguish question numbering (decimal) from option numbering (letter).
 *
 * Word numbering.xml structure:
 *  <w:abstractNum w:abstractNumId="0">
 *    <w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/></w:lvl>
 *  </w:abstractNum>
 *  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
 */
// eslint-disable-next-line @typescript-eslint/no-explicit-any
function buildNumberingFormatMap(numberingObj: any): Map<string, string> {
  const formatMap = new Map<string, string>();
  if (!numberingObj) return formatMap;

  const numbering = numberingObj["w:numbering"];
  if (!numbering) return formatMap;

  // Step 1: Build abstractNumId -> lvl formats
  const abstractNums = toArray(numbering["w:abstractNum"]);
  const abstractMap = new Map<string, Map<number, string>>(); // abstractNumId -> (ilvl -> numFmt)

  for (const abs of abstractNums) {
    const absId = abs?.["@_w:abstractNumId"];
    if (absId === undefined) continue;

    const lvls = toArray(abs["w:lvl"]);
    const lvlMap = new Map<number, string>();

    for (const lvl of lvls) {
      const ilvl = Number(lvl?.["@_w:ilvl"] ?? 0);
      const numFmt = lvl?.["w:numFmt"]?.["@_w:val"] || "decimal";
      lvlMap.set(ilvl, numFmt);
    }

    abstractMap.set(String(absId), lvlMap);
  }

  // Step 2: Map numId -> abstractNumId, then resolve formats
  const nums = toArray(numbering["w:num"]);
  for (const num of nums) {
    const numId = num?.["@_w:numId"];
    if (numId === undefined) continue;

    const absIdRef = num?.["w:abstractNumId"]?.["@_w:val"];
    if (absIdRef === undefined) continue;

    const lvlMap = abstractMap.get(String(absIdRef));
    if (!lvlMap) continue;

    for (const entry of Array.from(lvlMap.entries())) {
      const ilvl = entry[0];
      const fmt = entry[1];
      formatMap.set(`${numId}:${ilvl}`, fmt);
    }
  }

  return formatMap;
}

// ---------- Helpers ----------

function toArray<T>(val: T | T[] | undefined | null): T[] {
  if (val === undefined || val === null) return [];
  return Array.isArray(val) ? val : [val];
}

/**
 * Check if a run is bold by examining multiple bold indicators in rPr:
 * - <w:b/> (standard bold)
 * - <w:b w:val="true"/> or <w:b w:val="1"/> or <w:b/> (no val = means bold)
 * - <w:bCs/> (bold complex script — used for some languages/fonts)
 * - Inherited from paragraph-level rPr (handled at paragraph level)
 *
 * Also handles the case where fast-xml-parser returns:
 * - bNode = "" (empty string for self-closing <w:b/>)
 * - bNode = {} (empty object)
 * - bNode = null vs undefined
 */
// eslint-disable-next-line @typescript-eslint/no-explicit-any
function checkRunBold(rPr: any): boolean {
  if (!rPr) return false;

  // Check <w:b/>
  const bNode = rPr["w:b"];
  if (bNode !== undefined) {
    // <w:b/> with no val attribute means bold=true
    if (bNode === "" || bNode === null || bNode === true) return true;
    if (typeof bNode === "object") {
      const val = bNode["@_w:val"];
      // No val attribute = bold
      if (val === undefined) return true;
      // Explicit true/1
      if (val === "1" || val === "true" || val === true) return true;
      // Explicit false/0 = not bold
      if (val === "0" || val === "false" || val === false) return false;
      // Any other value (rare) = treat as bold
      return true;
    }
    // bNode is a truthy primitive (number, non-empty string, etc)
    return true;
  }

  // NOTE: Do NOT check <w:bCs/> (bold complex script) — this is used for
  // bidirectional/complex script text and does NOT represent visual bold
  // that the user sees. Only <w:b/> represents actual visual bold.

  return false;
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function extractParagraphs(documentObj: any): any[] {
  const body = documentObj?.["w:document"]?.["w:body"];
  if (!body) return [];
  
  // Collect paragraphs from top-level and from nested structures (sdt, tbl, tc, etc.)
  const result: any[] = [];
  collectParagraphs(body, result);
  return result;
}

/**
 * Recursively collect all <w:p> elements from the document body,
 * including those nested inside <w:sdt>, <w:tbl>, <w:tc>, etc.
 */
// eslint-disable-next-line @typescript-eslint/no-explicit-any
function collectParagraphs(node: any, result: any[]): void {
  if (!node || typeof node !== "object") return;

  // Direct paragraphs
  const ps = toArray(node["w:p"]);
  for (const p of ps) {
    result.push(p);
  }

  // Paragraphs inside structured document tags (w:sdt -> w:sdtContent)
  const sdts = toArray(node["w:sdt"]);
  for (const sdt of sdts) {
    const sdtContent = sdt?.["w:sdtContent"];
    if (sdtContent) {
      collectParagraphs(sdtContent, result);
    }
  }

  // Paragraphs inside tables (w:tbl -> w:tr -> w:tc -> w:p)
  const tbls = toArray(node["w:tbl"]);
  for (const tbl of tbls) {
    const trs = toArray(tbl["w:tr"]);
    for (const tr of trs) {
      const tcs = toArray(tr["w:tc"]);
      for (const tc of tcs) {
        collectParagraphs(tc, result);
      }
    }
  }
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function extractParagraphText(p: any): { text: string; isBold: boolean } {
  const runs = toArray(p["w:r"]);
  if (runs.length === 0) {
    return { text: "", isBold: false };
  }

  let combinedText = "";
  let boldRunCount = 0;
  let textRunCount = 0;

  let boldCharCount = 0;
  let totalCharCount = 0;

  for (const run of runs) {
    const tNode = run["w:t"];
    if (tNode === undefined) continue;

    let runText = "";
    if (typeof tNode === "string") {
      runText = tNode;
    } else if (typeof tNode === "number") {
      runText = String(tNode);
    } else if (tNode && typeof tNode === "object" && tNode["#text"] !== undefined) {
      runText = String(tNode["#text"]);
    }

    if (runText === "") continue;

    textRunCount++;
    combinedText += runText;

    const rPr = run["w:rPr"];
    const isRunBold = checkRunBold(rPr);

    // Count characters for bold detection, EXCLUDING option prefix runs (A. B. C. D.)
    // These prefixes are often not bold even when the answer text is bold
    const trimmedRun = runText.trim();
    const isOptionPrefix = /^[A-Da-d]\.\s*$/.test(trimmedRun) || /^[A-Da-d]$/.test(trimmedRun) || /^\.\s*$/.test(trimmedRun);
    
    if (!isOptionPrefix) {
      const charCount = runText.replace(/\s/g, "").length;
      totalCharCount += charCount;
      if (isRunBold) boldCharCount += charCount;
    }

    if (isRunBold) boldRunCount++;
  }

  // Bold if more than 50% of non-prefix, non-space characters are in bold runs
  // If no non-prefix chars found, fall back to run-based check
  const isBold = totalCharCount > 0 
    ? boldCharCount > totalCharCount * 0.5
    : (textRunCount > 0 && boldRunCount === textRunCount);
  return { text: combinedText.trim(), isBold };
}

/**
 * Check if rPr explicitly sets bold to OFF (w:b val="0" or val="false")
 */
// eslint-disable-next-line @typescript-eslint/no-explicit-any
function hasExplicitBoldOff(rPr: any): boolean {
  if (!rPr) return false;
  const bNode = rPr["w:b"];
  if (bNode !== undefined && typeof bNode === "object") {
    const val = bNode["@_w:val"];
    if (val === "0" || val === "false" || val === false) return true;
  }
  return false;
}

/**
 * Extract multiple text segments from a single paragraph, splitting on <w:br/> (soft line breaks).
 * Each segment gets its own bold status.
 * This handles the case where options A/B/C/D are separated by Shift+Enter within one <w:p>.
 *
 * Also handles the case where <w:br/> is a direct child of <w:p> (between runs).
 */
// eslint-disable-next-line @typescript-eslint/no-explicit-any
function extractParagraphSegments(p: any): { text: string; isBold: boolean }[] {
  // Collect all runs including those inside w:hyperlink
  let runs = toArray(p["w:r"]);
  const hyperlinks = toArray(p["w:hyperlink"]);
  for (const hl of hyperlinks) {
    runs = runs.concat(toArray(hl["w:r"]));
  }
  
  if (runs.length === 0) {
    return [];
  }

  // First pass: build an array of "chunks" with text + bold + isBr flag
  // This preserves per-run bold info for later splitting
  const chunks: { text: string; isBold: boolean; isBr: boolean }[] = [];

  for (const run of runs) {
    // Check if run contains a <w:br/> element (line break)
    const brNode = run["w:br"];
    const hasBr = brNode !== undefined;

    if (hasBr) {
      chunks.push({ text: "", isBold: false, isBr: true });
    }

    const tNode = run["w:t"];
    if (tNode === undefined) continue;

    let runText = "";
    if (typeof tNode === "string") {
      runText = tNode;
    } else if (typeof tNode === "number") {
      runText = String(tNode);
    } else if (Array.isArray(tNode)) {
      for (const t of tNode) {
        if (typeof t === "string") runText += t;
        else if (typeof t === "number") runText += String(t);
        else if (t && typeof t === "object" && t["#text"] !== undefined) runText += String(t["#text"]);
      }
    } else if (tNode && typeof tNode === "object" && tNode["#text"] !== undefined) {
      runText = String(tNode["#text"]);
    }

    if (runText === "") continue;

    const rPr = run["w:rPr"];
    const isRunBold = checkRunBold(rPr);

    chunks.push({ text: runText, isBold: isRunBold, isBr: false });
  }

  // Second pass: split chunks into segments by <w:br/>
  const segments: { text: string; isBold: boolean }[] = [];
  let segChunks: { text: string; isBold: boolean; isBr: boolean }[] = [];

  for (const chunk of chunks) {
    if (chunk.isBr) {
      // Flush current segment
      const seg = flushChunksToSegment(segChunks);
      if (seg) segments.push(seg);
      segChunks = [];
    } else {
      segChunks.push(chunk);
    }
  }
  // Flush last segment
  const lastSeg = flushChunksToSegment(segChunks);
  if (lastSeg) segments.push(lastSeg);

  // For each segment, check if it contains multiple concatenated options and split them
  const finalSegments: { text: string; isBold: boolean }[] = [];
  
  for (const seg of segments) {
    // Check if this segment contains multiple option markers (e.g. "A. xxxB. yyyC. zzz")
    const positions: number[] = [];
    for (let i = 0; i < seg.text.length - 1; i++) {
      const ch = seg.text[i];
      const next = seg.text[i + 1];
      if (ch >= "A" && ch <= "D" && next === ".") {
        if (i === 0 || ((seg.text[i - 1] >= "a" && seg.text[i - 1] <= "z") || seg.text[i - 1] === " " || seg.text[i - 1] === ")" || seg.text[i - 1] === "]")) {
          positions.push(i);
        }
      }
    }
    
    if (positions.length >= 2) {
      // This segment has multiple options concatenated — split using bold-aware function
      // Build chunks for just this segment's text range
      const segSplit = splitConcatenatedOptionsFromText(seg.text, chunks, segments, seg);
      if (segSplit.length > 1) {
        finalSegments.push(...segSplit);
        continue;
      }
    }
    
    finalSegments.push(seg);
  }

  return finalSegments;
}

/**
 * Split a single segment's text into multiple options, using the segment's own text for bold detection.
 * Since we may not have exact chunk mapping for individual segments, we use the segment's isBold as default
 * and try to detect bold from the full chunks array.
 */
function splitConcatenatedOptionsFromText(
  text: string,
  allChunks: { text: string; isBold: boolean; isBr: boolean }[],
  allSegments: { text: string; isBold: boolean }[],
  currentSeg: { text: string; isBold: boolean }
): { text: string; isBold: boolean }[] {
  // Try using the full chunk-based bold detection
  const textChunks = allChunks.filter(c => !c.isBr);
  
  // If there's only 1 segment total, use the full chunk-based approach
  if (allSegments.length === 1) {
    const split = splitConcatenatedOptionsWithBold(textChunks);
    if (split.length > 1) return split;
  }
  
  // Otherwise fall back to simple text splitting with inherited bold
  return splitConcatenatedOptions(text, currentSeg.isBold);
}

function flushChunksToSegment(chunks: { text: string; isBold: boolean }[]): { text: string; isBold: boolean } | null {
  const combinedText = chunks.map(c => c.text).join("").trim();
  if (!combinedText) return null;
  
  // Bold detection: exclude option prefix chunks (A. B. C. D.) from calculation
  let totalChars = 0;
  let boldChars = 0;
  for (const c of chunks) {
    const trimmed = c.text.trim();
    const isPrefix = /^[A-Da-d]\.\s*$/.test(trimmed) || /^[A-Da-d]$/.test(trimmed) || /^\.\s*$/.test(trimmed);
    if (!isPrefix) {
      const charCount = c.text.replace(/\s/g, "").length;
      totalChars += charCount;
      if (c.isBold) boldChars += charCount;
    }
  }
  
  // Bold if more than 50% of non-prefix characters are bold
  const textChunks = chunks.filter(c => c.text.trim());
  const isBold = totalChars > 0 
    ? boldChars > totalChars * 0.5 
    : (textChunks.length > 0 && textChunks.every(c => c.isBold));
  
  return { text: combinedText, isBold };
}

/**
 * Split concatenated options like "A. Radiasi...B. Perubahan..." preserving per-option bold status.
 * Uses the original chunks array to determine which option text was bold.
 */
function splitConcatenatedOptionsWithBold(chunks: { text: string; isBold: boolean }[]): { text: string; isBold: boolean }[] {
  // Build full text with bold mapping per character
  let fullText = "";
  const boldMap: boolean[] = [];
  
  for (const chunk of chunks) {
    for (let i = 0; i < chunk.text.length; i++) {
      fullText += chunk.text[i];
      boldMap.push(chunk.isBold);
    }
  }

  // Split on option markers (A. B. C. D.)
  // Match uppercase A-D followed by ". " or "." at word boundary positions
  // We find all positions where an option marker starts
  const matches: { index: number }[] = [];
  for (let i = 0; i < fullText.length - 1; i++) {
    const ch = fullText[i];
    const next = fullText[i + 1];
    // Check if this is an option marker: [A-D] followed by "."
    if (ch >= "A" && ch <= "D" && next === ".") {
      // Validate: must be at start OR preceded by a lowercase letter or space (not mid-word uppercase)
      if (i === 0) {
        matches.push({ index: i });
      } else {
        const prev = fullText[i - 1];
        // preceded by lowercase letter (concatenated), space, closing paren/bracket, or line start
        if ((prev >= "a" && prev <= "z") || prev === " " || prev === "\n" || prev === "\t" || prev === ")" || prev === "]") {
          matches.push({ index: i });
        }
      }
    }
  }

  if (matches.length < 2) return [];

  const results: { text: string; isBold: boolean }[] = [];
  
  // Check if there's text before the first option marker (could be soal text)
  if (matches[0].index > 0) {
    const preText = fullText.substring(0, matches[0].index).trim();
    if (preText) {
      const preBold = boldMap.slice(0, matches[0].index).some(b => b) && 
                      boldMap.slice(0, matches[0].index).filter((_, i) => fullText[i] !== " ").every(b => b);
      results.push({ text: preText, isBold: preBold });
    }
  }

  for (let i = 0; i < matches.length; i++) {
    const start = matches[i].index;
    const end = i + 1 < matches.length ? matches[i + 1].index : fullText.length;
    const optText = fullText.substring(start, end).trim();
    
    if (!optText) continue;

    // Determine bold: check if majority of non-space characters in this range are bold
    const rangeBold = boldMap.slice(start, end);
    const nonSpaceCount = optText.replace(/\s/g, "").length;
    const boldCount = rangeBold.filter((b, idx) => b && fullText[start + idx] !== " ").length;
    const isBold = nonSpaceCount > 0 && boldCount >= nonSpaceCount * 0.5;

    results.push({ text: optText, isBold });
  }

  return results.length > 1 ? results : [];
}

/**
 * Legacy split function (kept as final fallback when chunks aren't available)
 */
function splitConcatenatedOptions(text: string, defaultBold: boolean): { text: string; isBold: boolean }[] {
  // Find option marker positions using same logic as splitConcatenatedOptionsWithBold
  const positions: number[] = [];
  for (let i = 0; i < text.length - 1; i++) {
    const ch = text[i];
    const next = text[i + 1];
    if (ch >= "A" && ch <= "D" && next === ".") {
      if (i === 0 || ((text[i - 1] >= "a" && text[i - 1] <= "z") || text[i - 1] === " " || text[i - 1] === ")" || text[i - 1] === "]")) {
        positions.push(i);
      }
    }
  }
  
  if (positions.length < 2) return [{ text, isBold: defaultBold }];
  
  const parts: string[] = [];
  for (let i = 0; i < positions.length; i++) {
    const start = positions[i];
    const end = i + 1 < positions.length ? positions[i + 1] : text.length;
    parts.push(text.substring(start, end).trim());
  }
  
  return parts.filter(p => p).map((part) => ({
    text: part.trim(),
    isBold: defaultBold,
  }));
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function extractNumbering(p: any): {
  numId: number | null;
  ilvl: number | null;
} {
  const pPr = p["w:pPr"];
  const numPr = pPr?.["w:numPr"];
  if (!numPr) return { numId: null, ilvl: null };

  const numIdRaw = numPr["w:numId"]?.["@_w:val"];
  const ilvlRaw = numPr["w:ilvl"]?.["@_w:val"];

  return {
    numId: numIdRaw !== undefined ? Number(numIdRaw) : null,
    ilvl:
      ilvlRaw !== undefined
        ? Number(ilvlRaw)
        : numIdRaw !== undefined
          ? 0
          : null,
  };
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function extractPStyleId(p: any): string | null {
  const val = p["w:pPr"]?.["w:pStyle"]?.["@_w:val"];
  return val !== undefined ? String(val) : null;
}

// ---------- Step 2: Build ParsedParagraph[] dari XML ----------

export async function parseDocxViaXml(
  buffer: Buffer
): Promise<ParsedParagraph[]> {
  const { documentObj, numberingObj, stylesObj } = await loadDocxXml(buffer);
  const paragraphs = extractParagraphs(documentObj);
  const styleNumberingMap = extractStyleNumberingMap(stylesObj);
  const numberingFormatMap = buildNumberingFormatMap(numberingObj);

  const result: ParsedParagraph[] = [];

  for (const p of paragraphs) {
    const { numId, ilvl } = extractNumbering(p);

    let hasStyleNumbering = false;
    if (numId === null) {
      const styleId = extractPStyleId(p);
      if (styleId && styleNumberingMap.has(styleId)) {
        hasStyleNumbering = true;
      }
    }

    // Resolve numFmt from numbering.xml
    let numFormat: string | null = null;
    if (numId !== null && ilvl !== null) {
      numFormat = numberingFormatMap.get(`${numId}:${ilvl}`) || null;
    }

    // Extract segments (handles soft line breaks within a paragraph)
    const segments = extractParagraphSegments(p);

    if (segments.length === 0) continue;

    // If there's only 1 segment, treat as normal paragraph
    if (segments.length === 1) {
      const seg = segments[0];
      if (!seg.text) continue;
      result.push({ text: seg.text, isBold: seg.isBold, numId, ilvl, hasStyleNumbering, numFormat });
    } else {
      // Multiple segments from soft line breaks — each becomes its own ParsedParagraph
      // The first segment inherits the paragraph's numbering info
      // Subsequent segments get numId=null (they're sub-lines, classified by regex)
      for (let i = 0; i < segments.length; i++) {
        const seg = segments[i];
        if (!seg.text) continue;
        if (i === 0) {
          result.push({ text: seg.text, isBold: seg.isBold, numId, ilvl, hasStyleNumbering, numFormat });
        } else {
          result.push({ text: seg.text, isBold: seg.isBold, numId: null, ilvl: null, hasStyleNumbering: false, numFormat: null });
        }
      }
    }
  }

  return result;
}

// ---------- Step 3: Regex classifier ----------

// ---------- Numbering format helpers ----------

const ALPHA_FORMATS = new Set([
  "lowerLetter",
  "upperLetter",
  "lowerRoman",
  "upperRoman",
]);

const DECIMAL_FORMATS = new Set(["decimal", "decimalZero"]);

function isAlphaFormat(fmt: string | null): boolean {
  return fmt !== null && ALPHA_FORMATS.has(fmt);
}

function isDecimalFormat(fmt: string | null): boolean {
  return fmt !== null && DECIMAL_FORMATS.has(fmt);
}

// ---------- Step 3: Regex classifier ----------

type Role = "soal" | "opsi" | "ignore";

const SOAL_PATTERNS: RegExp[] = [
  /^\s*soal\s*[:.\-]?\s*\d+\s*[:.\-]?\s*/i,
  /^\s*no\.?\s*\d+\s*[:.\-]?\s*/i,
  /^\s*\d+\s*[.)\-:]\s*/,
];

const OPSI_PATTERNS: RegExp[] = [
  /^\s*[A-Da-d]\s*[.)\-:]\s*/,
  /^\s*\(([A-Da-d])\)\s*/,
  /^\s*[ivx]{1,4}\s*[.)]\s*/i,
  /^\s*[•▪●\-*]\s+/,
];

function stripFirstMatch(
  text: string,
  patterns: RegExp[]
): { matched: boolean; rest: string } {
  for (const re of patterns) {
    const m = text.match(re);
    if (m) {
      return { matched: true, rest: text.slice(m[0].length).trim() };
    }
  }
  return { matched: false, rest: text };
}

function classifyParagraph(p: ParsedParagraph): {
  role: Role;
  cleanText: string;
} {
  // --- Priority 1: Use numFormat from numbering.xml ---
  // If the paragraph has auto-numbering with alpha format → it's an option
  if (p.numId !== null && p.ilvl !== null && isAlphaFormat(p.numFormat)) {
    return { role: "opsi", cleanText: p.text };
  }
  // If the paragraph has auto-numbering with decimal format → it's a question
  if (p.numId !== null && p.ilvl !== null && isDecimalFormat(p.numFormat)) {
    return { role: "soal", cleanText: p.text };
  }

  // --- Priority 2: ilvl-based (for nested lists) ---
  if (p.ilvl === 1) return { role: "opsi", cleanText: p.text };
  // Only treat ilvl 0 as soal if no numFormat info AND no regex match for opsi
  if (p.ilvl === 0 && p.numId !== null && p.numFormat === null) {
    // Unknown format at ilvl 0 — try regex fallback before defaulting to soal
    const opsiMatch = stripFirstMatch(p.text, OPSI_PATTERNS);
    if (opsiMatch.matched) return { role: "opsi", cleanText: opsiMatch.rest };
    return { role: "soal", cleanText: p.text };
  }

  // --- Priority 3: Regex fallback (manual numbering typed by user) ---
  const soalMatch = stripFirstMatch(p.text, SOAL_PATTERNS);
  if (soalMatch.matched) return { role: "soal", cleanText: soalMatch.rest };

  const opsiMatch = stripFirstMatch(p.text, OPSI_PATTERNS);
  if (opsiMatch.matched) return { role: "opsi", cleanText: opsiMatch.rest };

  if (p.hasStyleNumbering) return { role: "opsi", cleanText: p.text };

  return { role: "ignore", cleanText: p.text };
}

// ---------- Step 4: Convert ParsedParagraph[] -> QuizItem[] ----------

export function buildQuizFromParagraphs(
  paragraphs: ParsedParagraph[]
): QuizItem[] {
  // First attempt: use classifyParagraph directly
  const items = buildQuizDirect(paragraphs);

  // If direct classification found results, return them
  if (items.length > 0) return items;

  // Fallback: try numId-based grouping
  // This handles cases where numbering.xml doesn't have format info
  // but soal and opsi use different numIds at ilvl 0
  return buildQuizByNumIdGrouping(paragraphs);
}

function buildQuizDirect(paragraphs: ParsedParagraph[]): QuizItem[] {
  const items: QuizItem[] = [];
  let current: QuizItem | null = null;
  let currentSoalBold = false;
  // Track which opsi indices are bold for post-processing
  let boldOpsiIndices: number[] = [];

  for (const p of paragraphs) {
    const { role, cleanText } = classifyParagraph(p);
    if (!cleanText) continue;

    if (role === "opsi" && current) {
      current.opsi.push(cleanText);
      if (p.isBold) {
        boldOpsiIndices.push(current.opsi.length - 1);
      }
      continue;
    }

    if (role === "soal") {
      // Finalize previous question
      if (current) {
        current.jawabanIndex = determineBestAnswer(boldOpsiIndices, currentSoalBold);
        items.push(current);
      }
      current = { soal: cleanText, opsi: [] };
      currentSoalBold = p.isBold;
      boldOpsiIndices = [];
      continue;
    }

    // role === "ignore": lanjutan soal kalau belum ada opsi
    if (current && current.opsi.length === 0 && role === "ignore") {
      current.soal += " " + p.text;
    }
  }

  // Finalize last question
  if (current) {
    current.jawabanIndex = determineBestAnswer(boldOpsiIndices, currentSoalBold);
    items.push(current);
  }

  // Filter: minimal 2 opsi baru dianggap soal valid
  return items.filter((item) => item.opsi.length >= 2);
}

/**
 * Determine the correct answer index from bold opsi indices.
 * 
 * Logic:
 * 1. No bold opsi → undefined (no answer)
 * 2. Exactly 1 bold opsi → that's the answer
 * 3. Multiple bold opsi AND soal was NOT bold → first bold opsi is answer
 * 4. Multiple bold opsi AND soal WAS bold → 
 *    - If first opsi (index 0) is bold, it's likely "spillover" from soal formatting
 *    - Skip index 0, use the NEXT bold opsi that's NOT at index 0
 *    - If ALL bold opsi are just index 0, use it anyway
 * 5. ALL opsi are bold → undefined (can't determine)
 */
function determineBestAnswer(boldIndices: number[], soalBold: boolean): number | undefined {
  if (boldIndices.length === 0) return undefined;
  if (boldIndices.length === 1) return boldIndices[0];
  
  // If all 4 opsi are bold, we can't determine
  // (threshold: if more than 3 are bold, it's probably all bold = no signal)
  if (boldIndices.length >= 4) return undefined;
  
  // If soal was bold and first opsi (index 0) is also bold, it might be spillover
  if (soalBold && boldIndices[0] === 0) {
    // Try to find a bold opsi that's NOT at index 0
    const nonFirstBold = boldIndices.filter(i => i !== 0);
    if (nonFirstBold.length === 1) {
      // Exactly one other bold opsi besides index 0 → that's likely the real answer
      return nonFirstBold[0];
    }
    if (nonFirstBold.length > 1) {
      // Multiple non-first bold opsi — just use first non-zero one
      return nonFirstBold[0];
    }
    // Only index 0 is bold — use it (it's the only signal we have)
    return 0;
  }
  
  // Soal not bold or first opsi not at index 0: use first bold opsi
  return boldIndices[0];
}

/**
 * Fallback strategy: group paragraphs by numId.
 * 
 * When all paragraphs are ilvl 0 with different numIds but no numFormat info,
 * we use a heuristic:
 * - Paragraphs with numId that appears 1x per "group" (longer text, question-like) = soal
 * - Paragraphs with numId that appears in clusters of 3-6 (shorter text) = opsi
 * 
 * Additional heuristic: if paragraphs alternate between two numIds in a pattern
 * like [numA, numB, numB, numB, numB, numA, numB, numB, numB, numB...] then
 * numA = soal, numB = opsi.
 */
function buildQuizByNumIdGrouping(paragraphs: ParsedParagraph[]): QuizItem[] {
  // Only consider paragraphs with numId
  const numbered = paragraphs.filter((p) => p.numId !== null);
  if (numbered.length === 0) return [];

  // Count how many distinct numIds exist
  const numIdSet = new Set(numbered.map((p) => p.numId));

  // If there are exactly 2 numIds, try the alternating pattern
  if (numIdSet.size === 2) {
    const numIdArr = Array.from(numIdSet);
    const [numIdA, numIdB] = numIdArr;
    const countA = numbered.filter((p) => p.numId === numIdA).length;
    const countB = numbered.filter((p) => p.numId === numIdB).length;

    // The numId with fewer paragraphs is likely the "soal" numId
    // (typically 1 soal followed by 4 opsi)
    const soalNumId = countA <= countB ? numIdA : numIdB;

    const items: QuizItem[] = [];
    let current: QuizItem | null = null;

    for (const p of numbered) {
      if (p.numId === soalNumId) {
        if (current) items.push(current);
        current = { soal: p.text, opsi: [] };
      } else if (current) {
        current.opsi.push(p.text);
        if (p.isBold && current.jawabanIndex === undefined) {
          current.jawabanIndex = current.opsi.length - 1;
        }
      }
    }

    if (current) items.push(current);
    const valid = items.filter((item) => item.opsi.length >= 2);
    if (valid.length > 0) return valid;
  }

  // If more than 2 numIds or the 2-numId approach failed,
  // try sequential grouping: a "soal" is a longer paragraph followed by
  // shorter paragraphs (regardless of numId)
  return buildQuizByLengthHeuristic(numbered);
}

/**
 * Last resort heuristic: group by text length pattern.
 * A question tends to be longer, followed by 3-5 shorter option texts.
 */
function buildQuizByLengthHeuristic(paragraphs: ParsedParagraph[]): QuizItem[] {
  if (paragraphs.length < 3) return [];

  const items: QuizItem[] = [];
  let current: QuizItem | null = null;
  let opsiNumId: number | null = null;

  for (let i = 0; i < paragraphs.length; i++) {
    const p = paragraphs[i];

    // If we already have a current question and this paragraph has the same numId as options
    if (current && opsiNumId !== null && p.numId === opsiNumId) {
      current.opsi.push(p.text);
      if (p.isBold && current.jawabanIndex === undefined) {
        current.jawabanIndex = current.opsi.length - 1;
      }
      continue;
    }

    // If we have a current question and this paragraph has a DIFFERENT numId
    // (switch from opsi back to soal)
    if (current && p.numId !== opsiNumId) {
      items.push(current);
      current = { soal: p.text, opsi: [] };
      // Look ahead: if next paragraph has a different numId, that's the opsi numId
      if (i + 1 < paragraphs.length && paragraphs[i + 1].numId !== p.numId) {
        opsiNumId = paragraphs[i + 1].numId;
      }
      continue;
    }

    // Start first question
    if (!current) {
      current = { soal: p.text, opsi: [] };
      // Look ahead for opsi numId
      if (i + 1 < paragraphs.length && paragraphs[i + 1].numId !== p.numId) {
        opsiNumId = paragraphs[i + 1].numId;
      }
      continue;
    }
  }

  if (current) items.push(current);
  return items.filter((item) => item.opsi.length >= 2);
}

// ---------- Step 5: Public API ----------

export async function parseQuizDocx(buffer: Buffer): Promise<QuizItem[]> {
  const paragraphs = await parseDocxViaXml(buffer);
  return buildQuizFromParagraphs(paragraphs);
}
