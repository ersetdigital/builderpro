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

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function extractParagraphs(documentObj: any): any[] {
  const body = documentObj?.["w:document"]?.["w:body"];
  if (!body) return [];
  return toArray(body["w:p"]);
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
    const bNode = rPr?.["w:b"];
    const isRunBold =
      bNode !== undefined &&
      bNode?.["@_w:val"] !== "0" &&
      bNode?.["@_w:val"] !== "false";

    if (isRunBold) boldRunCount++;
  }

  const isBold = textRunCount > 0 && boldRunCount === textRunCount;
  return { text: combinedText.trim(), isBold };
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
    const { text, isBold } = extractParagraphText(p);
    if (!text) continue;

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

    result.push({ text, isBold, numId, ilvl, hasStyleNumbering, numFormat });
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

  for (const p of paragraphs) {
    const { role, cleanText } = classifyParagraph(p);
    if (!cleanText) continue;

    if (role === "opsi" && current) {
      current.opsi.push(cleanText);
      if (p.isBold && current.jawabanIndex === undefined) {
        current.jawabanIndex = current.opsi.length - 1;
      }
      continue;
    }

    if (role === "soal") {
      if (current) items.push(current);
      current = { soal: cleanText, opsi: [] };
      continue;
    }

    // role === "ignore": lanjutan soal kalau belum ada opsi
    if (current && current.opsi.length === 0 && role === "ignore") {
      current.soal += " " + p.text;
    }
  }

  if (current) items.push(current);

  // Filter: minimal 2 opsi baru dianggap soal valid
  return items.filter((item) => item.opsi.length >= 2);
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
