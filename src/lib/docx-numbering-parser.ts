/**
 * Hybrid DOCX Quiz Parser
 * ------------------------------------------------------------
 * Strategy:
 *  1. Baca langsung word/document.xml + word/numbering.xml + word/styles.xml
 *  2. Classify setiap paragraph: soal vs opsi vs ignore
 *     - numPr ilvl 0 = SOAL
 *     - numPr ilvl 1 = OPSI
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
  const { documentObj, stylesObj } = await loadDocxXml(buffer);
  const paragraphs = extractParagraphs(documentObj);
  const styleNumberingMap = extractStyleNumberingMap(stylesObj);

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

    result.push({ text, isBold, numId, ilvl, hasStyleNumbering });
  }

  return result;
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
  if (p.ilvl === 0) return { role: "soal", cleanText: p.text };
  if (p.ilvl === 1) return { role: "opsi", cleanText: p.text };

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

// ---------- Step 5: Public API ----------

export async function parseQuizDocx(buffer: Buffer): Promise<QuizItem[]> {
  const paragraphs = await parseDocxViaXml(buffer);
  return buildQuizFromParagraphs(paragraphs);
}
