import mammoth from "mammoth";

export interface ParsedQuestion {
  number: number;
  question: string;
  options: { label: string; text: string; isBold: boolean }[];
  correctAnswer: string | null;
}

export interface ParseResult {
  questions: ParsedQuestion[];
  errors: string[];
}

/**
 * Parses a DOCX buffer and extracts quiz questions.
 */
export async function parseDocx(buffer: Buffer): Promise<ParseResult> {
  const htmlResult = await mammoth.convertToHtml({ buffer });
  const html = htmlResult.value;

  const items = flattenHtml(html);
  const questions = groupItems(items);

  const errors: string[] = [];
  const valid: ParsedQuestion[] = [];

  for (const q of questions) {
    if (q.options.length === 0) {
      errors.push(`Soal nomor ${q.number} tidak dapat diproses.`);
      continue;
    }
    if (!q.correctAnswer) {
      errors.push(`Soal nomor ${q.number} tidak memiliki kunci jawaban.`);
    }
    valid.push(q);
  }

  return { questions: valid, errors };
}

interface FlatItem {
  text: string;
  isBold: boolean;
  depth: number;
}

function flattenHtml(html: string): FlatItem[] {
  const items: FlatItem[] = [];
  let depth = 0;
  let buffer = "";

  const parts = html.split(/(<[^>]+>)/);

  for (const part of parts) {
    if (!part) continue;

    if (part.startsWith("<")) {
      const lower = part.toLowerCase();

      if (lower.startsWith("<ol") || lower.startsWith("<ul")) {
        depth++;
      } else if (lower === "</ol>" || lower === "</ul>") {
        depth = Math.max(0, depth - 1);
      } else if (lower.startsWith("<li")) {
        flushBuffer(buffer, depth, items);
        buffer = "";
      } else if (lower === "</li>") {
        flushBuffer(buffer, depth, items);
        buffer = "";
      } else if (lower.startsWith("<p")) {
        flushBuffer(buffer, depth, items);
        buffer = "";
      } else if (lower === "</p>") {
        flushBuffer(buffer, depth, items);
        buffer = "";
      } else {
        buffer += part;
      }
    } else {
      buffer += part;
    }
  }

  flushBuffer(buffer, depth, items);
  return items;
}

function flushBuffer(buffer: string, depth: number, items: FlatItem[]) {
  const text = buffer.replace(/<[^>]+>/g, "").trim();
  if (!text) return;
  const isBold = checkBold(buffer);
  items.push({ text, isBold, depth });
}

function checkBold(html: string): boolean {
  const text = html.replace(/<[^>]+>/g, "").trim();
  if (!text) return false;
  const boldRegex = /<strong>([\s\S]*?)<\/strong>|<b>([\s\S]*?)<\/b>/gi;
  let boldText = "";
  let match;
  while ((match = boldRegex.exec(html)) !== null) {
    boldText += (match[1] || match[2] || "").replace(/<[^>]+>/g, "");
  }
  return boldText.trim().length >= text.length * 0.5;
}

/**
 * Group items into questions by detecting question patterns.
 * 
 * Strategy: A question is any item that:
 * - Ends with : or ? 
 * - OR matches "N. <text>" pattern (paragraph question)
 * - OR is long text (>40 chars) followed by short items
 * 
 * After a question, the next 4 (or up to 6) items are its options,
 * UNTIL we hit another question.
 */
function groupItems(items: FlatItem[]): ParsedQuestion[] {
  const questions: ParsedQuestion[] = [];
  const optLabels = ["A", "B", "C", "D", "E", "F"];
  let questionNum = 0;

  let i = 0;
  while (i < items.length) {
    const item = items[i];

    // Check if this item is a question
    if (isQuestion(item, items, i)) {
      questionNum++;

      // Check if it's a numbered paragraph question like "7. Standart..."
      let qText = item.text;
      const numMatch = qText.match(/^(\d+)[.)]\s*(.+)/);
      if (numMatch) {
        questionNum = parseInt(numMatch[1]);
        qText = numMatch[2];
      }

      // Collect options: next items until another question or max 4
      const options: { label: string; text: string; isBold: boolean }[] = [];
      let j = i + 1;

      while (j < items.length && options.length < 4) {
        const next = items[j];

        // If this next item is also a question, stop collecting options
        if (options.length >= 2 && isQuestion(next, items, j)) {
          break;
        }

        options.push({
          label: optLabels[options.length],
          text: next.text,
          isBold: next.isBold,
        });
        j++;
      }

      const correctAnswer = options.find((o) => o.isBold)?.label || null;

      // Only add if we have at least 2 options (otherwise it's not a real question)
      if (options.length >= 2) {
        questions.push({
          number: questionNum,
          question: qText,
          options,
          correctAnswer,
        });
      }

      i = j;
    } else {
      i++;
    }
  }

  return questions;
}

/**
 * Determine if an item is a question.
 */
function isQuestion(item: FlatItem, items: FlatItem[], index: number): boolean {
  const text = item.text;

  // Numbered paragraph "7. Standart Tegangan..."
  const numMatch = text.match(/^(\d+)[.)]\s*(.+)/);
  if (numMatch && numMatch[2].length > 5) {
    return true;
  }

  // Ends with colon or question mark — strong indicator
  if (text.endsWith(":") || text.endsWith("?")) {
    return true;
  }

  // Contains question keywords AND is reasonably long
  const keywords = ["adalah", "berikut", "kecuali", "merupakan", "terletak", "sesuai", "mempunyai", "terdiri"];
  const hasKeyword = keywords.some((kw) => text.toLowerCase().includes(kw));
  if (hasKeyword && text.length > 20) {
    return true;
  }

  // Long text followed by at least 3 shorter items
  if (text.length > 40) {
    let shortFollowers = 0;
    for (let k = index + 1; k < Math.min(index + 5, items.length); k++) {
      if (items[k].text.length < text.length) {
        shortFollowers++;
      } else {
        break;
      }
    }
    if (shortFollowers >= 3) return true;
  }

  return false;
}
