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
 * Handles multiple HTML structures produced by mammoth from Word docs.
 */
export async function parseDocx(buffer: Buffer): Promise<ParseResult> {
  const htmlResult = await mammoth.convertToHtml({ buffer });
  const html = htmlResult.value;

  const questions = parseFromHtml(html);
  const errors: string[] = [];
  const validQuestions: ParsedQuestion[] = [];

  for (const q of questions) {
    if (q.options.length === 0) {
      errors.push(`Soal nomor ${q.number} tidak dapat diproses.`);
      continue;
    }

    if (!q.correctAnswer) {
      errors.push(`Soal nomor ${q.number} tidak memiliki kunci jawaban.`);
    }

    validQuestions.push(q);
  }

  return { questions: validQuestions, errors };
}

/**
 * Parse questions from mammoth HTML output.
 * Strategy: flatten all content into sequential items, then group into questions.
 */
function parseFromHtml(html: string): ParsedQuestion[] {
  // Step 1: Extract all content items in order with their properties
  const items = extractItems(html);

  // Step 2: Group items into questions
  const questions = groupIntoQuestions(items);

  return questions;
}

interface ContentItem {
  text: string;
  isBold: boolean;
  isNested: boolean; // true if inside a nested list (likely an option)
  depth: number;
}

/**
 * Walk through the HTML and extract content items with context.
 */
function extractItems(html: string): ContentItem[] {
  const items: ContentItem[] = [];

  // Split HTML into tokens (tags and text)
  const tokens = html.split(/(<[^>]+>)/g).filter(Boolean);

  let depth = 0; // ol/ul nesting depth
  let inLi = false;
  let currentText = "";
  let currentHtml = "";

  function flush() {
    const text = currentText.trim();
    if (text) {
      items.push({
        text,
        isBold: isTextBold(currentHtml),
        isNested: depth > 1,
        depth,
      });
    }
    currentText = "";
    currentHtml = "";
  }

  for (const token of tokens) {
    if (token.startsWith("<")) {
      const tagLower = token.toLowerCase();

      if (tagLower.startsWith("<ol") || tagLower.startsWith("<ul")) {
        depth++;
      } else if (tagLower === "</ol>" || tagLower === "</ul>") {
        depth--;
        if (depth < 0) depth = 0;
      } else if (tagLower.startsWith("<li")) {
        flush();
        inLi = true;
      } else if (tagLower === "</li>") {
        flush();
        inLi = false;
      } else if (tagLower.startsWith("<p")) {
        flush();
      } else if (tagLower === "</p>") {
        flush();
      }

      // Track HTML for bold detection
      currentHtml += token;
    } else {
      // Text content
      currentText += token;
      currentHtml += token;
    }
  }

  flush();
  return items;
}

/**
 * Group sequential items into questions.
 * A question is identified by:
 * - A non-nested item at depth <= 1 that looks like question text
 * - Followed by nested items (depth > 1) or sequential short items that are options
 */
function groupIntoQuestions(items: ContentItem[]): ParsedQuestion[] {
  const questions: ParsedQuestion[] = [];
  let questionNumber = 0;

  let i = 0;
  while (i < items.length) {
    const item = items[i];

    // Check if this looks like a question start
    // Questions are typically at depth 1 (first level li) or depth 0 (paragraph)
    // and contain meaningful text (not just a short option)
    const isQuestionStart =
      !item.isNested &&
      item.depth <= 1 &&
      item.text.length > 5 && // Questions are longer than options typically
      !isLikelyOption(item.text);

    // Also detect "7. Standart..." pattern in paragraphs
    const plainQuestionMatch = item.text.match(/^(\d+)[.)]\s*(.+)/);

    if (isQuestionStart || (plainQuestionMatch && !isLikelyOption(item.text))) {
      questionNumber++;
      let questionText = item.text;

      // If it matches a numbered pattern, extract the number
      if (plainQuestionMatch) {
        questionNumber = parseInt(plainQuestionMatch[1]);
        questionText = plainQuestionMatch[2];
      }

      // Collect options that follow
      const options: { label: string; text: string; isBold: boolean }[] = [];
      const optionLabels = ["A", "B", "C", "D", "E", "F"];
      let j = i + 1;

      while (j < items.length) {
        const next = items[j];

        // Stop if we hit another question
        if (
          !next.isNested &&
          next.depth <= 1 &&
          next.text.length > 5 &&
          !isLikelyOption(next.text) &&
          options.length >= 2 // We need at least 2 options before a new question
        ) {
          break;
        }

        // Also stop if we hit a numbered question pattern
        const nextQuestionMatch = next.text.match(/^(\d+)[.)]\s*(.+)/);
        if (nextQuestionMatch && !isLikelyOption(next.text) && options.length >= 2) {
          break;
        }

        // This is an option
        if (next.text.trim()) {
          const label = optionLabels[options.length] || "?";
          options.push({
            label,
            text: next.text.trim(),
            isBold: next.isBold,
          });
        }

        j++;

        // Max 6 options per question
        if (options.length >= 6) break;
      }

      // Determine correct answer
      let correctAnswer: string | null = null;
      const boldOptions = options.filter((o) => o.isBold);
      if (boldOptions.length > 0) {
        correctAnswer = boldOptions[0].label;
      }

      if (options.length > 0) {
        questions.push({
          number: questionNumber,
          question: questionText,
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
 * Check if text looks like an option rather than a question.
 */
function isLikelyOption(text: string): boolean {
  // Very short text is likely an option
  if (text.length <= 30 && !text.includes(":") && !text.includes("?")) {
    return true;
  }
  // Starts with option-like pattern
  if (/^[A-Da-d][.)]\s/.test(text)) {
    return true;
  }
  // Common option texts
  const optionPatterns = [
    /^semua\s+(benar|salah)/i,
    /^output\s/i,
    /^\d+\s*volt/i,
    /^(CPU|Input|Output|Rele|Transistor|TRIAC)/i,
  ];
  for (const p of optionPatterns) {
    if (p.test(text.trim())) return true;
  }
  return false;
}

/**
 * Check if HTML segment contains bold formatting for the majority of its text.
 */
function isTextBold(html: string): boolean {
  const textContent = html.replace(/<[^>]+>/g, "").trim();
  if (!textContent) return false;

  const boldMatches = html.match(/<strong>([\s\S]*?)<\/strong>|<b>([\s\S]*?)<\/b>/gi);
  if (!boldMatches) return false;

  const boldText = boldMatches
    .map((m) => m.replace(/<[^>]+>/g, ""))
    .join("")
    .trim();

  // More than 50% of the text is bold
  return boldText.length >= textContent.length * 0.5;
}
