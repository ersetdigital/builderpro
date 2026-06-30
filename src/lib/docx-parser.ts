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
 * Uses both raw text (for reliable structure) and HTML (for bold detection).
 */
export async function parseDocx(buffer: Buffer): Promise<ParseResult> {
  // Get HTML for bold detection
  const htmlResult = await mammoth.convertToHtml({ buffer });
  const html = htmlResult.value;

  // Get raw text — mammoth extractRawText preserves list numbering
  const textResult = await mammoth.extractRawText({ buffer });
  const rawText = textResult.value;

  // Extract all bold text snippets from HTML
  const boldTexts = extractAllBoldTexts(html);

  // Parse questions from raw text
  const parsed = parseRawText(rawText, boldTexts);

  return parsed;
}

/**
 * Parse the raw text from mammoth into questions.
 * Raw text from mammoth includes the text content of list items in order,
 * separated by newlines, but WITHOUT the numbering/lettering from Word lists.
 * 
 * However, based on actual testing, mammoth extractRawText DOES include
 * content in reading order. We need to figure out which lines are questions
 * and which are options based on the document's content patterns.
 */
function parseRawText(rawText: string, boldTexts: string[]): ParseResult {
  const lines = rawText
    .split(/\r?\n/)
    .map((l) => l.trim())
    .filter((l) => l.length > 0);

  const questions: ParsedQuestion[] = [];
  const errors: string[] = [];

  let currentQuestion: { number: number; question: string } | null = null;
  let currentOptions: { label: string; text: string; isBold: boolean }[] = [];
  const optLabels = ["A", "B", "C", "D", "E", "F"];

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];

    // Check for explicit question pattern: "1." "2." etc
    const questionMatch = line.match(/^(\d+)[.)]\s*(.*)/);
    // Check for explicit option pattern: "a." "b." "A." "B." etc
    const optionMatch = line.match(/^([A-Da-d])[.)]\s*(.*)/);

    if (optionMatch) {
      // Explicit option
      if (currentQuestion) {
        const label = optionMatch[1].toUpperCase();
        const text = optionMatch[2].trim();
        const isBold = isTextInBoldList(text, boldTexts) || isTextInBoldList(`${optionMatch[1]}. ${text}`, boldTexts);
        currentOptions.push({ label, text, isBold });
      }
    } else if (questionMatch) {
      const qNum = parseInt(questionMatch[1]);
      const qText = questionMatch[2].trim();

      // Edge case: "7." alone at end of soal 6 — skip empty question text
      if (!qText || qText.length === 0) {
        continue;
      }

      // Save previous question
      if (currentQuestion) {
        finalize(currentQuestion, currentOptions, questions, errors);
      }

      currentQuestion = { number: qNum, question: qText };
      currentOptions = [];
    } else {
      // Line without explicit prefix
      // If we have a current question and no options yet, this might be question continuation
      // If we have options already, this might be a stray option without prefix
      if (currentQuestion && currentOptions.length === 0) {
        // Append to question text
        currentQuestion.question += " " + line;
      } else if (currentQuestion && currentOptions.length > 0) {
        // After options started, a line without prefix is unusual
        // Could be continuation of last option or a new option without letter
        // For now, treat as additional option
        const label = optLabels[currentOptions.length] || "?";
        const isBold = isTextInBoldList(line, boldTexts);
        currentOptions.push({ label, text: line, isBold });
      }
    }
  }

  // Last question
  if (currentQuestion) {
    finalize(currentQuestion, currentOptions, questions, errors);
  }

  return { questions, errors };
}

function finalize(
  q: { number: number; question: string },
  options: { label: string; text: string; isBold: boolean }[],
  questions: ParsedQuestion[],
  errors: string[]
) {
  if (options.length === 0) {
    errors.push(`Soal nomor ${q.number} tidak dapat diproses.`);
    return;
  }

  let correctAnswer: string | null = null;
  const boldOpts = options.filter((o) => o.isBold);
  if (boldOpts.length > 0) {
    correctAnswer = boldOpts[0].label;
  }

  questions.push({
    number: q.number,
    question: q.question,
    options,
    correctAnswer,
  });
}

/**
 * Check if a text is contained in the bold texts list.
 * Uses fuzzy matching to handle minor differences.
 */
function isTextInBoldList(text: string, boldTexts: string[]): boolean {
  const normalized = text.toLowerCase().replace(/\s+/g, " ").trim();
  if (!normalized || normalized.length < 2) return false;

  for (const bold of boldTexts) {
    const boldNorm = bold.toLowerCase().replace(/\s+/g, " ").trim();
    if (!boldNorm) continue;

    // Exact match
    if (boldNorm === normalized) return true;

    // Bold contains the option text
    if (boldNorm.length > 3 && normalized.includes(boldNorm)) return true;

    // Option text contains the bold text
    if (normalized.length > 3 && boldNorm.includes(normalized)) return true;

    // Match without trailing punctuation
    const normNoPunct = normalized.replace(/[.,;:!?]+$/, "").trim();
    const boldNoPunct = boldNorm.replace(/[.,;:!?]+$/, "").trim();
    if (normNoPunct === boldNoPunct) return true;
    if (normNoPunct.length > 3 && boldNoPunct.includes(normNoPunct)) return true;
    if (boldNoPunct.length > 3 && normNoPunct.includes(boldNoPunct)) return true;
  }

  return false;
}

/**
 * Extract all bold text segments from HTML.
 */
function extractAllBoldTexts(html: string): string[] {
  const boldTexts: string[] = [];
  const regex = /<strong>([\s\S]*?)<\/strong>|<b>([\s\S]*?)<\/b>/gi;
  let match;

  while ((match = regex.exec(html)) !== null) {
    const text = (match[1] || match[2] || "").replace(/<[^>]+>/g, "").trim();
    if (text && text.length > 1) {
      boldTexts.push(text);
    }
  }

  return boldTexts;
}
