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
 *
 * Expected format:
 * 1. Question text here
 * A. Option A
 * B. Option B
 * **C. Correct option** (bold = correct answer)
 * D. Option D
 */
export async function parseDocx(buffer: Buffer): Promise<ParseResult> {
  const result = await mammoth.convertToHtml({ buffer });
  const html = result.value;

  // Split HTML into lines by block elements
  const lines = extractTextLines(html);

  const questions: ParsedQuestion[] = [];
  const errors: string[] = [];

  let currentQuestion: Partial<ParsedQuestion> | null = null;
  let currentOptions: { label: string; text: string; isBold: boolean }[] = [];

  for (const line of lines) {
    const trimmed = line.text.trim();
    if (!trimmed) continue;

    // Check if this is a new question (starts with a number followed by . or ))
    const questionMatch = trimmed.match(/^(\d+)[.)]\s*(.+)/);
    // Check if this is an option (starts with A-D followed by . or ))
    const optionMatch = trimmed.match(/^([A-Da-d])[.)]\s*(.+)/);

    if (questionMatch && !optionMatch) {
      // Save previous question if exists
      if (currentQuestion && currentQuestion.question) {
        finalizeQuestion(currentQuestion, currentOptions, questions, errors);
      }

      currentQuestion = {
        number: parseInt(questionMatch[1]),
        question: questionMatch[2],
      };
      currentOptions = [];
    } else if (optionMatch && currentQuestion) {
      const label = optionMatch[1].toUpperCase();
      const text = optionMatch[2];
      const isBold = line.isBold;

      currentOptions.push({ label, text, isBold });
    } else if (currentQuestion && !currentOptions.length) {
      // Continuation of question text
      currentQuestion.question =
        (currentQuestion.question || "") + " " + trimmed;
    }
  }

  // Don't forget the last question
  if (currentQuestion && currentQuestion.question) {
    finalizeQuestion(currentQuestion, currentOptions, questions, errors);
  }

  return { questions, errors };
}

function finalizeQuestion(
  q: Partial<ParsedQuestion>,
  options: { label: string; text: string; isBold: boolean }[],
  questions: ParsedQuestion[],
  errors: string[]
) {
  const number = q.number || questions.length + 1;

  if (options.length === 0) {
    errors.push(`Soal nomor ${number} tidak dapat diproses.`);
    return;
  }

  const boldOptions = options.filter((o) => o.isBold);
  let correctAnswer: string | null = null;

  if (boldOptions.length === 0) {
    errors.push(`Soal nomor ${number} tidak memiliki kunci jawaban.`);
  } else {
    correctAnswer = boldOptions[0].label;
  }

  questions.push({
    number,
    question: q.question || "",
    options,
    correctAnswer,
  });
}

interface TextLine {
  text: string;
  isBold: boolean;
}

function extractTextLines(html: string): TextLine[] {
  const lines: TextLine[] = [];

  // Split by paragraph/block tags
  const blocks = html.split(/<\/p>|<\/h[1-6]>|<br\s*\/?>/gi);

  for (const block of blocks) {
    const cleanBlock = block.replace(/<p[^>]*>|<h[1-6][^>]*>/gi, "");
    if (!cleanBlock.trim()) continue;

    // Check if the entire line content is wrapped in <strong> or <b>
    const stripped = cleanBlock.trim();
    const isBold = isLineBold(stripped);

    // Remove all HTML tags to get plain text
    const text = stripped.replace(/<[^>]+>/g, "").trim();

    if (text) {
      lines.push({ text, isBold });
    }
  }

  return lines;
}

function isLineBold(html: string): boolean {
  const textContent = html.replace(/<[^>]+>/g, "").trim();
  const boldContent = extractBoldText(html).trim();

  // If the bold content is the same as the full text content, the line is bold
  return boldContent.length > 0 && boldContent === textContent;
}

function extractBoldText(html: string): string {
  const boldMatches = html.match(
    /<strong>(.*?)<\/strong>|<b>(.*?)<\/b>/gi
  );
  if (!boldMatches) return "";

  return boldMatches
    .map((m) => m.replace(/<[^>]+>/g, ""))
    .join("")
    .trim();
}
