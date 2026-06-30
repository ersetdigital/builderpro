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
 * Uses both raw text (for structure) and HTML (for bold detection).
 */
export async function parseDocx(buffer: Buffer): Promise<ParseResult> {
  // Get HTML for bold detection
  const htmlResult = await mammoth.convertToHtml({ buffer });
  const html = htmlResult.value;

  // Get raw text for structure parsing
  const textResult = await mammoth.extractRawText({ buffer });
  const rawText = textResult.value;

  // Parse questions from raw text
  const questions = parseQuestionsFromText(rawText);

  // Detect bold answers from HTML
  const boldTexts = extractAllBoldTexts(html);

  // Match bold texts to options
  const result: ParsedQuestion[] = [];
  const errors: string[] = [];

  for (const q of questions) {
    if (q.options.length === 0) {
      errors.push(`Soal nomor ${q.number} tidak dapat diproses.`);
      continue;
    }

    // Find correct answer by matching bold text to option text
    let correctAnswer: string | null = null;

    for (const opt of q.options) {
      const optFullText = `${opt.label.toLowerCase()}. ${opt.text}`.toLowerCase().trim();
      const optTextOnly = opt.text.toLowerCase().trim();

      for (const boldText of boldTexts) {
        const boldLower = boldText.toLowerCase().trim();

        // Check if bold text matches the option (full or partial)
        if (
          boldLower === optFullText ||
          boldLower === optTextOnly ||
          boldLower === `${opt.label.toLowerCase()}.${opt.text.toLowerCase().trim()}` ||
          boldLower === `${opt.label.toLowerCase()}. ${opt.text.toLowerCase().trim()}` ||
          (optTextOnly.length > 3 && boldLower.includes(optTextOnly)) ||
          (boldLower.length > 3 && optTextOnly.includes(boldLower))
        ) {
          correctAnswer = opt.label;
          opt.isBold = true;
          break;
        }
      }
      if (correctAnswer) break;
    }

    if (!correctAnswer) {
      errors.push(`Soal nomor ${q.number} tidak memiliki kunci jawaban.`);
    }

    result.push({
      number: q.number,
      question: q.question,
      options: q.options,
      correctAnswer,
    });
  }

  return { questions: result, errors };
}

interface RawQuestion {
  number: number;
  question: string;
  options: { label: string; text: string; isBold: boolean }[];
}

/**
 * Parse questions from raw text extracted by mammoth.
 * Handles both plain numbered format and list format.
 */
function parseQuestionsFromText(rawText: string): RawQuestion[] {
  const lines = rawText.split("\n").map((l) => l.trim()).filter((l) => l.length > 0);
  const questions: RawQuestion[] = [];

  let currentQuestion: RawQuestion | null = null;

  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];

    // Match question: starts with number followed by . or )
    const questionMatch = line.match(/^(\d+)[.)]\s*(.+)/);
    // Match option: starts with a-d/A-D followed by . or )
    const optionMatch = line.match(/^([A-Da-d])[.)]\s*(.+)/);

    if (questionMatch && !optionMatch) {
      // New question
      if (currentQuestion) {
        questions.push(currentQuestion);
      }
      currentQuestion = {
        number: parseInt(questionMatch[1]),
        question: questionMatch[2].trim(),
        options: [],
      };
    } else if (optionMatch && currentQuestion) {
      // Option for current question
      const label = optionMatch[1].toUpperCase();
      const text = optionMatch[2].trim();
      currentQuestion.options.push({ label, text, isBold: false });
    } else if (currentQuestion && currentQuestion.options.length === 0) {
      // Continuation of question text
      currentQuestion.question += " " + line;
    }
  }

  // Don't forget last question
  if (currentQuestion) {
    questions.push(currentQuestion);
  }

  // Filter out "questions" that are actually just noise (like empty ones or "7." alone)
  return questions.filter((q) => q.question.trim().length > 0 && q.options.length > 0);
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
