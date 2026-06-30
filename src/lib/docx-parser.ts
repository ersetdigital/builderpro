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
 * Supports multiple formats:
 * - Numbered questions (1. 2. 3. or 1) 2) 3))
 * - Options with A/B/C/D or a/b/c/d (with . or ))
 * - Bold text as correct answer indicator
 * - Word's automatic numbering/lists
 */
export async function parseDocx(buffer: Buffer): Promise<ParseResult> {
  const result = await mammoth.convertToHtml({ buffer });
  const html = result.value;

  const lines = extractTextLines(html);

  const questions: ParsedQuestion[] = [];
  const errors: string[] = [];

  let currentQuestion: Partial<ParsedQuestion> | null = null;
  let currentOptions: { label: string; text: string; isBold: boolean }[] = [];

  for (const line of lines) {
    const trimmed = line.text.trim();
    if (!trimmed) continue;

    // Check if this is a question line (starts with number)
    const questionMatch = trimmed.match(/^(\d+)[.)]\s*(.+)/);
    // Check if this is an option line (starts with a-d or A-D)
    const optionMatch = trimmed.match(/^([A-Da-d])[.)]\s*(.+)/);

    if (questionMatch) {
      // Could be a question OR an option that starts with a number-like pattern
      // If it matches option pattern too, treat as option
      if (optionMatch && currentQuestion) {
        const label = optionMatch[1].toUpperCase();
        const text = optionMatch[2];
        currentOptions.push({ label, text, isBold: line.isBold });
      } else {
        // Save previous question
        if (currentQuestion && currentQuestion.question) {
          finalizeQuestion(currentQuestion, currentOptions, questions, errors);
        }

        currentQuestion = {
          number: parseInt(questionMatch[1]),
          question: questionMatch[2],
        };
        currentOptions = [];
      }
    } else if (optionMatch && currentQuestion) {
      const label = optionMatch[1].toUpperCase();
      const text = optionMatch[2];
      currentOptions.push({ label, text, isBold: line.isBold });
    } else if (currentQuestion && currentOptions.length === 0) {
      // Continuation of question text
      currentQuestion.question =
        (currentQuestion.question || "") + " " + trimmed;
    }
  }

  // Don't forget the last question
  if (currentQuestion && currentQuestion.question) {
    finalizeQuestion(currentQuestion, currentOptions, questions, errors);
  }

  // If standard parsing found nothing, try alternative parsing
  if (questions.length === 0 && errors.length === 0) {
    return parseAlternativeFormat(html);
  }

  return { questions, errors };
}

/**
 * Alternative parser for documents using Word's automatic numbering/list features.
 * Detects questions by looking at list structure and bold patterns.
 */
function parseAlternativeFormat(html: string): ParseResult {
  const questions: ParsedQuestion[] = [];
  const errors: string[] = [];

  // Extract all list items and paragraphs with their formatting
  const blocks = extractAllBlocks(html);

  let currentQuestion: Partial<ParsedQuestion> | null = null;
  let currentOptions: { label: string; text: string; isBold: boolean }[] = [];
  let questionNumber = 0;

  for (const block of blocks) {
    const trimmed = block.text.trim();
    if (!trimmed) continue;

    // Detect option patterns: a. b. c. d. or A. B. C. D. or a) b) c) d)
    const optionMatch = trimmed.match(/^([A-Da-d])[.)]\s*(.+)/);

    // Detect question patterns: numbered or by context (non-option text before options)
    const questionMatch = trimmed.match(/^(\d+)[.)]\s*(.+)/);

    if (optionMatch) {
      // This is an option
      if (!currentQuestion) {
        // No question yet, might be a standalone list — skip
        continue;
      }
      const label = optionMatch[1].toUpperCase();
      const text = optionMatch[2];
      currentOptions.push({ label, text, isBold: block.isBold });
    } else if (questionMatch && !optionMatch) {
      // Save previous question
      if (currentQuestion && currentQuestion.question) {
        finalizeQuestion(currentQuestion, currentOptions, questions, errors);
      }

      questionNumber = parseInt(questionMatch[1]);
      currentQuestion = {
        number: questionNumber,
        question: questionMatch[2],
      };
      currentOptions = [];
    } else if (currentOptions.length > 0 && currentQuestion) {
      // We have options collected and hit non-option text = new question without number
      finalizeQuestion(currentQuestion, currentOptions, questions, errors);
      questionNumber++;
      currentQuestion = {
        number: questionNumber,
        question: trimmed,
      };
      currentOptions = [];
    } else if (!currentQuestion || currentOptions.length === 0) {
      // Could be start of a new question or continuation
      if (currentQuestion && currentOptions.length === 0) {
        // Continuation of question text
        currentQuestion.question =
          (currentQuestion.question || "") + " " + trimmed;
      } else {
        // New question without explicit number
        if (currentQuestion && currentQuestion.question) {
          finalizeQuestion(currentQuestion, currentOptions, questions, errors);
        }
        questionNumber++;
        currentQuestion = {
          number: questionNumber,
          question: trimmed,
        };
        currentOptions = [];
      }
    }
  }

  // Last question
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

interface TextBlock {
  text: string;
  isBold: boolean;
  isListItem: boolean;
}

function extractTextLines(html: string): TextBlock[] {
  const lines: TextBlock[] = [];

  // Split by paragraph/block/list-item tags
  const blocks = html.split(
    /<\/p>|<\/h[1-6]>|<\/li>|<br\s*\/?>/gi
  );

  for (const block of blocks) {
    const cleanBlock = block
      .replace(/<p[^>]*>|<h[1-6][^>]*>|<li[^>]*>|<ul[^>]*>|<ol[^>]*>|<\/ul>|<\/ol>/gi, "");
    if (!cleanBlock.trim()) continue;

    const stripped = cleanBlock.trim();
    const isBold = isContentBold(stripped);
    const isListItem = /<li/i.test(block);

    // Remove all HTML tags to get plain text
    const text = stripped.replace(/<[^>]+>/g, "").trim();

    if (text) {
      lines.push({ text, isBold, isListItem });
    }
  }

  return lines;
}

function extractAllBlocks(html: string): TextBlock[] {
  const blocks: TextBlock[] = [];

  // More aggressive splitting — handle nested lists, paragraphs, etc.
  const segments = html.split(
    /<\/p>|<\/h[1-6]>|<\/li>|<br\s*\/?>/gi
  );

  for (const segment of segments) {
    const clean = segment
      .replace(/<p[^>]*>|<h[1-6][^>]*>|<li[^>]*>|<ul[^>]*>|<ol[^>]*>|<\/ul>|<\/ol>/gi, "")
      .trim();
    if (!clean) continue;

    const isBold = isContentBold(clean);
    const text = clean.replace(/<[^>]+>/g, "").trim();

    if (text) {
      blocks.push({ text, isBold, isListItem: false });
    }
  }

  return blocks;
}

function isContentBold(html: string): boolean {
  const textContent = html.replace(/<[^>]+>/g, "").trim();
  if (!textContent) return false;

  const boldContent = extractBoldText(html).trim();

  // If most of the text content is bold, consider the line bold
  // Using > 50% threshold to handle minor formatting differences
  return boldContent.length > 0 && boldContent.length >= textContent.length * 0.5;
}

function extractBoldText(html: string): string {
  const boldMatches = html.match(
    /<strong>([\s\S]*?)<\/strong>|<b>([\s\S]*?)<\/b>/gi
  );
  if (!boldMatches) return "";

  return boldMatches
    .map((m) => m.replace(/<[^>]+>/g, ""))
    .join("")
    .trim();
}
