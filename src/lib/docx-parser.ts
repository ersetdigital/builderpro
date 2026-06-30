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

interface TextBlock {
  text: string;
  isBold: boolean;
  nestLevel: number;
}

/**
 * Parses a DOCX buffer and extracts quiz questions.
 * Supports:
 * - Plain text numbered questions (1. 2. 3.)
 * - Word auto-numbered lists (ol/li)
 * - Options: a/b/c/d or A/B/C/D (with . or ))
 * - Options as nested/indented list items
 * - Bold text = correct answer
 */
export async function parseDocx(buffer: Buffer): Promise<ParseResult> {
  const result = await mammoth.convertToHtml({ buffer });
  const html = result.value;

  // Try structured list parsing first (handles Word numbered lists)
  const listResult = parseFromHtmlStructure(html);
  if (listResult.questions.length > 0) {
    return listResult;
  }

  // Fallback: line-by-line text parsing
  const lineResult = parseFromTextLines(html);
  return lineResult;
}

/**
 * Parse using HTML structure — handles Word's <ol>/<li> numbered lists
 * where questions are top-level list items and options are nested lists.
 */
function parseFromHtmlStructure(html: string): ParseResult {
  const questions: ParsedQuestion[] = [];
  const errors: string[] = [];

  // Extract all blocks with nesting level info
  const blocks = extractBlocksWithNesting(html);

  let currentQuestion: Partial<ParsedQuestion> | null = null;
  let currentOptions: { label: string; text: string; isBold: boolean }[] = [];
  let questionNumber = 0;
  let lastTopLevelIndex = -1;

  for (let i = 0; i < blocks.length; i++) {
    const block = blocks[i];
    const trimmed = block.text.trim();
    if (!trimmed) continue;

    // Detect explicit option pattern: a. / b. / c. / d. or A. / B. etc
    const optionMatch = trimmed.match(/^([A-Da-d])[.)]\s*(.*)/);

    // Detect explicit question number pattern
    const questionMatch = trimmed.match(/^(\d+)[.)]\s*(.*)/);

    if (optionMatch) {
      // This is definitely an option
      const label = optionMatch[1].toUpperCase();
      const text = optionMatch[2] || trimmed;
      currentOptions.push({ label, text, isBold: block.isBold });
    } else if (block.nestLevel > 0 && currentQuestion) {
      // Nested item without explicit a/b/c/d prefix — treat as option by position
      const optionLabels = ["A", "B", "C", "D", "E"];
      const label = optionLabels[currentOptions.length] || "?";
      currentOptions.push({ label, text: trimmed, isBold: block.isBold });
    } else if (questionMatch && block.nestLevel === 0) {
      // Top level numbered item = new question
      if (currentQuestion && currentQuestion.question) {
        finalizeQuestion(currentQuestion, currentOptions, questions, errors);
      }
      questionNumber = parseInt(questionMatch[1]);
      currentQuestion = {
        number: questionNumber,
        question: questionMatch[2] || trimmed,
      };
      currentOptions = [];
      lastTopLevelIndex = i;
    } else if (block.nestLevel === 0 && currentOptions.length > 0) {
      // New top-level text after options = finalize and start new question
      if (currentQuestion && currentQuestion.question) {
        finalizeQuestion(currentQuestion, currentOptions, questions, errors);
      }
      questionNumber++;
      currentQuestion = {
        number: questionNumber,
        question: trimmed,
      };
      currentOptions = [];
      lastTopLevelIndex = i;
    } else if (block.nestLevel === 0 && !currentQuestion) {
      // First top-level item, start as question
      questionNumber++;
      currentQuestion = {
        number: questionNumber,
        question: trimmed,
      };
      currentOptions = [];
      lastTopLevelIndex = i;
    } else if (currentQuestion && currentOptions.length === 0 && block.nestLevel === 0) {
      // Continuation of question text at same level
      currentQuestion.question =
        (currentQuestion.question || "") + " " + trimmed;
    }
  }

  // Last question
  if (currentQuestion && currentQuestion.question) {
    finalizeQuestion(currentQuestion, currentOptions, questions, errors);
  }

  return { questions, errors };
}

/**
 * Fallback: parse as flat text lines (for simpler formatted documents)
 */
function parseFromTextLines(html: string): ParseResult {
  const questions: ParsedQuestion[] = [];
  const errors: string[] = [];

  const lines = extractFlatLines(html);

  let currentQuestion: Partial<ParsedQuestion> | null = null;
  let currentOptions: { label: string; text: string; isBold: boolean }[] = [];

  for (const line of lines) {
    const trimmed = line.text.trim();
    if (!trimmed) continue;

    const questionMatch = trimmed.match(/^(\d+)[.)]\s*(.+)/);
    const optionMatch = trimmed.match(/^([A-Da-d])[.)]\s*(.+)/);

    if (questionMatch && !optionMatch) {
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
      currentOptions.push({ label, text, isBold: line.isBold });
    } else if (currentQuestion && currentOptions.length === 0) {
      currentQuestion.question =
        (currentQuestion.question || "") + " " + trimmed;
    }
  }

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

/**
 * Extract blocks with nesting level from HTML.
 * Tracks <ol>/<ul> depth to determine if something is a question or option.
 */
function extractBlocksWithNesting(html: string): TextBlock[] {
  const blocks: TextBlock[] = [];

  // Process HTML by tracking list nesting
  let nestLevel = 0;
  const parts = html.split(/(<\/?(?:ol|ul|li|p|h[1-6])[^>]*>)/gi);

  let currentText = "";
  let currentBold = false;
  let inListItem = false;

  for (const part of parts) {
    const lowerPart = part.toLowerCase();

    if (lowerPart.startsWith("<ol") || lowerPart.startsWith("<ul")) {
      nestLevel++;
    } else if (lowerPart === "</ol>" || lowerPart === "</ul>") {
      nestLevel--;
      if (nestLevel < 0) nestLevel = 0;
    } else if (lowerPart.startsWith("<li")) {
      // Start of list item — flush previous
      if (currentText.trim()) {
        const text = currentText.replace(/<[^>]+>/g, "").trim();
        if (text) {
          blocks.push({
            text,
            isBold: isContentBold(currentText),
            nestLevel: Math.max(0, nestLevel - 1),
          });
        }
      }
      currentText = "";
      inListItem = true;
    } else if (lowerPart === "</li>") {
      // End of list item
      if (currentText.trim()) {
        const text = currentText.replace(/<[^>]+>/g, "").trim();
        if (text) {
          blocks.push({
            text,
            isBold: isContentBold(currentText),
            nestLevel: Math.max(0, nestLevel - 1),
          });
        }
      }
      currentText = "";
      inListItem = false;
    } else if (lowerPart.startsWith("<p") || lowerPart.startsWith("<h")) {
      // Paragraph/heading start — flush previous
      if (currentText.trim()) {
        const text = currentText.replace(/<[^>]+>/g, "").trim();
        if (text) {
          blocks.push({
            text,
            isBold: isContentBold(currentText),
            nestLevel: 0,
          });
        }
      }
      currentText = "";
    } else if (lowerPart === "</p>" || /^<\/h[1-6]>$/.test(lowerPart)) {
      // Paragraph/heading end
      if (currentText.trim()) {
        const text = currentText.replace(/<[^>]+>/g, "").trim();
        if (text) {
          blocks.push({
            text,
            isBold: isContentBold(currentText),
            nestLevel: 0,
          });
        }
      }
      currentText = "";
    } else {
      currentText += part;
    }
  }

  // Flush remaining
  if (currentText.trim()) {
    const text = currentText.replace(/<[^>]+>/g, "").trim();
    if (text) {
      blocks.push({
        text,
        isBold: isContentBold(currentText),
        nestLevel: 0,
      });
    }
  }

  return blocks;
}

function extractFlatLines(html: string): TextBlock[] {
  const lines: TextBlock[] = [];
  const segments = html.split(/<\/p>|<\/h[1-6]>|<\/li>|<br\s*\/?>/gi);

  for (const segment of segments) {
    const clean = segment
      .replace(/<p[^>]*>|<h[1-6][^>]*>|<li[^>]*>|<ul[^>]*>|<ol[^>]*>|<\/ul>|<\/ol>/gi, "")
      .trim();
    if (!clean) continue;

    const text = clean.replace(/<[^>]+>/g, "").trim();
    if (text) {
      lines.push({ text, isBold: isContentBold(clean), nestLevel: 0 });
    }
  }

  return lines;
}

function isContentBold(html: string): boolean {
  const textContent = html.replace(/<[^>]+>/g, "").trim();
  if (!textContent) return false;

  const boldContent = extractBoldText(html).trim();

  // If more than 50% of text is bold, consider it bold
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
