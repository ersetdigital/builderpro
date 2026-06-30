import JSZip from "jszip";

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

interface Paragraph {
  text: string;
  isBold: boolean;
}

/**
 * Parse a DOCX file by reading the XML directly.
 * A .docx is a ZIP containing word/document.xml.
 * This gives us full control over text and formatting detection.
 */
export async function parseDocx(buffer: Buffer): Promise<ParseResult> {
  const zip = await JSZip.loadAsync(buffer);
  const docXml = await zip.file("word/document.xml")?.async("string");

  if (!docXml) {
    return { questions: [], errors: ["Could not read document.xml from the DOCX file."] };
  }

  // Extract paragraphs with their text and bold status
  const paragraphs = extractParagraphs(docXml);

  // Parse into questions
  const questions = parseQuestions(paragraphs);

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

/**
 * Extract paragraphs from the DOCX XML.
 * Each <w:p> is a paragraph. Inside it, <w:r> are runs of text.
 * Bold is indicated by <w:b/> or <w:b w:val="true"/> in <w:rPr>.
 */
function extractParagraphs(xml: string): Paragraph[] {
  const paragraphs: Paragraph[] = [];

  // Split by paragraph tags
  const pMatches = xml.match(/<w:p[ >][\s\S]*?<\/w:p>/g);
  if (!pMatches) return [];

  for (const pXml of pMatches) {
    // Check for paragraph-level bold (in pPr > rPr)
    const pPrBold = hasParagraphBold(pXml);

    // Extract all runs in this paragraph
    const runs = pXml.match(/<w:r[ >][\s\S]*?<\/w:r>/g);
    if (!runs) continue;

    let paragraphText = "";
    let boldCharCount = 0;
    let totalCharCount = 0;

    for (const run of runs) {
      // Get text from <w:t> tags
      const textMatches = run.match(/<w:t[^>]*>([\s\S]*?)<\/w:t>/g);
      if (!textMatches) continue;

      const runText = textMatches
        .map((t) => t.replace(/<[^>]+>/g, ""))
        .join("");

      if (!runText) continue;

      // Check if this run is bold
      const runIsBold = pPrBold || isRunBold(run);

      paragraphText += runText;
      totalCharCount += runText.length;
      if (runIsBold) boldCharCount += runText.length;
    }

    const text = paragraphText.trim();
    if (text) {
      const isBold = totalCharCount > 0 && boldCharCount >= totalCharCount * 0.5;
      paragraphs.push({ text, isBold });
    }
  }

  return paragraphs;
}

/**
 * Check if a run (<w:r>) has bold formatting.
 */
function isRunBold(runXml: string): boolean {
  // Look for <w:rPr> containing <w:b/> or <w:b w:val="1"/> or <w:b w:val="true"/>
  const rPrMatch = runXml.match(/<w:rPr>([\s\S]*?)<\/w:rPr>/);
  if (!rPrMatch) return false;

  const rPr = rPrMatch[1];

  // <w:b/> or <w:b w:val="..."/> (where val is not "0" or "false")
  if (/<w:b\/>/.test(rPr)) return true;
  if (/<w:b\s+w:val="0"/.test(rPr)) return false;
  if (/<w:b\s+w:val="false"/.test(rPr)) return false;
  if (/<w:b[\s>]/.test(rPr)) return true;

  return false;
}

/**
 * Check if a paragraph has bold set at the paragraph level (pPr > rPr > b).
 */
function hasParagraphBold(pXml: string): boolean {
  const pPrMatch = pXml.match(/<w:pPr>([\s\S]*?)<\/w:pPr>/);
  if (!pPrMatch) return false;

  const pPr = pPrMatch[1];
  const rPrMatch = pPr.match(/<w:rPr>([\s\S]*?)<\/w:rPr>/);
  if (!rPrMatch) return false;

  const rPr = rPrMatch[1];
  if (/<w:b\/>/.test(rPr)) return true;
  if (/<w:b\s+w:val="0"/.test(rPr)) return false;
  if (/<w:b\s+w:val="false"/.test(rPr)) return false;
  if (/<w:b[\s>]/.test(rPr)) return true;

  return false;
}

/**
 * Parse paragraphs into questions.
 * Simple and reliable: look for numbered lines as questions,
 * lettered lines as options.
 */
function parseQuestions(paragraphs: Paragraph[]): ParsedQuestion[] {
  const questions: ParsedQuestion[] = [];
  let currentQuestion: { number: number; question: string } | null = null;
  let currentOptions: { label: string; text: string; isBold: boolean }[] = [];

  for (const para of paragraphs) {
    const text = para.text.trim();
    if (!text) continue;

    // Match question: starts with number + . or )
    const questionMatch = text.match(/^(\d+)[.)]\s*(.*)/);
    // Match option: starts with a-d/A-D + . or )
    const optionMatch = text.match(/^([A-Da-d])[.)]\s*(.*)/);

    if (optionMatch && currentQuestion) {
      // This is an option
      const label = optionMatch[1].toUpperCase();
      const optText = optionMatch[2].trim();
      currentOptions.push({ label, text: optText, isBold: para.isBold });
    } else if (questionMatch) {
      // Save previous question
      if (currentQuestion) {
        finalizeQuestion(currentQuestion, currentOptions, questions);
      }

      const qNum = parseInt(questionMatch[1]);
      const qText = questionMatch[2].trim();

      // Skip empty question text (like stray "7." at end of soal 6)
      if (!qText) continue;

      currentQuestion = { number: qNum, question: qText };
      currentOptions = [];
    } else if (currentQuestion && currentOptions.length === 0) {
      // Continuation of question text
      currentQuestion.question += " " + text;
    }
  }

  // Last question
  if (currentQuestion) {
    finalizeQuestion(currentQuestion, currentOptions, questions);
  }

  return questions;
}

function finalizeQuestion(
  q: { number: number; question: string },
  options: { label: string; text: string; isBold: boolean }[],
  questions: ParsedQuestion[]
) {
  if (options.length === 0) return;

  const correctAnswer = options.find((o) => o.isBold)?.label || null;

  questions.push({
    number: q.number,
    question: q.question,
    options,
    correctAnswer,
  });
}
