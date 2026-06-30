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
 * Works by parsing the HTML tree structure from mammoth.
 */
export async function parseDocx(buffer: Buffer): Promise<ParseResult> {
  const htmlResult = await mammoth.convertToHtml({ buffer });
  const html = htmlResult.value;

  const questions = parseHtmlToQuestions(html);
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

interface Node {
  type: "element" | "text";
  tag?: string;
  children?: Node[];
  text?: string;
  html?: string;
}

/**
 * Simple HTML parser to build a tree structure.
 */
function parseHtmlTree(html: string): Node[] {
  const nodes: Node[] = [];
  const stack: Node[] = [];

  // Split into tags and text
  const regex = /(<\/?[a-z][a-z0-9]*[^>]*>)/gi;
  const parts = html.split(regex);

  for (const part of parts) {
    if (!part) continue;

    if (part.startsWith("</")) {
      // Closing tag - pop from stack
      if (stack.length > 0) {
        const node = stack.pop()!;
        if (stack.length > 0) {
          stack[stack.length - 1].children!.push(node);
        } else {
          nodes.push(node);
        }
      }
    } else if (part.startsWith("<")) {
      // Opening tag
      const tagMatch = part.match(/^<([a-z][a-z0-9]*)/i);
      if (tagMatch) {
        const tag = tagMatch[1].toLowerCase();
        // Self-closing tags
        if (part.endsWith("/>") || ["br", "hr", "img"].includes(tag)) {
          const node: Node = { type: "element", tag, children: [] };
          if (stack.length > 0) {
            stack[stack.length - 1].children!.push(node);
          } else {
            nodes.push(node);
          }
        } else {
          stack.push({ type: "element", tag, children: [] });
        }
      }
    } else {
      // Text node
      if (part.trim()) {
        const textNode: Node = { type: "text", text: part };
        if (stack.length > 0) {
          stack[stack.length - 1].children!.push(textNode);
        } else {
          nodes.push(textNode);
        }
      }
    }
  }

  // Flush remaining stack
  while (stack.length > 0) {
    const node = stack.pop()!;
    if (stack.length > 0) {
      stack[stack.length - 1].children!.push(node);
    } else {
      nodes.push(node);
    }
  }

  return nodes;
}

/**
 * Get plain text from a node tree.
 */
function getNodeText(node: Node): string {
  if (node.type === "text") return node.text || "";
  if (!node.children) return "";
  return node.children.map(getNodeText).join("");
}

/**
 * Check if a node contains bold (strong/b) formatting.
 */
function isNodeBold(node: Node): boolean {
  if (node.type === "text") return false;
  if (node.tag === "strong" || node.tag === "b") return true;
  if (!node.children) return false;

  const totalText = getNodeText(node).trim();
  if (!totalText) return false;

  const boldText = getBoldText(node).trim();
  return boldText.length >= totalText.length * 0.5;
}

function getBoldText(node: Node): string {
  if (node.type === "text") return "";
  if (node.tag === "strong" || node.tag === "b") return getNodeText(node);
  if (!node.children) return "";
  return node.children.map(getBoldText).join("");
}

/**
 * Main parsing logic: walk the HTML tree and extract questions.
 */
function parseHtmlToQuestions(html: string): ParsedQuestion[] {
  const tree = parseHtmlTree(html);
  const questions: ParsedQuestion[] = [];
  let questionCounter = 0;

  // Process all top-level nodes
  for (let i = 0; i < tree.length; i++) {
    const node = tree[i];

    if (node.type === "element" && (node.tag === "ol" || node.tag === "ul")) {
      // This is a top-level list — each <li> is potentially a question
      processTopLevelList(node, questions, questionCounter);
      questionCounter = questions.length;
    } else if (node.type === "element" && node.tag === "p") {
      // Paragraph — might be a question like "7. Standart..."
      const text = getNodeText(node).trim();
      const questionMatch = text.match(/^(\d+)[.)]\s*(.+)/);

      if (questionMatch) {
        // This is a numbered question in a paragraph
        const qNum = parseInt(questionMatch[1]);
        const qText = questionMatch[2];

        // Look ahead for options (next node might be ul/ol with options)
        const options: { label: string; text: string; isBold: boolean }[] = [];
        let j = i + 1;

        while (j < tree.length) {
          const next = tree[j];
          if (next.type === "element" && (next.tag === "ol" || next.tag === "ul")) {
            // Collect options from this list
            const listOptions = extractOptionsFromList(next);
            options.push(...listOptions);
            j++;
            break;
          } else if (next.type === "element" && next.tag === "p") {
            // Paragraph options (like soal 10)
            const pText = getNodeText(next).trim();
            if (pText && pText.length < 100) {
              const optLabels = ["A", "B", "C", "D", "E", "F"];
              const label = optLabels[options.length] || "?";
              options.push({ label, text: pText, isBold: isNodeBold(next) });
              j++;
            } else {
              break;
            }
          } else {
            break;
          }
        }

        if (options.length > 0) {
          const correctAnswer = options.find((o) => o.isBold)?.label || null;
          questions.push({ number: qNum, question: qText, options, correctAnswer });
          questionCounter = questions.length;
          i = j - 1; // Skip processed nodes
        }
      }
    }
  }

  return questions;
}

/**
 * Process a top-level <ol> where each <li> can be a question.
 */
function processTopLevelList(
  listNode: Node,
  questions: ParsedQuestion[],
  startNumber: number
): void {
  if (!listNode.children) return;

  const items = listNode.children.filter(
    (n) => n.type === "element" && n.tag === "li"
  );

  let questionNum = startNumber;

  for (const li of items) {
    if (!li.children) continue;

    // Separate the <li> content into:
    // - Direct text/inline content (the question)
    // - Nested <ol>/<ul> (the options)
    const questionParts: string[] = [];
    let questionIsBold = false;
    const nestedLists: Node[] = [];
    const directChildren: Node[] = [];

    for (const child of li.children) {
      if (child.type === "element" && (child.tag === "ol" || child.tag === "ul")) {
        nestedLists.push(child);
      } else {
        directChildren.push(child);
        questionParts.push(getNodeText(child));
      }
    }

    const questionText = questionParts.join("").trim();

    // If there's no question text and no nested list, this might be a flat option
    // that was incorrectly placed at top level (like soal 5-6 broken structure)
    if (!questionText && nestedLists.length === 0) continue;

    // If there IS question text but NO nested list, this <li> might be
    // a stray option from a broken structure. Skip short items without nested options.
    if (questionText && nestedLists.length === 0 && questionText.length < 50) {
      // This is likely a stray option, not a question. Skip it.
      continue;
    }

    // We have a question
    if (nestedLists.length > 0) {
      questionNum++;
      const options = extractOptionsFromList(nestedLists[0]);

      // For deeply nested lists (ul > li > ol > li pattern from soal 3)
      if (options.length === 0 && nestedLists[0].children) {
        for (const nestedChild of nestedLists[0].children) {
          if (nestedChild.type === "element" && (nestedChild.tag === "ol" || nestedChild.tag === "ul")) {
            const deepOptions = extractOptionsFromList(nestedChild);
            options.push(...deepOptions);
          } else if (nestedChild.type === "element" && nestedChild.tag === "li") {
            // li inside ul that contains another ol
            if (nestedChild.children) {
              for (const deepChild of nestedChild.children) {
                if (deepChild.type === "element" && (deepChild.tag === "ol" || deepChild.tag === "ul")) {
                  const deepOptions = extractOptionsFromList(deepChild);
                  options.push(...deepOptions);
                }
              }
            }
          }
        }
      }

      const correctAnswer = options.find((o) => o.isBold)?.label || null;
      questions.push({
        number: questionNum,
        question: questionText,
        options,
        correctAnswer,
      });
    } else if (questionText.length >= 50) {
      // Long text without nested list — question without properly nested options
      // Look for options in subsequent top-level <li> items (broken structure)
      questionNum++;
      questions.push({
        number: questionNum,
        question: questionText,
        options: [],
        correctAnswer: null,
      });
    }
  }
}

/**
 * Extract options from a nested list (ol/ul).
 * Each <li> is one option.
 */
function extractOptionsFromList(listNode: Node): { label: string; text: string; isBold: boolean }[] {
  if (!listNode.children) return [];

  const options: { label: string; text: string; isBold: boolean }[] = [];
  const optLabels = ["A", "B", "C", "D", "E", "F"];

  const items = listNode.children.filter(
    (n) => n.type === "element" && n.tag === "li"
  );

  for (const li of items) {
    const text = getNodeText(li).trim();
    if (!text) continue;

    const label = optLabels[options.length] || "?";
    const isBold = isNodeBold(li);

    options.push({ label, text, isBold });
  }

  return options;
}
