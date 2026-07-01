import { NextResponse } from "next/server";
import JSZip from "jszip";

/**
 * Generate a proper .docx template file for download.
 * Creates a minimal valid .docx with sample questions in correct format.
 */
export async function GET() {
  const zip = new JSZip();

  // [Content_Types].xml
  zip.file(
    "[Content_Types].xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`
  );

  // _rels/.rels
  zip.file(
    "_rels/.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`
  );

  // word/_rels/document.xml.rels
  zip.file(
    "word/_rels/document.xml.rels",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
</Relationships>`
  );

  // word/document.xml - the actual content
  const bodyContent = generateTemplateContent();
  zip.file(
    "word/document.xml",
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
${bodyContent}
  </w:body>
</w:document>`
  );

  const buffer = await zip.generateAsync({ type: "nodebuffer" });

  return new NextResponse(buffer as unknown as BodyInit, {
    headers: {
      "Content-Type":
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      "Content-Disposition":
        'attachment; filename="Template_Soal_Quiz.docx"',
    },
  });
}

function generateTemplateContent(): string {
  const paragraphs: string[] = [];

  // Title
  paragraphs.push(makeParagraph("TEMPLATE SOAL QUIZ", true));
  paragraphs.push(makeParagraph(""));
  paragraphs.push(
    makeParagraph(
      "Petunjuk: Ketik soal dengan format di bawah. BOLD jawaban yang benar (Ctrl+B)."
    )
  );
  paragraphs.push(makeParagraph("Hapus petunjuk ini sebelum upload."));
  paragraphs.push(makeParagraph(""));

  // Sample Question 1
  paragraphs.push(makeParagraph("1. Apa ibukota Indonesia?"));
  paragraphs.push(makeParagraph("A. Bangkok"));
  paragraphs.push(makeParagraph("B. Jakarta", true)); // Bold = correct
  paragraphs.push(makeParagraph("C. Manila"));
  paragraphs.push(makeParagraph("D. Kuala Lumpur"));
  paragraphs.push(makeParagraph(""));

  // Sample Question 2
  paragraphs.push(makeParagraph("2. Planet terbesar di tata surya adalah:"));
  paragraphs.push(makeParagraph("A. Mars"));
  paragraphs.push(makeParagraph("B. Saturnus"));
  paragraphs.push(makeParagraph("C. Jupiter", true)); // Bold = correct
  paragraphs.push(makeParagraph("D. Neptunus"));
  paragraphs.push(makeParagraph(""));

  // Sample Question 3
  paragraphs.push(makeParagraph("3. Hasil dari 5 x 8 adalah:"));
  paragraphs.push(makeParagraph("A. 35"));
  paragraphs.push(makeParagraph("B. 40", true)); // Bold = correct
  paragraphs.push(makeParagraph("C. 45"));
  paragraphs.push(makeParagraph("D. 50"));
  paragraphs.push(makeParagraph(""));

  // Instructions for user
  paragraphs.push(makeParagraph(""));
  paragraphs.push(
    makeParagraph("--- Tambahkan soal Anda di bawah ini ---")
  );
  paragraphs.push(makeParagraph(""));

  return paragraphs.join("\n");
}

function makeParagraph(text: string, bold = false): string {
  if (!text) {
    return `    <w:p><w:r><w:t></w:t></w:r></w:p>`;
  }

  const rPr = bold ? `<w:rPr><w:b/></w:rPr>` : "";
  // xml:space="preserve" to keep spaces
  return `    <w:p><w:r>${rPr}<w:t xml:space="preserve">${escapeXml(text)}</w:t></w:r></w:p>`;
}

function escapeXml(text: string): string {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&apos;");
}
