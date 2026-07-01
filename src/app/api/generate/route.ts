import { NextRequest, NextResponse } from "next/server";
import { auth } from "@/lib/auth";
import { google } from "googleapis";
import { QuizItem } from "@/lib/docx-numbering-parser";

const MAX_FILE_SIZE = 50 * 1024 * 1024; // 50 MB

export async function POST(request: NextRequest) {
  try {
    const session = await auth();

    if (!session?.accessToken) {
      return NextResponse.json(
        { error: "Unauthorized. Please login with Google." },
        { status: 401 }
      );
    }

    const formData = await request.formData();
    const file = formData.get("file") as File | null;
    const title = (formData.get("title") as string) || "Quiz";
    const subtitle = (formData.get("subtitle") as string) || "";
    const customFieldsRaw = (formData.get("customFields") as string) || "";
    const customFields = customFieldsRaw.split(",").map(f => f.trim()).filter(f => f);

    if (!file) {
      return NextResponse.json(
        { error: "No file uploaded." },
        { status: 400 }
      );
    }

    // Validate file type
    if (!file.name.endsWith(".docx")) {
      return NextResponse.json(
        { error: "Invalid file type. Only .docx files are supported." },
        { status: 400 }
      );
    }

    // Validate file size
    if (file.size > MAX_FILE_SIZE) {
      return NextResponse.json(
        { error: "File size exceeds the 50 MB limit." },
        { status: 400 }
      );
    }

    // Parse the DOCX file using XML-based parser
    const arrayBuffer = await file.arrayBuffer();
    const buffer = Buffer.from(arrayBuffer);

    let items: QuizItem[];
    let debugParagraphs: { text: string; role: string; numId: number | null; ilvl: number | null; numFormat: string | null; isBold: boolean; rawBoldInfo?: string }[] = [];
    try {
      const { parseDocxViaXml, buildQuizFromParagraphs, parseDocxRawDebug } = await import("@/lib/docx-numbering-parser");
      const paragraphs = await parseDocxViaXml(buffer);
      
      // Get raw debug info for bold analysis
      let rawDebug: { text: string; rPrKeys: string }[] = [];
      try {
        rawDebug = await parseDocxRawDebug(buffer);
      } catch { /* ignore */ }
      
      // Build debug info (show first 50 paragraphs with longer text)
      debugParagraphs = paragraphs.slice(0, 50).map((p, idx) => ({
        text: p.text.substring(0, 120),
        role: "pending",
        numId: p.numId,
        ilvl: p.ilvl,
        numFormat: p.numFormat,
        isBold: p.isBold,
        rawBoldInfo: rawDebug[idx]?.rPrKeys || undefined,
      }));
      
      items = buildQuizFromParagraphs(paragraphs);
    } catch (err) {
      console.error("DOCX parse error:", err);
      return NextResponse.json(
        { error: "Gagal membaca file. Pastikan file tidak corrupt." },
        { status: 422 }
      );
    }

    if (items.length === 0) {
      return NextResponse.json(
        {
          error:
            "Tidak ada soal terdeteksi. Pastikan format soal & opsi bernomor/berhuruf.",
          debug: debugParagraphs,
        },
        { status: 422 }
      );
    }

    // Log for monitoring (keep for first weeks in production)
    console.log(
      `[parse-quiz] File: ${file.name} | Soal detected: ${items.length} | Soal tanpa jawaban: ${items.filter((i) => i.jawabanIndex === undefined).length}`
    );

    // Create Google Form
    const formResult = await createGoogleFormQuiz(
      session.accessToken,
      title || file.name.replace(/\.docx$/i, ""),
      subtitle,
      customFields,
      items
    );

    // Build warnings
    const errors: string[] = [];
    items.forEach((item, idx) => {
      if (item.jawabanIndex === undefined) {
        errors.push(`Soal nomor ${idx + 1} tidak memiliki kunci jawaban.`);
      }
    });

    return NextResponse.json({
      success: true,
      formUrl: formResult.formUrl,
      editUrl: formResult.editUrl,
      formId: formResult.formId,
      questionsProcessed: items.length,
      errors,
      debug: debugParagraphs,
      parsedQuiz: items.map((item, idx) => ({
        no: idx + 1,
        soal: item.soal.substring(0, 80),
        opsiCount: item.opsi.length,
        opsi: item.opsi.map((o, i) => ({ text: o.substring(0, 60), isAnswer: item.jawabanIndex === i })),
        jawabanIndex: item.jawabanIndex ?? null,
      })),
    });
  } catch (error: unknown) {
    console.error("Generate error:", error);
    const message =
      error instanceof Error ? error.message : "An unexpected error occurred.";
    return NextResponse.json(
      { error: `Failed to generate form: ${message}` },
      { status: 500 }
    );
  }
}

// ---------- Google Form Creation ----------

interface FormResult {
  formId: string;
  formUrl: string;
  editUrl: string;
}

async function createGoogleFormQuiz(
  accessToken: string,
  title: string,
  subtitle: string,
  customFields: string[],
  items: QuizItem[]
): Promise<FormResult> {
  const oAuth = new google.auth.OAuth2();
  oAuth.setCredentials({ access_token: accessToken });

  const forms = google.forms({ version: "v1", auth: oAuth });

  // Step 1: Create form
  const createRes = await forms.forms.create({
    requestBody: {
      info: { 
        title,
        documentTitle: title,
      },
    },
  });

  const formId = createRes.data.formId!;

  // Step 2: Update description if subtitle provided
  if (subtitle) {
    await forms.forms.batchUpdate({
      formId,
      requestBody: {
        requests: [{
          updateFormInfo: {
            info: { description: subtitle },
            updateMask: "description",
          },
        }],
      },
    });
  }

  // Step 3: Enable quiz mode + add custom fields + section + questions
  const requests: object[] = [
    {
      updateSettings: {
        settings: { 
          quizSettings: { isQuiz: true },
        },
        updateMask: "quizSettings.isQuiz",
      },
    },
  ];

  let currentIndex = 0;

  // Add custom fields (e.g. Email, Nama Lengkap)
  for (const fieldName of customFields) {
    requests.push({
      createItem: {
        item: {
          title: fieldName,
          questionItem: {
            question: {
              required: true,
              textQuestion: {
                paragraph: false,
              },
            },
          },
        },
        location: { index: currentIndex++ },
      },
    });
  }

  // Add section header before quiz questions
  requests.push({
    createItem: {
      item: {
        title: "Soal Pilihan Ganda",
        description: "Pilihlah salah satu jawaban yang paling tepat!",
        pageBreakItem: {},
      },
      location: { index: currentIndex++ },
    },
  });

  // Add quiz questions
  items.forEach((item, idx) => {
    requests.push({
      createItem: {
        item: {
          title: `${idx + 1}. ${item.soal}`,
          questionItem: {
            question: {
              required: true,
              choiceQuestion: {
                type: "RADIO",
                options: item.opsi.map((opsiText, opsiIdx) => ({
                  value: `${String.fromCharCode(65 + opsiIdx)}. ${opsiText}`,
                })),
              },
              grading:
                item.jawabanIndex !== undefined
                  ? {
                      pointValue: 1,
                      correctAnswers: {
                        answers: [
                          {
                            value: `${String.fromCharCode(65 + item.jawabanIndex)}. ${item.opsi[item.jawabanIndex]}`,
                          },
                        ],
                      },
                    }
                  : undefined,
            },
          },
        },
        location: { index: currentIndex++ },
      },
    });
  });

  await forms.forms.batchUpdate({
    formId,
    requestBody: { requests },
  });

  const formUrl = `https://docs.google.com/forms/d/${formId}/viewform`;
  const editUrl = `https://docs.google.com/forms/d/${formId}/edit`;

  return { formId, formUrl, editUrl };
}
