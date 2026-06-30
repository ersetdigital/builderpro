import { NextRequest, NextResponse } from "next/server";
import { auth } from "@/lib/auth";
import { parseDocx } from "@/lib/docx-parser";
import { createGoogleForm } from "@/lib/google-forms";

const MAX_FILE_SIZE = 10 * 1024 * 1024; // 10 MB

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

    if (!file) {
      return NextResponse.json(
        { error: "No file uploaded." },
        { status: 400 }
      );
    }

    // Validate file type
    const validTypes = [
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    ];
    if (!validTypes.includes(file.type) && !file.name.endsWith(".docx")) {
      return NextResponse.json(
        { error: "Invalid file type. Only .docx files are supported." },
        { status: 400 }
      );
    }

    // Validate file size
    if (file.size > MAX_FILE_SIZE) {
      return NextResponse.json(
        { error: "File size exceeds the 10 MB limit." },
        { status: 400 }
      );
    }

    // Parse the DOCX file
    const arrayBuffer = await file.arrayBuffer();
    const buffer = Buffer.from(arrayBuffer);
    const parseResult = await parseDocx(buffer);

    if (parseResult.questions.length === 0) {
      return NextResponse.json(
        {
          error:
            "No questions found in the document. Please check the format.",
          details: parseResult.errors,
        },
        { status: 422 }
      );
    }

    // Create Google Form
    const formResult = await createGoogleForm(
      session.accessToken,
      title,
      parseResult.questions
    );

    return NextResponse.json({
      success: true,
      formUrl: formResult.formUrl,
      editUrl: formResult.editUrl,
      formId: formResult.formId,
      questionsProcessed: parseResult.questions.length,
      errors: parseResult.errors,
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
