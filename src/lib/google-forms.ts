import { google } from "googleapis";
import { ParsedQuestion } from "./docx-parser";

export interface FormResult {
  formId: string;
  formUrl: string;
  editUrl: string;
}

/**
 * Creates a Google Form in Quiz mode with the parsed questions.
 */
export async function createGoogleForm(
  accessToken: string,
  title: string,
  questions: ParsedQuestion[]
): Promise<FormResult> {
  const auth = new google.auth.OAuth2();
  auth.setCredentials({ access_token: accessToken });

  const forms = google.forms({ version: "v1", auth });

  // Step 1: Create the form
  const createResponse = await forms.forms.create({
    requestBody: {
      info: {
        title,
      },
    },
  });

  const formId = createResponse.data.formId!;

  // Step 2: Enable Quiz mode
  await forms.forms.batchUpdate({
    formId,
    requestBody: {
      requests: [
        {
          updateSettings: {
            settings: {
              quizSettings: {
                isQuiz: true,
              },
            },
            updateMask: "quizSettings.isQuiz",
          },
        },
      ],
    },
  });

  // Step 3: Add questions
  const questionRequests = questions.map((q, index) => {
    const options = q.options.map((opt) => ({
      value: opt.text,
    }));

    const correctIndex = q.options.findIndex(
      (opt) => opt.label === q.correctAnswer
    );

    return {
      createItem: {
        item: {
          title: `${q.number}. ${q.question}`,
          questionItem: {
            question: {
              required: true,
              grading: {
                pointValue: 1,
                correctAnswers: {
                  answers:
                    correctIndex >= 0
                      ? [{ value: options[correctIndex].value }]
                      : [],
                },
              },
              choiceQuestion: {
                type: "RADIO" as const,
                options,
              },
            },
          },
        },
        location: {
          index,
        },
      },
    };
  });

  if (questionRequests.length > 0) {
    await forms.forms.batchUpdate({
      formId,
      requestBody: {
        requests: questionRequests,
      },
    });
  }

  const formUrl = `https://docs.google.com/forms/d/${formId}/viewform`;
  const editUrl = `https://docs.google.com/forms/d/${formId}/edit`;

  return { formId, formUrl, editUrl };
}
