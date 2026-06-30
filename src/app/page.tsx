"use client";

import { useSession, signIn, signOut } from "next-auth/react";
import { useState, useRef } from "react";

interface GenerateResult {
  success: boolean;
  formUrl?: string;
  editUrl?: string;
  formId?: string;
  questionsProcessed?: number;
  errors?: string[];
  error?: string;
  details?: string[];
}

export default function Home() {
  const { data: session, status } = useSession();
  const [file, setFile] = useState<File | null>(null);
  const [title, setTitle] = useState("");
  const [loading, setLoading] = useState(false);
  const [progress, setProgress] = useState("");
  const [result, setResult] = useState<GenerateResult | null>(null);
  const [error, setError] = useState<string | null>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const selected = e.target.files?.[0];
    if (selected) {
      if (selected.size > 10 * 1024 * 1024) {
        setError("File size exceeds the 10 MB limit.");
        setFile(null);
        return;
      }
      if (!selected.name.endsWith(".docx")) {
        setError("Only .docx files are supported.");
        setFile(null);
        return;
      }
      setFile(selected);
      setError(null);
      setResult(null);
    }
  };

  const handleSubmit = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!file) return;

    setLoading(true);
    setError(null);
    setResult(null);
    setProgress("Uploading file...");

    try {
      const formData = new FormData();
      formData.append("file", file);
      formData.append("title", title || file.name.replace(".docx", ""));

      setProgress("Parsing document...");

      const response = await fetch("/api/generate", {
        method: "POST",
        body: formData,
      });

      setProgress("Generating Google Form...");

      const data: GenerateResult = await response.json();

      if (!response.ok) {
        setError(data.error || "Failed to generate form.");
        if (data.details) {
          setError(
            (data.error || "") + "\n" + data.details.join("\n")
          );
        }
      } else {
        setResult(data);
      }
    } catch (err) {
      setError(
        err instanceof Error ? err.message : "An unexpected error occurred."
      );
    } finally {
      setLoading(false);
      setProgress("");
    }
  };

  const resetForm = () => {
    setFile(null);
    setTitle("");
    setResult(null);
    setError(null);
    if (fileInputRef.current) {
      fileInputRef.current.value = "";
    }
  };

  // Loading auth state
  if (status === "loading") {
    return (
      <main className="flex-1 flex items-center justify-center">
        <div className="animate-pulse text-gray-500">Loading...</div>
      </main>
    );
  }

  // Not authenticated
  if (!session) {
    return (
      <main className="flex-1 flex items-center justify-center p-4">
        <div className="w-full max-w-md text-center space-y-6">
          <div className="space-y-2">
            <h1 className="text-3xl font-bold text-gray-900">
              Word → Google Form
            </h1>
            <p className="text-gray-600">
              Convert your Word documents into Google Forms quizzes
              automatically.
            </p>
          </div>

          <div className="bg-white rounded-xl shadow-sm border border-gray-200 p-8 space-y-4">
            <p className="text-sm text-gray-600">
              Sign in with your Google account to get started. We need
              permission to create Google Forms on your behalf.
            </p>
            <button
              onClick={() => signIn("google")}
              className="w-full flex items-center justify-center gap-3 px-4 py-3 bg-white border border-gray-300 rounded-lg font-medium text-gray-700 hover:bg-gray-50 transition-colors cursor-pointer"
            >
              <svg className="w-5 h-5" viewBox="0 0 24 24">
                <path
                  fill="#4285F4"
                  d="M22.56 12.25c0-.78-.07-1.53-.2-2.25H12v4.26h5.92c-.26 1.37-1.04 2.53-2.21 3.31v2.77h3.57c2.08-1.92 3.28-4.74 3.28-8.09z"
                />
                <path
                  fill="#34A853"
                  d="M12 23c2.97 0 5.46-.98 7.28-2.66l-3.57-2.77c-.98.66-2.23 1.06-3.71 1.06-2.86 0-5.29-1.93-6.16-4.53H2.18v2.84C3.99 20.53 7.7 23 12 23z"
                />
                <path
                  fill="#FBBC05"
                  d="M5.84 14.09c-.22-.66-.35-1.36-.35-2.09s.13-1.43.35-2.09V7.07H2.18C1.43 8.55 1 10.22 1 12s.43 3.45 1.18 4.93l2.85-2.22.81-.62z"
                />
                <path
                  fill="#EA4335"
                  d="M12 5.38c1.62 0 3.06.56 4.21 1.64l3.15-3.15C17.45 2.09 14.97 1 12 1 7.7 1 3.99 3.47 2.18 7.07l3.66 2.84c.87-2.6 3.3-4.53 6.16-4.53z"
                />
              </svg>
              Sign in with Google
            </button>
          </div>
        </div>
      </main>
    );
  }

  // Authenticated - show upload form or result
  return (
    <main className="flex-1 flex flex-col">
      {/* Header */}
      <header className="border-b border-gray-200 bg-white">
        <div className="max-w-3xl mx-auto px-4 py-4 flex items-center justify-between">
          <h1 className="text-lg font-semibold text-gray-900">
            Word → Google Form
          </h1>
          <div className="flex items-center gap-3">
            <span className="text-sm text-gray-600">
              {session.user?.email}
            </span>
            <button
              onClick={() => signOut()}
              className="text-sm text-gray-500 hover:text-gray-700 cursor-pointer"
            >
              Sign out
            </button>
          </div>
        </div>
      </header>

      {/* Content */}
      <div className="flex-1 flex items-center justify-center p-4">
        <div className="w-full max-w-xl">
          {result?.success ? (
            // Success result
            <div className="bg-white rounded-xl shadow-sm border border-gray-200 p-8 space-y-6">
              <div className="text-center space-y-2">
                <div className="w-12 h-12 bg-green-100 rounded-full flex items-center justify-center mx-auto">
                  <svg
                    className="w-6 h-6 text-green-600"
                    fill="none"
                    viewBox="0 0 24 24"
                    stroke="currentColor"
                  >
                    <path
                      strokeLinecap="round"
                      strokeLinejoin="round"
                      strokeWidth={2}
                      d="M5 13l4 4L19 7"
                    />
                  </svg>
                </div>
                <h2 className="text-xl font-semibold text-gray-900">
                  Form Created Successfully!
                </h2>
                <p className="text-sm text-gray-600">
                  {result.questionsProcessed} questions processed
                </p>
              </div>

              {result.errors && result.errors.length > 0 && (
                <div className="bg-yellow-50 border border-yellow-200 rounded-lg p-4">
                  <p className="text-sm font-medium text-yellow-800 mb-1">
                    Warnings:
                  </p>
                  <ul className="text-sm text-yellow-700 space-y-1">
                    {result.errors.map((err, i) => (
                      <li key={i}>{err}</li>
                    ))}
                  </ul>
                </div>
              )}

              <div className="space-y-3">
                <div className="bg-gray-50 rounded-lg p-3">
                  <label className="text-xs font-medium text-gray-500 uppercase tracking-wider">
                    Google Form URL
                  </label>
                  <a
                    href={result.formUrl}
                    target="_blank"
                    rel="noopener noreferrer"
                    className="block text-sm text-blue-600 hover:text-blue-800 break-all mt-1"
                  >
                    {result.formUrl}
                  </a>
                </div>

                <a
                  href={result.formUrl}
                  target="_blank"
                  rel="noopener noreferrer"
                  className="block w-full text-center px-4 py-3 bg-blue-600 text-white font-medium rounded-lg hover:bg-blue-700 transition-colors"
                >
                  Open Google Form
                </a>

                <button
                  onClick={resetForm}
                  className="block w-full text-center px-4 py-3 bg-gray-100 text-gray-700 font-medium rounded-lg hover:bg-gray-200 transition-colors cursor-pointer"
                >
                  Generate Another Form
                </button>
              </div>
            </div>
          ) : (
            // Upload form
            <form
              onSubmit={handleSubmit}
              className="bg-white rounded-xl shadow-sm border border-gray-200 p-8 space-y-6"
            >
              <div className="space-y-1">
                <h2 className="text-xl font-semibold text-gray-900">
                  Upload Word Document
                </h2>
                <p className="text-sm text-gray-600">
                  Upload a .docx file with quiz questions. Bold options will
                  be marked as correct answers.
                </p>
              </div>

              {/* Title input */}
              <div className="space-y-2">
                <label
                  htmlFor="title"
                  className="block text-sm font-medium text-gray-700"
                >
                  Form Title
                </label>
                <input
                  id="title"
                  type="text"
                  value={title}
                  onChange={(e) => setTitle(e.target.value)}
                  placeholder="e.g., Pre-Test Module 1"
                  className="w-full px-3 py-2 border border-gray-300 rounded-lg text-sm focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent"
                />
              </div>

              {/* File upload */}
              <div className="space-y-2">
                <label className="block text-sm font-medium text-gray-700">
                  Word File (.docx)
                </label>
                <div
                  className="border-2 border-dashed border-gray-300 rounded-lg p-6 text-center hover:border-gray-400 transition-colors cursor-pointer"
                  onClick={() => fileInputRef.current?.click()}
                >
                  <input
                    ref={fileInputRef}
                    type="file"
                    accept=".docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    onChange={handleFileChange}
                    className="hidden"
                  />
                  {file ? (
                    <div className="space-y-1">
                      <svg
                        className="w-8 h-8 mx-auto text-blue-500"
                        fill="none"
                        viewBox="0 0 24 24"
                        stroke="currentColor"
                      >
                        <path
                          strokeLinecap="round"
                          strokeLinejoin="round"
                          strokeWidth={2}
                          d="M9 12h6m-6 4h6m2 5H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"
                        />
                      </svg>
                      <p className="text-sm font-medium text-gray-900">
                        {file.name}
                      </p>
                      <p className="text-xs text-gray-500">
                        {(file.size / 1024).toFixed(1)} KB
                      </p>
                    </div>
                  ) : (
                    <div className="space-y-1">
                      <svg
                        className="w-8 h-8 mx-auto text-gray-400"
                        fill="none"
                        viewBox="0 0 24 24"
                        stroke="currentColor"
                      >
                        <path
                          strokeLinecap="round"
                          strokeLinejoin="round"
                          strokeWidth={2}
                          d="M7 16a4 4 0 01-.88-7.903A5 5 0 1115.9 6L16 6a5 5 0 011 9.9M15 13l-3-3m0 0l-3 3m3-3v12"
                        />
                      </svg>
                      <p className="text-sm text-gray-600">
                        Click to select a file or drag & drop
                      </p>
                      <p className="text-xs text-gray-400">
                        .docx files up to 10 MB
                      </p>
                    </div>
                  )}
                </div>
              </div>

              {/* Error display */}
              {error && (
                <div className="bg-red-50 border border-red-200 rounded-lg p-4">
                  <p className="text-sm text-red-700 whitespace-pre-line">
                    {error}
                  </p>
                </div>
              )}

              {/* Progress */}
              {loading && (
                <div className="flex items-center gap-3 p-4 bg-blue-50 rounded-lg">
                  <svg
                    className="animate-spin w-5 h-5 text-blue-600"
                    fill="none"
                    viewBox="0 0 24 24"
                  >
                    <circle
                      className="opacity-25"
                      cx="12"
                      cy="12"
                      r="10"
                      stroke="currentColor"
                      strokeWidth="4"
                    />
                    <path
                      className="opacity-75"
                      fill="currentColor"
                      d="M4 12a8 8 0 018-8V0C5.373 0 0 5.373 0 12h4zm2 5.291A7.962 7.962 0 014 12H0c0 3.042 1.135 5.824 3 7.938l3-2.647z"
                    />
                  </svg>
                  <span className="text-sm text-blue-700">{progress}</span>
                </div>
              )}

              {/* Submit */}
              <button
                type="submit"
                disabled={!file || loading}
                className="w-full px-4 py-3 bg-blue-600 text-white font-medium rounded-lg hover:bg-blue-700 transition-colors disabled:opacity-50 disabled:cursor-not-allowed cursor-pointer"
              >
                {loading ? "Generating..." : "Generate Google Form"}
              </button>

              {/* Format guide */}
              <details className="text-sm text-gray-600" open>
                <summary className="cursor-pointer font-medium hover:text-gray-800">
                  Format Dokumen (PENTING)
                </summary>
                <div className="mt-3 bg-gray-50 rounded-lg p-4 space-y-3 text-xs">
                  <div className="bg-yellow-50 border border-yellow-200 rounded p-2">
                    <p className="font-semibold text-yellow-800">Aturan:</p>
                    <ul className="list-disc list-inside text-yellow-700 mt-1 space-y-1">
                      <li>Ketik nomor soal MANUAL (1. 2. 3.) — jangan pakai auto-numbering Word</li>
                      <li>Ketik huruf opsi MANUAL (A. B. C. D.) — jangan pakai bullet/list otomatis</li>
                      <li>BOLD jawaban yang benar (Select → Ctrl+B)</li>
                      <li>Setiap soal harus punya 4 opsi</li>
                    </ul>
                  </div>
                  <p className="font-medium">Contoh format:</p>
                  <pre className="bg-gray-100 p-3 rounded text-xs overflow-x-auto">
                    {`1. Apa ibukota Indonesia?
A. Bangkok
B. Jakarta    ← Bold ini di Word
C. Manila
D. Kuala Lumpur

2. Bahasa pemrograman untuk web:
A. Python
B. JavaScript    ← Bold ini di Word
C. C++
D. Semua benar`}
                  </pre>
                  <a
                    href="/template-guide.txt"
                    download
                    className="inline-block mt-2 text-blue-600 hover:text-blue-800 font-medium"
                  >
                    Download Template Guide →
                  </a>
                </div>
              </details>
            </form>
          )}
        </div>
      </div>
    </main>
  );
}
