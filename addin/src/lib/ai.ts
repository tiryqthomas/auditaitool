import type { ExtractedField, Extractor } from "./extraction";
import type { SheetAnswer } from "./spreadsheet";

// Extraction runs in a Supabase Edge Function (supabase/functions/extract) that
// holds the Anthropic API key. The anon key below is public by design: it only
// lets the add-in call the function, never read the API key.
const FUNCTIONS = "https://xzxdmicwrtuqhiefrtlw.supabase.co/functions/v1";
const SUPABASE_ANON_KEY =
  "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Inh6eGRtaWN3cnR1cWhpZWZydGx3Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3OTI0NTM3MDUsImV4cCI6MjEwODAyOTcwNX0.T3KwolwVOSfBuQfiD6mZ6RTqyKC-7_GychechDUMkR8";

async function callFunction<T>(name: string, payload: unknown): Promise<T> {
  let res: Response;
  try {
    res = await fetch(`${FUNCTIONS}/${name}`, {
      method: "POST",
      headers: { Authorization: `Bearer ${SUPABASE_ANON_KEY}`, "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
  } catch {
    throw new Error("Couldn't reach the AI service. Check your internet connection.");
  }
  const body = (await res.json().catch(() => ({}))) as T & { error?: string };
  if (!res.ok) throw new Error(body.error ?? `The AI service returned an error (${res.status}).`);
  return body;
}

const supabaseExtractor: Extractor = {
  async extract(pdfBase64, fields, instructions) {
    const body = await callFunction<{ fields?: ExtractedField[] }>("extract", { pdfBase64, fields, instructions });
    return body.fields ?? [];
  },
};

/** Asks a plain-language question about a spreadsheet; each answer names its source cell. */
export async function askSheet(workbookText: string, question: string): Promise<SheetAnswer[]> {
  const body = await callFunction<{ answers?: SheetAnswer[] }>("ask-sheet", { workbookText, question });
  return body.answers ?? [];
}

export function getExtractor(): Extractor | null {
  return supabaseExtractor;
}
