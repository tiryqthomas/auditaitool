import type { ExtractedField, Extractor } from "./extraction";

// Extraction runs in a Supabase Edge Function (supabase/functions/extract) that
// holds the Anthropic API key. The anon key below is public by design: it only
// lets the add-in call the function, never read the API key.
const FUNCTION_URL = "https://xzxdmicwrtuqhiefrtlw.supabase.co/functions/v1/extract";
const SUPABASE_ANON_KEY =
  "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Inh6eGRtaWN3cnR1cWhpZWZydGx3Iiwicm9sZSI6ImFub24iLCJpYXQiOjE3OTI0NTM3MDUsImV4cCI6MjEwODAyOTcwNX0.T3KwolwVOSfBuQfiD6mZ6RTqyKC-7_GychechDUMkR8";

const supabaseExtractor: Extractor = {
  async extract(pdfBase64, fields, instructions) {
    let res: Response;
    try {
      res = await fetch(FUNCTION_URL, {
        method: "POST",
        headers: { Authorization: `Bearer ${SUPABASE_ANON_KEY}`, "Content-Type": "application/json" },
        body: JSON.stringify({ pdfBase64, fields, instructions }),
      });
    } catch {
      throw new Error("Couldn't reach the AI service. Check your internet connection.");
    }
    const body = (await res.json().catch(() => ({}))) as { fields?: ExtractedField[]; error?: string };
    if (!res.ok) throw new Error(body.error ?? `The AI service returned an error (${res.status}).`);
    return body.fields ?? [];
  },
};

export function getExtractor(): Extractor | null {
  return supabaseExtractor;
}
