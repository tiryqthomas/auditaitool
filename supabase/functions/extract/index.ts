// Supabase Edge Function: the add-in's only route to Claude. The Anthropic API
// key lives in this project's secrets (ANTHROPIC_API_KEY), never in Excel.
import Anthropic from "@anthropic-ai/sdk";
import { ExtractError, extractFields, type ExtractRequest } from "./core.ts";

const ALLOWED_ORIGINS = ["https://tiryqthomas.github.io", "https://localhost:3000"];
const MAX_FIELDS = 30;
const MAX_PDF_BASE64 = 20 * 1024 * 1024; // ~15 MB PDF

const client = new Anthropic({ apiKey: Deno.env.get("ANTHROPIC_API_KEY") });

function cors(req: Request): Record<string, string> {
  const origin = req.headers.get("Origin") ?? "";
  return {
    "Access-Control-Allow-Origin": ALLOWED_ORIGINS.includes(origin) ? origin : ALLOWED_ORIGINS[0],
    "Access-Control-Allow-Headers": "authorization, apikey, content-type, x-client-info",
    "Access-Control-Allow-Methods": "POST, OPTIONS",
    Vary: "Origin",
  };
}

function json(req: Request, status: number, body: unknown): Response {
  return new Response(JSON.stringify(body), {
    status,
    headers: { ...cors(req), "Content-Type": "application/json" },
  });
}

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response("ok", { headers: cors(req) });
  if (req.method !== "POST") return json(req, 405, { error: "Use POST." });
  if (!Deno.env.get("ANTHROPIC_API_KEY")) return json(req, 500, { error: "The Anthropic API key hasn't been added to Supabase yet." });

  let body: ExtractRequest;
  try {
    body = await req.json();
  } catch {
    return json(req, 400, { error: "Request body must be JSON." });
  }
  const fields = Array.isArray(body.fields) ? body.fields.filter((f) => typeof f === "string" && f.trim()) : [];
  if (typeof body.pdfBase64 !== "string" || !body.pdfBase64) return json(req, 400, { error: "No document was sent." });
  if (body.pdfBase64.length > MAX_PDF_BASE64) return json(req, 413, { error: "That PDF is too large (limit about 15 MB)." });
  if (fields.length === 0 || fields.length > MAX_FIELDS) return json(req, 400, { error: `Send between 1 and ${MAX_FIELDS} fields.` });

  try {
    const results = await extractFields(client, {
      pdfBase64: body.pdfBase64,
      fields,
      instructions: typeof body.instructions === "string" ? body.instructions.slice(0, 2000) : undefined,
    });
    return json(req, 200, { fields: results });
  } catch (e) {
    if (e instanceof ExtractError) return json(req, 422, { error: e.message });
    if (e instanceof Anthropic.APIError) {
      console.error("Anthropic API error", e.status, e.message);
      const message =
        e.status === 401 ? "The Anthropic API key in Supabase isn't valid."
        : e.status === 429 ? "Too many requests to the AI right now. Try again in a minute."
        : "The AI service returned an error. Try again.";
      return json(req, 502, { error: message });
    }
    console.error(e);
    return json(req, 500, { error: "Something went wrong reading the document." });
  }
});
