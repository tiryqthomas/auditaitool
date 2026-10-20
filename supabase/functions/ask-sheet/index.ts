// Supabase Edge Function: answers questions about a client spreadsheet. The
// Anthropic API key lives in this project's secrets (ANTHROPIC_API_KEY).
import Anthropic from "@anthropic-ai/sdk";
import { AskSheetError, askSheet, type AskSheetRequest } from "./core.ts";

const ALLOWED_ORIGINS = ["https://tiryqthomas.github.io", "https://localhost:3000"];
const MAX_WORKBOOK_TEXT = 500_000;

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

  let body: AskSheetRequest;
  try {
    body = await req.json();
  } catch {
    return json(req, 400, { error: "Request body must be JSON." });
  }
  if (typeof body.workbookText !== "string" || !body.workbookText.trim()) return json(req, 400, { error: "No spreadsheet was sent." });
  if (body.workbookText.length > MAX_WORKBOOK_TEXT) return json(req, 413, { error: "That spreadsheet is too large to ask about in one go." });
  if (typeof body.question !== "string" || !body.question.trim()) return json(req, 400, { error: "Type a question first." });

  try {
    const answers = await askSheet(client, { workbookText: body.workbookText, question: body.question.slice(0, 2000) });
    return json(req, 200, { answers });
  } catch (e) {
    if (e instanceof AskSheetError) return json(req, 422, { error: e.message });
    if (e instanceof Anthropic.APIError) {
      console.error("Anthropic API error", e.status, e.message);
      const message =
        e.status === 401 ? "The Anthropic API key in Supabase isn't valid."
        : e.status === 429 ? "Too many requests to the AI right now. Try again in a minute."
        : "The AI service returned an error. Try again.";
      return json(req, 502, { error: message });
    }
    console.error(e);
    return json(req, 500, { error: "Something went wrong reading the spreadsheet." });
  }
});
