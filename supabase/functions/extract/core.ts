import type Anthropic from "@anthropic-ai/sdk";

export interface ExtractRequest {
  pdfBase64: string;
  fields: string[];
  instructions?: string;
}

export interface ExtractedField {
  field: string;
  value: string | null;
  quote: string | null;
  page: number | null;
}

const nullable = (type: "string" | "integer") => ({ anyOf: [{ type }, { type: "null" }] });

const OUTPUT_SCHEMA = {
  type: "object",
  properties: {
    fields: {
      type: "array",
      items: {
        type: "object",
        properties: {
          field: { type: "string" },
          value: nullable("string"),
          quote: nullable("string"),
          page: nullable("integer"),
        },
        required: ["field", "value", "quote", "page"],
        additionalProperties: false,
      },
    },
  },
  required: ["fields"],
  additionalProperties: false,
};

const SYSTEM = `You extract information from client documents for a financial statement auditor.
Every answer is checked word for word against the document, and the auditor relies on it as audit evidence.

For each requested field:
- value: the value exactly as the document states it. Do not compute, convert or infer.
- quote: a short span copied character for character from the document that contains the value, such as the label and value on one line ("Total due $12,480.00"). Keep the document's spelling, punctuation and number formatting.
- page: the 1-based page number where the quote appears.
If the document does not clearly state a field, return null for value, quote and page rather than guessing.
Return one entry per requested field, using the field names exactly as given.`;

export class ExtractError extends Error {}

/** Asks Claude for the requested fields, each with a verbatim quote and page. */
export async function extractFields(client: Anthropic, req: ExtractRequest): Promise<ExtractedField[]> {
  const fieldList = req.fields.map((f) => `- ${f}`).join("\n");
  const response = await client.beta.messages.create({
    model: "claude-opus-5-5",
    max_tokens: 16000,
    betas: ["server-side-fallback-2026-07-01"],
    fallbacks: "default",
    output_config: { effort: "high", format: { type: "json_schema", schema: OUTPUT_SCHEMA } },
    system: SYSTEM,
    messages: [
      {
        role: "user",
        content: [
          { type: "document", source: { type: "base64", media_type: "application/pdf", data: req.pdfBase64 } },
          {
            type: "text",
            text: `Fields to find:\n${fieldList}${req.instructions ? `\n\nAuditor's instructions: ${req.instructions}` : ""}`,
          },
        ],
      },
    ],
  });

  if (response.stop_reason === "refusal") throw new ExtractError("The AI declined to read this document.");
  if (response.stop_reason === "max_tokens") throw new ExtractError("The AI's answer was cut off. Try fewer fields.");
  const text = response.content.flatMap((b) => (b.type === "text" ? [b.text] : [])).join("");
  const parsed = JSON.parse(text) as { fields: ExtractedField[] };
  return parsed.fields;
}
