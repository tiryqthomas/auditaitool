import type Anthropic from "@anthropic-ai/sdk";

export interface AskSheetRequest {
  workbookText: string;
  question: string;
}

export interface SheetAnswer {
  answer: string;
  sheet: string;
  cell: string;
  reason: string;
}

const OUTPUT_SCHEMA = {
  type: "object",
  properties: {
    answers: {
      type: "array",
      items: {
        type: "object",
        properties: {
          answer: { type: "string" },
          sheet: { type: "string" },
          cell: { type: "string" },
          reason: { type: "string" },
        },
        required: ["answer", "sheet", "cell", "reason"],
        additionalProperties: false,
      },
    },
  },
  required: ["answers"],
  additionalProperties: false,
};

const SYSTEM = `You answer questions about a client's spreadsheet for a financial statement auditor.
The spreadsheet is given as text: each sheet starts with "=== Sheet: <name> ===", then one line per row listing cells as ADDRESS="displayed value".
Every answer is checked against the named cell, and the auditor relies on it as audit evidence.

For each answer:
- answer: the cell's displayed value exactly as given. Do not compute, sum, convert or reformat it.
- sheet: the sheet name exactly as given.
- cell: the cell address holding that value, such as D14.
- reason: one short sentence saying why this cell answers the question, naming its row and column labels.
When the question asks for several items (for example every invoice over a threshold), return one answer per matching cell.
If the spreadsheet does not contain the answer in a single cell, return an empty list rather than guessing or calculating.`;

export class AskSheetError extends Error {}

/** Asks Claude which cells answer the question. */
export async function askSheet(client: Anthropic, req: AskSheetRequest): Promise<SheetAnswer[]> {
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
          { type: "text", text: `<spreadsheet>\n${req.workbookText}\n</spreadsheet>` },
          { type: "text", text: `Question: ${req.question}` },
        ],
      },
    ],
  });

  if (response.stop_reason === "refusal") throw new AskSheetError("The AI declined to answer this question.");
  if (response.stop_reason === "max_tokens") throw new AskSheetError("The AI's answer was cut off. Try a narrower question.");
  const text = response.content.flatMap((b) => (b.type === "text" ? [b.text] : [])).join("");
  return (JSON.parse(text) as { answers: SheetAnswer[] }).answers;
}
