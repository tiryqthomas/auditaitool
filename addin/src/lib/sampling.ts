import { asAmount, findHeaderRow, type SourceCell, type SourceSheet } from "./spreadsheet";

// Sample selection is done in code, never by the AI: the auditor has to show
// the selection was unbiased and can be reproduced, so every run with the same
// population, settings and seed picks the same items.

export interface PopulationItem {
  row: number; // 0-based row in the source sheet
  key: string;
  keyAddress: string; // e.g. A9
  amount: number | null;
  cells: (SourceCell | null)[]; // aligned with Population.headers
}

export interface Population {
  sheet: string;
  headers: string[];
  keyHeader: string;
  amountHeader: string;
  items: PopulationItem[];
  skipped: { row: number; reason: string }[];
}

const TOTAL_ROW = /^(grand\s+|sub\s*)?total\b/i;

/** The column headings of a listing, from its heading row. */
export function listingHeaders(sheet: SourceSheet): string[] {
  const h = findHeaderRow(sheet);
  return h < 0 ? [] : sheet.rows[h].map((c) => c?.text.trim() ?? "");
}

export function guessAmountHeader(headers: string[]): string {
  return headers.find((h) => /amount|total|balance|value|\$/i.test(h)) ?? "";
}

/**
 * Reads the items of a listing below its heading row. Blank rows, total rows
 * and rows with no key are skipped and reported, never silently dropped.
 */
export function readPopulation(sheet: SourceSheet, keyHeader: string, amountHeader: string): Population {
  const headerRow = findHeaderRow(sheet);
  if (headerRow < 0) throw new Error(`Couldn't find a row of column headings on "${sheet.name}".`);
  const headers = listingHeaders(sheet);
  const keyCol = headers.indexOf(keyHeader);
  const amountCol = headers.indexOf(amountHeader);
  if (keyCol < 0) throw new Error(`"${keyHeader}" isn't a column heading on "${sheet.name}".`);

  const items: PopulationItem[] = [];
  const skipped: Population["skipped"] = [];
  for (let r = headerRow + 1; r < sheet.rows.length; r++) {
    const row = sheet.rows[r] ?? [];
    if (row.every((c) => !c)) continue;
    if (row.some((c) => c && TOTAL_ROW.test(c.text.trim()))) {
      skipped.push({ row: r, reason: "total row" });
      continue;
    }
    const key = row[keyCol];
    if (!key) {
      skipped.push({ row: r, reason: `no ${keyHeader}` });
      continue;
    }
    const amountCell = amountCol >= 0 ? row[amountCol] : null;
    const amount = !amountCell ? null : typeof amountCell.value === "number" ? amountCell.value : asAmount(amountCell.text);
    items.push({ row: r, key: key.text, keyAddress: key.address, amount, cells: headers.map((_, c) => row[c] ?? null) });
  }
  return { sheet: sheet.name, headers, keyHeader, amountHeader, items, skipped };
}

/** Small seeded generator (mulberry32): the same seed always gives the same sequence. */
export function seededRandom(seed: number): () => number {
  let a = seed >>> 0;
  return () => {
    a = (a + 0x6d2b79f5) >>> 0;
    let t = a;
    t = Math.imul(t ^ (t >>> 15), t | 1);
    t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
    return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
  };
}

export type SampleMethod = "random" | "mus";

export interface SampleSettings {
  method: SampleMethod;
  size: number; // items to select in addition to key items
  seed: number;
  keyItemThreshold?: number; // every item at or above this amount is tested
}

export interface Selection {
  items: (PopulationItem & { reason: "key item" | "random" | "monetary unit" })[];
  keyItems: number;
  interval?: number; // MUS sampling interval
  randomStart?: number; // MUS random start
  notSelectable: number; // items MUS can't pick (zero, negative or no amount)
  populationTotal: number;
}

const round2 = (n: number) => Math.round(n * 100) / 100;

/** Selects the sample: key items first, then random or monetary-unit selection from the rest. */
export function selectSample(pop: Population, s: SampleSettings): Selection {
  if (!Number.isInteger(s.size) || s.size < 0) throw new Error("Sample size must be a whole number.");
  const populationTotal = round2(pop.items.reduce((t, i) => t + (i.amount ?? 0), 0));
  const threshold = s.keyItemThreshold;
  const isKey = (i: PopulationItem) => threshold !== undefined && i.amount !== null && i.amount >= threshold;
  const keys = pop.items.filter(isKey);
  const rest = pop.items.filter((i) => !isKey(i));
  const rand = seededRandom(s.seed);
  const picked: Selection["items"] = keys.map((i) => ({ ...i, reason: "key item" }));
  let interval: number | undefined;
  let randomStart: number | undefined;
  let notSelectable = 0;

  if (s.method === "random") {
    const pool = [...rest];
    const n = Math.min(s.size, pool.length);
    // Partial Fisher-Yates shuffle: each item is equally likely to be chosen.
    for (let k = 0; k < n; k++) {
      const j = k + Math.floor(rand() * (pool.length - k));
      [pool[k], pool[j]] = [pool[j], pool[k]];
      picked.push({ ...pool[k], reason: "random" });
    }
  } else {
    // Monetary unit sampling: systematic selection of dollars with a random start,
    // so larger items are more likely to be picked.
    const selectable = rest.filter((i) => i.amount !== null && i.amount > 0);
    notSelectable = rest.length - selectable.length;
    const total = selectable.reduce((t, i) => t + i.amount!, 0);
    if (s.size > 0 && total > 0) {
      interval = total / s.size;
      randomStart = rand() * interval;
      let next = randomStart;
      let cumulative = 0;
      for (const item of selectable) {
        cumulative += item.amount!;
        if (cumulative > next) {
          picked.push({ ...item, reason: "monetary unit" });
          // An item bigger than the interval can hold several hits; it is tested once.
          while (next < cumulative) next += interval;
        }
      }
      interval = round2(interval);
      randomStart = round2(randomStart);
    }
  }

  picked.sort((a, b) => a.row - b.row);
  return { items: picked, keyItems: keys.length, interval, randomStart, notSelectable, populationTotal };
}

export function newSeed(): number {
  return 1 + Math.floor(Math.random() * 999_999);
}
