import * as fs from "fs";
import * as path from "path";

import type { JlptLevel } from "../../types";

export interface JlptVocabLookupOptions {
  searchPaths: string[];
  log: (message: string) => void;
}

const JLPT_BANK_FILES: { level: JlptLevel; filename: string }[] = [
  { level: "N1", filename: "term_meta_bank_1.json" },
  { level: "N2", filename: "term_meta_bank_2.json" },
  { level: "N3", filename: "term_meta_bank_3.json" },
  { level: "N4", filename: "term_meta_bank_4.json" },
  { level: "N5", filename: "term_meta_bank_5.json" },
];
const JLPT_LEVEL_PRECEDENCE: Record<JlptLevel, number> = {
  N1: 5,
  N2: 4,
  N3: 3,
  N4: 2,
  N5: 1,
};

const NOOP_LOOKUP = (): null => null;

function normalizeJlptTerm(value: string): string {
  return value.trim();
}

function hasFrequencyDisplayValue(meta: unknown): boolean {
  if (!meta || typeof meta !== "object") return false;
  const frequency = (meta as { frequency?: unknown }).frequency;
  if (!frequency || typeof frequency !== "object") return false;
  return Object.prototype.hasOwnProperty.call(
    frequency as Record<string, unknown>,
    "displayValue",
  );
}

function addEntriesToMap(
  rawEntries: unknown,
  level: JlptLevel,
  terms: Map<string, JlptLevel>,
  log: (message: string) => void,
): void {
  const shouldUpdateLevel = (
    existingLevel: JlptLevel | undefined,
    incomingLevel: JlptLevel,
  ): boolean =>
    existingLevel === undefined ||
    JLPT_LEVEL_PRECEDENCE[incomingLevel] >
      JLPT_LEVEL_PRECEDENCE[existingLevel];

  if (!Array.isArray(rawEntries)) {
    return;
  }

  for (const rawEntry of rawEntries) {
    if (!Array.isArray(rawEntry)) {
      continue;
    }

    const [term, _entryId, meta] = rawEntry as [unknown, unknown, unknown];
    if (typeof term !== "string") {
      continue;
    }

    const normalizedTerm = normalizeJlptTerm(term);
    if (!normalizedTerm) {
      continue;
    }

    if (!hasFrequencyDisplayValue(meta)) {
      continue;
    }

    const existingLevel = terms.get(normalizedTerm);
    if (shouldUpdateLevel(existingLevel, level)) {
      terms.set(normalizedTerm, level);
      continue;
    }

    log(
      `JLPT dictionary already has ${normalizedTerm} as ${existingLevel}; keeping that level instead of ${level}`,
    );
  }
}

function collectDictionaryFromPath(
  dictionaryPath: string,
  log: (message: string) => void,
): Map<string, JlptLevel> {
  const terms = new Map<string, JlptLevel>();

  for (const bank of JLPT_BANK_FILES) {
    const bankPath = path.join(dictionaryPath, bank.filename);
    if (!fs.existsSync(bankPath)) {
      continue;
    }

    let rawText: string;
    try {
      rawText = fs.readFileSync(bankPath, "utf-8");
    } catch {
      continue;
    }

    let rawEntries: unknown;
    try {
      rawEntries = JSON.parse(rawText) as unknown;
    } catch {
      continue;
    }

    addEntriesToMap(rawEntries, bank.level, terms, log);
  }

  return terms;
}

export async function createJlptVocabularyLookupService(
  options: JlptVocabLookupOptions,
): Promise<(term: string) => JlptLevel | null> {
  const attemptedPaths: string[] = [];
  let foundDirectoryCount = 0;
  let foundBankCount = 0;
  for (const dictionaryPath of options.searchPaths) {
    attemptedPaths.push(dictionaryPath);
    if (!fs.existsSync(dictionaryPath)) {
      continue;
    }

    if (!fs.statSync(dictionaryPath).isDirectory()) {
      continue;
    }

    foundDirectoryCount += 1;

    const terms = collectDictionaryFromPath(dictionaryPath, options.log);
    if (terms.size > 0) {
      foundBankCount += 1;
      options.log(
        `JLPT dictionary loaded from ${dictionaryPath} (${terms.size} entries)`,
      );
      return (term: string): JlptLevel | null => {
        if (!term) return null;
        const normalized = normalizeJlptTerm(term);
        return normalized ? terms.get(normalized) ?? null : null;
      };
    }

    options.log(
      `JLPT dictionary directory exists but contains no readable term_meta_bank_*.json files: ${dictionaryPath}`,
    );
  }

  options.log(
    `JLPT dictionary not found. Searched ${attemptedPaths.length} candidate path(s): ${attemptedPaths.join(", ")}`,
  );
  if (foundDirectoryCount > 0 && foundBankCount === 0) {
    options.log(
      "JLPT dictionary directories found, but none contained valid term_meta_bank_*.json files.",
    );
  }
  return NOOP_LOOKUP;
}
