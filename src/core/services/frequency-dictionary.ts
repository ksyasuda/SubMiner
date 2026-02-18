import * as fs from "node:fs";
import * as path from "node:path";

export interface FrequencyDictionaryLookupOptions {
  searchPaths: string[];
  log: (message: string) => void;
}

interface FrequencyDictionaryEntry {
  rank: number;
  term: string;
}

const FREQUENCY_BANK_FILE_GLOB = /^term_meta_bank_.*\.json$/;
const NOOP_LOOKUP = (): null => null;

function normalizeFrequencyTerm(value: string): string {
  return value.trim().toLowerCase();
}

function extractFrequencyDisplayValue(meta: unknown): number | null {
  if (!meta || typeof meta !== "object") return null;
  const frequency = (meta as { frequency?: unknown }).frequency;
  if (!frequency || typeof frequency !== "object") return null;
  const displayValue = (frequency as { displayValue?: unknown }).displayValue;
  if (typeof displayValue === "number") {
    if (!Number.isFinite(displayValue) || displayValue <= 0) return null;
    return Math.floor(displayValue);
  }
  if (typeof displayValue === "string") {
    const normalized = displayValue.trim().replace(/,/g, "");
    const parsed = Number.parseInt(normalized, 10);
    if (!Number.isFinite(parsed) || parsed <= 0) return null;
    return parsed;
  }

  return null;
}

function asFrequencyDictionaryEntry(
  entry: unknown,
): FrequencyDictionaryEntry | null {
  if (!Array.isArray(entry) || entry.length < 3) {
    return null;
  }

  const [term, _id, meta] = entry as [unknown, unknown, unknown];
  if (typeof term !== "string") {
    return null;
  }

  const frequency = extractFrequencyDisplayValue(meta);
  if (frequency === null) return null;

  const normalizedTerm = normalizeFrequencyTerm(term);
  if (!normalizedTerm) return null;

  return {
    term: normalizedTerm,
    rank: frequency,
  };
}

function addEntriesToMap(
  rawEntries: unknown,
  terms: Map<string, number>,
  log: (message: string) => void,
): void {
  if (!Array.isArray(rawEntries)) {
    return;
  }

  for (const rawEntry of rawEntries) {
    const entry = asFrequencyDictionaryEntry(rawEntry);
    if (!entry) {
      continue;
    }
    const currentRank = terms.get(entry.term);
    if (currentRank === undefined || entry.rank < currentRank) {
      terms.set(entry.term, entry.rank);
      continue;
    }

    log(
      `Frequency dictionary duplicate term ${entry.term} with weaker rank ${entry.rank}; keeping ${currentRank}.`,
    );
  }
}

function collectDictionaryFromPath(
  dictionaryPath: string,
  log: (message: string) => void,
): Map<string, number> {
  const terms = new Map<string, number>();

  let fileNames: string[];
  try {
    fileNames = fs.readdirSync(dictionaryPath);
  } catch (error) {
    log(
      `Failed to read frequency dictionary directory ${dictionaryPath}: ${String(error)}`,
    );
    return terms;
  }

  const bankFiles = fileNames
    .filter((name) => FREQUENCY_BANK_FILE_GLOB.test(name))
    .sort();

  if (bankFiles.length === 0) {
    return terms;
  }

  for (const bankFile of bankFiles) {
    const bankPath = path.join(dictionaryPath, bankFile);
    let rawText: string;
    try {
      rawText = fs.readFileSync(bankPath, "utf-8");
    } catch {
      log(`Failed to read frequency dictionary file ${bankPath}`);
      continue;
    }

    let rawEntries: unknown;
    try {
      rawEntries = JSON.parse(rawText) as unknown;
    } catch {
      log(`Failed to parse frequency dictionary file as JSON: ${bankPath}`);
      continue;
    }

    const beforeSize = terms.size;
    addEntriesToMap(rawEntries, terms, log);
    if (terms.size === beforeSize) {
      log(
        `Frequency dictionary file contained no extractable entries: ${bankPath}`,
      );
    }
  }

  return terms;
}

export async function createFrequencyDictionaryLookup(
  options: FrequencyDictionaryLookupOptions,
): Promise<(term: string) => number | null> {
  const attemptedPaths: string[] = [];
  let foundDictionaryPathCount = 0;

  for (const dictionaryPath of options.searchPaths) {
    attemptedPaths.push(dictionaryPath);
    let isDirectory = false;

    try {
      if (!fs.existsSync(dictionaryPath)) {
        continue;
      }
      isDirectory = fs.statSync(dictionaryPath).isDirectory();
    } catch (error) {
      options.log(
        `Failed to inspect frequency dictionary path ${dictionaryPath}: ${String(error)}`,
      );
      continue;
    }

    if (!isDirectory) {
      continue;
    }

    foundDictionaryPathCount += 1;
    const terms = collectDictionaryFromPath(dictionaryPath, options.log);
    if (terms.size > 0) {
      options.log(
        `Frequency dictionary loaded from ${dictionaryPath} (${terms.size} entries)`,
      );
      return (term: string): number | null => {
        const normalized = normalizeFrequencyTerm(term);
        if (!normalized) return null;
        return terms.get(normalized) ?? null;
      };
    }

    options.log(
      `Frequency dictionary directory exists but contains no readable term_meta_bank_*.json files: ${dictionaryPath}`,
    );
  }

  options.log(
    `Frequency dictionary not found. Searched ${attemptedPaths.length} candidate path(s): ${attemptedPaths.join(", ")}`,
  );
  if (foundDictionaryPathCount > 0) {
    options.log(
      "Frequency dictionary directories found, but no usable term_meta_bank_*.json files were loaded.",
    );
  }

  return NOOP_LOOKUP;
}
