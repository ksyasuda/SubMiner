import * as fs from 'node:fs/promises';
import * as path from 'node:path';

export interface FrequencyDictionaryLookupOptions {
  searchPaths: string[];
  log: (message: string) => void;
}

type FrequencyDictionaryMode = 'occurrence-based' | 'rank-based';

interface FrequencyDictionaryEntry {
  rank: number;
  term: string;
}

const FREQUENCY_BANK_FILE_GLOB = /^term_meta_bank_.*\.json$/;
const NOOP_LOOKUP = (): null => null;
const ENTRY_YIELD_INTERVAL = 5000;

function isErrorCode(error: unknown, code: string): boolean {
  return Boolean(error && typeof error === 'object' && (error as { code?: unknown }).code === code);
}

async function yieldToEventLoop(): Promise<void> {
  await new Promise<void>((resolve) => {
    setImmediate(resolve);
  });
}

function normalizeFrequencyTerm(value: string): string {
  return value.trim().toLowerCase();
}

async function readDictionaryMetadata(
  dictionaryPath: string,
  log: (message: string) => void,
): Promise<{ title: string | null; frequencyMode: FrequencyDictionaryMode | null }> {
  const indexPath = path.join(dictionaryPath, 'index.json');
  let rawText: string;
  try {
    rawText = await fs.readFile(indexPath, 'utf-8');
  } catch (error) {
    if (isErrorCode(error, 'ENOENT')) {
      return { title: null, frequencyMode: null };
    }
    log(`Failed to read frequency dictionary index ${indexPath}: ${String(error)}`);
    return { title: null, frequencyMode: null };
  }

  let rawIndex: unknown;
  try {
    rawIndex = JSON.parse(rawText) as unknown;
  } catch {
    log(`Failed to parse frequency dictionary index as JSON: ${indexPath}`);
    return { title: null, frequencyMode: null };
  }

  if (!rawIndex || typeof rawIndex !== 'object') {
    return { title: null, frequencyMode: null };
  }

  const titleRaw = (rawIndex as { title?: unknown }).title;
  const frequencyModeRaw = (rawIndex as { frequencyMode?: unknown }).frequencyMode;
  return {
    title: typeof titleRaw === 'string' && titleRaw.trim().length > 0 ? titleRaw.trim() : null,
    frequencyMode:
      frequencyModeRaw === 'occurrence-based' || frequencyModeRaw === 'rank-based'
        ? frequencyModeRaw
        : null,
  };
}

function parsePositiveFrequencyString(value: string): number | null {
  const trimmed = value.trim();
  if (!trimmed) {
    return null;
  }

  const numericMatch = trimmed.match(/[+-]?(\d+(\.\d*)?|\.\d+)([eE][+-]?\d+)?/)?.[0];
  if (!numericMatch) {
    return null;
  }

  const parsed = Number.parseFloat(numericMatch);
  if (!Number.isFinite(parsed) || parsed <= 0) {
    return null;
  }

  const normalized = Math.floor(parsed);
  if (!Number.isFinite(normalized) || normalized <= 0) {
    return null;
  }

  return normalized;
}

function parsePositiveFrequencyNumber(value: unknown): number | null {
  if (typeof value === 'number') {
    if (!Number.isFinite(value) || value <= 0) return null;
    return Math.floor(value);
  }

  if (typeof value === 'string') {
    return parsePositiveFrequencyString(value);
  }

  return null;
}

function parseDisplayFrequencyNumber(value: unknown): number | null {
  if (typeof value === 'string') {
    const leadingDigits = value.trim().match(/^\d+/)?.[0];
    if (!leadingDigits) {
      return null;
    }
    const parsed = Number.parseInt(leadingDigits, 10);
    return Number.isFinite(parsed) && parsed > 0 ? parsed : null;
  }

  return parsePositiveFrequencyNumber(value);
}

function extractFrequencyDisplayValue(meta: unknown): number | null {
  if (!meta || typeof meta !== 'object') return null;
  const frequency = (meta as { frequency?: unknown }).frequency;
  if (!frequency || typeof frequency !== 'object') return null;
  const rawValue = (frequency as { value?: unknown }).value;
  const parsedRawValue = parsePositiveFrequencyNumber(rawValue);
  const displayValue = (frequency as { displayValue?: unknown }).displayValue;
  const parsedDisplayValue = parseDisplayFrequencyNumber(displayValue);
  if (parsedDisplayValue !== null) {
    return parsedDisplayValue;
  }

  return parsedRawValue;
}

function asFrequencyDictionaryEntry(entry: unknown): FrequencyDictionaryEntry | null {
  if (!Array.isArray(entry) || entry.length < 3) {
    return null;
  }

  const [term, _id, meta] = entry as [unknown, unknown, unknown];
  if (typeof term !== 'string') {
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

async function addEntriesToMap(
  rawEntries: unknown,
  terms: Map<string, number>,
): Promise<{ duplicateCount: number }> {
  if (!Array.isArray(rawEntries)) {
    return { duplicateCount: 0 };
  }

  let duplicateCount = 0;
  let processedCount = 0;
  for (const rawEntry of rawEntries) {
    processedCount += 1;
    if (processedCount % ENTRY_YIELD_INTERVAL === 0) {
      await yieldToEventLoop();
    }

    const entry = asFrequencyDictionaryEntry(rawEntry);
    if (!entry) {
      continue;
    }
    const currentRank = terms.get(entry.term);
    if (currentRank === undefined || entry.rank < currentRank) {
      terms.set(entry.term, entry.rank);
      continue;
    }

    duplicateCount += 1;
  }

  return { duplicateCount };
}

async function collectDictionaryFromPath(
  dictionaryPath: string,
  log: (message: string) => void,
): Promise<Map<string, number>> {
  const terms = new Map<string, number>();
  const metadata = await readDictionaryMetadata(dictionaryPath, log);
  if (metadata.frequencyMode === 'occurrence-based') {
    log(
      `Skipping occurrence-based frequency dictionary ${
        metadata.title ?? dictionaryPath
      }; SubMiner frequency tags require rank-based values.`,
    );
    return terms;
  }

  let fileNames: string[];
  try {
    fileNames = await fs.readdir(dictionaryPath);
  } catch (error) {
    log(`Failed to read frequency dictionary directory ${dictionaryPath}: ${String(error)}`);
    return terms;
  }

  const bankFiles = fileNames.filter((name) => FREQUENCY_BANK_FILE_GLOB.test(name)).sort();

  if (bankFiles.length === 0) {
    return terms;
  }

  for (const bankFile of bankFiles) {
    const bankPath = path.join(dictionaryPath, bankFile);
    let rawText: string;
    try {
      rawText = await fs.readFile(bankPath, 'utf-8');
    } catch {
      log(`Failed to read frequency dictionary file ${bankPath}`);
      continue;
    }

    let rawEntries: unknown;
    try {
      await yieldToEventLoop();
      rawEntries = JSON.parse(rawText) as unknown;
    } catch {
      log(`Failed to parse frequency dictionary file as JSON: ${bankPath}`);
      continue;
    }

    const beforeSize = terms.size;
    const { duplicateCount } = await addEntriesToMap(rawEntries, terms);
    if (duplicateCount > 0) {
      log(
        `Frequency dictionary ignored ${duplicateCount} duplicate term entr${
          duplicateCount === 1 ? 'y' : 'ies'
        } in ${bankPath} (kept strongest rank per term).`,
      );
    }
    if (terms.size === beforeSize) {
      log(`Frequency dictionary file contained no extractable entries: ${bankPath}`);
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
      isDirectory = (await fs.stat(dictionaryPath)).isDirectory();
    } catch (error) {
      if (isErrorCode(error, 'ENOENT')) {
        continue;
      }
      options.log(
        `Failed to inspect frequency dictionary path ${dictionaryPath}: ${String(error)}`,
      );
      continue;
    }

    if (!isDirectory) {
      continue;
    }

    foundDictionaryPathCount += 1;
    const terms = await collectDictionaryFromPath(dictionaryPath, options.log);
    if (terms.size > 0) {
      options.log(`Frequency dictionary loaded from ${dictionaryPath} (${terms.size} entries)`);
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
    `Frequency dictionary not found. Searched ${attemptedPaths.length} candidate path(s): ${attemptedPaths.join(', ')}`,
  );
  if (foundDictionaryPathCount > 0) {
    options.log(
      'Frequency dictionary directories found, but no usable term_meta_bank_*.json files were loaded.',
    );
  }

  return NOOP_LOOKUP;
}
