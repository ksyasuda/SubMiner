import * as fs from 'node:fs/promises';
import * as path from 'path';

import type { JlptLevel } from '../../types';

export interface JlptVocabLookupOptions {
  searchPaths: string[];
  log: (message: string) => void;
}

const JLPT_BANK_FILES: { level: JlptLevel; filename: string }[] = [
  { level: 'N1', filename: 'term_meta_bank_1.json' },
  { level: 'N2', filename: 'term_meta_bank_2.json' },
  { level: 'N3', filename: 'term_meta_bank_3.json' },
  { level: 'N4', filename: 'term_meta_bank_4.json' },
  { level: 'N5', filename: 'term_meta_bank_5.json' },
];
const JLPT_LEVEL_PRECEDENCE: Record<JlptLevel, number> = {
  N1: 5,
  N2: 4,
  N3: 3,
  N4: 2,
  N5: 1,
};
const JLPT_DUPLICATE_LOG_EXAMPLE_LIMIT = 5;

const NOOP_LOOKUP = (): null => null;
const ENTRY_YIELD_INTERVAL = 5000;

interface JlptDuplicateStats {
  duplicateEntryCount: number;
  duplicateTerms: Set<string>;
  exampleTerms: string[];
}

function isErrorCode(error: unknown, code: string): boolean {
  return Boolean(error && typeof error === 'object' && (error as { code?: unknown }).code === code);
}

async function yieldToEventLoop(): Promise<void> {
  await new Promise<void>((resolve) => {
    setImmediate(resolve);
  });
}

function normalizeJlptTerm(value: string): string {
  return value.trim();
}

function hasFrequencyDisplayValue(meta: unknown): boolean {
  if (!meta || typeof meta !== 'object') return false;
  const frequency = (meta as { frequency?: unknown }).frequency;
  if (!frequency || typeof frequency !== 'object') return false;
  return Object.prototype.hasOwnProperty.call(frequency as Record<string, unknown>, 'displayValue');
}

function createJlptDuplicateStats(): JlptDuplicateStats {
  return {
    duplicateEntryCount: 0,
    duplicateTerms: new Set<string>(),
    exampleTerms: [],
  };
}

function recordJlptDuplicate(stats: JlptDuplicateStats, term: string): void {
  stats.duplicateEntryCount += 1;
  stats.duplicateTerms.add(term);
  if (
    stats.exampleTerms.length < JLPT_DUPLICATE_LOG_EXAMPLE_LIMIT &&
    !stats.exampleTerms.includes(term)
  ) {
    stats.exampleTerms.push(term);
  }
}

async function addEntriesToMap(
  rawEntries: unknown,
  level: JlptLevel,
  terms: Map<string, JlptLevel>,
  duplicateStats: JlptDuplicateStats,
): Promise<void> {
  const shouldUpdateLevel = (
    existingLevel: JlptLevel | undefined,
    incomingLevel: JlptLevel,
  ): boolean =>
    existingLevel === undefined ||
    JLPT_LEVEL_PRECEDENCE[incomingLevel] > JLPT_LEVEL_PRECEDENCE[existingLevel];

  if (!Array.isArray(rawEntries)) {
    return;
  }

  let processedCount = 0;
  for (const rawEntry of rawEntries) {
    processedCount += 1;
    if (processedCount % ENTRY_YIELD_INTERVAL === 0) {
      await yieldToEventLoop();
    }

    if (!Array.isArray(rawEntry)) {
      continue;
    }

    const [term, _entryId, meta] = rawEntry as [unknown, unknown, unknown];
    if (typeof term !== 'string') {
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
    if (existingLevel !== undefined) {
      recordJlptDuplicate(duplicateStats, normalizedTerm);
    }

    if (shouldUpdateLevel(existingLevel, level)) {
      terms.set(normalizedTerm, level);
    }
  }
}

async function collectDictionaryFromPath(
  dictionaryPath: string,
  log: (message: string) => void,
): Promise<Map<string, JlptLevel>> {
  const terms = new Map<string, JlptLevel>();
  const duplicateStats = createJlptDuplicateStats();

  for (const bank of JLPT_BANK_FILES) {
    const bankPath = path.join(dictionaryPath, bank.filename);
    try {
      if (!(await fs.stat(bankPath)).isFile()) {
        log(`JLPT bank file missing for ${bank.level}: ${bankPath}`);
        continue;
      }
    } catch (error) {
      if (isErrorCode(error, 'ENOENT')) {
        log(`JLPT bank file missing for ${bank.level}: ${bankPath}`);
        continue;
      }
      log(`Failed to inspect JLPT bank file ${bankPath}: ${String(error)}`);
      continue;
    }

    let rawText: string;
    try {
      rawText = await fs.readFile(bankPath, 'utf-8');
    } catch {
      log(`Failed to read JLPT bank file ${bankPath}`);
      continue;
    }

    let rawEntries: unknown;
    try {
      await yieldToEventLoop();
      rawEntries = JSON.parse(rawText) as unknown;
    } catch {
      log(`Failed to parse JLPT bank file as JSON: ${bankPath}`);
      continue;
    }

    if (!Array.isArray(rawEntries)) {
      log(`JLPT bank file has unsupported format (expected JSON array): ${bankPath}`);
      continue;
    }

    const beforeSize = terms.size;
    await addEntriesToMap(rawEntries, bank.level, terms, duplicateStats);
    if (terms.size === beforeSize) {
      log(`JLPT bank file contained no extractable entries: ${bankPath}`);
    }
  }

  if (duplicateStats.duplicateEntryCount > 0) {
    const examples =
      duplicateStats.exampleTerms.length > 0
        ? `; examples: ${duplicateStats.exampleTerms.join(', ')}`
        : '';
    log(
      `JLPT dictionary collapsed ${duplicateStats.duplicateEntryCount} duplicate JLPT entries across ${duplicateStats.duplicateTerms.size} terms; keeping highest-precedence level per surface form${examples}`,
    );
  }

  return terms;
}

export async function createJlptVocabularyLookup(
  options: JlptVocabLookupOptions,
): Promise<(term: string) => JlptLevel | null> {
  const attemptedPaths: string[] = [];
  let foundDictionaryPathCount = 0;
  let foundBankCount = 0;
  const resolvedBanks: string[] = [];
  for (const dictionaryPath of options.searchPaths) {
    attemptedPaths.push(dictionaryPath);
    let isDirectory = false;
    try {
      isDirectory = (await fs.stat(dictionaryPath)).isDirectory();
    } catch (error) {
      if (isErrorCode(error, 'ENOENT')) {
        continue;
      }
      options.log(`Failed to inspect JLPT dictionary path ${dictionaryPath}: ${String(error)}`);
      continue;
    }
    if (!isDirectory) continue;

    foundDictionaryPathCount += 1;

    const terms = await collectDictionaryFromPath(dictionaryPath, options.log);
    if (terms.size > 0) {
      resolvedBanks.push(dictionaryPath);
      foundBankCount += 1;
      options.log(`JLPT dictionary loaded from ${dictionaryPath} (${terms.size} entries)`);
      return (term: string): JlptLevel | null => {
        if (!term) return null;
        const normalized = normalizeJlptTerm(term);
        return normalized ? (terms.get(normalized) ?? null) : null;
      };
    }

    options.log(
      `JLPT dictionary directory exists but contains no readable term_meta_bank_*.json files: ${dictionaryPath}`,
    );
  }

  options.log(
    `JLPT dictionary not found. Searched ${attemptedPaths.length} candidate path(s): ${attemptedPaths.join(', ')}`,
  );
  if (foundDictionaryPathCount > 0 && foundBankCount === 0) {
    options.log(
      'JLPT dictionary directories found, but none contained valid term_meta_bank_*.json files.',
    );
  }
  if (resolvedBanks.length > 0 && foundBankCount > 0) {
    options.log(`JLPT dictionary search matched path(s): ${resolvedBanks.join(', ')}`);
  }
  return NOOP_LOOKUP;
}
