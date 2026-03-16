#!/usr/bin/env bun
/**
 * Backfill frequency_rank in imm_words from a Yomitan-format frequency dictionary.
 *
 * Usage:
 *   bun update-frequency.ts <path-to-frequency-dictionary-directory>
 *
 * The directory should contain term_meta_bank_*.json files (Yomitan format)
 * and optionally an index.json with metadata.
 *
 * Example dictionaries: JPDB, BCCWJ, Innocent Corpus (in Yomitan format).
 */

import { readFileSync, readdirSync, existsSync } from 'node:fs';
import { join } from 'node:path';
import Database from 'libsql';

const DB_PATH = join(process.env.HOME ?? '~', '.config/SubMiner/immersion.sqlite');

function parsePositiveNumber(value: unknown): number | null {
  if (typeof value !== 'number' || !Number.isFinite(value) || value <= 0) return null;
  return Math.floor(value);
}

function parseDisplayValue(value: unknown): number | null {
  if (typeof value === 'string') {
    const match = value.trim().match(/^\d+/)?.[0];
    if (!match) return null;
    const n = Number.parseInt(match, 10);
    return Number.isFinite(n) && n > 0 ? n : null;
  }
  return parsePositiveNumber(value);
}

function extractRank(meta: unknown): number | null {
  if (!meta || typeof meta !== 'object') return null;
  const freq = (meta as Record<string, unknown>).frequency;
  if (!freq || typeof freq !== 'object') return null;
  const f = freq as Record<string, unknown>;
  return parseDisplayValue(f.displayValue) ?? parsePositiveNumber(f.value);
}

function loadDictionary(dirPath: string): Map<string, number> {
  const terms = new Map<string, number>();

  const files = readdirSync(dirPath)
    .filter((f) => /^term_meta_bank.*\.json$/.test(f))
    .sort();

  if (files.length === 0) {
    console.error(`No term_meta_bank_*.json files found in ${dirPath}`);
    process.exit(1);
  }

  for (const file of files) {
    const raw = JSON.parse(readFileSync(join(dirPath, file), 'utf-8')) as unknown[];
    for (const entry of raw) {
      if (!Array.isArray(entry) || entry.length < 3) continue;
      const [term, , meta] = entry;
      if (typeof term !== 'string') continue;
      const rank = extractRank(meta);
      if (rank === null) continue;
      const normalized = term.trim().toLowerCase();
      if (!normalized) continue;
      const existing = terms.get(normalized);
      if (existing === undefined || rank < existing) {
        terms.set(normalized, rank);
      }
    }
    console.log(`  Loaded ${file} (${terms.size} terms total)`);
  }

  return terms;
}

function main() {
  const dictPath = process.argv[2];
  if (!dictPath) {
    console.error('Usage: bun update-frequency.ts <path-to-frequency-dictionary-directory>');
    console.error('');
    console.error('The directory should contain Yomitan term_meta_bank_*.json files.');
    console.error('Examples: JPDB, BCCWJ, Innocent Corpus frequency lists.');
    process.exit(1);
  }

  if (!existsSync(dictPath)) {
    console.error(`Directory not found: ${dictPath}`);
    process.exit(1);
  }

  if (!existsSync(DB_PATH)) {
    console.error(`Database not found: ${DB_PATH}`);
    process.exit(1);
  }

  console.log(`Loading frequency dictionary from ${dictPath}...`);
  const dict = loadDictionary(dictPath);
  console.log(`Loaded ${dict.size} terms from frequency dictionary.\n`);

  console.log(`Opening database: ${DB_PATH}`);
  const db = new Database(DB_PATH);
  db.exec('PRAGMA journal_mode = WAL');
  db.exec('PRAGMA foreign_keys = ON');

  const words = db.prepare('SELECT id, headword, word FROM imm_words').all() as Array<{
    id: number;
    headword: string;
    word: string;
  }>;
  console.log(`Found ${words.length} words in imm_words.\n`);

  const updateStmt = db.prepare(
    'UPDATE imm_words SET frequency_rank = ? WHERE id = ? AND (frequency_rank IS NULL OR frequency_rank > ?)',
  );

  let updated = 0;
  let matched = 0;

  for (const w of words) {
    const headwordNorm = w.headword.trim().toLowerCase();
    const wordNorm = w.word.trim().toLowerCase();

    const rank = dict.get(headwordNorm) ?? dict.get(wordNorm) ?? null;
    if (rank === null) continue;

    matched++;
    const result = updateStmt.run(rank, w.id, rank);
    if (result.changes > 0) updated++;
  }

  console.log(`Matched: ${matched}/${words.length} words found in frequency dictionary`);
  console.log(`Updated: ${updated} rows with new or better frequency_rank`);

  db.close();
  console.log('Done.');
}

main();
