import { spawnSync } from 'node:child_process';
import fs from 'node:fs';
import path from 'node:path';

type FileBudgetResult = {
  file: string;
  lines: number;
};

const DEFAULT_LINE_LIMIT = 500;
const TARGET_DIRS = ['src', 'launcher'];
const TARGET_EXTENSIONS = new Set(['.ts']);
const IGNORE_NAMES = new Set(['.DS_Store']);

function parseArgs(argv: string[]): { strict: boolean; limit: number } {
  let strict = false;
  let limit = DEFAULT_LINE_LIMIT;

  for (let i = 0; i < argv.length; i += 1) {
    const arg = argv[i];
    if (arg === '--strict') {
      strict = true;
      continue;
    }
    if (arg === '--limit') {
      const raw = argv[i + 1];
      const parsed = Number.parseInt(raw ?? '', 10);
      if (!Number.isFinite(parsed) || parsed <= 0) {
        throw new Error(`Invalid --limit value: ${raw ?? '<missing>'}`);
      }
      limit = parsed;
      i += 1;
      continue;
    }
  }

  return { strict, limit };
}

function resolveFilesWithRipgrep(): string[] {
  const rg = spawnSync('rg', ['--files', ...TARGET_DIRS], {
    cwd: process.cwd(),
    encoding: 'utf8',
  });
  if (rg.status !== 0) {
    throw new Error(`rg --files failed:\n${rg.stderr || rg.stdout}`);
  }

  return rg.stdout
    .split('\n')
    .map((line) => line.trim())
    .filter((line) => line.length > 0)
    .filter((file) => TARGET_EXTENSIONS.has(path.extname(file)))
    .filter((file) => !IGNORE_NAMES.has(path.basename(file)));
}

function countLines(content: string): number {
  if (content.length === 0) return 0;
  return content.split('\n').length;
}

function collectOverBudgetFiles(files: string[], limit: number): FileBudgetResult[] {
  const results: FileBudgetResult[] = [];
  for (const file of files) {
    const content = fs.readFileSync(file, 'utf8');
    const lines = countLines(content);
    if (lines > limit) {
      results.push({ file, lines });
    }
  }
  return results.sort((a, b) => b.lines - a.lines || a.file.localeCompare(b.file));
}

function printReport(overBudget: FileBudgetResult[], limit: number, strict: boolean): void {
  const mode = strict ? 'strict' : 'warning';
  if (overBudget.length === 0) {
    console.log(`[OK] file budget check (${mode}) — no files over ${limit} LOC`);
    return;
  }

  const heading = strict ? '[FAIL]' : '[WARN]';
  console.log(`${heading} file budget check (${mode}) — ${overBudget.length} files over ${limit} LOC`);
  for (const item of overBudget) {
    console.log(`  - ${item.file}: ${item.lines} LOC`);
  }
  console.log('  Hint: split by runtime/domain boundaries; keep composition roots thin.');
}

function main(): void {
  const { strict, limit } = parseArgs(process.argv.slice(2));
  const files = resolveFilesWithRipgrep();
  const overBudget = collectOverBudgetFiles(files, limit);
  printReport(overBudget, limit, strict);

  if (strict && overBudget.length > 0) {
    process.exitCode = 1;
  }
}

main();
