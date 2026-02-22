import { spawnSync } from 'node:child_process';
import fs from 'node:fs';
import path from 'node:path';

type FileBudgetResult = {
  file: string;
  lines: number;
};

type HotspotBudget = {
  file: string;
  baseline: number;
  limit: number;
};

type HotspotStatus = HotspotBudget & {
  lines: number;
  delta: number;
  overLimit: boolean;
};

const DEFAULT_LINE_LIMIT = 500;
const TARGET_DIRS = ['src', 'launcher'];
const TARGET_EXTENSIONS = new Set(['.ts']);
const IGNORE_NAMES = new Set(['.DS_Store']);
const HOTSPOT_BUDGETS: readonly HotspotBudget[] = [
  { file: 'src/main.ts', baseline: 2978, limit: 3150 },
  { file: 'src/config/service.ts', baseline: 116, limit: 140 },
  { file: 'src/core/services/tokenizer.ts', baseline: 232, limit: 260 },
  { file: 'launcher/main.ts', baseline: 101, limit: 130 },
  { file: 'src/config/resolve.ts', baseline: 33, limit: 55 },
  { file: 'src/main/runtime/composers/contracts.ts', baseline: 13, limit: 30 },
];

function parseArgs(argv: string[]): { strict: boolean; limit: number; enforceGlobal: boolean } {
  let strict = false;
  let limit = DEFAULT_LINE_LIMIT;
  let enforceGlobal = false;

  for (let i = 0; i < argv.length; i += 1) {
    const arg = argv[i];
    if (arg === '--strict') {
      strict = true;
      continue;
    }
    if (arg === '--enforce-global') {
      enforceGlobal = true;
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

  return { strict, limit, enforceGlobal };
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

function collectHotspotStatuses(hotspots: readonly HotspotBudget[]): HotspotStatus[] {
  return hotspots.map((hotspot) => {
    const content = fs.readFileSync(hotspot.file, 'utf8');
    const lines = countLines(content);
    return {
      ...hotspot,
      lines,
      delta: lines - hotspot.baseline,
      overLimit: lines > hotspot.limit,
    };
  });
}

function printGlobalReport(
  overBudget: FileBudgetResult[],
  limit: number,
  strict: boolean,
  enforceGlobal: boolean,
): void {
  const mode = strict ? 'strict' : 'warning';
  if (overBudget.length === 0) {
    console.log(`[OK] file budget check (${mode}) — no files over ${limit} LOC`);
    return;
  }

  const heading = strict && enforceGlobal ? '[FAIL]' : '[WARN]';
  console.log(
    `${heading} file budget check (${mode}) — ${overBudget.length} files over ${limit} LOC`,
  );
  for (const item of overBudget) {
    console.log(`  - ${item.file}: ${item.lines} LOC`);
  }
  console.log('  Hint: split by runtime/domain boundaries; keep composition roots thin.');
}

function printHotspotReport(hotspots: readonly HotspotStatus[], strict: boolean): void {
  const mode = strict ? 'strict' : 'warning';
  const overLimit = hotspots.filter((hotspot) => hotspot.overLimit);

  if (overLimit.length === 0) {
    console.log(`[OK] hotspot budget check (${mode}) — no hotspots over limit`);
  } else {
    const heading = strict ? '[FAIL]' : '[WARN]';
    console.log(
      `${heading} hotspot budget check (${mode}) — ${overLimit.length} hotspots over limit`,
    );
  }

  for (const hotspot of hotspots) {
    const status = hotspot.overLimit ? 'OVER' : 'OK';
    const delta = hotspot.delta >= 0 ? `+${hotspot.delta}` : `${hotspot.delta}`;
    console.log(
      `  - [${status}] ${hotspot.file}: baseline=${hotspot.baseline}, limit=${hotspot.limit}, current=${hotspot.lines}, delta=${delta}`,
    );
  }
}

function main(): void {
  const { strict, limit, enforceGlobal } = parseArgs(process.argv.slice(2));
  const files = resolveFilesWithRipgrep();
  const overBudget = collectOverBudgetFiles(files, limit);
  const hotspotStatuses = collectHotspotStatuses(HOTSPOT_BUDGETS);

  printGlobalReport(overBudget, limit, strict, enforceGlobal);
  printHotspotReport(hotspotStatuses, strict);

  const overGlobalLimit = overBudget.length > 0;
  const hotspotOverLimit = hotspotStatuses.some((hotspot) => hotspot.overLimit);
  const strictFailure = hotspotOverLimit || (strict && enforceGlobal && overGlobalLimit);

  if (strict && strictFailure) {
    process.exitCode = 1;
  }
}

main();
