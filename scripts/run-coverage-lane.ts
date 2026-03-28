import { existsSync, mkdirSync, readFileSync, readdirSync, rmSync, writeFileSync } from 'node:fs';
import { spawnSync } from 'node:child_process';
import { fileURLToPath } from 'node:url';
import { join, relative, resolve } from 'node:path';

type LaneConfig = {
  roots: string[];
  include: string[];
  exclude: Set<string>;
};

type LcovRecord = {
  sourceFile: string;
  functions: Map<string, number>;
  functionHits: Map<string, number>;
  lines: Map<number, number>;
  branches: Map<string, { line: number; block: string; branch: string; hits: number | null }>;
};

const repoRoot = resolve(fileURLToPath(new URL('..', import.meta.url)));

const lanes: Record<string, LaneConfig> = {
  'bun-src-full': {
    roots: ['src'],
    include: ['.test.ts', '.type-test.ts'],
    exclude: new Set([
      'src/core/services/anki-jimaku-ipc.test.ts',
      'src/core/services/ipc.test.ts',
      'src/core/services/overlay-manager.test.ts',
      'src/main/config-validation.test.ts',
      'src/main/runtime/registry.test.ts',
      'src/main/runtime/startup-config.test.ts',
    ]),
  },
  'bun-launcher-unit': {
    roots: ['launcher'],
    include: ['.test.ts'],
    exclude: new Set(['launcher/smoke.e2e.test.ts']),
  },
};

function collectFiles(rootDir: string, includeSuffixes: string[], excludeSet: Set<string>): string[] {
  const out: string[] = [];
  const visit = (currentDir: string) => {
    for (const entry of readdirSync(currentDir, { withFileTypes: true })) {
      const fullPath = resolve(currentDir, entry.name);
      if (entry.isDirectory()) {
        visit(fullPath);
        continue;
      }
      const relPath = relative(repoRoot, fullPath).replaceAll('\\', '/');
      if (excludeSet.has(relPath)) continue;
      if (includeSuffixes.some((suffix) => relPath.endsWith(suffix))) {
        out.push(relPath);
      }
    }
  };

  visit(resolve(repoRoot, rootDir));
  out.sort();
  return out;
}

function getLaneFiles(laneName: string): string[] {
  const lane = lanes[laneName];
  if (!lane) {
    throw new Error(`Unknown coverage lane: ${laneName}`);
  }
  const files = lane.roots.flatMap((rootDir) => collectFiles(rootDir, lane.include, lane.exclude));
  if (files.length === 0) {
    throw new Error(`No test files found for coverage lane: ${laneName}`);
  }
  return files;
}

function parseCoverageDirArg(argv: string[]): string {
  for (let index = 0; index < argv.length; index += 1) {
    if (argv[index] === '--coverage-dir') {
      const next = argv[index + 1];
      if (!next) {
        throw new Error('Missing value for --coverage-dir');
      }
      return next;
    }
  }
  return 'coverage';
}

function parseLcovReport(report: string): LcovRecord[] {
  const records: LcovRecord[] = [];
  let current: LcovRecord | null = null;

  const ensureCurrent = (): LcovRecord => {
    if (!current) {
      throw new Error('Malformed lcov report: record data before SF');
    }
    return current;
  };

  for (const rawLine of report.split(/\r?\n/)) {
    const line = rawLine.trim();
    if (!line) continue;
    if (line.startsWith('TN:')) {
      continue;
    }
    if (line.startsWith('SF:')) {
      current = {
        sourceFile: line.slice(3),
        functions: new Map(),
        functionHits: new Map(),
        lines: new Map(),
        branches: new Map(),
      };
      continue;
    }
    if (line === 'end_of_record') {
      if (current) {
        records.push(current);
        current = null;
      }
      continue;
    }
    if (line.startsWith('FN:')) {
      const [lineNumber, ...nameParts] = line.slice(3).split(',');
      ensureCurrent().functions.set(nameParts.join(','), Number(lineNumber));
      continue;
    }
    if (line.startsWith('FNDA:')) {
      const [hits, ...nameParts] = line.slice(5).split(',');
      ensureCurrent().functionHits.set(nameParts.join(','), Number(hits));
      continue;
    }
    if (line.startsWith('DA:')) {
      const [lineNumber, hits] = line.slice(3).split(',');
      ensureCurrent().lines.set(Number(lineNumber), Number(hits));
      continue;
    }
    if (line.startsWith('BRDA:')) {
      const [lineNumber, block, branch, hits] = line.slice(5).split(',');
      ensureCurrent().branches.set(`${lineNumber}:${block}:${branch}`, {
        line: Number(lineNumber),
        block,
        branch,
        hits: hits === '-' ? null : Number(hits),
      });
    }
  }

  if (current) {
    records.push(current);
  }

  return records;
}

export function mergeLcovReports(reports: string[]): string {
  const merged = new Map<string, LcovRecord>();

  for (const report of reports) {
    for (const record of parseLcovReport(report)) {
      let target = merged.get(record.sourceFile);
      if (!target) {
        target = {
          sourceFile: record.sourceFile,
          functions: new Map(),
          functionHits: new Map(),
          lines: new Map(),
          branches: new Map(),
        };
        merged.set(record.sourceFile, target);
      }

      for (const [name, line] of record.functions) {
        if (!target.functions.has(name)) {
          target.functions.set(name, line);
        }
      }

      for (const [name, hits] of record.functionHits) {
        target.functionHits.set(name, (target.functionHits.get(name) ?? 0) + hits);
      }

      for (const [lineNumber, hits] of record.lines) {
        target.lines.set(lineNumber, (target.lines.get(lineNumber) ?? 0) + hits);
      }

      for (const [branchKey, branchRecord] of record.branches) {
        const existing = target.branches.get(branchKey);
        if (!existing) {
          target.branches.set(branchKey, { ...branchRecord });
          continue;
        }
        if (branchRecord.hits === null) {
          continue;
        }
        existing.hits = (existing.hits ?? 0) + branchRecord.hits;
      }
    }
  }

  const chunks: string[] = [];
  for (const sourceFile of [...merged.keys()].sort()) {
    const record = merged.get(sourceFile)!;
    chunks.push(`SF:${record.sourceFile}`);

    const functions = [...record.functions.entries()].sort((a, b) =>
      a[1] === b[1] ? a[0].localeCompare(b[0]) : a[1] - b[1],
    );
    for (const [name, line] of functions) {
      chunks.push(`FN:${line},${name}`);
    }
    for (const [name] of functions) {
      chunks.push(`FNDA:${record.functionHits.get(name) ?? 0},${name}`);
    }
    chunks.push(`FNF:${functions.length}`);
    chunks.push(`FNH:${functions.filter(([name]) => (record.functionHits.get(name) ?? 0) > 0).length}`);

    const branches = [...record.branches.values()].sort((a, b) =>
      a.line === b.line
        ? a.block === b.block
          ? a.branch.localeCompare(b.branch)
          : a.block.localeCompare(b.block)
        : a.line - b.line,
    );
    for (const branch of branches) {
      chunks.push(
        `BRDA:${branch.line},${branch.block},${branch.branch},${branch.hits === null ? '-' : branch.hits}`,
      );
    }
    chunks.push(`BRF:${branches.length}`);
    chunks.push(`BRH:${branches.filter((branch) => (branch.hits ?? 0) > 0).length}`);

    const lines = [...record.lines.entries()].sort((a, b) => a[0] - b[0]);
    for (const [lineNumber, hits] of lines) {
      chunks.push(`DA:${lineNumber},${hits}`);
    }
    chunks.push(`LF:${lines.length}`);
    chunks.push(`LH:${lines.filter(([, hits]) => hits > 0).length}`);
    chunks.push('end_of_record');
  }

  return chunks.length > 0 ? `${chunks.join('\n')}\n` : '';
}

function runCoverageLane(): number {
  const laneName = process.argv[2];
  if (!laneName) {
    process.stderr.write('Missing coverage lane name\n');
    return 1;
  }

  const coverageDir = resolve(repoRoot, parseCoverageDirArg(process.argv.slice(3)));
  const shardRoot = join(coverageDir, '.shards');
  mkdirSync(coverageDir, { recursive: true });
  rmSync(shardRoot, { recursive: true, force: true });
  mkdirSync(shardRoot, { recursive: true });

  const files = getLaneFiles(laneName);
  const reports: string[] = [];

  for (const [index, file] of files.entries()) {
    const shardDir = join(shardRoot, `${String(index + 1).padStart(3, '0')}`);
    const result = spawnSync(
      'bun',
      ['test', '--coverage', '--coverage-reporter=lcov', '--coverage-dir', shardDir, `./${file}`],
      {
        cwd: repoRoot,
        stdio: 'inherit',
      },
    );

    if (result.error) {
      throw result.error;
    }
    if ((result.status ?? 1) !== 0) {
      return result.status ?? 1;
    }

    const lcovPath = join(shardDir, 'lcov.info');
    if (!existsSync(lcovPath)) {
      process.stdout.write(`Skipping empty coverage shard for ${file}\n`);
      continue;
    }

    reports.push(readFileSync(lcovPath, 'utf8'));
  }

  writeFileSync(join(coverageDir, 'lcov.info'), mergeLcovReports(reports), 'utf8');
  rmSync(shardRoot, { recursive: true, force: true });
  process.stdout.write(`Merged LCOV written to ${relative(repoRoot, join(coverageDir, 'lcov.info'))}\n`);
  return 0;
}

if (import.meta.main) {
  process.exit(runCoverageLane());
}
