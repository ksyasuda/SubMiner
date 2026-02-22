import fs from 'node:fs';
import path from 'node:path';

export type ParsedArgs = {
  strict: boolean;
  root: string;
};

export type RuntimeCycleCheckResult = {
  root: string;
  mode: 'warning' | 'strict';
  fileCount: number;
  hasCycle: boolean;
  cyclePath: string[];
};

const DEFAULT_ROOT = path.join('src', 'main', 'runtime');

export function parseArgs(argv: string[]): ParsedArgs {
  let strict = false;
  let root = DEFAULT_ROOT;

  for (let i = 0; i < argv.length; i += 1) {
    const arg = argv[i];
    if (arg === '--strict') {
      strict = true;
      continue;
    }
    if (arg === '--root') {
      const raw = argv[i + 1];
      if (!raw || raw.startsWith('--')) {
        throw new Error('Missing value for --root. Usage: --root <path>');
      }
      root = raw;
      i += 1;
      continue;
    }
    throw new Error(`Unknown argument: ${arg}`);
  }

  return { strict, root };
}

function walkTsFiles(root: string): string[] {
  const entries = fs.readdirSync(root, { withFileTypes: true });
  const files: string[] = [];
  for (const entry of entries) {
    const fullPath = path.join(root, entry.name);
    if (entry.isDirectory()) {
      files.push(...walkTsFiles(fullPath));
      continue;
    }
    if (entry.isFile() && path.extname(entry.name) === '.ts') {
      files.push(fullPath);
    }
  }
  return files;
}

function extractRelativeSpecifiers(source: string): string[] {
  const specifiers = new Set<string>();
  const pattern = /(?:import|export)\s+(?:[^'";]+\s+from\s+)?['"]([^'"]+)['"]/g;
  for (const match of source.matchAll(pattern)) {
    const specifier = match[1];
    if (specifier.startsWith('.')) {
      specifiers.add(specifier);
    }
  }
  return Array.from(specifiers);
}

function resolveRelativeTarget(importerFile: string, specifier: string): string | null {
  const importerDir = path.dirname(importerFile);
  const baseTarget = path.resolve(importerDir, specifier);
  const extension = path.extname(baseTarget);

  const candidates = extension
    ? [baseTarget]
    : [`${baseTarget}.ts`, path.join(baseTarget, 'index.ts')];

  for (const candidate of candidates) {
    if (fs.existsSync(candidate) && fs.statSync(candidate).isFile()) {
      return path.normalize(candidate);
    }
  }

  return null;
}

function toRelativePosix(root: string, filePath: string): string {
  return path.relative(root, filePath).split(path.sep).join('/');
}

function detectCycle(graph: Map<string, string[]>): string[] {
  const state = new Map<string, 0 | 1 | 2>();
  const stack: string[] = [];

  const visit = (node: string): string[] => {
    state.set(node, 1);
    stack.push(node);

    const neighbors = graph.get(node) ?? [];
    for (const next of neighbors) {
      const nextState = state.get(next) ?? 0;
      if (nextState === 0) {
        const cycle = visit(next);
        if (cycle.length > 0) {
          return cycle;
        }
      } else if (nextState === 1) {
        const startIndex = stack.indexOf(next);
        if (startIndex >= 0) {
          return [...stack.slice(startIndex), next];
        }
      }
    }

    stack.pop();
    state.set(node, 2);
    return [];
  };

  const nodes = Array.from(graph.keys()).sort();
  for (const node of nodes) {
    if ((state.get(node) ?? 0) !== 0) {
      continue;
    }
    const cycle = visit(node);
    if (cycle.length > 0) {
      return cycle;
    }
  }

  return [];
}

export function runRuntimeCycleCheck(rootArg: string, strict: boolean): RuntimeCycleCheckResult {
  const root = path.resolve(process.cwd(), rootArg);
  if (!fs.existsSync(root) || !fs.statSync(root).isDirectory()) {
    throw new Error(
      `Runtime root not found: ${rootArg}. Use --root <path> to point at a directory.`,
    );
  }

  const files = walkTsFiles(root)
    .map((file) => path.normalize(file))
    .sort();
  const fileSet = new Set(files);
  const graph = new Map<string, string[]>();

  for (const file of files) {
    const source = fs.readFileSync(file, 'utf8');
    const specifiers = extractRelativeSpecifiers(source);
    const edges = new Set<string>();
    for (const specifier of specifiers) {
      const target = resolveRelativeTarget(file, specifier);
      if (!target) {
        continue;
      }
      if (fileSet.has(target)) {
        edges.add(target);
      }
    }
    graph.set(file, Array.from(edges).sort());
  }

  const cycle = detectCycle(graph);
  return {
    root: rootArg,
    mode: strict ? 'strict' : 'warning',
    fileCount: files.length,
    hasCycle: cycle.length > 0,
    cyclePath: cycle.map((file) => toRelativePosix(root, file)),
  };
}

export function formatSummary(result: RuntimeCycleCheckResult): string[] {
  const lines: string[] = [];
  if (!result.hasCycle) {
    lines.push(
      `[OK] runtime cycle check (${result.mode}) - ${result.fileCount} files scanned, no cycles detected`,
    );
    return lines;
  }

  const heading = result.mode === 'strict' ? '[FAIL]' : '[WARN]';
  lines.push(`${heading} runtime cycle check (${result.mode}) - cycle detected`);
  lines.push(`  root: ${result.root}`);
  lines.push(`  cycle: ${result.cyclePath.join(' -> ')}`);
  lines.push('  Hint: break bidirectional imports by moving shared logic to leaf modules.');
  return lines;
}

export function main(argv: string[] = process.argv.slice(2)): void {
  try {
    const { strict, root } = parseArgs(argv);
    const result = runRuntimeCycleCheck(root, strict);
    for (const line of formatSummary(result)) {
      console.log(line);
    }

    if (strict && result.hasCycle) {
      process.exitCode = 1;
    }
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    console.log(`[FAIL] runtime cycle check (strict) - ${message}`);
    process.exitCode = 1;
  }
}

if (import.meta.main) {
  main();
}
