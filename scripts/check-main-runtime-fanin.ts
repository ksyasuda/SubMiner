import fs from 'node:fs';
import path from 'node:path';

type ParsedArgs = {
  strict: boolean;
  uniquePathLimit: number;
  importLineLimit: number;
};

const DEFAULT_UNIQUE_PATH_LIMIT = 11;
const DEFAULT_IMPORT_LINE_LIMIT = 110;
const MAIN_PATH = path.join(process.cwd(), 'src', 'main.ts');

function parseArgs(argv: string[]): ParsedArgs {
  let strict = false;
  let uniquePathLimit = DEFAULT_UNIQUE_PATH_LIMIT;
  let importLineLimit = DEFAULT_IMPORT_LINE_LIMIT;

  for (let i = 0; i < argv.length; i += 1) {
    const arg = argv[i];
    if (arg === '--strict') {
      strict = true;
      continue;
    }
    if (arg === '--unique-path-limit') {
      const raw = argv[i + 1];
      const parsed = Number.parseInt(raw ?? '', 10);
      if (!Number.isFinite(parsed) || parsed <= 0) {
        throw new Error(`Invalid --unique-path-limit value: ${raw ?? '<missing>'}`);
      }
      uniquePathLimit = parsed;
      i += 1;
      continue;
    }
    if (arg === '--import-line-limit') {
      const raw = argv[i + 1];
      const parsed = Number.parseInt(raw ?? '', 10);
      if (!Number.isFinite(parsed) || parsed <= 0) {
        throw new Error(`Invalid --import-line-limit value: ${raw ?? '<missing>'}`);
      }
      importLineLimit = parsed;
      i += 1;
      continue;
    }
  }

  return { strict, uniquePathLimit, importLineLimit };
}

function collectRuntimeImportStats(source: string): {
  importLines: number;
  uniquePaths: string[];
} {
  const pathMatches = Array.from(source.matchAll(/from '(\.\/main\/runtime[^']*)';/g)).map(
    (match) => match[1],
  );
  const uniquePaths = new Set<string>();

  for (const runtimeImportPath of pathMatches) {
    uniquePaths.add(runtimeImportPath);
  }

  return {
    importLines: pathMatches.length,
    uniquePaths: Array.from(uniquePaths).sort(),
  };
}

function main(): void {
  const { strict, uniquePathLimit, importLineLimit } = parseArgs(process.argv.slice(2));
  const source = fs.readFileSync(MAIN_PATH, 'utf8');
  const { importLines, uniquePaths } = collectRuntimeImportStats(source);
  const overUniquePathLimit = uniquePaths.length > uniquePathLimit;
  const overImportLineLimit = importLines > importLineLimit;
  const hasFailure = overUniquePathLimit || overImportLineLimit;
  const mode = strict ? 'strict' : 'warning';

  if (!hasFailure) {
    console.log(
      `[OK] main runtime fan-in (${mode}) — ${importLines} import lines, ${uniquePaths.length} unique runtime paths`,
    );
    return;
  }

  const heading = strict ? '[FAIL]' : '[WARN]';
  console.log(
    `${heading} main runtime fan-in (${mode}) — ${importLines} import lines, ${uniquePaths.length} unique runtime paths`,
  );
  console.log(`  limits: import lines <= ${importLineLimit}, unique paths <= ${uniquePathLimit}`);
  console.log('  runtime import paths:');
  for (const runtimeImportPath of uniquePaths) {
    console.log(`  - ${runtimeImportPath}`);
  }
  console.log(
    '  Hint: keep main.ts focused on boot wiring; move domain imports behind domain barrels/registries.',
  );

  if (strict) {
    process.exitCode = 1;
  }
}

main();
