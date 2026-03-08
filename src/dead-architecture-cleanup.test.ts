import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';

const DEAD_MODULE_PATHS = [
  'src/translators/index.ts',
  'src/subsync/engines.ts',
  'src/subtitle/pipeline.ts',
  'src/subtitle/stages/merge.ts',
  'src/subtitle/stages/normalize.ts',
  'src/subtitle/stages/normalize.test.ts',
  'src/subtitle/stages/tokenize.ts',
  'src/tokenizers/index.ts',
  'src/token-mergers/index.ts',
] as const;

const FORBIDDEN_IMPORT_PATTERNS = [
  /from ['"]\.\.?\/tokenizers['"]/,
  /from ['"]\.\.?\/token-mergers['"]/,
  /from ['"]\.\.?\/subtitle\/pipeline['"]/,
  /from ['"]\.\.?\/subsync\/engines['"]/,
  /from ['"]\.\.?\/translators['"]/,
] as const;

function readWorkspaceFile(relativePath: string): string {
  return fs.readFileSync(path.join(process.cwd(), relativePath), 'utf8');
}

function collectSourceFiles(rootDir: string): string[] {
  const absoluteRoot = path.join(process.cwd(), rootDir);
  const out: string[] = [];

  const visit = (currentDir: string) => {
    for (const entry of fs.readdirSync(currentDir, { withFileTypes: true })) {
      const fullPath = path.join(currentDir, entry.name);
      if (entry.isDirectory()) {
        visit(fullPath);
        continue;
      }
      if (!fullPath.endsWith('.ts') && !fullPath.endsWith('.tsx')) {
        continue;
      }
      out.push(path.relative(process.cwd(), fullPath).replaceAll('\\', '/'));
    }
  };

  visit(absoluteRoot);
  out.sort();
  return out;
}

test('dead registry and pipeline modules stay removed from the repository', () => {
  for (const relativePath of DEAD_MODULE_PATHS) {
    assert.equal(
      fs.existsSync(path.join(process.cwd(), relativePath)),
      false,
      `${relativePath} should stay deleted`,
    );
  }
});

test('live source tree no longer imports dead registry and pipeline modules', () => {
  for (const relativePath of collectSourceFiles('src')) {
    const source = readWorkspaceFile(relativePath);
    for (const pattern of FORBIDDEN_IMPORT_PATTERNS) {
      assert.doesNotMatch(source, pattern, `${relativePath} should not import ${pattern.source}`);
    }
  }
});
