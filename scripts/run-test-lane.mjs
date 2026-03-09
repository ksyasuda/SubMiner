import { readdirSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { relative, resolve } from 'node:path';
import { spawnSync } from 'node:child_process';

const repoRoot = resolve(fileURLToPath(new URL('..', import.meta.url)));

const lanes = {
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

function collectFiles(rootDir, includeSuffixes, excludeSet) {
  const out = [];
  const visit = (currentDir) => {
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

const lane = lanes[process.argv[2]];

if (!lane) {
  process.stderr.write(`Unknown test lane: ${process.argv[2] ?? '(missing)'}\n`);
  process.exit(1);
}

const files = lane.roots.flatMap((rootDir) => collectFiles(rootDir, lane.include, lane.exclude));

if (files.length === 0) {
  process.stderr.write(`No test files found for lane: ${process.argv[2]}\n`);
  process.exit(1);
}

const result = spawnSync('bun', ['test', ...files.map((file) => `./${file}`)], {
  cwd: repoRoot,
  stdio: 'inherit',
});

if (result.error) {
  throw result.error;
}

process.exit(result.status ?? 1);
