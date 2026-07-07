import { readdirSync } from 'node:fs';
import { relative, resolve } from 'node:path';

export type TestLane = {
  roots: string[];
  include: string[];
  exclude?: string[];
  extraFiles?: string[];
};

// Single source of truth for test-lane membership. Consumed by
// scripts/run-test-lane.mjs (plain runs) and scripts/run-coverage-lane.ts
// (per-file coverage shards). Lanes discover files by directory so new test
// files join their lane automatically.
export const testLanes: Record<string, TestLane> = {
  'bun-src-full': {
    roots: ['src'],
    include: ['.test.ts', '.type-test.ts'],
    // Node-compat suites; their dist builds run via test:runtime:compat.
    exclude: [
      'src/core/services/anki-jimaku-ipc.test.ts',
      'src/core/services/ipc.test.ts',
      'src/core/services/overlay-manager.test.ts',
      'src/main/config-validation.test.ts',
      'src/main/runtime/registry.test.ts',
      'src/main/runtime/startup-config.test.ts',
    ],
  },
  config: {
    roots: ['src/config'],
    include: ['.test.ts'],
    extraFiles: ['src/generate-config-example.test.ts', 'src/verify-config-example.test.ts'],
  },
  launcher: {
    roots: ['launcher'],
    include: ['.test.ts'],
  },
  'bun-launcher-unit': {
    roots: ['launcher'],
    include: ['.test.ts'],
    exclude: ['launcher/smoke.e2e.test.ts'],
  },
  scripts: {
    roots: ['scripts'],
    include: ['.test.ts'],
  },
  stats: {
    roots: ['stats/src'],
    include: ['.test.ts', '.test.tsx'],
  },
};

function collectFiles(
  repoRoot: string,
  rootDir: string,
  includeSuffixes: string[],
  excludeSet: Set<string>,
): string[] {
  const out: string[] = [];
  const visit = (currentDir: string): void => {
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

export function collectLaneFiles(repoRoot: string, laneName: string): string[] {
  const lane = testLanes[laneName];
  if (!lane) {
    throw new Error(`Unknown test lane: ${laneName}`);
  }
  const excludeSet = new Set(lane.exclude ?? []);
  const files = lane.roots.flatMap((rootDir) =>
    collectFiles(repoRoot, rootDir, lane.include, excludeSet),
  );
  for (const extra of lane.extraFiles ?? []) {
    if (!files.includes(extra)) files.push(extra);
  }
  files.sort();
  if (files.length === 0) {
    throw new Error(`No test files found for lane: ${laneName}`);
  }
  return files;
}
