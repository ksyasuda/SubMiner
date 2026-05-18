import { existsSync, lstatSync, readdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { join, relative } from 'node:path';

const textOutputPattern = /\.(?:css|html|js|json|map|mjs|txt|xml)$/;

function normalizeBase(base: string): string {
  return base === '/' ? '/' : `${base.replace(/\/+$/, '')}/`;
}

export function collectSharedAssetPaths(publicAssetsDir: string): Set<string> {
  if (!existsSync(publicAssetsDir)) {
    return new Set();
  }

  return new Set(
    walkFiles(publicAssetsDir).map((file) => relative(publicAssetsDir, file).split('\\').join('/')),
  );
}

export function rewriteSharedAssetReferences(
  content: string,
  base: string,
  sharedAssetPaths: Set<string>,
): string {
  const normalizedBase = normalizeBase(base);
  if (normalizedBase === '/') {
    return content;
  }

  let rewritten = content;
  for (const assetPath of sharedAssetPaths) {
    rewritten = rewritten
      .split(`${normalizedBase}assets/${assetPath}`)
      .join(`/assets/${assetPath}`);
  }
  return rewritten;
}

function walkFiles(root: string): string[] {
  const files: string[] = [];

  function walk(dir: string) {
    for (const entry of readdirSync(dir)) {
      const path = join(dir, entry);
      const stat = lstatSync(path);
      if (stat.isSymbolicLink()) {
        continue;
      }
      if (stat.isDirectory()) {
        walk(path);
        continue;
      }
      files.push(path);
    }
  }

  walk(root);
  return files;
}

export function dedupeVersionedPublicAssets(options: {
  outDir: string;
  base: string;
  sharedAssetPaths: Set<string>;
}): {
  removedAssetsDir: boolean;
  rewrittenFiles: string[];
} {
  const normalizedBase = normalizeBase(options.base);
  const rewrittenFiles: string[] = [];

  for (const file of walkFiles(options.outDir)) {
    if (!textOutputPattern.test(file)) {
      continue;
    }

    const before = readFileSync(file, 'utf8');
    const after = rewriteSharedAssetReferences(before, normalizedBase, options.sharedAssetPaths);
    if (after === before) {
      continue;
    }

    writeFileSync(file, after);
    rewrittenFiles.push(file);
  }

  const assetsDir = join(options.outDir, 'assets');
  if (normalizedBase !== '/') {
    for (const assetPath of options.sharedAssetPaths) {
      rmSync(join(assetsDir, assetPath), { force: true });
    }
    removeEmptyDirectories(assetsDir);
  }

  const removedAssetsDir = !existsSync(assetsDir);
  return { removedAssetsDir, rewrittenFiles };
}

function removeEmptyDirectories(root: string) {
  if (!existsSync(root) || !lstatSync(root).isDirectory()) {
    return;
  }

  for (const entry of readdirSync(root)) {
    const path = join(root, entry);
    if (lstatSync(path).isDirectory()) {
      removeEmptyDirectories(path);
    }
  }

  if (readdirSync(root).length === 0) {
    rmSync(root, { recursive: true, force: true });
  }
}

export function pruneArchiveCacheGenerations(options: {
  cacheRoot: string;
  activeCacheKey: string;
}): string[] {
  if (!existsSync(options.cacheRoot)) {
    return [];
  }

  const activePrefix = options.activeCacheKey.slice(0, 12);
  const removed: string[] = [];

  for (const entry of readdirSync(options.cacheRoot)) {
    const path = join(options.cacheRoot, entry);
    if (!lstatSync(path).isDirectory()) {
      continue;
    }
    if (entry.startsWith(`${activePrefix}-`)) {
      continue;
    }

    rmSync(path, { recursive: true, force: true });
    removed.push(path);
  }

  return removed;
}
