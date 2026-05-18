import { spawnSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import { cpSync, existsSync, lstatSync, mkdirSync, readFileSync, readdirSync, rmSync, symlinkSync, writeFileSync } from 'node:fs';
import { join, resolve } from 'node:path';
import {
  buildVersionManifest,
  stableTagsWithDocs,
  versionArchiveCacheName,
  versionOutputPath,
  versionPath,
} from './docs-versioning';

const repoRoot = resolve(__dirname, '..');
const currentDocsSite = join(repoRoot, 'docs-site');
const buildRoot = join(repoRoot, '.tmp/docs-versioned-build');
const aggregateOutDir = join(repoRoot, '.tmp/docs-versioned-site');
const archiveCacheRoot = join(repoRoot, '.tmp/docs-versioned-archive-cache');
const maxCloudflareFiles = 20_000;
const maxCloudflareFileBytes = 25 * 1024 * 1024;

function run(command: string, args: string[], options: { cwd?: string; env?: NodeJS.ProcessEnv } = {}) {
  const result = spawnSync(command, args, {
    cwd: options.cwd ?? repoRoot,
    env: options.env ?? process.env,
    stdio: 'inherit',
  });

  if (result.status !== 0) {
    throw new Error(`Command failed: ${command} ${args.join(' ')}`);
  }
}

function capture(command: string, args: string[]): string {
  const result = spawnSync(command, args, {
    cwd: repoRoot,
    encoding: 'utf8',
  });

  if (result.status !== 0) {
    throw new Error(result.stderr || `Command failed: ${command} ${args.join(' ')}`);
  }

  return result.stdout;
}

function archiveDocsSite(ref: string, targetDir: string) {
  mkdirSync(targetDir, { recursive: true });
  const archive = spawnSync('git', ['archive', '--format=tar', ref, 'docs-site'], {
    cwd: repoRoot,
    encoding: 'buffer',
    maxBuffer: 1024 * 1024 * 1024,
  });

  if (archive.status !== 0 || !archive.stdout) {
    throw new Error(`Unable to archive docs-site from ${ref}`);
  }

  const extract = spawnSync('tar', ['-x', '-C', targetDir], {
    input: archive.stdout,
    stdio: ['pipe', 'inherit', 'inherit'],
  });

  if (extract.status !== 0) {
    throw new Error(`Unable to extract docs-site archive from ${ref}`);
  }
}

function copyCurrentDocsSite(targetDir: string) {
  mkdirSync(targetDir, { recursive: true });
  cpSync(currentDocsSite, join(targetDir, 'docs-site'), {
    recursive: true,
    dereference: false,
    filter: (source) =>
      !/[\\/]node_modules([\\/]|$)/.test(source) &&
      !/[\\/]\\.vitepress[\\/]dist([\\/]|$)/.test(source),
  });
}

function overlayCurrentVitePress(snapshotDocsSite: string) {
  const targetVitePress = join(snapshotDocsSite, '.vitepress');
  rmSync(targetVitePress, { recursive: true, force: true });
  cpSync(join(currentDocsSite, '.vitepress'), targetVitePress, {
    recursive: true,
    filter: (source) => !isGeneratedVitePressPath(source),
  });

  const currentThemeFonts = join(currentDocsSite, 'public/assets/fonts');
  if (existsSync(currentThemeFonts)) {
    cpSync(currentThemeFonts, join(snapshotDocsSite, 'public/assets/fonts'), {
      recursive: true,
      force: true,
    });
  }
}

function linkDocsDependencies(snapshotDocsSite: string) {
  const currentNodeModules = join(currentDocsSite, 'node_modules');
  const targetNodeModules = join(snapshotDocsSite, 'node_modules');

  if (!existsSync(currentNodeModules) || existsSync(targetNodeModules)) {
    return;
  }

  symlinkSync(currentNodeModules, targetNodeModules, 'dir');
}

function prepareSnapshot(name: string, ref?: string): string {
  const snapshotRoot = join(buildRoot, name);
  rmSync(snapshotRoot, { recursive: true, force: true });

  if (ref) {
    archiveDocsSite(ref, snapshotRoot);
  } else {
    copyCurrentDocsSite(snapshotRoot);
  }

  const snapshotDocsSite = join(snapshotRoot, 'docs-site');
  overlayCurrentVitePress(snapshotDocsSite);
  linkDocsDependencies(snapshotDocsSite);
  return snapshotDocsSite;
}

function tagHasDocsSite(tag: string): boolean {
  const result = spawnSync('git', ['cat-file', '-e', `${tag}:docs-site/package.json`], {
    cwd: repoRoot,
  });
  return result.status === 0;
}

function getStableVersions(): string[] {
  const tags = capture('git', ['tag', '--list', 'v*'])
    .split('\n')
    .map((tag) => tag.trim())
    .filter(Boolean);
  return stableTagsWithDocs(tags, tagHasDocsSite);
}

function buildDocs(options: {
  snapshotDocsSite: string;
  base: string;
  outDir: string;
  channel: string;
  version?: string;
  latestStable: string;
  manifestJson: string;
}) {
  run('bun', ['run', '--cwd', currentDocsSite, 'vitepress', 'build', options.snapshotDocsSite], {
    cwd: repoRoot,
    env: {
      ...process.env,
      SUBMINER_DOCS_BASE: options.base,
      SUBMINER_DOCS_OUT_DIR: options.outDir,
      SUBMINER_DOCS_CHANNEL: options.channel,
      SUBMINER_DOCS_VERSION: options.version ?? '',
      SUBMINER_DOCS_LATEST_STABLE: options.latestStable,
      SUBMINER_DOCS_VERSION_MANIFEST: options.manifestJson,
      VITE_EXTRA_EXTENSIONS: 'jsonc',
    },
  });
}

function updateHashWithPath(hash: ReturnType<typeof createHash>, path: string) {
  if (isGeneratedVitePressPath(path)) {
    return;
  }

  const stat = lstatSync(path);
  hash.update(path.replace(repoRoot, ''));
  hash.update(String(stat.mode));

  if (stat.isSymbolicLink()) {
    return;
  }

  if (stat.isDirectory()) {
    for (const entry of readdirSync(path).sort()) {
      updateHashWithPath(hash, join(path, entry));
    }
    return;
  }

  hash.update(readFileSync(path));
}

function isGeneratedVitePressPath(path: string): boolean {
  return /[\\/]\\.vitepress[\\/](cache|dist)([\\/]|$)/.test(path);
}

function computeSharedInternalsHash(): string {
  const hash = createHash('sha256');
  const paths = [
    join(currentDocsSite, '.vitepress'),
    join(currentDocsSite, 'public/assets/fonts'),
    join(currentDocsSite, 'package.json'),
    join(currentDocsSite, 'bun.lock'),
    join(repoRoot, 'scripts/build-versioned-docs.ts'),
    join(repoRoot, 'scripts/docs-versioning.ts'),
  ];

  for (const path of paths) {
    if (existsSync(path)) {
      updateHashWithPath(hash, path);
    }
  }

  return hash.digest('hex');
}

function archiveCachePath(version: string, sharedInternalsHash: string): string {
  return join(archiveCacheRoot, versionArchiveCacheName(version, sharedInternalsHash));
}

function restoreCachedArchive(version: string, sharedInternalsHash: string): boolean {
  const cachedArchive = archiveCachePath(version, sharedInternalsHash);
  if (!existsSync(cachedArchive)) {
    return false;
  }

  cpSync(cachedArchive, join(aggregateOutDir, versionOutputPath(version)), {
    recursive: true,
    force: true,
  });
  return true;
}

function saveArchiveCache(version: string, sharedInternalsHash: string) {
  const outputPath = join(aggregateOutDir, versionOutputPath(version));
  if (!existsSync(outputPath)) {
    return;
  }

  const cachedArchive = archiveCachePath(version, sharedInternalsHash);
  rmSync(cachedArchive, { recursive: true, force: true });
  mkdirSync(archiveCacheRoot, { recursive: true });
  cpSync(outputPath, cachedArchive, { recursive: true, force: true });
}

function assertCloudflarePagesLimits(root: string) {
  let fileCount = 0;
  const oversizedFiles: string[] = [];

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

      fileCount += 1;
      if (stat.size > maxCloudflareFileBytes) {
        oversizedFiles.push(path);
      }
    }
  }

  walk(root);

  if (fileCount > maxCloudflareFiles) {
    throw new Error(
      `Versioned docs output has ${fileCount} files; Cloudflare Pages free plan limit is ${maxCloudflareFiles}.`,
    );
  }

  if (oversizedFiles.length > 0) {
    throw new Error(
      `Versioned docs output has files over 25 MiB:\n${oversizedFiles.join('\n')}`,
    );
  }
}

function main() {
  const stableVersions = getStableVersions();
  const latestStable = stableVersions[0];

  if (!latestStable) {
    throw new Error('No stable release tags with docs-site/package.json found.');
  }

  const manifest = buildVersionManifest({ latestStable, stableVersions });
  const manifestJson = JSON.stringify(manifest);
  const sharedInternalsHash = computeSharedInternalsHash();

  rmSync(buildRoot, { recursive: true, force: true });
  rmSync(aggregateOutDir, { recursive: true, force: true });
  mkdirSync(buildRoot, { recursive: true });
  mkdirSync(aggregateOutDir, { recursive: true });

  const latestStableSnapshot = prepareSnapshot(latestStable, latestStable);
  buildDocs({
    snapshotDocsSite: latestStableSnapshot,
    base: '/',
    outDir: aggregateOutDir,
    channel: 'stable-root',
    version: latestStable,
    latestStable,
    manifestJson,
  });

  for (const version of stableVersions) {
    if (restoreCachedArchive(version, sharedInternalsHash)) {
      continue;
    }

    const snapshot = version === latestStable ? latestStableSnapshot : prepareSnapshot(version, version);
    buildDocs({
      snapshotDocsSite: snapshot,
      base: versionPath(version),
      outDir: join(aggregateOutDir, versionOutputPath(version)),
      channel: 'stable-archive',
      version,
      latestStable,
      manifestJson,
    });
    saveArchiveCache(version, sharedInternalsHash);
  }

  const mainSnapshot = prepareSnapshot('main');
  buildDocs({
    snapshotDocsSite: mainSnapshot,
    base: '/main/',
    outDir: join(aggregateOutDir, 'main'),
    channel: 'main',
    version: 'main',
    latestStable,
    manifestJson,
  });

  writeFileSync(join(aggregateOutDir, 'versions.json'), `${JSON.stringify(manifest, null, 2)}\n`);
  assertCloudflarePagesLimits(aggregateOutDir);
}

main();
