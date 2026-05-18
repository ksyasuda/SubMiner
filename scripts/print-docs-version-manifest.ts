import { spawnSync } from 'node:child_process';
import { resolve } from 'node:path';
import { buildVersionManifest, stableTagsWithDocs } from './docs-versioning';

const repoRoot = resolve(__dirname, '..');

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

function tagHasDocsSite(tag: string): boolean {
  const result = spawnSync('git', ['cat-file', '-e', `${tag}:docs-site/package.json`], {
    cwd: repoRoot,
  });
  return result.status === 0;
}

const stableVersions = stableTagsWithDocs(
  capture('git', ['tag', '--list', 'v*'])
    .split('\n')
    .map((tag) => tag.trim())
    .filter(Boolean),
  tagHasDocsSite,
);

const latestStable = stableVersions[0];

if (!latestStable) {
  throw new Error('No stable release tags with docs-site/package.json found.');
}

process.stdout.write(JSON.stringify(buildVersionManifest({ latestStable, stableVersions })));
