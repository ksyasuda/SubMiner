import type { UpdateChannel } from '../../../types/config';

export interface GitHubReleaseAsset {
  name: string;
  browser_download_url: string;
  size?: number;
}

export interface GitHubRelease {
  tag_name: string;
  name?: string;
  draft?: boolean;
  prerelease?: boolean;
  html_url?: string;
  assets: GitHubReleaseAsset[];
}

export interface FetchResponseLike {
  ok: boolean;
  status: number;
  statusText?: string;
  json: () => Promise<unknown>;
  text: () => Promise<string>;
  arrayBuffer: () => Promise<ArrayBuffer>;
}

export type FetchLike = (url: string, init?: Record<string, unknown>) => Promise<FetchResponseLike>;

export function parseSha256Sums(text: string): Map<string, string> {
  const sums = new Map<string, string>();
  for (const line of text.split(/\r?\n/)) {
    const trimmed = line.trim();
    if (!trimmed || trimmed.startsWith('#')) continue;
    const match = trimmed.match(/^([a-fA-F0-9]{64})\s+\*?(.+)$/);
    if (!match) continue;
    const [, hash, name] = match;
    if (!hash || !name) continue;
    sums.set(name.trim().split(/[\\/]/).pop() ?? name.trim(), hash.toLowerCase());
  }
  return sums;
}

export function selectLatestStableRelease(
  releases: GitHubRelease[],
  channel: UpdateChannel = 'stable',
): GitHubRelease | null {
  return (
    releases.find(
      (release) => !release.draft && (channel === 'prerelease' || !release.prerelease),
    ) ?? null
  );
}

export function findReleaseAsset(
  release: Pick<GitHubRelease, 'assets'>,
  assetName: string,
): GitHubReleaseAsset | null {
  return release.assets.find((asset) => asset.name === assetName) ?? null;
}

function assertRelease(value: unknown): GitHubRelease | null {
  if (!value || typeof value !== 'object') return null;
  const release = value as Partial<GitHubRelease>;
  if (typeof release.tag_name !== 'string' || !Array.isArray(release.assets)) return null;
  return {
    tag_name: release.tag_name,
    name: typeof release.name === 'string' ? release.name : undefined,
    draft: release.draft === true,
    prerelease: release.prerelease === true,
    html_url: typeof release.html_url === 'string' ? release.html_url : undefined,
    assets: release.assets
      .filter((asset): asset is GitHubReleaseAsset => {
        const candidate = asset as Partial<GitHubReleaseAsset>;
        return (
          typeof candidate.name === 'string' && typeof candidate.browser_download_url === 'string'
        );
      })
      .map((asset) => ({
        name: asset.name,
        browser_download_url: asset.browser_download_url,
        size: typeof asset.size === 'number' ? asset.size : undefined,
      })),
  };
}

export async function fetchLatestStableRelease(options: {
  fetch: FetchLike;
  owner?: string;
  repo?: string;
  channel?: UpdateChannel;
}): Promise<GitHubRelease | null> {
  const owner = options.owner ?? 'ksyasuda';
  const repo = options.repo ?? 'SubMiner';
  const response = await options.fetch(`https://api.github.com/repos/${owner}/${repo}/releases`, {
    headers: {
      Accept: 'application/vnd.github+json',
      'User-Agent': 'SubMiner updater',
    },
  });
  if (!response.ok) {
    throw new Error(`GitHub releases request failed with ${response.status}`);
  }
  const body = await response.json();
  if (!Array.isArray(body)) return null;
  return selectLatestStableRelease(
    body.map(assertRelease).filter((item): item is GitHubRelease => item !== null),
    options.channel,
  );
}

export async function fetchReleaseAssetText(fetch: FetchLike, assetUrl: string): Promise<string> {
  const response = await fetch(assetUrl);
  if (!response.ok) {
    throw new Error(`Release asset request failed with ${response.status}`);
  }
  return await response.text();
}

export async function fetchReleaseAssetBuffer(fetch: FetchLike, assetUrl: string): Promise<Buffer> {
  const response = await fetch(assetUrl);
  if (!response.ok) {
    throw new Error(`Release asset request failed with ${response.status}`);
  }
  return Buffer.from(await response.arrayBuffer());
}

export function parseReleaseVersion(
  release: Pick<GitHubRelease, 'tag_name'> | null,
): string | null {
  if (!release) return null;
  return release.tag_name.replace(/^v/i, '');
}

export function compareSemverLike(a: string, b: string): number {
  const parse = (
    value: string,
  ): {
    core: number[];
    prerelease: Array<number | string>;
  } => {
    const normalized = value.replace(/^v/i, '');
    const [coreText = '', ...prereleaseParts] = normalized.split('-');
    const core = coreText
      .split('.')
      .slice(0, 3)
      .map((part) => Number.parseInt(part, 10) || 0);
    while (core.length < 3) core.push(0);
    const prereleaseText = prereleaseParts.join('-');
    return {
      core,
      prerelease: prereleaseText
        ? prereleaseText.split('.').map((part) => {
            const numeric = Number.parseInt(part, 10);
            return /^\d+$/.test(part) ? numeric : part;
          })
        : [],
    };
  };
  const left = parse(a);
  const right = parse(b);
  for (let i = 0; i < 3; i += 1) {
    const diff = (left.core[i] ?? 0) - (right.core[i] ?? 0);
    if (diff !== 0) return diff;
  }

  if (left.prerelease.length === 0 && right.prerelease.length === 0) return 0;
  if (left.prerelease.length === 0) return 1;
  if (right.prerelease.length === 0) return -1;

  const length = Math.max(left.prerelease.length, right.prerelease.length);
  for (let i = 0; i < length; i += 1) {
    const leftPart = left.prerelease[i];
    const rightPart = right.prerelease[i];
    if (leftPart === undefined && rightPart === undefined) return 0;
    if (leftPart === undefined) return -1;
    if (rightPart === undefined) return 1;
    if (leftPart === rightPart) continue;
    if (typeof leftPart === 'number' && typeof rightPart === 'number') {
      return leftPart - rightPart;
    }
    if (typeof leftPart === 'number') return -1;
    if (typeof rightPart === 'number') return 1;
    return leftPart > rightPart ? 1 : -1;
  }
  return 0;
}
