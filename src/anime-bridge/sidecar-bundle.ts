import { readdir, readFile, stat, writeFile } from 'node:fs/promises';
import path from 'node:path';

/**
 * Locates the M-Extension-Server release bundle for the host platform. Each
 * bundle ships a matching JRE alongside the server JAR, so no system JDK is
 * required.
 *
 * Releases are taken from upstream's GitHub releases, newest first, the same
 * way Mangatan does it. Upstream publishes no checksums, so the download is
 * trusted on the strength of TLS to GitHub and the maintainer's account, which
 * is the same trust running their code implies in the first place.
 */

const BUNDLE_REPO_API = 'https://api.github.com/repos/1Selxo/M-Extension-Server';

/** Enough of the list to skip the iOS-runtime releases that carry no desktop bundle. */
const RELEASE_PAGE_SIZE = 20;

export function bundleReleaseUrl(): string {
  return `${BUNDLE_REPO_API}/releases?per_page=${RELEASE_PAGE_SIZE}`;
}

/**
 * The oldest server this client is known to work with. Releases below it are
 * skipped even when they are the only ones on offer; anything newer is taken,
 * and the readiness probe in `AnimeBridgeClient.isReady` catches a server that
 * has stopped speaking our protocol.
 */
export const MIN_BUNDLE_VERSION = 'v1.0.6.0';

/** `v1.0.6.2` → `[1, 0, 6, 2]`; null for tags that are not dotted numbers. */
export function parseBundleVersion(tag: string): number[] | null {
  const match = /^v?(\d+(?:\.\d+)*)(?:-|$)/.exec(tag.trim());
  return match ? match[1]!.split('.').map(Number) : null;
}

/** Numeric, segment-wise comparison; a missing segment reads as zero. */
export function compareBundleVersions(a: string, b: string): number {
  const left = parseBundleVersion(a) ?? [];
  const right = parseBundleVersion(b) ?? [];
  const length = Math.max(left.length, right.length);
  for (let index = 0; index < length; index += 1) {
    const difference = (left[index] ?? 0) - (right[index] ?? 0);
    if (difference !== 0) return Math.sign(difference);
  }
  return 0;
}

/** Release asset name for a platform/arch pair, or null when unsupported. */
export function resolveBundleAssetName(platform: string, arch: string): string | null {
  if (platform === 'linux') return arch === 'x64' ? 'linux-x64-bundle.zip' : null;
  if (platform === 'win32') return arch === 'x64' ? 'windows-x64-bundle.zip' : null;
  if (platform === 'darwin') {
    if (arch === 'arm64') return 'macOS-arm64-bundle.zip';
    if (arch === 'x64') return 'macOS-x64-bundle.zip';
  }
  return null;
}

export interface BundleBinaries {
  javaPath: string;
  jarPath: string;
}

async function walk(dir: string, depth: number, onFile: (file: string) => void): Promise<void> {
  if (depth < 0) return;
  let entries;
  try {
    entries = await readdir(dir, { withFileTypes: true });
  } catch {
    return;
  }
  for (const entry of entries) {
    const full = path.join(dir, entry.name);
    if (entry.isDirectory()) await walk(full, depth - 1, onFile);
    else onFile(full);
  }
}

/**
 * Find the bundled `java` executable and `MExtensionServer-*.jar` inside an
 * extracted bundle. The archive nests them a few levels deep and the exact
 * layout differs per platform, so this searches rather than assuming a path.
 */
export async function findBundleBinaries(rootDir: string): Promise<BundleBinaries | null> {
  const javaCandidates: string[] = [];
  const jarCandidates: string[] = [];
  await walk(rootDir, 6, (file) => {
    const base = path.basename(file);
    if (base === 'java' || base === 'java.exe') javaCandidates.push(file);
    else if (/^MExtensionServer.*\.jar$/.test(base)) jarCandidates.push(file);
  });
  // Prefer the shallowest match so a nested duplicate never shadows the real one.
  const byDepth = (a: string, b: string) => a.split(path.sep).length - b.split(path.sep).length;
  const javaPath = javaCandidates.sort(byDepth)[0];
  const jarPath = jarCandidates.sort(byDepth)[0];
  if (!javaPath || !jarPath) return null;
  return { javaPath, jarPath };
}

export async function isExecutableFile(file: string): Promise<boolean> {
  try {
    const info = await stat(file);
    return info.isFile();
  } catch {
    return false;
  }
}

export interface BundleAsset {
  tagName: string;
  assetName: string;
  downloadUrl: string;
  sizeBytes: number;
}

interface GithubRelease {
  tag_name?: string;
  draft?: boolean;
  prerelease?: boolean;
  assets?: Array<{ name?: string; browser_download_url?: string; size?: number }>;
}

/**
 * Pick the newest release that ships this platform's bundle and meets the
 * minimum version. Sorted by parsed version rather than list order, so a
 * re-published older release cannot shadow the current one.
 */
export function selectBundleAsset(
  releases: unknown,
  assetName: string,
  minimumVersion: string = MIN_BUNDLE_VERSION,
): BundleAsset | null {
  // Accepts either a single release or the list endpoint's array.
  const candidates = Array.isArray(releases)
    ? releases
    : releases && typeof releases === 'object'
      ? [releases]
      : [];
  const usable: BundleAsset[] = [];
  for (const release of candidates as GithubRelease[]) {
    if (!release.tag_name || release.draft || release.prerelease) continue;
    if (parseBundleVersion(release.tag_name) === null) continue;
    if (compareBundleVersions(release.tag_name, minimumVersion) < 0) continue;
    const asset = release.assets?.find((candidate) => candidate.name === assetName);
    if (!asset?.browser_download_url) continue;
    usable.push({
      tagName: release.tag_name,
      assetName,
      downloadUrl: asset.browser_download_url,
      sizeBytes: asset.size ?? 0,
    });
  }
  usable.sort((a, b) => compareBundleVersions(b.tagName, a.tagName));
  return usable[0] ?? null;
}

/**
 * Where a package manager may have put the same bundle. Only Arch has a
 * package today (AUR `mangatan-extension-server`, installed for Mangatan); it
 * unpacks the upstream layout verbatim, so `findBundleBinaries` reads it as-is.
 */
export function systemBundleDirs(platform: string): string[] {
  if (platform === 'linux') return ['/usr/share/mangatan/extension_server'];
  return [];
}

/**
 * The release version a server jar carries in its name, normalised to the tag
 * form (`MExtensionServer-v1.0.6.0-r1.jar` → `v1.0.6.0`). Upstream appends a
 * `-rN` rebuild suffix to some jars that the release tag does not carry.
 */
export function bundleVersionFromJar(jarPath: string): string | null {
  const match = /^MExtensionServer-v?(\d+(?:\.\d+)*)/.exec(path.basename(jarPath));
  return match ? `v${match[1]}` : null;
}

/**
 * Records which release SubMiner unpacked into a managed install directory.
 * Older installs predate the marker; they fall back to the jar name.
 */
export const BUNDLE_MARKER_FILE = 'bundle.json';

export async function writeBundleMarker(installDir: string, tag: string): Promise<void> {
  await writeFile(path.join(installDir, BUNDLE_MARKER_FILE), JSON.stringify({ tag }, null, 2));
}

export async function readBundleMarker(installDir: string): Promise<string | null> {
  try {
    const parsed: unknown = JSON.parse(
      await readFile(path.join(installDir, BUNDLE_MARKER_FILE), 'utf8'),
    );
    const tag = (parsed as { tag?: unknown } | null)?.tag;
    return typeof tag === 'string' && tag.length > 0 ? tag : null;
  } catch {
    return null;
  }
}
