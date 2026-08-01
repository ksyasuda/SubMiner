/**
 * Client for Aniyomi-format extension repositories.
 *
 * SubMiner ships no repositories and performs no discovery. A repository only
 * exists once the user adds its index URL, and only extensions from those
 * repositories are ever listed or downloaded.
 */

/** Aniyomi extension packages carry this prefix; manga packages are ignored. */
const ANIME_PACKAGE_PREFIX = 'eu.kanade.tachiyomi.animeextension';

/**
 * Repos are identified by their index URL. The file name is not fixed:
 * `index.min.json` is the Aniyomi convention, but repositories publish under
 * other names too (e.g. `video.min.json`), so only https and a `.json` file
 * name are required.
 */
const INDEX_URL_PATTERN = /^https:\/\/[^\s/]+(?:\/[^\s/]*)*\/[^\s/]+\.json$/;

export interface RepoExtension {
  /** Package name, the stable identity of an extension across versions. */
  pkg: string;
  name: string;
  lang: string;
  version: string;
  /** Monotonic version code; the comparison basis for updates. */
  versionCode: number;
  nsfw: boolean;
  apkUrl: string;
  iconUrl: string;
  /** Index URL of the repo this came from. */
  repoUrl: string;
  /** Source names the package provides, when the index declares them. */
  sourceNames: string[];
}

/** `true` when `url` is a usable Aniyomi index URL. */
export function isValidRepoUrl(url: string): boolean {
  return INDEX_URL_PATTERN.test(url.trim());
}

/** Strip the index file name to get the repo root. */
export function repoBaseUrl(indexUrl: string): string {
  return indexUrl.trim().replace(/\/[^/]*$/, '');
}

interface RawEntry {
  name?: unknown;
  pkg?: unknown;
  apk?: unknown;
  lang?: unknown;
  code?: unknown;
  version?: unknown;
  nsfw?: unknown;
  sources?: unknown;
}

function readSourceNames(sources: unknown): string[] {
  if (!Array.isArray(sources)) return [];
  return sources
    .map((source) =>
      source !== null && typeof source === 'object'
        ? (source as { name?: unknown }).name
        : undefined,
    )
    .filter((name): name is string => typeof name === 'string' && name.length > 0);
}

/**
 * Parse an index payload into anime extensions.
 *
 * Entries that are malformed, or that are manga rather than anime packages,
 * are skipped rather than failing the whole repo.
 */
export function parseRepoIndex(indexUrl: string, payload: unknown): RepoExtension[] {
  if (!Array.isArray(payload)) return [];
  const base = repoBaseUrl(indexUrl);

  const extensions: RepoExtension[] = [];
  for (const raw of payload as RawEntry[]) {
    if (raw === null || typeof raw !== 'object') continue;
    const pkg = typeof raw.pkg === 'string' ? raw.pkg : '';
    const apk = typeof raw.apk === 'string' ? raw.apk : '';
    if (!pkg.startsWith(ANIME_PACKAGE_PREFIX) || apk.length === 0) continue;

    const versionCode = Number(raw.code);
    extensions.push({
      pkg,
      // Repo entries are prefixed "Aniyomi: "; the app supplies its own context.
      name: (typeof raw.name === 'string' ? raw.name : pkg).replace(/^Aniyomi:\s*/, ''),
      lang: typeof raw.lang === 'string' ? raw.lang : 'all',
      version: typeof raw.version === 'string' ? raw.version : '0',
      versionCode: Number.isFinite(versionCode) ? versionCode : 0,
      nsfw: Number(raw.nsfw) === 1,
      apkUrl: `${base}/apk/${apk}`,
      iconUrl: `${base}/icon/${pkg}.png`,
      repoUrl: indexUrl,
      sourceNames: readSourceNames(raw.sources),
    });
  }
  return extensions;
}

export interface FetchRepoOptions {
  fetchImpl?: typeof fetch;
  signal?: AbortSignal;
}

/** Fetch and parse one repository index. */
export async function fetchRepoIndex(
  indexUrl: string,
  options: FetchRepoOptions = {},
): Promise<RepoExtension[]> {
  if (!isValidRepoUrl(indexUrl)) {
    throw new Error(`Not a valid repository index URL: ${indexUrl}`);
  }
  const fetchImpl = options.fetchImpl ?? fetch;
  const response = await fetchImpl(indexUrl.trim(), {
    headers: { Accept: 'application/json' },
    ...(options.signal ? { signal: options.signal } : {}),
  });
  if (!response.ok) {
    throw new Error(`Repository returned ${response.status} for ${indexUrl}`);
  }
  return parseRepoIndex(indexUrl, await response.json());
}

export interface RepoFetchFailure {
  repoUrl: string;
  error: string;
}

export interface RepoCatalogue {
  extensions: RepoExtension[];
  failures: RepoFetchFailure[];
}

/**
 * Fetch every configured repository.
 *
 * When two repos publish the same package, the higher version code wins, so a
 * user's preferred repo ordering does not silently pin an older build.
 */
export async function fetchRepoCatalogue(
  indexUrls: string[],
  options: FetchRepoOptions = {},
): Promise<RepoCatalogue> {
  const failures: RepoFetchFailure[] = [];
  const byPackage = new Map<string, RepoExtension>();

  const results = await Promise.all(
    indexUrls.map(async (indexUrl) => {
      try {
        return { indexUrl, extensions: await fetchRepoIndex(indexUrl, options) };
      } catch (error) {
        failures.push({
          repoUrl: indexUrl,
          error: error instanceof Error ? error.message : String(error),
        });
        return { indexUrl, extensions: [] as RepoExtension[] };
      }
    }),
  );

  for (const { extensions } of results) {
    for (const extension of extensions) {
      const existing = byPackage.get(extension.pkg);
      if (!existing || extension.versionCode > existing.versionCode) {
        byPackage.set(extension.pkg, extension);
      }
    }
  }

  return {
    extensions: [...byPackage.values()].sort((a, b) => a.name.localeCompare(b.name)),
    failures,
  };
}

/** File name an extension is stored under, so updates replace in place. */
export function extensionFileName(pkg: string): string {
  return `${pkg}.apk`;
}
