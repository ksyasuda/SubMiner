export type DocsVersionEntry = {
  version: string;
  path: string;
};

export type DocsChannelEntry = {
  label: string;
  path: string;
};

export type DocsVersionManifest = {
  latestStable: string;
  channels: DocsChannelEntry[];
  versions: DocsVersionEntry[];
};

const STABLE_TAG_PATTERN = /^v\d+\.\d+\.\d+$/;

export function isStableReleaseTag(tag: string): boolean {
  return STABLE_TAG_PATTERN.test(tag);
}

function parseStableVersion(tag: string): [number, number, number] {
  const match = /^v(\d+)\.(\d+)\.(\d+)$/.exec(tag);
  if (!match) {
    throw new Error(`Invalid stable SubMiner version tag: ${tag}`);
  }

  return [Number(match[1]), Number(match[2]), Number(match[3])];
}

export function compareStableVersionsDesc(a: string, b: string): number {
  if (!isStableReleaseTag(a) && !isStableReleaseTag(b)) return a.localeCompare(b);
  if (!isStableReleaseTag(a)) return 1;
  if (!isStableReleaseTag(b)) return -1;

  const parsedA = parseStableVersion(a);
  const parsedB = parseStableVersion(b);

  for (let index = 0; index < parsedA.length; index += 1) {
    const difference = parsedB[index]! - parsedA[index]!;
    if (difference !== 0) return difference;
  }

  return 0;
}

export function versionPath(version: string): string {
  return `/v/${version.replace(/^v/, '')}/`;
}

export function versionOutputPath(version: string): string {
  return `v/${version.replace(/^v/, '')}`;
}

export function stableTagsWithDocs(
  tags: string[],
  hasDocsSite: (tag: string) => boolean,
): string[] {
  return tags.filter(isStableReleaseTag).filter(hasDocsSite).sort(compareStableVersionsDesc);
}

export function buildVersionManifest(options: {
  latestStable: string;
  stableVersions: string[];
}): DocsVersionManifest {
  return {
    latestStable: options.latestStable,
    channels: [
      { label: 'Latest stable', path: '/' },
      { label: 'main', path: '/main/' },
    ],
    versions: options.stableVersions.map((version) => ({
      version,
      path: versionPath(version),
    })),
  };
}

// Markdown for the root-only `/versions` page. Archives link here instead of baking the
// release list into their nav, so an archive never needs a rebuild when a new tag ships.
// Raw anchors with `target="_self"` keep VitePress from treating the other builds as
// dead links or routing to them client-side.
export function renderVersionsPage(manifest: DocsVersionManifest): string {
  const link = (path: string, text: string) => `<a href="${path}" target="_self">${text}</a>`;
  return [
    '---',
    'title: Documentation versions',
    'description: Every published version of the SubMiner documentation.',
    '---',
    '',
    '# Documentation versions',
    '',
    `- ${link('/', `Latest stable (${manifest.latestStable})`)}`,
    `- ${link('/main/', 'main')}: development docs, may describe unreleased behavior`,
    '',
    '## Stable releases',
    '',
    ...manifest.versions.map((entry) => `- ${link(entry.path, entry.version)}`),
    '',
  ].join('\n');
}
