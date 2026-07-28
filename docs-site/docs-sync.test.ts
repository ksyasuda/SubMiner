import { expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';

const rootChangelogContents = readFileSync(new URL('../CHANGELOG.md', import.meta.url), 'utf8');
const readmeContents = readFileSync(new URL('./README.md', import.meta.url), 'utf8');
const usageContents = readFileSync(new URL('./usage.md', import.meta.url), 'utf8');
const installationContents = readFileSync(new URL('./installation.md', import.meta.url), 'utf8');
const mpvPluginContents = readFileSync(new URL('./mpv-plugin.md', import.meta.url), 'utf8');
const developmentContents = readFileSync(new URL('./development.md', import.meta.url), 'utf8');
const changelogContents = readFileSync(new URL('./changelog.md', import.meta.url), 'utf8');
const docsPackageContents = readFileSync(new URL('./package.json', import.meta.url), 'utf8');
const ankiIntegrationContents = readFileSync(
  new URL('./anki-integration.md', import.meta.url),
  'utf8',
);
const configurationContents = readFileSync(new URL('./configuration.md', import.meta.url), 'utf8');
const troubleshootingContents = readFileSync(
  new URL('./troubleshooting.md', import.meta.url),
  'utf8',
);

function extractReleaseHeadings(content: string, count: number): string[] {
  return Array.from(content.matchAll(/^## v[^\n]+$/gm))
    .map(([heading]) => heading)
    .slice(0, count);
}

function extractCurrentMinorHeadings(content: string): string[] {
  const allHeadings = Array.from(content.matchAll(/^## v(\d+\.\d+)\.\d+[^\n]*$/gm));
  if (allHeadings.length === 0) return [];
  const currentMinor = allHeadings[0]![1];
  return allHeadings.filter(([, minor]) => minor === currentMinor).map(([heading]) => heading);
}

test('docs reflect current launcher and release surfaces', () => {
  expect(usageContents).not.toContain('--mode preprocess');
  expect(usageContents).not.toContain('"automatic" (default)');
  expect(usageContents).toContain('during startup while mpv is paused');

  expect(installationContents).toContain('bun run build:appimage');
  expect(installationContents).toContain('bun run build:win');

  expect(mpvPluginContents).toContain('\\\\.\\pipe\\subminer-socket');

  expect(readmeContents).toContain('Automatic production and preview deployments: disabled');
  expect(readmeContents).toContain('/main/');
  expect(readmeContents).toContain('GitHub Actions direct upload with Wrangler');
  expect(developmentContents).not.toContain('../subminer-docs');
  expect(developmentContents).toContain('bun run docs:build');
  expect(developmentContents).toContain('bun run docs:test');
  expect(developmentContents).toContain('bun run docs:build:versioned');
  expect(developmentContents).not.toContain('test:subtitle:dist');
  expect(developmentContents).toContain('bun run build:win');

  expect(ankiIntegrationContents).not.toContain('alwaysUseAiTranslation');
  expect(ankiIntegrationContents).not.toContain('targetLanguage');
  expect(configurationContents).not.toContain('youtubeSubgen": {\n    "mode"');
  expect(configurationContents).not.toContain('youtubeSubgen.primarySubLanguages');
  expect(configurationContents).toContain('youtube.primarySubLanguages');
  expect(configurationContents).toContain('### Shared AI Provider');

  expect(changelogContents).toContain('v0.5.1 (2026-03-09)');
});

test('docs document config surfaces that are easy to miss when they ship', () => {
  // Anki maturity-based known-word highlighting (#172) landed in
  // subtitle-annotations.md but was missing from the config reference.
  expect(configurationContents).toContain('ankiConnect.knownWords.maturityEnabled');
  expect(configurationContents).toContain('ankiConnect.knownWords.matureThresholdDays');

  // Every top-level config block should be reachable from the config reference.
  expect(configurationContents).toContain('### TsukiHime');
  expect(configurationContents).toContain('tsukihime.maxSearchResults');

  // xz is a hard runtime dependency of the TsukiHime download path.
  expect(installationContents).toContain('xz');
});

test('docs state the real secondary-subtitle and Anki field-matching behavior', () => {
  // secondarySub auto-load is off by default; the config example previously
  // implied otherwise while youtube-integration.md documented it correctly.
  expect(configurationContents).toContain('Secondary subtitles do **not** auto-load by default');
  expect(configurationContents).toContain('default: `false`');

  // Anki field names are matched case-insensitively (src/anki-integration.ts
  // resolveFieldName: exact match first, then a lowercase comparison).
  expect(usageContents).not.toContain('exactly (case-sensitive)');
  expect(troubleshootingContents).not.toContain('exactly (case-sensitive)');
  expect(ankiIntegrationContents).toContain('case-insensitively');
});

test('docs dev server links version navigation to local dev routes', () => {
  expect(docsPackageContents).toContain('scripts/build-versioned-docs.ts');
  expect(docsPackageContents).toContain(
    'SUBMINER_DOCS_VERSION_LINK_ORIGIN=local bun run ../scripts/build-versioned-docs.ts',
  );
  expect(docsPackageContents).toContain('SUBMINER_DOCS_VERSION_LINK_ORIGIN=local');
  expect(docsPackageContents).toContain('SUBMINER_DOCS_VERSION_MANIFEST');
});

test('docs changelog keeps the current minor release headings aligned with the root changelog', () => {
  const docsHeadings = extractCurrentMinorHeadings(changelogContents);
  expect(docsHeadings.length).toBeGreaterThan(0);
  expect(docsHeadings).toEqual(extractReleaseHeadings(rootChangelogContents, docsHeadings.length));
});
