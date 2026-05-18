import { expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';

const rootChangelogContents = readFileSync(new URL('../CHANGELOG.md', import.meta.url), 'utf8');
const readmeContents = readFileSync(new URL('./README.md', import.meta.url), 'utf8');
const usageContents = readFileSync(new URL('./usage.md', import.meta.url), 'utf8');
const installationContents = readFileSync(new URL('./installation.md', import.meta.url), 'utf8');
const mpvPluginContents = readFileSync(new URL('./mpv-plugin.md', import.meta.url), 'utf8');
const developmentContents = readFileSync(new URL('./development.md', import.meta.url), 'utf8');
const changelogContents = readFileSync(new URL('./changelog.md', import.meta.url), 'utf8');
const ankiIntegrationContents = readFileSync(
  new URL('./anki-integration.md', import.meta.url),
  'utf8',
);
const configurationContents = readFileSync(new URL('./configuration.md', import.meta.url), 'utf8');

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

test('docs changelog keeps the current minor release headings aligned with the root changelog', () => {
  const docsHeadings = extractCurrentMinorHeadings(changelogContents);
  expect(docsHeadings.length).toBeGreaterThan(0);
  expect(docsHeadings).toEqual(extractReleaseHeadings(rootChangelogContents, docsHeadings.length));
});
