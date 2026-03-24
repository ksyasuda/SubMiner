import { expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';

const rootChangelogContents = readFileSync(new URL('../CHANGELOG.md', import.meta.url), 'utf8');
const readmeContents = readFileSync(new URL('./README.md', import.meta.url), 'utf8');
const usageContents = readFileSync(new URL('./usage.md', import.meta.url), 'utf8');
const installationContents = readFileSync(new URL('./installation.md', import.meta.url), 'utf8');
const mpvPluginContents = readFileSync(new URL('./mpv-plugin.md', import.meta.url), 'utf8');
const developmentContents = readFileSync(new URL('./development.md', import.meta.url), 'utf8');
const changelogContents = readFileSync(new URL('./changelog.md', import.meta.url), 'utf8');
const ankiIntegrationContents = readFileSync(new URL('./anki-integration.md', import.meta.url), 'utf8');
const configurationContents = readFileSync(new URL('./configuration.md', import.meta.url), 'utf8');

function extractReleaseHeadings(content: string, count: number): string[] {
  return Array.from(content.matchAll(/^## v[^\n]+$/gm))
    .map(([heading]) => heading)
    .slice(0, count);
}

test('docs reflect current launcher and release surfaces', () => {
  expect(usageContents).not.toContain('--mode preprocess');
  expect(usageContents).not.toContain('"automatic" (default)');
  expect(usageContents).toContain('during startup while mpv is paused');

  expect(installationContents).toContain('bun run build:appimage');
  expect(installationContents).toContain('bun run build:win');

  expect(mpvPluginContents).toContain('\\\\.\\pipe\\subminer-socket');

  expect(readmeContents).toContain('Root directory: `docs-site`');
  expect(readmeContents).toContain('Build output directory: `.vitepress/dist`');
  expect(readmeContents).toContain('Build watch paths: `docs-site/*`');
  expect(developmentContents).not.toContain('../subminer-docs');
  expect(developmentContents).toContain('bun run docs:build');
  expect(developmentContents).toContain('bun run docs:test');
  expect(developmentContents).toContain('Build watch paths: `docs-site/*`');
  expect(developmentContents).not.toContain('test:subtitle:dist');
  expect(developmentContents).toContain('bun run build:win');

  expect(ankiIntegrationContents).not.toContain('alwaysUseAiTranslation');
  expect(ankiIntegrationContents).not.toContain('targetLanguage');
  expect(configurationContents).not.toContain('youtubeSubgen": {\n    "mode"');
  expect(configurationContents).toContain('### Shared AI Provider');

  expect(changelogContents).toContain('## v0.5.1 (2026-03-09)');
});

test('docs changelog keeps the newest release headings aligned with the root changelog', () => {
  expect(extractReleaseHeadings(changelogContents, 3)).toEqual(extractReleaseHeadings(rootChangelogContents, 3));
});
