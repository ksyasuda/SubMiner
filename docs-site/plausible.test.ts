import { expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';

const docsConfigPath = new URL('./.vitepress/config.ts', import.meta.url);
const docsThemePath = new URL('./.vitepress/theme/index.ts', import.meta.url);
const docsPackagePath = new URL('./package.json', import.meta.url);
const versionedBuildPath = new URL('../scripts/build-versioned-docs.ts', import.meta.url);
const docsConfigContents = readFileSync(docsConfigPath, 'utf8');
const docsThemeContents = readFileSync(docsThemePath, 'utf8');
const docsPackageContents = readFileSync(docsPackagePath, 'utf8');
const versionedBuildContents = readFileSync(versionedBuildPath, 'utf8');

test('docs site loads the docs.subminer.moe Plausible script through the analytics proxy', () => {
  expect(docsConfigContents).toContain("const DOCS_HOSTNAME = 'https://docs.subminer.moe'");
  expect(docsConfigContents).toContain(
    "const PLAUSIBLE_PROXY_HOSTNAME = 'https://worker.sudacode.com'",
  );
  expect(docsConfigContents).toContain(
    "const PLAUSIBLE_SITE_SCRIPT_PATH = '/js/pa-h28Pn9ppgTJRmiSJlyPT6.js'",
  );
  expect(docsConfigContents).toContain(
    'const PLAUSIBLE_ENDPOINT = `${PLAUSIBLE_PROXY_HOSTNAME}/api/event`',
  );
  expect(docsConfigContents).toContain('hostname: DOCS_HOSTNAME');
  expect(docsConfigContents).toContain("rel: 'preconnect'");
  expect(docsConfigContents).toContain('href: PLAUSIBLE_PROXY_HOSTNAME');
  expect(docsConfigContents).toContain("async: ''");
  expect(docsConfigContents).toContain(
    'src: `${PLAUSIBLE_PROXY_HOSTNAME}${PLAUSIBLE_SITE_SCRIPT_PATH}`',
  );
  expect(docsConfigContents).toContain('plausible.init({ endpoint:');
  expect(docsConfigContents).toContain('PLAUSIBLE_ENDPOINT');
  expect(docsConfigContents).not.toContain("'data-domain'");
  expect(docsConfigContents).not.toContain("'data-api'");
  expect(docsThemeContents).not.toContain('@plausible-analytics/tracker');
  expect(docsThemeContents).not.toContain('initPlausibleTracker');
  expect(docsPackageContents).not.toContain('@plausible-analytics/tracker');
});

test('versioned docs reuse current VitePress internals for old page snapshots', () => {
  expect(versionedBuildContents).toContain("cpSync(join(currentDocsSite, '.vitepress')");
  expect(versionedBuildContents).toContain('overlayCurrentVitePress(snapshotDocsSite)');
});
