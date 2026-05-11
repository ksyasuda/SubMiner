import { expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';

const docsConfigPath = new URL('./.vitepress/config.ts', import.meta.url);
const docsThemePath = new URL('./.vitepress/theme/index.ts', import.meta.url);
const docsConfigContents = readFileSync(docsConfigPath, 'utf8');
const docsThemeContents = readFileSync(docsThemePath, 'utf8');

test('docs site keeps docs hostname while sending plausible events to subminer.moe via worker.subminer.moe capture endpoint', () => {
  expect(docsConfigContents).toContain("const DOCS_HOSTNAME = 'https://docs.subminer.moe'");
  expect(docsConfigContents).toContain('hostname: DOCS_HOSTNAME');
  expect(docsThemeContents).toContain("const PLAUSIBLE_DOMAIN = 'subminer.moe'");
  expect(docsThemeContents).toContain('const PLAUSIBLE_ENABLED_HOSTNAMES = new Set([');
  expect(docsThemeContents).toContain("'docs.subminer.moe'");
  expect(docsThemeContents).toContain(
    "const PLAUSIBLE_ENDPOINT = 'https://worker.subminer.moe/api/capture'",
  );
  expect(docsThemeContents).toContain('@plausible-analytics/tracker');
  expect(docsThemeContents).toContain('const { init } = await import');
  expect(docsThemeContents).toContain('!PLAUSIBLE_ENABLED_HOSTNAMES.has(window.location.hostname)');
  expect(docsThemeContents).toContain('domain: PLAUSIBLE_DOMAIN');
  expect(docsThemeContents).toContain('endpoint: PLAUSIBLE_ENDPOINT');
  expect(docsThemeContents).toContain('outboundLinks: true');
  expect(docsThemeContents).toContain('fileDownloads: true');
  expect(docsThemeContents).toContain('formSubmissions: true');
  expect(docsThemeContents).toContain('captureOnLocalhost: false');
});
