import { expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';

const docsThemePath = new URL('./.vitepress/theme/index.ts', import.meta.url);
const docsThemeContents = readFileSync(docsThemePath, 'utf8');

test('docs theme has no plausible analytics wiring', () => {
  expect(docsThemeContents).not.toContain('@plausible-analytics/tracker');
  expect(docsThemeContents).not.toContain('initPlausibleTracker');
  expect(docsThemeContents).not.toContain('worker.subminer.moe');
  expect(docsThemeContents).not.toContain('domain:');
  expect(docsThemeContents).not.toContain('outboundLinks: true');
  expect(docsThemeContents).not.toContain('fileDownloads: true');
  expect(docsThemeContents).not.toContain('formSubmissions: true');
});
