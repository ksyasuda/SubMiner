import { expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';

const docsThemePath = new URL('./.vitepress/theme/index.ts', import.meta.url);
const docsThemeContents = readFileSync(docsThemePath, 'utf8');

test('docs theme configures plausible tracker for subminer.moe via worker.subminer.moe api endpoint', () => {
  expect(docsThemeContents).toContain('@plausible-analytics/tracker');
  expect(docsThemeContents).toContain('const { init } = await import');
  expect(docsThemeContents).toContain("domain: 'subminer.moe'");
  expect(docsThemeContents).toContain("endpoint: 'https://worker.subminer.moe/api/event'");
  expect(docsThemeContents).toContain('outboundLinks: true');
  expect(docsThemeContents).toContain('fileDownloads: true');
  expect(docsThemeContents).toContain('formSubmissions: true');
});
