import { expect, test } from 'bun:test';
import { formatStatusLineFilePath } from './status-line';

test('status line file path formats root home as index markdown', () => {
  expect(formatStatusLineFilePath('/')).toBe('index.md');
});

test('status line file path formats version archive home without trailing slash', () => {
  expect(formatStatusLineFilePath('/v/0.12.0/')).toBe('v/0.12.0.md');
});

test('status line file path keeps normal docs routes as markdown files', () => {
  expect(formatStatusLineFilePath('/v/0.12.0/configuration')).toBe('v/0.12.0/configuration.md');
});
