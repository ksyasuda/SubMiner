import { expect, test } from 'bun:test';
import { formatStatusLineDate, formatStatusLineFilePath } from './status-line';

test('status line file path formats root home as index markdown', () => {
  expect(formatStatusLineFilePath('/')).toBe('index.md');
});

test('status line file path formats version archive home without trailing slash', () => {
  expect(formatStatusLineFilePath('/v/0.12.0/')).toBe('v/0.12.0.md');
});

test('status line file path keeps normal docs routes as markdown files', () => {
  expect(formatStatusLineFilePath('/v/0.12.0/configuration')).toBe('v/0.12.0/configuration.md');
});

test('status line date uses the local calendar day, zero padded', () => {
  expect(formatStatusLineDate(new Date(2026, 0, 5, 23, 59))).toBe('2026-01-05');
});
