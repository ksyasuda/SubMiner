import assert from 'node:assert/strict';
import test from 'node:test';

import { formatRelativeDate } from './formatters';

test('formatRelativeDate: future timestamps return "just now"', () => {
  assert.equal(formatRelativeDate(Date.now() + 60_000), 'just now');
});

test('formatRelativeDate: 0ms ago returns "just now"', () => {
  assert.equal(formatRelativeDate(Date.now()), 'just now');
});

test('formatRelativeDate: 30s ago returns "just now"', () => {
  assert.equal(formatRelativeDate(Date.now() - 30_000), 'just now');
});

test('formatRelativeDate: 5 minutes ago returns "5m ago"', () => {
  assert.equal(formatRelativeDate(Date.now() - 5 * 60_000), '5m ago');
});

test('formatRelativeDate: 59 minutes ago returns "59m ago"', () => {
  assert.equal(formatRelativeDate(Date.now() - 59 * 60_000), '59m ago');
});

test('formatRelativeDate: 2 hours ago returns "2h ago"', () => {
  assert.equal(formatRelativeDate(Date.now() - 2 * 3_600_000), '2h ago');
});

test('formatRelativeDate: 23 hours ago returns "23h ago"', () => {
  assert.equal(formatRelativeDate(Date.now() - 23 * 3_600_000), '23h ago');
});

test('formatRelativeDate: 36 hours ago returns "Yesterday"', () => {
  assert.equal(formatRelativeDate(Date.now() - 36 * 3_600_000), 'Yesterday');
});

test('formatRelativeDate: 5 days ago returns "5d ago"', () => {
  assert.equal(formatRelativeDate(Date.now() - 5 * 86_400_000), '5d ago');
});

test('formatRelativeDate: 10 days ago returns locale date string', () => {
  const ts = Date.now() - 10 * 86_400_000;
  assert.equal(formatRelativeDate(ts), new Date(ts).toLocaleDateString());
});
