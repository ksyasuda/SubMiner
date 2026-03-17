import assert from 'node:assert/strict';
import test from 'node:test';

import { epochMsFromDbTimestamp, formatRelativeDate, formatSessionDayLabel } from './formatters';

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

test('formatRelativeDate: same calendar day can return "23h ago"', () => {
  const realNow = Date.now;
  const now = new Date(2026, 2, 16, 23, 30, 0).getTime();
  const sameDayMorning = new Date(2026, 2, 16, 0, 30, 0).getTime();
  Date.now = () => now;
  try {
    assert.equal(formatRelativeDate(sameDayMorning), '23h ago');
  } finally {
    Date.now = realNow;
  }
});

test('formatRelativeDate: two calendar days ago returns "2d ago"', () => {
  const realNow = Date.now;
  const now = new Date(2026, 2, 16, 12, 0, 0).getTime();
  const twoDaysAgo = new Date(2026, 2, 14, 0, 0, 0).getTime();
  Date.now = () => now;
  try {
    assert.equal(formatRelativeDate(twoDaysAgo), '2d ago');
  } finally {
    Date.now = realNow;
  }
});

test('formatRelativeDate: 5 days ago returns "5d ago"', () => {
  assert.equal(formatRelativeDate(Date.now() - 5 * 86_400_000), '5d ago');
});

test('formatRelativeDate: 10 days ago returns locale date string', () => {
  const ts = Date.now() - 10 * 86_400_000;
  assert.equal(formatRelativeDate(ts), new Date(ts).toLocaleDateString());
});

test('formatRelativeDate: prior calendar day under 24h returns "Yesterday"', () => {
  const realNow = Date.now;
  const now = new Date(2026, 2, 16, 0, 30, 0).getTime();
  const previousDayLate = new Date(2026, 2, 15, 23, 45, 0).getTime();
  Date.now = () => now;
  try {
    assert.equal(formatRelativeDate(previousDayLate), 'Yesterday');
  } finally {
    Date.now = realNow;
  }
});

test('epochMsFromDbTimestamp converts seconds to ms', () => {
  assert.equal(epochMsFromDbTimestamp(1_700_000_000), 1_700_000_000_000);
});

test('epochMsFromDbTimestamp keeps ms timestamps as-is', () => {
  assert.equal(epochMsFromDbTimestamp(1_700_000_000_000), 1_700_000_000_000);
});

test('formatSessionDayLabel formats today and yesterday', () => {
  const now = Date.now();
  const oneDayMs = 24 * 60 * 60_000;
  assert.equal(formatSessionDayLabel(now), 'Today');
  assert.equal(formatSessionDayLabel(now - oneDayMs), 'Yesterday');
});

test('formatSessionDayLabel includes year for past-year dates', () => {
  const now = new Date();
  const sameDayLastYear = new Date(now.getFullYear() - 1, now.getMonth(), now.getDate()).getTime();
  const label = formatSessionDayLabel(sameDayLastYear);
  const year = new Date(sameDayLastYear).getFullYear();
  assert.ok(label.includes(String(year)));
  const withoutYear = new Date(sameDayLastYear).toLocaleDateString(undefined, {
    month: 'long',
    day: 'numeric',
  });
  assert.notEqual(label, withoutYear);
});
