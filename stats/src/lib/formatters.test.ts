import assert from 'node:assert/strict';
import test from 'node:test';

import { epochMsFromDbTimestamp, formatRelativeDate, formatSessionDayLabel } from './formatters';

const FIXED_NOW = new Date(2026, 2, 16, 12, 0, 0).getTime();

function withFixedNow(run: (now: number) => void, now = FIXED_NOW): void {
  const realNow = Date.now;
  Date.now = () => now;
  try {
    run(now);
  } finally {
    Date.now = realNow;
  }
}

test('formatRelativeDate: future timestamps return "just now"', () => {
  withFixedNow((now) => {
    assert.equal(formatRelativeDate(now + 60_000), 'just now');
  });
});

test('formatRelativeDate: 0ms ago returns "just now"', () => {
  withFixedNow((now) => {
    assert.equal(formatRelativeDate(now), 'just now');
  });
});

test('formatRelativeDate: 30s ago returns "just now"', () => {
  withFixedNow((now) => {
    assert.equal(formatRelativeDate(now - 30_000), 'just now');
  });
});

test('formatRelativeDate: 5 minutes ago returns "5m ago"', () => {
  withFixedNow((now) => {
    assert.equal(formatRelativeDate(now - 5 * 60_000), '5m ago');
  });
});

test('formatRelativeDate: 59 minutes ago returns "59m ago"', () => {
  withFixedNow((now) => {
    assert.equal(formatRelativeDate(now - 59 * 60_000), '59m ago');
  });
});

test('formatRelativeDate: 2 hours ago returns "2h ago"', () => {
  withFixedNow((now) => {
    assert.equal(formatRelativeDate(now - 2 * 3_600_000), '2h ago');
  });
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
  withFixedNow((now) => {
    assert.equal(formatRelativeDate(now - 5 * 86_400_000), '5d ago');
  });
});

test('formatRelativeDate: 10 days ago returns locale date string', () => {
  withFixedNow((now) => {
    const ts = now - 10 * 86_400_000;
    assert.equal(formatRelativeDate(ts), new Date(ts).toLocaleDateString());
  });
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
  withFixedNow((now) => {
    const oneDayMs = 24 * 60 * 60_000;
    assert.equal(formatSessionDayLabel(now), 'Today');
    assert.equal(formatSessionDayLabel(now - oneDayMs), 'Yesterday');
  });
});

test('formatSessionDayLabel includes year for past-year dates', () => {
  const fixedNow = new Date(2027, 2, 16, 12, 0, 0).getTime();
  withFixedNow((now) => {
    const current = new Date(now);
    const sameDayLastYear = new Date(
      current.getFullYear() - 1,
      current.getMonth(),
      current.getDate(),
    ).getTime();
    const label = formatSessionDayLabel(sameDayLastYear);
    const year = new Date(sameDayLastYear).getFullYear();
    assert.ok(label.includes(String(year)));
    const withoutYear = new Date(sameDayLastYear).toLocaleDateString(undefined, {
      month: 'long',
      day: 'numeric',
    });
    assert.notEqual(label, withoutYear);
  }, fixedNow);
});
