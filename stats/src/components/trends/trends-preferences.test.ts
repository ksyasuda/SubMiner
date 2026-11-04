import assert from 'node:assert/strict';
import test from 'node:test';
import {
  loadGroupBy,
  loadRange,
  loadRunningTotalMetric,
  saveGroupBy,
  saveRange,
  saveRunningTotalMetric,
} from './trends-preferences';

function withStorage(run: () => void, initial: Record<string, string> = {}) {
  const values = new Map(Object.entries(initial));
  const previous = Object.getOwnPropertyDescriptor(globalThis, 'localStorage');
  Object.defineProperty(globalThis, 'localStorage', {
    configurable: true,
    value: {
      getItem: (key: string) => values.get(key) ?? null,
      setItem: (key: string, value: string) => values.set(key, value),
    },
  });
  try {
    run();
  } finally {
    if (previous) Object.defineProperty(globalThis, 'localStorage', previous);
    else Reflect.deleteProperty(globalThis, 'localStorage');
  }
}

test('range and grouping survive a reload, including custom ranges', () => {
  withStorage(() => {
    saveRange('90d');
    assert.equal(loadRange(), '90d');
    saveRange({ startDay: 20_000, endDay: 20_030 });
    assert.deepEqual(loadRange(), { startDay: 20_000, endDay: 20_030 });
    saveGroupBy('week');
    assert.equal(loadGroupBy(), 'week');
    saveRunningTotalMetric('cards');
    assert.equal(loadRunningTotalMetric(), 'cards');
  });
});

test('unreadable stored values fall back to defaults', () => {
  withStorage(
    () => {
      assert.equal(loadRange(), '30d');
      assert.equal(loadGroupBy(), 'day');
      assert.equal(loadRunningTotalMetric(), 'watchTime');
    },
    {
      'subminer-stats-trends-range': '9:1',
      'subminer-stats-trends-group-by': 'year',
      'subminer-stats-trends-running-total-metric': 'nope',
    },
  );
});
