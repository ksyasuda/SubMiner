import test from 'node:test';
import assert from 'node:assert/strict';
import { interleave, mapSourcesConcurrently } from './multi-source-search';

const source = (id: string) => ({ id, name: `Source ${id}` });

test('mapSourcesConcurrently returns results in source order, not completion order', async () => {
  const sources = [source('a'), source('b'), source('c')];
  const delays: Record<string, number> = { a: 20, b: 0, c: 10 };

  const { results, failures } = await mapSourcesConcurrently(sources, async (target) => {
    await new Promise((resolve) => setTimeout(resolve, delays[target.id]));
    return target.id;
  });

  assert.deepEqual(results, ['a', 'b', 'c']);
  assert.deepEqual(failures, []);
});

test('a failing source is reported without losing the others', async () => {
  const sources = [source('a'), source('b'), source('c')];

  const { results, failures } = await mapSourcesConcurrently(sources, async (target) => {
    if (target.id === 'b') throw new Error('login required');
    return target.id;
  });

  assert.deepEqual(results, ['a', 'c']);
  assert.deepEqual(failures, [{ sourceId: 'b', sourceName: 'Source b', error: 'login required' }]);
});

test('mapSourcesConcurrently never runs more than the concurrency limit at once', async () => {
  const sources = ['a', 'b', 'c', 'd', 'e'].map(source);
  let running = 0;
  let peak = 0;

  await mapSourcesConcurrently(
    sources,
    async () => {
      running += 1;
      peak = Math.max(peak, running);
      await new Promise((resolve) => setTimeout(resolve, 5));
      running -= 1;
    },
    2,
  );

  assert.equal(peak, 2);
});

test('mapSourcesConcurrently handles an empty source list', async () => {
  const { results, failures } = await mapSourcesConcurrently([], async () => 'x');
  assert.deepEqual(results, []);
  assert.deepEqual(failures, []);
});

test('interleave takes one from each source before taking a second', () => {
  assert.deepEqual(interleave([['a1', 'a2', 'a3'], ['b1'], ['c1', 'c2']]), [
    'a1',
    'b1',
    'c1',
    'a2',
    'c2',
    'a3',
  ]);
});

test('interleave ignores empty groups', () => {
  assert.deepEqual(interleave([[], ['b1', 'b2'], []]), ['b1', 'b2']);
  assert.deepEqual(interleave([]), []);
});
