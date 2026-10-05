import assert from 'node:assert/strict';
import test from 'node:test';
import { createSetFieldGroupingResolverHandler } from './field-grouping-resolver';

test('set field grouping resolver clears resolver when null is provided', () => {
  let current: ((choice: unknown) => void) | null = () => undefined;
  const setResolver = createSetFieldGroupingResolverHandler({
    setResolver: (resolver) => {
      current = resolver as never;
    },
    nextSequence: () => 1,
    getSequence: () => 1,
  });

  setResolver(null);
  assert.equal(current, null);
});

test('set field grouping resolver wraps resolver and ignores stale sequence', () => {
  const calls: string[] = [];
  let current: ((choice: unknown) => void) | null = null;
  let sequence = 0;

  const setResolver = createSetFieldGroupingResolverHandler({
    setResolver: (resolver) => {
      current = resolver as never;
    },
    nextSequence: () => {
      sequence += 1;
      return sequence;
    },
    getSequence: () => sequence,
  });

  setResolver((choice) => calls.push(`new:${choice}`));
  const firstWrapped = current!;
  setResolver((choice) => calls.push(`latest:${choice}`));
  const latestWrapped = current!;

  firstWrapped('A');
  latestWrapped('B');

  assert.deepEqual(calls, ['latest:B']);
});
