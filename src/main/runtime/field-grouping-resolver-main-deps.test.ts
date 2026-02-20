import assert from 'node:assert/strict';
import test from 'node:test';
import {
  createBuildGetFieldGroupingResolverMainDepsHandler,
  createBuildSetFieldGroupingResolverMainDepsHandler,
} from './field-grouping-resolver-main-deps';

test('get field grouping resolver main deps builder maps callbacks', () => {
  const resolver = () => undefined;
  const deps = createBuildGetFieldGroupingResolverMainDepsHandler({
    getResolver: () => resolver,
  })();
  assert.equal(deps.getResolver(), resolver);
});

test('set field grouping resolver main deps builder maps callbacks', () => {
  const calls: string[] = [];
  const wrapped = (choice: unknown) => calls.push(String(choice));
  const deps = createBuildSetFieldGroupingResolverMainDepsHandler({
    setResolver: (resolver) => {
      if (resolver) {
        resolver('x' as never);
      }
    },
    nextSequence: () => 2,
    getSequence: () => 2,
  })();

  assert.equal(deps.nextSequence(), 2);
  assert.equal(deps.getSequence(), 2);
  deps.setResolver(wrapped as never);
  assert.deepEqual(calls, ['x']);
});
