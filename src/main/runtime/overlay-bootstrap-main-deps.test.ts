import assert from 'node:assert/strict';
import test from 'node:test';
import {
  createBuildOverlayContentMeasurementStoreMainDepsHandler,
  createBuildOverlayModalRuntimeMainDepsHandler,
} from './overlay-bootstrap-main-deps';

test('overlay content measurement store main deps builder maps callbacks', () => {
  const calls: string[] = [];
  const deps = createBuildOverlayContentMeasurementStoreMainDepsHandler({
    now: () => 42,
    warn: (message) => calls.push(`warn:${message}`),
  })();

  assert.equal(deps.now(), 42);
  deps.warn('bad payload');
  assert.deepEqual(calls, ['warn:bad payload']);
});

test('overlay modal runtime main deps builder maps window resolvers', () => {
  const mainWindow = { id: 'main' };
  const invisibleWindow = { id: 'invisible' };
  const deps = createBuildOverlayModalRuntimeMainDepsHandler({
    getMainWindow: () => mainWindow as never,
    getInvisibleWindow: () => invisibleWindow as never,
  })();

  assert.equal(deps.getMainWindow(), mainWindow);
  assert.equal(deps.getInvisibleWindow(), invisibleWindow);
});
