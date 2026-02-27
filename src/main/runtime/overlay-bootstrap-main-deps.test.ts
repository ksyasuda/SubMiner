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
  const modalWindow = { id: 'modal' };
  const calls: string[] = [];
  const deps = createBuildOverlayModalRuntimeMainDepsHandler({
    getMainWindow: () => mainWindow as never,
    getModalWindow: () => modalWindow as never,
    createModalWindow: () => modalWindow as never,
    getModalGeometry: () => ({ x: 1, y: 2, width: 3, height: 4 }),
    setModalWindowBounds: (geometry) =>
      calls.push(`modal-bounds:${geometry.x},${geometry.y},${geometry.width},${geometry.height}`),
  })();

  assert.equal(deps.getMainWindow(), mainWindow);
  assert.equal(deps.getModalWindow(), modalWindow);
  assert.equal(deps.createModalWindow(), modalWindow);
  assert.deepEqual(deps.getModalGeometry(), { x: 1, y: 2, width: 3, height: 4 });
  deps.setModalWindowBounds({ x: 10, y: 20, width: 30, height: 40 });
  assert.deepEqual(calls, ['modal-bounds:10,20,30,40']);
});
