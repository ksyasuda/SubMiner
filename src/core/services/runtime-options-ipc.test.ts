import test from 'node:test';
import assert from 'node:assert/strict';
import {
  applyRuntimeOptionResultRuntime,
  setRuntimeOptionFromIpcRuntime,
} from './runtime-options-ipc';

test('applyRuntimeOptionResultRuntime emits success OSD message', () => {
  const osd: string[] = [];
  const result = applyRuntimeOptionResultRuntime({ ok: true, osdMessage: 'Updated' }, (text) => {
    osd.push(text);
  });

  assert.equal(result.ok, true);
  assert.deepEqual(osd, ['Updated']);
});

test('setRuntimeOptionFromIpcRuntime returns unavailable when manager missing', () => {
  const osd: string[] = [];
  const result = setRuntimeOptionFromIpcRuntime(null, 'anki.autoUpdateNewCards', true, (text) => {
    osd.push(text);
  });
  assert.equal(result.ok, false);
  assert.equal(result.error, 'Runtime options manager unavailable');
  assert.deepEqual(osd, []);
});
