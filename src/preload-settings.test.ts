import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import test from 'node:test';

test('settings preload stays sandbox-compatible by avoiding local runtime imports', () => {
  const source = fs.readFileSync(path.join(process.cwd(), 'src', 'preload-settings.ts'), 'utf8');

  assert.doesNotMatch(source, /from\s+['"]\.\/shared\/ipc\/contracts(?:\.(?:js|ts))?['"]/);
});

test('settings preload exposes Anki lookup helpers', () => {
  const source = fs.readFileSync(path.join(process.cwd(), 'src', 'preload-settings.ts'), 'utf8');

  for (const method of [
    'getAnkiDeckNames',
    'getAnkiDeckFieldNames',
    'getAnkiDeckModelNames',
    'getAnkiModelNames',
    'getAnkiModelFieldNames',
    'getYomitanAnkiDeckName',
  ]) {
    assert.match(source, new RegExp(`${method}:`));
  }
});

test('overlay preload buffers only latest subtitle state until renderer listener registration', () => {
  const source = fs.readFileSync(path.join(process.cwd(), 'src', 'preload.ts'), 'utf8');

  assert.match(
    source,
    /const onSubtitleSetEvent =\s*createLatestValueIpcListenerWithPayload<SubtitleData>\(\s*IPC_CHANNELS\.event\.subtitleSet,/,
  );
  assert.match(source, /onSubtitle:\s*\(callback:[\s\S]+?onSubtitleSetEvent\(callback\);/);
});

test('overlay preload does not expose the old mining image toast IPC path', () => {
  const source = fs.readFileSync(path.join(process.cwd(), 'src', 'preload.ts'), 'utf8');

  assert.doesNotMatch(source, /MiningImagePayload|onMiningImage|IPC_CHANNELS\.event\.miningImage/);
});

test('overlay preload exposes queued pointer recovery requests', () => {
  const source = fs.readFileSync(path.join(process.cwd(), 'src', 'preload.ts'), 'utf8');

  assert.match(
    source,
    /const onOverlayPointerRecoveryRequestEvent =\s*createQueuedIpcListener\(\s*IPC_CHANNELS\.event\.overlayPointerRecoveryRequest,/,
  );
  assert.match(
    source,
    /onOverlayPointerRecoveryRequested:\s*onOverlayPointerRecoveryRequestEvent,/,
  );
});
