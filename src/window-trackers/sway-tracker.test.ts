import test from 'node:test';
import assert from 'node:assert/strict';
import { SwayWindowTracker } from './sway-tracker';

test('SwayWindowTracker tracks the mpv window bound to the target socket', async () => {
  const tree = {
    nodes: [
      { pid: 10, app_id: 'mpv', rect: { x: 0, y: 0, width: 800, height: 600 } },
      { pid: 20, app_id: 'mpv', rect: { x: 100, y: 50, width: 1280, height: 720 } },
    ],
  };
  const commandLines = new Map([
    ['10', 'mpv --input-ipc-server=/tmp/other.sock a.mkv'],
    ['20', 'mpv --input-ipc-server=/tmp/subminer.sock b.mkv'],
  ]);
  const tracker = new SwayWindowTracker('/tmp/subminer.sock', async (command, args) =>
    command === 'swaymsg' ? JSON.stringify(tree) : (commandLines.get(args[1]!) ?? ''),
  );

  await tracker.refreshNow();

  assert.deepEqual(tracker.getGeometry(), { x: 100, y: 50, width: 1280, height: 720 });
});
