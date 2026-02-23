import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';
import { appendClipboardVideoToQueueRuntime } from './clipboard-queue';

test('appendClipboardVideoToQueueRuntime returns disconnected when mpv unavailable', () => {
  const result = appendClipboardVideoToQueueRuntime({
    getMpvClient: () => null,
    readClipboardText: () => '',
    showMpvOsd: () => {},
    sendMpvCommand: () => {},
  });
  assert.deepEqual(result, { ok: false, message: 'MPV is not connected.' });
});

test('appendClipboardVideoToQueueRuntime rejects unsupported clipboard path', () => {
  const osdMessages: string[] = [];
  const result = appendClipboardVideoToQueueRuntime({
    getMpvClient: () => ({ connected: true }),
    readClipboardText: () => 'not a media path',
    showMpvOsd: (text) => osdMessages.push(text),
    sendMpvCommand: () => {},
  });
  assert.equal(result.ok, false);
  assert.equal(osdMessages[0], 'Clipboard does not contain a supported video path.');
});

test('appendClipboardVideoToQueueRuntime queues readable media file', () => {
  const tempPath = path.join(process.cwd(), 'dist', 'clipboard-queue-test-video.mkv');
  fs.writeFileSync(tempPath, 'stub');

  const commands: Array<(string | number)[]> = [];
  const osdMessages: string[] = [];
  const result = appendClipboardVideoToQueueRuntime({
    getMpvClient: () => ({ connected: true }),
    readClipboardText: () => tempPath,
    showMpvOsd: (text) => osdMessages.push(text),
    sendMpvCommand: (command) => commands.push(command),
  });

  assert.equal(result.ok, true);
  assert.deepEqual(commands[0], ['loadfile', tempPath, 'append']);
  assert.equal(osdMessages[0], `Queued from clipboard: ${path.basename(tempPath)}`);

  fs.unlinkSync(tempPath);
});
