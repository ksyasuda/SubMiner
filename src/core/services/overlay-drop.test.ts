import assert from 'node:assert/strict';
import test from 'node:test';
import {
  buildMpvLoadfileCommands,
  buildMpvSubtitleAddCommands,
  collectDroppedSubtitlePaths,
  collectDroppedVideoPaths,
  parseClipboardVideoPath,
  type DropDataTransferLike,
} from './overlay-drop';

function makeTransfer(data: Partial<DropDataTransferLike>): DropDataTransferLike {
  return {
    files: data.files,
    getData: data.getData,
  };
}

test('collectDroppedVideoPaths keeps supported dropped file paths in order', () => {
  const transfer = makeTransfer({
    files: [
      { path: '/videos/ep02.mkv' },
      { path: '/videos/notes.txt' },
      { path: '/videos/ep03.MP4' },
    ],
  });

  const result = collectDroppedVideoPaths(transfer);

  assert.deepEqual(result, ['/videos/ep02.mkv', '/videos/ep03.MP4']);
});

test('collectDroppedVideoPaths parses text/uri-list entries and de-duplicates', () => {
  const transfer = makeTransfer({
    getData: (format: string) =>
      format === 'text/uri-list'
        ? '#comment\nfile:///tmp/ep01.mkv\nfile:///tmp/ep01.mkv\nfile:///tmp/ep02.webm\nfile:///tmp/readme.md\n'
        : '',
  });

  const result = collectDroppedVideoPaths(transfer);

  assert.deepEqual(result, ['/tmp/ep01.mkv', '/tmp/ep02.webm']);
});

test('collectDroppedSubtitlePaths keeps supported dropped subtitle paths in order', () => {
  const transfer = makeTransfer({
    files: [{ path: '/subs/ep02.ass' }, { path: '/subs/readme.txt' }, { path: '/subs/ep03.SRT' }],
  });

  const result = collectDroppedSubtitlePaths(transfer);

  assert.deepEqual(result, ['/subs/ep02.ass', '/subs/ep03.SRT']);
});

test('collectDroppedSubtitlePaths parses text/uri-list entries and de-duplicates', () => {
  const transfer = makeTransfer({
    getData: (format: string) =>
      format === 'text/uri-list'
        ? '#comment\nfile:///tmp/ep01.ass\nfile:///tmp/ep01.ass\nfile:///tmp/ep02.vtt\nfile:///tmp/readme.md\n'
        : '',
  });

  const result = collectDroppedSubtitlePaths(transfer);

  assert.deepEqual(result, ['/tmp/ep01.ass', '/tmp/ep02.vtt']);
});

test('buildMpvLoadfileCommands replaces first file and appends remainder by default', () => {
  const commands = buildMpvLoadfileCommands(['/tmp/ep01.mkv', '/tmp/ep02.mkv'], false);

  assert.deepEqual(commands, [
    ['loadfile', '/tmp/ep01.mkv', 'replace'],
    ['loadfile', '/tmp/ep02.mkv', 'append'],
  ]);
});

test('buildMpvLoadfileCommands uses append mode when shift-drop is used', () => {
  const commands = buildMpvLoadfileCommands(['/tmp/ep01.mkv', '/tmp/ep02.mkv'], true);

  assert.deepEqual(commands, [
    ['loadfile', '/tmp/ep01.mkv', 'append'],
    ['loadfile', '/tmp/ep02.mkv', 'append'],
  ]);
});

test('buildMpvSubtitleAddCommands selects first subtitle and adds remainder', () => {
  const commands = buildMpvSubtitleAddCommands(['/tmp/ep01.ass', '/tmp/ep02.srt']);

  assert.deepEqual(commands, [
    ['sub-add', '/tmp/ep01.ass', 'select'],
    ['sub-add', '/tmp/ep02.srt'],
  ]);
});

test('parseClipboardVideoPath accepts quoted local paths', () => {
  assert.equal(parseClipboardVideoPath('"/tmp/ep10.mkv"'), '/tmp/ep10.mkv');
});

test('parseClipboardVideoPath accepts file URI and rejects non-video', () => {
  assert.equal(parseClipboardVideoPath('file:///tmp/ep11.mp4'), '/tmp/ep11.mp4');
  assert.equal(parseClipboardVideoPath('/tmp/notes.txt'), null);
});
