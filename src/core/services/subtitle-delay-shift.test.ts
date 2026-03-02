import assert from 'node:assert/strict';
import test from 'node:test';
import { createShiftSubtitleDelayToAdjacentCueHandler } from './subtitle-delay-shift';

function createMpvClient(props: Record<string, unknown>) {
  return {
    connected: true,
    requestProperty: async (name: string) => props[name],
  };
}

test('shift subtitle delay to next cue using active external srt track', async () => {
  const commands: Array<Array<string | number>> = [];
  const osd: string[] = [];
  let loadCount = 0;
  const handler = createShiftSubtitleDelayToAdjacentCueHandler({
    getMpvClient: () =>
      createMpvClient({
        'track-list': [
          {
            type: 'sub',
            id: 2,
            external: true,
            'external-filename': '/tmp/subs.srt',
          },
        ],
        sid: 2,
        'sub-start': 3.0,
      }),
    loadSubtitleSourceText: async () => {
      loadCount += 1;
      return `1
00:00:01,000 --> 00:00:02,000
line-1

2
00:00:03,000 --> 00:00:04,000
line-2

3
00:00:05,000 --> 00:00:06,000
line-3`;
    },
    sendMpvCommand: (command) => commands.push(command),
    showMpvOsd: (text) => osd.push(text),
  });

  await handler('next');
  await handler('next');

  assert.equal(loadCount, 1);
  assert.equal(commands.length, 2);
  const delta = commands[0]?.[2];
  assert.equal(commands[0]?.[0], 'add');
  assert.equal(commands[0]?.[1], 'sub-delay');
  assert.equal(typeof delta, 'number');
  assert.equal(Math.abs((delta as number) - 2) < 0.0001, true);
  assert.deepEqual(osd, ['Subtitle delay: ${sub-delay}', 'Subtitle delay: ${sub-delay}']);
});

test('shift subtitle delay to previous cue using active external ass track', async () => {
  const commands: Array<Array<string | number>> = [];
  const handler = createShiftSubtitleDelayToAdjacentCueHandler({
    getMpvClient: () =>
      createMpvClient({
        'track-list': [
          {
            type: 'sub',
            id: 4,
            external: true,
            'external-filename': '/tmp/subs.ass',
          },
        ],
        sid: 4,
        'sub-start': 2.0,
      }),
    loadSubtitleSourceText: async () => `[Events]
Dialogue: 0,0:00:00.50,0:00:01.50,Default,,0,0,0,,line-1
Dialogue: 0,0:00:02.00,0:00:03.00,Default,,0,0,0,,line-2
Dialogue: 0,0:00:04.00,0:00:05.00,Default,,0,0,0,,line-3`,
    sendMpvCommand: (command) => commands.push(command),
    showMpvOsd: () => {},
  });

  await handler('previous');

  const delta = commands[0]?.[2];
  assert.equal(typeof delta, 'number');
  assert.equal(Math.abs((delta as number) + 1.5) < 0.0001, true);
});

test('shift subtitle delay throws when no next cue exists', async () => {
  const handler = createShiftSubtitleDelayToAdjacentCueHandler({
    getMpvClient: () =>
      createMpvClient({
        'track-list': [
          {
            type: 'sub',
            id: 1,
            external: true,
            'external-filename': '/tmp/subs.vtt',
          },
        ],
        sid: 1,
        'sub-start': 5.0,
      }),
    loadSubtitleSourceText: async () => `WEBVTT

00:00:01.000 --> 00:00:02.000
line-1

00:00:03.000 --> 00:00:04.000
line-2

00:00:05.000 --> 00:00:06.000
line-3`,
    sendMpvCommand: () => {},
    showMpvOsd: () => {},
  });

  await assert.rejects(() => handler('next'), /No next subtitle cue found/);
});
