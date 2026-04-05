import assert from 'node:assert/strict';
import test from 'node:test';

import {
  createManagedLocalSubtitleSelectionRuntime,
  resolveManagedLocalSubtitleSelection,
} from './local-subtitle-selection';

const mixedLanguageTrackList = [
  { type: 'sub', id: 1, lang: 'pt', title: '[Infinite]', external: false, selected: true },
  { type: 'sub', id: 2, lang: 'pt', title: '[Moshi Moshi]', external: false },
  { type: 'sub', id: 3, lang: 'en', title: '(Vivid)', external: false },
  { type: 'sub', id: 9, lang: 'en', title: 'English(US)', external: false },
  { type: 'sub', id: 11, lang: 'en', title: 'en.srt', external: true },
  { type: 'sub', id: 12, lang: 'ja', title: 'ja.srt', external: true },
];

const unlabeledExternalSidecarTrackList = [
  { type: 'sub', id: 1, lang: 'eng', title: 'English ASS', external: false, selected: true },
  { type: 'sub', id: 2, title: 'srt', external: true },
];

test('resolveManagedLocalSubtitleSelection prefers default Japanese primary and English secondary tracks', () => {
  const result = resolveManagedLocalSubtitleSelection({
    trackList: mixedLanguageTrackList,
    primaryLanguages: [],
    secondaryLanguages: [],
  });

  assert.equal(result.primaryTrackId, 12);
  assert.equal(result.secondaryTrackId, 11);
});

test('resolveManagedLocalSubtitleSelection respects configured language overrides', () => {
  const result = resolveManagedLocalSubtitleSelection({
    trackList: mixedLanguageTrackList,
    primaryLanguages: ['pt'],
    secondaryLanguages: ['ja'],
  });

  assert.equal(result.primaryTrackId, 1);
  assert.equal(result.secondaryTrackId, 12);
});

test('resolveManagedLocalSubtitleSelection promotes a single unlabeled external sidecar to primary', () => {
  const result = resolveManagedLocalSubtitleSelection({
    trackList: unlabeledExternalSidecarTrackList,
    primaryLanguages: [],
    secondaryLanguages: [],
  });

  assert.equal(result.primaryTrackId, 2);
  assert.equal(result.secondaryTrackId, 1);
});

test('resolveManagedLocalSubtitleSelection does not guess between multiple unlabeled external sidecars', () => {
  const result = resolveManagedLocalSubtitleSelection({
    trackList: [
      ...unlabeledExternalSidecarTrackList,
      { type: 'sub', id: 3, title: 'subrip', external: true },
    ],
    primaryLanguages: [],
    secondaryLanguages: [],
  });

  assert.equal(result.primaryTrackId, null);
  assert.equal(result.secondaryTrackId, 1);
});

test('managed local subtitle selection runtime applies preferred tracks once for a local media path', async () => {
  const commands: Array<Array<string | number>> = [];
  const scheduled: Array<() => void> = [];

  const runtime = createManagedLocalSubtitleSelectionRuntime({
    getCurrentMediaPath: () => '/videos/example.mkv',
    getMpvClient: () =>
      ({
        connected: true,
        requestProperty: async (name: string) => {
          if (name === 'track-list') {
            return mixedLanguageTrackList;
          }
          throw new Error(`Unexpected property: ${name}`);
        },
      }) as never,
    getPrimarySubtitleLanguages: () => [],
    getSecondarySubtitleLanguages: () => [],
    sendMpvCommand: (command) => {
      commands.push(command);
    },
    schedule: (callback) => {
      scheduled.push(callback);
      return 1 as never;
    },
    clearScheduled: () => {},
  });

  runtime.handleMediaPathChange('/videos/example.mkv');
  scheduled.shift()?.();
  await new Promise((resolve) => setTimeout(resolve, 0));
  runtime.handleSubtitleTrackListChange(mixedLanguageTrackList);

  assert.deepEqual(commands, [
    ['set_property', 'sid', 12],
    ['set_property', 'secondary-sid', 11],
  ]);
});

test('managed local subtitle selection runtime promotes a single unlabeled external sidecar over embedded english', async () => {
  const commands: Array<Array<string | number>> = [];
  const scheduled: Array<() => void> = [];

  const runtime = createManagedLocalSubtitleSelectionRuntime({
    getCurrentMediaPath: () => '/videos/example.mkv',
    getMpvClient: () =>
      ({
        connected: true,
        requestProperty: async (name: string) => {
          if (name === 'track-list') {
            return unlabeledExternalSidecarTrackList;
          }
          throw new Error(`Unexpected property: ${name}`);
        },
      }) as never,
    getPrimarySubtitleLanguages: () => [],
    getSecondarySubtitleLanguages: () => [],
    sendMpvCommand: (command) => {
      commands.push(command);
    },
    schedule: (callback) => {
      scheduled.push(callback);
      return 1 as never;
    },
    clearScheduled: () => {},
  });

  runtime.handleMediaPathChange('/videos/example.mkv');
  scheduled.shift()?.();
  await new Promise((resolve) => setTimeout(resolve, 0));
  runtime.handleSubtitleTrackListChange(unlabeledExternalSidecarTrackList);

  assert.deepEqual(commands, [
    ['set_property', 'sid', 2],
    ['set_property', 'secondary-sid', 1],
  ]);
});
