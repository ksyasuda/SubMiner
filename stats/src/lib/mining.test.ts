import assert from 'node:assert/strict';
import test from 'node:test';
import {
  buildStatsMineCardParams,
  getStatsMineCardError,
  getStatsMineCardUnavailableReason,
} from './mining';
import type { SentenceSearchResult } from '../types/stats';

function makeResult(overrides: Partial<SentenceSearchResult> = {}): SentenceSearchResult {
  return {
    animeId: null,
    animeTitle: 'Little Witch Academia',
    videoId: 4,
    videoTitle: 'Episode 4',
    sourcePath: '/tmp/lwa.mkv',
    secondaryText: 'Magic is gone',
    sessionId: 7,
    lineIndex: 12,
    segmentStartMs: 5_000,
    segmentEndMs: 6_000,
    text: '魔法がなくなった',
    ...overrides,
  };
}

test('buildStatsMineCardParams maps sentence result context to the shared mining payload', () => {
  assert.deepEqual(buildStatsMineCardParams(makeResult(), '魔法', 'sentence'), {
    sourcePath: '/tmp/lwa.mkv',
    startMs: 5_000,
    endMs: 6_000,
    sentence: '魔法がなくなった',
    word: '魔法',
    secondaryText: 'Magic is gone',
    videoTitle: 'Episode 4',
    mode: 'sentence',
  });
});

test('buildStatsMineCardParams returns null when media context is incomplete', () => {
  assert.equal(
    buildStatsMineCardParams(makeResult({ sourcePath: null }), '魔法', 'sentence'),
    null,
  );
  assert.equal(
    buildStatsMineCardParams(makeResult({ segmentStartMs: null }), '魔法', 'sentence'),
    null,
  );
  assert.equal(
    buildStatsMineCardParams(makeResult({ segmentEndMs: null }), '魔法', 'sentence'),
    null,
  );
});

test('buildStatsMineCardParams returns null when stored timing has no positive duration', () => {
  assert.equal(
    buildStatsMineCardParams(
      makeResult({ segmentStartMs: 5_000, segmentEndMs: 4_900 }),
      '魔法',
      'sentence',
    ),
    null,
  );
  assert.equal(
    getStatsMineCardUnavailableReason(makeResult({ segmentStartMs: 5_000, segmentEndMs: 5_000 })),
    'This line has invalid segment timing.',
  );
});

test('getStatsMineCardError surfaces partial media failures', () => {
  assert.equal(
    getStatsMineCardError({ noteId: 1, errors: ['audio: ffmpeg failed'] }),
    'audio: ffmpeg failed',
  );
  assert.equal(getStatsMineCardError({ error: 'File not found' }), 'File not found');
  assert.equal(getStatsMineCardError({ noteId: 1 }), null);
});
