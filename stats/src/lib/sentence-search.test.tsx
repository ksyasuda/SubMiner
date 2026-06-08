import assert from 'node:assert/strict';
import test from 'node:test';
import { renderToStaticMarkup } from 'react-dom/server';
import {
  findExactSentenceMatches,
  getSentenceSearchMineAvailability,
  renderSentenceWithMatches,
} from './sentence-search';
import type { SentenceSearchResult } from '../types/stats';

function makeResult(over: Partial<SentenceSearchResult>): SentenceSearchResult {
  return {
    animeId: null,
    animeTitle: null,
    videoId: 1,
    videoTitle: 'Episode 1',
    sourcePath: '/tmp/video.mkv',
    secondaryText: null,
    sessionId: 10,
    lineIndex: 3,
    segmentStartMs: 1000,
    segmentEndMs: 2500,
    text: '猫が猫を見た',
    ...over,
  };
}

test('findExactSentenceMatches returns every exact searched-word range', () => {
  assert.deepEqual(findExactSentenceMatches('猫が猫を見た', '猫'), [
    { start: 0, end: 1 },
    { start: 2, end: 3 },
  ]);
});

test('findExactSentenceMatches keeps source-text ranges under case folding', () => {
  assert.deepEqual(findExactSentenceMatches('İstanbul', 'İ'), [{ start: 0, end: 1 }]);
});

test('getSentenceSearchMineAvailability gates word and audio mining on exact sentence match', () => {
  const result = makeResult({});

  assert.deepEqual(getSentenceSearchMineAvailability(result, '猫'), {
    canMineSentence: true,
    canMineWordAudio: true,
    exactMatch: true,
    unavailableReason: null,
  });

  assert.deepEqual(getSentenceSearchMineAvailability(result, '犬'), {
    canMineSentence: true,
    canMineWordAudio: false,
    exactMatch: false,
    unavailableReason: null,
  });
});

test('getSentenceSearchMineAvailability disables every mining mode without source timing', () => {
  const result = makeResult({ sourcePath: null, segmentEndMs: null });

  assert.deepEqual(getSentenceSearchMineAvailability(result, '猫'), {
    canMineSentence: false,
    canMineWordAudio: false,
    exactMatch: true,
    unavailableReason: 'This source has no local file path.',
  });
});

test('getSentenceSearchMineAvailability disables every mining mode with invalid source timing', () => {
  const result = makeResult({ segmentStartMs: 2500, segmentEndMs: 2400 });

  assert.deepEqual(getSentenceSearchMineAvailability(result, '猫'), {
    canMineSentence: false,
    canMineWordAudio: false,
    exactMatch: true,
    unavailableReason: 'This line has invalid segment timing.',
  });
});

test('renderSentenceWithMatches highlights exact searched-word matches', () => {
  const markup = renderToStaticMarkup(<>{renderSentenceWithMatches('猫が寝る', '猫')}</>);

  assert.match(markup, /<mark/);
  assert.match(markup, />猫<\/mark>/);
});
