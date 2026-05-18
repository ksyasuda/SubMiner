import assert from 'node:assert/strict';
import test from 'node:test';
import type { SubtitleData } from '../../types';
import { resolveCurrentSubtitleForRenderer } from './current-subtitle-snapshot';

function withTiming(payload: SubtitleData): SubtitleData {
  return {
    ...payload,
    startTime: 1,
    endTime: 2,
  };
}

test('renderer current subtitle snapshot reuses cached payload for first paint', async () => {
  const payload = await resolveCurrentSubtitleForRenderer({
    currentSubText: '字幕',
    currentSubtitleData: { text: '字幕', tokens: [{ text: '字' } as never] },
    withCurrentSubtitleTiming: withTiming,
  });

  assert.equal(payload.text, '字幕');
  assert.equal(payload.startTime, 1);
  assert.deepEqual(payload.tokens, [{ text: '字' }]);
});

test('renderer current subtitle snapshot does not block on tokenizer for empty text', async () => {
  const payload = await resolveCurrentSubtitleForRenderer({
    currentSubText: '',
    currentSubtitleData: null,
    withCurrentSubtitleTiming: withTiming,
  });

  assert.equal(payload.text, '');
  assert.equal(payload.tokens, null);
});

test('renderer current subtitle snapshot falls back to raw text for uncached subtitles', async () => {
  const payload = await resolveCurrentSubtitleForRenderer({
    currentSubText: 'まだキャッシュされていない字幕',
    currentSubtitleData: null,
    withCurrentSubtitleTiming: withTiming,
  });

  assert.equal(payload.text, 'まだキャッシュされていない字幕');
  assert.equal(payload.startTime, 1);
  assert.equal(payload.tokens, null);
});

test('renderer current subtitle snapshot tokenizes uncached subtitles when tokenizer is available', async () => {
  const payload = await resolveCurrentSubtitleForRenderer({
    currentSubText: '新しい字幕',
    currentSubtitleData: null,
    withCurrentSubtitleTiming: withTiming,
    tokenizeSubtitle: async (text) => ({ text, tokens: [{ text: '新' } as never] }),
  });

  assert.equal(payload.text, '新しい字幕');
  assert.equal(payload.startTime, 1);
  assert.deepEqual(payload.tokens, [{ text: '新' }]);
});
