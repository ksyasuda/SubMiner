import assert from 'node:assert/strict';
import test from 'node:test';
import type { SubtitleData } from '../../types';
import {
  primeVisibleOverlaySubtitleFromMpv,
  resolveCurrentSubtitleForRenderer,
} from './current-subtitle-snapshot';

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

test('renderer current subtitle snapshot can skip cold tokenizer for first paint', async () => {
  let tokenizerCalled = false;
  const payload = await resolveCurrentSubtitleForRenderer({
    currentSubText: 'まだキャッシュされていない字幕',
    currentSubtitleData: null,
    withCurrentSubtitleTiming: withTiming,
    tokenizeUncached: false,
    tokenizeSubtitle: async (text) => {
      tokenizerCalled = true;
      return { text, tokens: [{ text: 'ま' } as never] };
    },
  });

  assert.equal(tokenizerCalled, false);
  assert.equal(payload.text, 'まだキャッシュされていない字幕');
  assert.equal(payload.startTime, 1);
  assert.equal(payload.tokens, null);
});

test('renderer current subtitle snapshot reports resolved payload for startup readiness', async () => {
  const resolvedPayloads: SubtitleData[] = [];
  const payload = await resolveCurrentSubtitleForRenderer({
    currentSubText: '起動字幕',
    currentSubtitleData: null,
    withCurrentSubtitleTiming: withTiming,
    tokenizeSubtitle: async (text) => ({ text, tokens: [{ text: '起' } as never] }),
    onResolvedSubtitle: (resolved) => {
      resolvedPayloads.push(resolved);
    },
  });

  assert.deepEqual(resolvedPayloads, [payload]);
});

test('visible overlay subtitle prime refreshes current text from mpv before showing overlay', async () => {
  const calls: string[] = [];

  await primeVisibleOverlaySubtitleFromMpv({
    getMpvClient: () => ({
      connected: true,
      requestProperty: async (name) => {
        calls.push(`request:${name}`);
        return '国内外から';
      },
    }),
    setCurrentSubText: (text) => calls.push(`set:${text}`),
    getCurrentSubtitleData: () => null,
    consumeCachedSubtitle: () => null,
    onSubtitleChange: (text) => calls.push(`change:${text}`),
    refreshCurrentSubtitle: (text) => calls.push(`refresh:${text}`),
    emitSubtitle: (payload) => calls.push(`emit:${payload.text}`),
  });

  assert.deepEqual(calls, ['request:sub-text', 'set:国内外から', 'refresh:国内外から']);
});

test('visible overlay subtitle prime can defer uncached tokenization until after first paint', async () => {
  const calls: string[] = [];

  await primeVisibleOverlaySubtitleFromMpv({
    getMpvClient: () => ({
      connected: true,
      requestProperty: async (name) => {
        calls.push(`request:${name}`);
        return '国内外から';
      },
    }),
    setCurrentSubText: (text) => calls.push(`set:${text}`),
    getCurrentSubtitleData: () => null,
    consumeCachedSubtitle: () => null,
    onSubtitleChange: (text) => calls.push(`change:${text}`),
    refreshCurrentSubtitle: (text) => calls.push(`refresh:${text}`),
    emitSubtitle: (payload) => calls.push(`emit:${payload.text}`),
    deferUncachedRefresh: true,
  });

  assert.deepEqual(calls, ['request:sub-text', 'set:国内外から']);
});

test('visible overlay subtitle prime repaints cached current subtitle immediately', async () => {
  const calls: string[] = [];
  const cachedPayload: SubtitleData = { text: '字幕', tokens: [{ text: '字' } as never] };

  await primeVisibleOverlaySubtitleFromMpv({
    getMpvClient: () => ({
      connected: true,
      requestProperty: async () => '字幕',
    }),
    setCurrentSubText: (text) => calls.push(`set:${text}`),
    getCurrentSubtitleData: () => cachedPayload,
    consumeCachedSubtitle: () => null,
    onSubtitleChange: (text) => calls.push(`change:${text}`),
    refreshCurrentSubtitle: (text) => calls.push(`refresh:${text}`),
    emitSubtitle: (payload) => calls.push(`emit:${payload.text}:${payload.tokens?.length ?? 0}`),
  });

  assert.deepEqual(calls, ['set:字幕', 'emit:字幕:1', 'refresh:字幕']);
});

test('visible overlay subtitle prime clears stale subtitle when mpv has no current text', async () => {
  const calls: string[] = [];

  await primeVisibleOverlaySubtitleFromMpv({
    getMpvClient: () => ({
      connected: true,
      requestProperty: async () => '',
    }),
    setCurrentSubText: (text) => calls.push(`set:${text}`),
    getCurrentSubtitleData: () => ({ text: 'old', tokens: null }),
    consumeCachedSubtitle: () => null,
    onSubtitleChange: (text) => calls.push(`change:${text}`),
    refreshCurrentSubtitle: (text) => calls.push(`refresh:${text}`),
    emitSubtitle: (payload) => calls.push(`emit:${payload.text}:${payload.tokens}`),
  });

  assert.deepEqual(calls, ['set:', 'change:', 'emit::null']);
});

test('visible overlay subtitle prime refreshes secondary subtitle when available', async () => {
  const calls: string[] = [];

  await primeVisibleOverlaySubtitleFromMpv({
    getMpvClient: () => ({
      connected: true,
      requestProperty: async (name) => {
        calls.push(`request:${name}`);
        return name === 'secondary-sub-text' ? 'from abroad' : '国内外から';
      },
    }),
    setCurrentSubText: (text) => calls.push(`set:${text}`),
    getCurrentSubtitleData: () => null,
    consumeCachedSubtitle: () => null,
    onSubtitleChange: (text) => calls.push(`change:${text}`),
    refreshCurrentSubtitle: (text) => calls.push(`refresh:${text}`),
    emitSubtitle: (payload) => calls.push(`emit:${payload.text}`),
    setCurrentSecondarySubText: (text) => calls.push(`set-secondary:${text}`),
    emitSecondarySubtitle: (text) => calls.push(`emit-secondary:${text}`),
  });

  assert.deepEqual(calls, [
    'request:sub-text',
    'set:国内外から',
    'refresh:国内外から',
    'request:secondary-sub-text',
    'set-secondary:from abroad',
    'emit-secondary:from abroad',
  ]);
});
