import test from 'node:test';
import assert from 'node:assert/strict';
import {
  createCopyCurrentSubtitleHandler,
  createHandleMineSentenceDigitHandler,
  createHandleMultiCopyDigitHandler,
} from './mining-actions';

test('multi-copy digit handler forwards tracker/clipboard/osd deps', () => {
  const calls: string[] = [];
  const tracker = {};
  const handleMultiCopyDigit = createHandleMultiCopyDigitHandler({
    getSubtitleTimingTracker: () => tracker,
    writeClipboardText: (text) => calls.push(`clipboard:${text}`),
    showMpvOsd: (text) => calls.push(`osd:${text}`),
    handleMultiCopyDigitCore: (count, options) => {
      assert.equal(count, 3);
      assert.equal(options.subtitleTimingTracker, tracker);
      options.writeClipboardText('copied');
      options.showMpvOsd('done');
    },
  });

  handleMultiCopyDigit(3);
  assert.deepEqual(calls, ['clipboard:copied', 'osd:done']);
});

test('copy current subtitle handler forwards tracker/clipboard/osd deps', () => {
  const calls: string[] = [];
  const tracker = {};
  const copyCurrentSubtitle = createCopyCurrentSubtitleHandler({
    getSubtitleTimingTracker: () => tracker,
    writeClipboardText: (text) => calls.push(`clipboard:${text}`),
    showMpvOsd: (text) => calls.push(`osd:${text}`),
    copyCurrentSubtitleCore: (options) => {
      assert.equal(options.subtitleTimingTracker, tracker);
      options.writeClipboardText('subtitle');
      options.showMpvOsd('copied');
    },
  });

  copyCurrentSubtitle();
  assert.deepEqual(calls, ['clipboard:subtitle', 'osd:copied']);
});

test('mine sentence digit handler forwards all dependencies', () => {
  const calls: string[] = [];
  const tracker = {};
  const integration = {};
  const handleMineSentenceDigit = createHandleMineSentenceDigitHandler({
    getSubtitleTimingTracker: () => tracker,
    getAnkiIntegration: () => integration,
    getCurrentSecondarySubText: () => 'secondary',
    showMpvOsd: (text) => calls.push(`osd:${text}`),
    logError: (message) => calls.push(`err:${message}`),
    onCardsMined: (count) => calls.push(`cards:${count}`),
    handleMineSentenceDigitCore: (count, options) => {
      assert.equal(count, 4);
      assert.equal(options.subtitleTimingTracker, tracker);
      assert.equal(options.ankiIntegration, integration);
      assert.equal(options.getCurrentSecondarySubText(), 'secondary');
      options.showMpvOsd('mine');
      options.logError('boom', new Error('x'));
      options.onCardsMined(2);
    },
  });

  handleMineSentenceDigit(4);
  assert.deepEqual(calls, ['osd:mine', 'err:boom', 'cards:2']);
});
