import assert from 'node:assert/strict';
import test from 'node:test';
import {
  createBuildCopyCurrentSubtitleMainDepsHandler,
  createBuildHandleMineSentenceDigitMainDepsHandler,
  createBuildHandleMultiCopyDigitMainDepsHandler,
} from './mining-actions-main-deps';

test('mining action main deps builders map callbacks', () => {
  const calls: string[] = [];

  const multiCopy = createBuildHandleMultiCopyDigitMainDepsHandler({
    getSubtitleTimingTracker: () => ({ track: true }),
    writeClipboardText: (text) => calls.push(`clip:${text}`),
    showMpvOsd: (text) => calls.push(`osd:${text}`),
    handleMultiCopyDigitCore: () => calls.push('multi-copy'),
  })();
  assert.deepEqual(multiCopy.getSubtitleTimingTracker(), { track: true });
  multiCopy.writeClipboardText('x');
  multiCopy.showMpvOsd('y');
  multiCopy.handleMultiCopyDigitCore(2, {
    subtitleTimingTracker: { track: true },
    writeClipboardText: () => {},
    showMpvOsd: () => {},
  });

  const copyCurrent = createBuildCopyCurrentSubtitleMainDepsHandler({
    getSubtitleTimingTracker: () => ({ track: true }),
    writeClipboardText: (text) => calls.push(`copy:${text}`),
    showMpvOsd: (text) => calls.push(`copy-osd:${text}`),
    copyCurrentSubtitleCore: () => calls.push('copy-current'),
  })();
  assert.deepEqual(copyCurrent.getSubtitleTimingTracker(), { track: true });
  copyCurrent.writeClipboardText('a');
  copyCurrent.showMpvOsd('b');
  copyCurrent.copyCurrentSubtitleCore({
    subtitleTimingTracker: { track: true },
    writeClipboardText: () => {},
    showMpvOsd: () => {},
  });

  const mineDigit = createBuildHandleMineSentenceDigitMainDepsHandler({
    getSubtitleTimingTracker: () => ({ track: true }),
    getAnkiIntegration: () => ({ enabled: true }),
    getCurrentSecondarySubText: () => 'sub',
    showMpvOsd: (text) => calls.push(`mine-osd:${text}`),
    logError: (message) => calls.push(`err:${message}`),
    onCardsMined: (count) => calls.push(`cards:${count}`),
    handleMineSentenceDigitCore: () => calls.push('mine-digit'),
  })();
  assert.equal(mineDigit.getCurrentSecondarySubText(), 'sub');
  mineDigit.showMpvOsd('done');
  mineDigit.logError('bad', null);
  mineDigit.onCardsMined(2);
  mineDigit.handleMineSentenceDigitCore(2, {
    subtitleTimingTracker: { track: true },
    ankiIntegration: { enabled: true },
    getCurrentSecondarySubText: () => 'sub',
    showMpvOsd: () => {},
    logError: () => {},
    onCardsMined: () => {},
  });

  assert.deepEqual(calls, [
    'clip:x',
    'osd:y',
    'multi-copy',
    'copy:a',
    'copy-osd:b',
    'copy-current',
    'mine-osd:done',
    'err:bad',
    'cards:2',
    'mine-digit',
  ]);
});
