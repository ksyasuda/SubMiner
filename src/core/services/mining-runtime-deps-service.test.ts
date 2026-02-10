import test from "node:test";
import assert from "node:assert/strict";
import {
  createCopyCurrentSubtitleDepsRuntimeService,
  createHandleMineSentenceDigitDepsRuntimeService,
  createHandleMultiCopyDigitDepsRuntimeService,
  createMarkLastCardAsAudioCardDepsRuntimeService,
  createMineSentenceCardDepsRuntimeService,
  createTriggerFieldGroupingDepsRuntimeService,
  createUpdateLastCardFromClipboardDepsRuntimeService,
} from "./mining-runtime-deps-service";

test("mining runtime deps builders preserve references", () => {
  const showMpvOsd = (_text: string) => {};
  const writeClipboardText = (_text: string) => {};
  const readClipboardText = () => "x";
  const logError = (_message: string, _err: unknown) => {};
  const subtitleTimingTracker = null;
  const ankiIntegration = null;
  const mpvClient = null;

  const multiCopy = createHandleMultiCopyDigitDepsRuntimeService({
    subtitleTimingTracker,
    writeClipboardText,
    showMpvOsd,
  });
  const copyCurrent = createCopyCurrentSubtitleDepsRuntimeService({
    subtitleTimingTracker,
    writeClipboardText,
    showMpvOsd,
  });
  const updateLast = createUpdateLastCardFromClipboardDepsRuntimeService({
    ankiIntegration,
    readClipboardText,
    showMpvOsd,
  });
  const fieldGrouping = createTriggerFieldGroupingDepsRuntimeService({
    ankiIntegration,
    showMpvOsd,
  });
  const markAudio = createMarkLastCardAsAudioCardDepsRuntimeService({
    ankiIntegration,
    showMpvOsd,
  });
  const mineCard = createMineSentenceCardDepsRuntimeService({
    ankiIntegration,
    mpvClient,
    showMpvOsd,
  });
  const mineDigit = createHandleMineSentenceDigitDepsRuntimeService({
    subtitleTimingTracker,
    ankiIntegration,
    getCurrentSecondarySubText: () => undefined,
    showMpvOsd,
    logError,
  });

  assert.equal(multiCopy.writeClipboardText, writeClipboardText);
  assert.equal(copyCurrent.showMpvOsd, showMpvOsd);
  assert.equal(updateLast.readClipboardText, readClipboardText);
  assert.equal(fieldGrouping.ankiIntegration, ankiIntegration);
  assert.equal(markAudio.showMpvOsd, showMpvOsd);
  assert.equal(mineCard.mpvClient, mpvClient);
  assert.equal(mineDigit.logError, logError);
});
