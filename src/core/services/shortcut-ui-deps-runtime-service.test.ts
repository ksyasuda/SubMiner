import test from "node:test";
import assert from "node:assert/strict";
import {
  runOverlayShortcutLocalFallbackRuntimeService,
} from "./shortcut-ui-deps-runtime-service";

function makeOptions() {
  return {
    getConfiguredShortcuts: () => ({
      toggleVisibleOverlayGlobal: null,
      toggleInvisibleOverlayGlobal: null,
      copySubtitle: null,
      copySubtitleMultiple: null,
      updateLastCardFromClipboard: null,
      triggerFieldGrouping: null,
      triggerSubsync: null,
      mineSentence: null,
      mineSentenceMultiple: null,
      multiCopyTimeoutMs: 5000,
      toggleSecondarySub: null,
      markAudioCard: null,
      openRuntimeOptions: "Ctrl+R",
      openJimaku: null,
    }),
    getOverlayShortcutFallbackHandlers: () => ({
      openRuntimeOptions: () => {},
      openJimaku: () => {},
      markAudioCard: () => {},
      copySubtitleMultiple: () => {},
      copySubtitle: () => {},
      toggleSecondarySub: () => {},
      updateLastCardFromClipboard: () => {},
      triggerFieldGrouping: () => {},
      triggerSubsync: () => {},
      mineSentence: () => {},
      mineSentenceMultiple: () => {},
    }),
    shortcutMatcher: () => false,
  };
}

test("runOverlayShortcutLocalFallbackRuntimeService delegates and returns boolean", () => {
  const options = {
    ...makeOptions(),
    shortcutMatcher: () => true,
  };

  const handled = runOverlayShortcutLocalFallbackRuntimeService(
    {
      key: "r",
      code: "KeyR",
      alt: false,
      control: true,
      shift: false,
      meta: false,
      type: "keyDown",
    } as Electron.Input,
    options,
  );

  assert.equal(handled, true);
});
