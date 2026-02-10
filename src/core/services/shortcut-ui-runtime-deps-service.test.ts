import test from "node:test";
import assert from "node:assert/strict";
import {
  createGlobalShortcutRegistrationDepsRuntimeService,
  createSecondarySubtitleCycleDepsRuntimeService,
  createYomitanSettingsWindowDepsRuntimeService,
  runOverlayShortcutLocalFallbackRuntimeService,
} from "./shortcut-ui-runtime-deps-service";

function makeOptions() {
  return {
    yomitanExt: null,
    getYomitanSettingsWindow: () => null,
    setYomitanSettingsWindow: () => {},

    shortcuts: {
      toggleVisibleOverlayGlobal: "Ctrl+Shift+O",
      toggleInvisibleOverlayGlobal: "Ctrl+Alt+O",
    },
    onToggleVisibleOverlay: () => {},
    onToggleInvisibleOverlay: () => {},
    onOpenYomitanSettings: () => {},
    isDev: false,
    getMainWindow: () => null,

    getSecondarySubMode: () => "hover" as const,
    setSecondarySubMode: () => {},
    getLastSecondarySubToggleAtMs: () => 0,
    setLastSecondarySubToggleAtMs: () => {},
    broadcastSecondarySubMode: () => {},
    showMpvOsd: () => {},

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

test("shortcut ui deps builders return expected adapters", () => {
  const options = makeOptions();
  const yomitan = createYomitanSettingsWindowDepsRuntimeService(options);
  const globalShortcuts = createGlobalShortcutRegistrationDepsRuntimeService(options);
  const secondary = createSecondarySubtitleCycleDepsRuntimeService(options);

  assert.equal(yomitan.yomitanExt, null);
  assert.equal(typeof globalShortcuts.onOpenYomitanSettings, "function");
  assert.equal(secondary.getSecondarySubMode(), "hover");
});

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
