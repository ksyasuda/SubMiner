import test from "node:test";
import assert from "node:assert/strict";
import { createAppLifecycleDepsRuntimeService } from "./app-lifecycle-deps-runtime-service";

test("createAppLifecycleDepsRuntimeService maps app methods and platform", async () => {
  const events: Record<string, (...args: unknown[]) => void> = {};
  let lockRequested = 0;
  let quitCalled = 0;
  const deps = createAppLifecycleDepsRuntimeService({
    app: {
      requestSingleInstanceLock: () => {
        lockRequested += 1;
        return true;
      },
      quit: () => {
        quitCalled += 1;
      },
      on: (event, listener) => {
        events[event] = listener;
      },
      whenReady: async () => {},
    },
    platform: "darwin",
    shouldStartApp: () => true,
    parseArgs: () => ({
      start: false,
      stop: false,
      toggle: false,
      toggleVisibleOverlay: false,
      toggleInvisibleOverlay: false,
      settings: false,
      show: false,
      hide: false,
      showVisibleOverlay: false,
      hideVisibleOverlay: false,
      showInvisibleOverlay: false,
      hideInvisibleOverlay: false,
      copySubtitle: false,
      copySubtitleMultiple: false,
      mineSentence: false,
      mineSentenceMultiple: false,
      updateLastCardFromClipboard: false,
      toggleSecondarySub: false,
      triggerFieldGrouping: false,
      triggerSubsync: false,
      markAudioCard: false,
      openRuntimeOptions: false,
      texthooker: false,
      help: false,
      autoStartOverlay: false,
      generateConfig: false,
      backupOverwrite: false,
      verbose: false,
    }),
    handleCliCommand: () => {},
    printHelp: () => {},
    logNoRunningInstance: () => {},
    onReady: async () => {},
    onWillQuitCleanup: () => {},
    shouldRestoreWindowsOnActivate: () => false,
    restoreWindowsOnActivate: () => {},
  });

  assert.equal(deps.requestSingleInstanceLock(), true);
  deps.quitApp();
  assert.equal(lockRequested, 1);
  assert.equal(quitCalled, 1);
  assert.equal(deps.isDarwinPlatform(), true);

  let callbackRan = false;
  deps.whenReady(async () => {
    callbackRan = true;
  });
  await new Promise((resolve) => setImmediate(resolve));
  assert.equal(callbackRan, true);

  deps.onActivate(() => {});
  assert.equal(typeof events["activate"], "function");
});
