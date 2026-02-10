import test from "node:test";
import assert from "node:assert/strict";
import { createCliCommandDepsRuntimeService } from "./cli-command-deps-runtime-service";

test("createCliCommandDepsRuntimeService wires runtime helpers", () => {
  let socketPath = "/tmp/mpv";
  let setClientSocketPath: string | null = null;
  let connectCalls = 0;
  let texthookerPort = 7000;
  let texthookerRunning = false;
  let texthookerStartPort: number | null = null;
  let overlayVisible = false;
  let overlayInvisible = false;
  let openYomitanAfterDelay: number | null = null;

  const deps = createCliCommandDepsRuntimeService({
    mpv: {
      getSocketPath: () => socketPath,
      setSocketPath: (next) => {
        socketPath = next;
      },
      getClient: () => ({
        setSocketPath: (next) => {
          setClientSocketPath = next;
        },
        connect: () => {
          connectCalls += 1;
        },
      }),
      showOsd: () => {},
    },
    texthooker: {
      service: {
        isRunning: () => texthookerRunning,
        start: (port) => {
          texthookerRunning = true;
          texthookerStartPort = port;
        },
      },
      getPort: () => texthookerPort,
      setPort: (port) => {
        texthookerPort = port;
      },
      shouldOpenBrowser: () => true,
      openInBrowser: () => {},
    },
    overlay: {
      isInitialized: () => false,
      initialize: () => {},
      toggleVisible: () => {
        overlayVisible = !overlayVisible;
      },
      toggleInvisible: () => {
        overlayInvisible = !overlayInvisible;
      },
      setVisible: (visible) => {
        overlayVisible = visible;
      },
      setInvisible: (visible) => {
        overlayInvisible = visible;
      },
    },
    mining: {
      copyCurrentSubtitle: () => {},
      startPendingMultiCopy: () => {},
      mineSentenceCard: async () => {},
      startPendingMineSentenceMultiple: () => {},
      updateLastCardFromClipboard: async () => {},
      triggerFieldGrouping: async () => {},
      triggerSubsyncFromConfig: async () => {},
      markLastCardAsAudioCard: async () => {},
    },
    ui: {
      openYomitanSettings: () => {},
      cycleSecondarySubMode: () => {},
      openRuntimeOptionsPalette: () => {},
      printHelp: () => {},
    },
    app: {
      stop: () => {},
      hasMainWindow: () => true,
    },
    getMultiCopyTimeoutMs: () => 2500,
    schedule: (_fn, delayMs) => {
      openYomitanAfterDelay = delayMs;
      return null;
    },
    log: () => {},
    warn: () => {},
    error: () => {},
  });

  deps.setMpvSocketPath("/tmp/new");
  deps.setMpvClientSocketPath("/tmp/new");
  deps.connectMpvClient();
  deps.ensureTexthookerRunning(9000);
  deps.openYomitanSettingsDelayed(1000);
  deps.toggleVisibleOverlay();
  deps.toggleInvisibleOverlay();

  assert.equal(deps.getMpvSocketPath(), "/tmp/new");
  assert.equal(setClientSocketPath, "/tmp/new");
  assert.equal(connectCalls, 1);
  assert.equal(texthookerStartPort, 9000);
  assert.equal(texthookerPort, 7000);
  assert.equal(openYomitanAfterDelay, 1000);
  assert.equal(overlayVisible, true);
  assert.equal(overlayInvisible, true);
  assert.equal(deps.getMultiCopyTimeoutMs(), 2500);
});
