import test from "node:test";
import assert from "node:assert/strict";
import { CliArgs } from "../../cli/args";
import { CliCommandServiceDeps, handleCliCommandService } from "./cli-command-service";

function makeArgs(overrides: Partial<CliArgs> = {}): CliArgs {
  return {
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
    ...overrides,
  };
}

function createDeps(overrides: Partial<CliCommandServiceDeps> = {}) {
  const calls: string[] = [];
  let mpvSocketPath = "/tmp/subminer.sock";
  let texthookerPort = 5174;
  const osd: string[] = [];

  const deps: CliCommandServiceDeps = {
    getMpvSocketPath: () => mpvSocketPath,
    setMpvSocketPath: (socketPath) => {
      mpvSocketPath = socketPath;
      calls.push(`setMpvSocketPath:${socketPath}`);
    },
    setMpvClientSocketPath: (socketPath) => {
      calls.push(`setMpvClientSocketPath:${socketPath}`);
    },
    hasMpvClient: () => true,
    connectMpvClient: () => {
      calls.push("connectMpvClient");
    },
    isTexthookerRunning: () => false,
    setTexthookerPort: (port) => {
      texthookerPort = port;
      calls.push(`setTexthookerPort:${port}`);
    },
    getTexthookerPort: () => texthookerPort,
    shouldOpenTexthookerBrowser: () => true,
    ensureTexthookerRunning: (port) => {
      calls.push(`ensureTexthookerRunning:${port}`);
    },
    openTexthookerInBrowser: (url) => {
      calls.push(`openTexthookerInBrowser:${url}`);
    },
    stopApp: () => {
      calls.push("stopApp");
    },
    isOverlayRuntimeInitialized: () => false,
    initializeOverlayRuntime: () => {
      calls.push("initializeOverlayRuntime");
    },
    toggleVisibleOverlay: () => {
      calls.push("toggleVisibleOverlay");
    },
    toggleInvisibleOverlay: () => {
      calls.push("toggleInvisibleOverlay");
    },
    openYomitanSettingsDelayed: (delayMs) => {
      calls.push(`openYomitanSettingsDelayed:${delayMs}`);
    },
    setVisibleOverlayVisible: (visible) => {
      calls.push(`setVisibleOverlayVisible:${visible}`);
    },
    setInvisibleOverlayVisible: (visible) => {
      calls.push(`setInvisibleOverlayVisible:${visible}`);
    },
    copyCurrentSubtitle: () => {
      calls.push("copyCurrentSubtitle");
    },
    startPendingMultiCopy: (timeoutMs) => {
      calls.push(`startPendingMultiCopy:${timeoutMs}`);
    },
    mineSentenceCard: async () => {
      calls.push("mineSentenceCard");
    },
    startPendingMineSentenceMultiple: (timeoutMs) => {
      calls.push(`startPendingMineSentenceMultiple:${timeoutMs}`);
    },
    updateLastCardFromClipboard: async () => {
      calls.push("updateLastCardFromClipboard");
    },
    cycleSecondarySubMode: () => {
      calls.push("cycleSecondarySubMode");
    },
    triggerFieldGrouping: async () => {
      calls.push("triggerFieldGrouping");
    },
    triggerSubsyncFromConfig: async () => {
      calls.push("triggerSubsyncFromConfig");
    },
    markLastCardAsAudioCard: async () => {
      calls.push("markLastCardAsAudioCard");
    },
    openRuntimeOptionsPalette: () => {
      calls.push("openRuntimeOptionsPalette");
    },
    printHelp: () => {
      calls.push("printHelp");
    },
    hasMainWindow: () => true,
    getMultiCopyTimeoutMs: () => 2500,
    showMpvOsd: (text) => {
      osd.push(text);
    },
    log: (message) => {
      calls.push(`log:${message}`);
    },
    warn: (message) => {
      calls.push(`warn:${message}`);
    },
    error: (message) => {
      calls.push(`error:${message}`);
    },
    ...overrides,
  };

  return { deps, calls, osd };
}

test("handleCliCommandService ignores --start for second-instance without actions", () => {
  const { deps, calls } = createDeps();
  const args = makeArgs({ start: true });

  handleCliCommandService(args, "second-instance", deps);

  assert.ok(calls.includes("log:Ignoring --start because SubMiner is already running."));
  assert.equal(calls.some((value) => value.includes("connectMpvClient")), false);
});

test("handleCliCommandService runs texthooker flow with browser open", () => {
  const { deps, calls } = createDeps();
  const args = makeArgs({ texthooker: true });

  handleCliCommandService(args, "initial", deps);

  assert.ok(calls.includes("ensureTexthookerRunning:5174"));
  assert.ok(
    calls.includes("openTexthookerInBrowser:http://127.0.0.1:5174"),
  );
});

test("handleCliCommandService reports async mine errors to OSD", async () => {
  const { deps, calls, osd } = createDeps({
    mineSentenceCard: async () => {
      throw new Error("boom");
    },
  });

  handleCliCommandService(makeArgs({ mineSentence: true }), "initial", deps);
  await new Promise((resolve) => setImmediate(resolve));

  assert.ok(calls.some((value) => value.startsWith("error:mineSentenceCard failed:")));
  assert.ok(osd.some((value) => value.includes("Mine sentence failed: boom")));
});
