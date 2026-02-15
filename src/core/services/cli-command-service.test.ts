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
    refreshKnownWords: false,
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
    refreshKnownWords: async () => {
      calls.push("refreshKnownWords");
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

test("handleCliCommandService applies socket path and connects on start", () => {
  const { deps, calls } = createDeps();

  handleCliCommandService(
    makeArgs({ start: true, socketPath: "/tmp/custom.sock" }),
    "initial",
    deps,
  );

  assert.ok(calls.includes("setMpvSocketPath:/tmp/custom.sock"));
  assert.ok(calls.includes("setMpvClientSocketPath:/tmp/custom.sock"));
  assert.ok(calls.includes("connectMpvClient"));
});

test("handleCliCommandService warns when texthooker port override used while running", () => {
  const { deps, calls } = createDeps({
    isTexthookerRunning: () => true,
  });

  handleCliCommandService(
    makeArgs({ texthookerPort: 9999, texthooker: true }),
    "initial",
    deps,
  );

  assert.ok(
    calls.includes(
      "warn:Ignoring --port override because the texthooker server is already running.",
    ),
  );
  assert.equal(calls.some((value) => value === "setTexthookerPort:9999"), false);
});

test("handleCliCommandService prints help and stops app when no window exists", () => {
  const { deps, calls } = createDeps({
    hasMainWindow: () => false,
  });

  handleCliCommandService(makeArgs({ help: true }), "initial", deps);

  assert.ok(calls.includes("printHelp"));
  assert.ok(calls.includes("stopApp"));
});

test("handleCliCommandService reports async trigger-subsync errors to OSD", async () => {
  const { deps, calls, osd } = createDeps({
    triggerSubsyncFromConfig: async () => {
      throw new Error("subsync boom");
    },
  });

  handleCliCommandService(makeArgs({ triggerSubsync: true }), "initial", deps);
  await new Promise((resolve) => setImmediate(resolve));

  assert.ok(
    calls.some((value) => value.startsWith("error:triggerSubsyncFromConfig failed:")),
  );
  assert.ok(osd.some((value) => value.includes("Subsync failed: subsync boom")));
});

test("handleCliCommandService stops app for --stop command", () => {
  const { deps, calls } = createDeps();
  handleCliCommandService(makeArgs({ stop: true }), "initial", deps);
  assert.ok(calls.includes("log:Stopping SubMiner..."));
  assert.ok(calls.includes("stopApp"));
});

test("handleCliCommandService still runs non-start actions on second-instance", () => {
  const { deps, calls } = createDeps();
  handleCliCommandService(
    makeArgs({ start: true, toggleVisibleOverlay: true }),
    "second-instance",
    deps,
  );
  assert.ok(calls.includes("toggleVisibleOverlay"));
  assert.equal(calls.some((value) => value === "connectMpvClient"), true);
});

test("handleCliCommandService handles visibility and utility command dispatches", () => {
  const cases: Array<{
    args: Partial<CliArgs>;
    expected: string;
  }> = [
    { args: { toggleInvisibleOverlay: true }, expected: "toggleInvisibleOverlay" },
    { args: { settings: true }, expected: "openYomitanSettingsDelayed:1000" },
    { args: { showVisibleOverlay: true }, expected: "setVisibleOverlayVisible:true" },
    { args: { hideVisibleOverlay: true }, expected: "setVisibleOverlayVisible:false" },
    { args: { showInvisibleOverlay: true }, expected: "setInvisibleOverlayVisible:true" },
    { args: { hideInvisibleOverlay: true }, expected: "setInvisibleOverlayVisible:false" },
    { args: { copySubtitle: true }, expected: "copyCurrentSubtitle" },
    { args: { copySubtitleMultiple: true }, expected: "startPendingMultiCopy:2500" },
    {
      args: { mineSentenceMultiple: true },
      expected: "startPendingMineSentenceMultiple:2500",
    },
    { args: { toggleSecondarySub: true }, expected: "cycleSecondarySubMode" },
    { args: { openRuntimeOptions: true }, expected: "openRuntimeOptionsPalette" },
  ];

  for (const entry of cases) {
    const { deps, calls } = createDeps();
    handleCliCommandService(makeArgs(entry.args), "initial", deps);
    assert.ok(
      calls.includes(entry.expected),
      `expected call missing for args ${JSON.stringify(entry.args)}: ${entry.expected}`,
    );
  }
});

test("handleCliCommandService runs refresh-known-words command", () => {
  const { deps, calls } = createDeps();

  handleCliCommandService(makeArgs({ refreshKnownWords: true }), "initial", deps);

  assert.ok(calls.includes("refreshKnownWords"));
});

test("handleCliCommandService reports async refresh-known-words errors to OSD", async () => {
  const { deps, calls, osd } = createDeps({
    refreshKnownWords: async () => {
      throw new Error("refresh boom");
    },
  });

  handleCliCommandService(makeArgs({ refreshKnownWords: true }), "initial", deps);
  await new Promise((resolve) => setImmediate(resolve));

  assert.ok(
    calls.some((value) => value.startsWith("error:refreshKnownWords failed:")),
  );
  assert.ok(osd.some((value) => value.includes("Refresh known words failed: refresh boom")));
});
