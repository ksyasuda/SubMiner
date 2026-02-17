import test from "node:test";
import assert from "node:assert/strict";
import {
  runStartupBootstrapRuntimeService,
} from "./startup-service";
import { CliArgs } from "../../cli/args";

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
    refreshKnownWords: false,
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
    debug: false,
    ...overrides,
  };
}

test("runStartupBootstrapRuntimeService configures startup state and starts lifecycle", () => {
  const calls: string[] = [];
  const args = makeArgs({
    logLevel: "debug",
    socketPath: "/tmp/custom.sock",
    texthookerPort: 9001,
    backend: "x11",
    autoStartOverlay: true,
    texthooker: true,
  });

  const result = runStartupBootstrapRuntimeService({
    argv: ["node", "main.ts", "--log-level", "debug"],
    parseArgs: () => args,
    setLogLevel: (level, source) => calls.push(`setLog:${level}:${source}`),
    forceX11Backend: () => calls.push("forceX11"),
    enforceUnsupportedWaylandMode: () => calls.push("enforceWayland"),
    getDefaultSocketPath: () => "/tmp/default.sock",
    defaultTexthookerPort: 5174,
    runGenerateConfigFlow: () => false,
    startAppLifecycle: () => calls.push("startLifecycle"),
  });

  assert.equal(result.initialArgs, args);
  assert.equal(result.mpvSocketPath, "/tmp/custom.sock");
  assert.equal(result.texthookerPort, 9001);
  assert.equal(result.backendOverride, "x11");
  assert.equal(result.autoStartOverlay, true);
  assert.equal(result.texthookerOnlyMode, true);
  assert.deepEqual(calls, [
    "setLog:debug:cli",
    "forceX11",
    "enforceWayland",
    "startLifecycle",
  ]);
});

test("runStartupBootstrapRuntimeService keeps log-level precedence for repeated calls", () => {
  const calls: string[] = [];
  const args = makeArgs({
    logLevel: "warn",
  });

  runStartupBootstrapRuntimeService({
    argv: ["node", "main.ts", "--log-level", "warn"],
    parseArgs: () => args,
    setLogLevel: (level, source) => calls.push(`setLog:${level}:${source}`),
    forceX11Backend: () => calls.push("forceX11"),
    enforceUnsupportedWaylandMode: () => calls.push("enforceWayland"),
    getDefaultSocketPath: () => "/tmp/default.sock",
    defaultTexthookerPort: 5174,
    runGenerateConfigFlow: () => false,
    startAppLifecycle: () => calls.push("startLifecycle"),
  });

  assert.deepEqual(calls.slice(0, 3), [
    "setLog:warn:cli",
    "forceX11",
    "enforceWayland",
  ]);
});

test("runStartupBootstrapRuntimeService keeps --debug separate from log verbosity", () => {
  const calls: string[] = [];
  const args = makeArgs({
    debug: true,
  });

  runStartupBootstrapRuntimeService({
    argv: ["node", "main.ts", "--debug"],
    parseArgs: () => args,
    setLogLevel: (level, source) => calls.push(`setLog:${level}:${source}`),
    forceX11Backend: () => calls.push("forceX11"),
    enforceUnsupportedWaylandMode: () => calls.push("enforceWayland"),
    getDefaultSocketPath: () => "/tmp/default.sock",
    defaultTexthookerPort: 5174,
    runGenerateConfigFlow: () => false,
    startAppLifecycle: () => calls.push("startLifecycle"),
  });

  assert.deepEqual(calls, ["forceX11", "enforceWayland", "startLifecycle"]);
});

test("runStartupBootstrapRuntimeService skips lifecycle when generate-config flow handled", () => {
  const calls: string[] = [];
  const args = makeArgs({ generateConfig: true, logLevel: "warn" });

  const result = runStartupBootstrapRuntimeService({
    argv: ["node", "main.ts", "--generate-config"],
    parseArgs: () => args,
    setLogLevel: (level, source) => calls.push(`setLog:${level}:${source}`),
    forceX11Backend: () => calls.push("forceX11"),
    enforceUnsupportedWaylandMode: () => calls.push("enforceWayland"),
    getDefaultSocketPath: () => "/tmp/default.sock",
    defaultTexthookerPort: 5174,
    runGenerateConfigFlow: () => true,
    startAppLifecycle: () => calls.push("startLifecycle"),
  });

  assert.equal(result.mpvSocketPath, "/tmp/default.sock");
  assert.equal(result.texthookerPort, 5174);
  assert.equal(result.backendOverride, null);
  assert.deepEqual(calls, [
    "setLog:warn:cli",
    "forceX11",
    "enforceWayland",
  ]);
});
