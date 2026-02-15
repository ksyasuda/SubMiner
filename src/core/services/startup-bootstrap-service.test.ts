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
    verbose: false,
    ...overrides,
  };
}

test("runStartupBootstrapRuntimeService configures startup state and starts lifecycle", () => {
  const calls: string[] = [];
  const args = makeArgs({
    verbose: true,
    socketPath: "/tmp/custom.sock",
    texthookerPort: 9001,
    backend: "x11",
    autoStartOverlay: true,
    texthooker: true,
  });

  const result = runStartupBootstrapRuntimeService({
    argv: ["node", "main.ts", "--verbose"],
    parseArgs: () => args,
    setLogLevelEnv: (level) => calls.push(`setLog:${level}`),
    enableVerboseLogging: () => calls.push("enableVerbose"),
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
    "enableVerbose",
    "forceX11",
    "enforceWayland",
    "startLifecycle",
  ]);
});

test("runStartupBootstrapRuntimeService skips lifecycle when generate-config flow handled", () => {
  const calls: string[] = [];
  const args = makeArgs({ generateConfig: true, logLevel: "warn" });

  const result = runStartupBootstrapRuntimeService({
    argv: ["node", "main.ts", "--generate-config"],
    parseArgs: () => args,
    setLogLevelEnv: (level) => calls.push(`setLog:${level}`),
    enableVerboseLogging: () => calls.push("enableVerbose"),
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
    "setLog:warn",
    "forceX11",
    "enforceWayland",
  ]);
});
