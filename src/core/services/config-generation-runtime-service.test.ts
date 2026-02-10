import test from "node:test";
import assert from "node:assert/strict";
import { runGenerateConfigFlowRuntimeService } from "./config-generation-runtime-service";
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

test("runGenerateConfigFlowRuntimeService starts flow when generateConfig is set and app should not start", async () => {
  const calls: string[] = [];
  const handled = runGenerateConfigFlowRuntimeService(
    makeArgs({ generateConfig: true }),
    {
      shouldStartApp: () => false,
      generateConfig: async () => 7,
      onSuccess: (code) => calls.push(`success:${code}`),
      onError: () => calls.push("error"),
    },
  );
  assert.equal(handled, true);
  await new Promise((resolve) => setImmediate(resolve));
  assert.deepEqual(calls, ["success:7"]);
});

test("runGenerateConfigFlowRuntimeService returns false when flow should not run", () => {
  const handled = runGenerateConfigFlowRuntimeService(
    makeArgs({ generateConfig: true, start: true }),
    {
      shouldStartApp: () => true,
      generateConfig: async () => 0,
      onSuccess: () => {},
      onError: () => {},
    },
  );
  assert.equal(handled, false);
});
