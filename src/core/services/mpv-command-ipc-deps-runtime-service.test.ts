import test from "node:test";
import assert from "node:assert/strict";
import { SPECIAL_COMMANDS } from "../../config";
import { createMpvCommandIpcDepsRuntimeService } from "./mpv-command-ipc-deps-runtime-service";

test("createMpvCommandIpcDepsRuntimeService wires runtime-options cycle and manager availability", () => {
  const osd: string[] = [];
  const deps = createMpvCommandIpcDepsRuntimeService({
    specialCommands: SPECIAL_COMMANDS,
    triggerSubsyncFromConfig: () => {},
    openRuntimeOptionsPalette: () => {},
    getRuntimeOptionsManager: () => ({
      cycleOption: () => ({ ok: true, osdMessage: "cycled" }),
    }),
    showMpvOsd: (text) => {
      osd.push(text);
    },
    mpvReplaySubtitle: () => {},
    mpvPlayNextSubtitle: () => {},
    mpvSendCommand: () => {},
    isMpvConnected: () => true,
  });

  const result = deps.runtimeOptionsCycle("subtitles.secondaryMode" as never, 1);
  assert.equal(result.ok, true);
  assert.equal(deps.hasRuntimeOptionsManager(), true);
  assert.ok(osd.includes("cycled"));
});
