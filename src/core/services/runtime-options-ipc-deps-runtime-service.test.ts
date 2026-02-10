import test from "node:test";
import assert from "node:assert/strict";
import { createRuntimeOptionsIpcDepsRuntimeService } from "./runtime-options-ipc-deps-runtime-service";

test("createRuntimeOptionsIpcDepsRuntimeService delegates set/cycle with osd", () => {
  const osd: string[] = [];
  const deps = createRuntimeOptionsIpcDepsRuntimeService({
    getRuntimeOptionsManager: () => ({
      setOptionValue: () => ({ ok: true, osdMessage: "set ok" }),
      cycleOption: () => ({ ok: true, osdMessage: "cycle ok" }),
    }),
    showMpvOsd: (text) => {
      osd.push(text);
    },
  });

  const setResult = deps.setRuntimeOption("subtitles.secondaryMode", "hidden") as {
    ok: boolean;
  };
  const cycleResult = deps.cycleRuntimeOption("subtitles.secondaryMode", 1) as {
    ok: boolean;
  };

  assert.equal(setResult.ok, true);
  assert.equal(cycleResult.ok, true);
  assert.ok(osd.includes("set ok"));
  assert.ok(osd.includes("cycle ok"));
});
