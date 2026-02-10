import test from "node:test";
import assert from "node:assert/strict";
import { createRuntimeOptionsManagerRuntimeService } from "./runtime-options-manager-runtime-service";

test("createRuntimeOptionsManagerRuntimeService wires patch + options changed callbacks", () => {
  const patches: unknown[] = [];
  const changedSnapshots: unknown[] = [];
  const manager = createRuntimeOptionsManagerRuntimeService({
    getAnkiConfig: () => ({
      behavior: { autoUpdateNewCards: true },
      isKiku: { fieldGrouping: "manual" },
    }),
    applyAnkiPatch: (patch) => {
      patches.push(patch);
    },
    onOptionsChanged: (options) => {
      changedSnapshots.push(options);
    },
  });

  const result = manager.setOptionValue("anki.autoUpdateNewCards", false);
  assert.equal(result.ok, true);
  assert.equal(patches.length > 0, true);
  assert.equal(changedSnapshots.length > 0, true);
});
