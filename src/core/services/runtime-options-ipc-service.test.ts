import test from "node:test";
import assert from "node:assert/strict";
import {
  applyRuntimeOptionResultRuntimeService,
  cycleRuntimeOptionFromIpcRuntimeService,
  setRuntimeOptionFromIpcRuntimeService,
} from "./runtime-options-ipc-service";

test("applyRuntimeOptionResultRuntimeService emits success OSD message", () => {
  const osd: string[] = [];
  const result = applyRuntimeOptionResultRuntimeService(
    { ok: true, osdMessage: "Updated" },
    (text) => {
      osd.push(text);
    },
  );

  assert.equal(result.ok, true);
  assert.deepEqual(osd, ["Updated"]);
});

test("setRuntimeOptionFromIpcRuntimeService returns unavailable when manager missing", () => {
  const osd: string[] = [];
  const result = setRuntimeOptionFromIpcRuntimeService(
    null,
    "anki.autoUpdateNewCards",
    true,
    (text) => {
      osd.push(text);
    },
  );
  assert.equal(result.ok, false);
  assert.equal(result.error, "Runtime options manager unavailable");
  assert.deepEqual(osd, []);
});

test("cycleRuntimeOptionFromIpcRuntimeService reports errors once", () => {
  const osd: string[] = [];
  const result = cycleRuntimeOptionFromIpcRuntimeService(
    {
      setOptionValue: () => ({ ok: true }),
      cycleOption: () => ({ ok: false, error: "bad option" }),
    },
    "anki.kikuFieldGrouping",
    1,
    (text) => {
      osd.push(text);
    },
  );
  assert.equal(result.ok, false);
  assert.equal(result.error, "bad option");
  assert.deepEqual(osd, ["bad option"]);
});
