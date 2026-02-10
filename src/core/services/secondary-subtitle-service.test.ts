import test from "node:test";
import assert from "node:assert/strict";
import { SecondarySubMode } from "../../types";
import { cycleSecondarySubModeService } from "./secondary-subtitle-service";

test("cycleSecondarySubModeService cycles and emits broadcast + OSD", () => {
  let mode: SecondarySubMode = "hover";
  let lastToggleAt = 0;
  const broadcasts: SecondarySubMode[] = [];
  const osd: string[] = [];

  cycleSecondarySubModeService({
    getSecondarySubMode: () => mode,
    setSecondarySubMode: (next) => {
      mode = next;
    },
    getLastSecondarySubToggleAtMs: () => lastToggleAt,
    setLastSecondarySubToggleAtMs: (value) => {
      lastToggleAt = value;
    },
    broadcastSecondarySubMode: (next) => {
      broadcasts.push(next);
    },
    showMpvOsd: (text) => {
      osd.push(text);
    },
    now: () => 1000,
  });

  assert.equal(mode, "hidden");
  assert.deepEqual(broadcasts, ["hidden"]);
  assert.deepEqual(osd, ["Secondary subtitle: hidden"]);
  assert.equal(lastToggleAt, 1000);
});

test("cycleSecondarySubModeService obeys debounce window", () => {
  let mode: SecondarySubMode = "visible";
  let lastToggleAt = 950;
  let broadcasted = false;
  let osdShown = false;

  cycleSecondarySubModeService({
    getSecondarySubMode: () => mode,
    setSecondarySubMode: (next) => {
      mode = next;
    },
    getLastSecondarySubToggleAtMs: () => lastToggleAt,
    setLastSecondarySubToggleAtMs: (value) => {
      lastToggleAt = value;
    },
    broadcastSecondarySubMode: () => {
      broadcasted = true;
    },
    showMpvOsd: () => {
      osdShown = true;
    },
    now: () => 1000,
  });

  assert.equal(mode, "visible");
  assert.equal(lastToggleAt, 950);
  assert.equal(broadcasted, false);
  assert.equal(osdShown, false);
});
