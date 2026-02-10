import test from "node:test";
import assert from "node:assert/strict";
import {
  addOverlayModalRestoreFlagService,
  handleOverlayModalClosedService,
} from "./overlay-modal-restore-service";

test("overlay modal restore service adds modal restore flag", () => {
  const restore = new Set<"runtime-options" | "subsync">();
  addOverlayModalRestoreFlagService(restore, "runtime-options");
  assert.equal(restore.has("runtime-options"), true);
});

test("overlay modal restore service hides overlay only when last modal closes", () => {
  const restore = new Set<"runtime-options" | "subsync">();
  const visibility: boolean[] = [];

  addOverlayModalRestoreFlagService(restore, "runtime-options");
  addOverlayModalRestoreFlagService(restore, "subsync");

  handleOverlayModalClosedService(restore, "runtime-options", (visible) => {
    visibility.push(visible);
  });
  assert.equal(visibility.length, 0);

  handleOverlayModalClosedService(restore, "subsync", (visible) => {
    visibility.push(visible);
  });
  assert.deepEqual(visibility, [false]);
});
