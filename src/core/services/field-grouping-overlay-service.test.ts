import test from "node:test";
import assert from "node:assert/strict";
import { KikuFieldGroupingChoice } from "../../types";
import { createFieldGroupingOverlayRuntimeService } from "./field-grouping-overlay-service";

test("createFieldGroupingOverlayRuntimeService sends overlay messages and sets restore flag", () => {
  const sent: unknown[][] = [];
  let visible = false;
  const restore = new Set<"runtime-options" | "subsync">();

  const runtime =
    createFieldGroupingOverlayRuntimeService<"runtime-options" | "subsync">({
    getMainWindow: () => ({
      isDestroyed: () => false,
      webContents: {
        isLoading: () => false,
        send: (...args: unknown[]) => {
          sent.push(args);
        },
      },
    }),
    getVisibleOverlayVisible: () => visible,
    getInvisibleOverlayVisible: () => false,
    setVisibleOverlayVisible: (next) => {
      visible = next;
    },
    setInvisibleOverlayVisible: () => {},
    getResolver: () => null,
    setResolver: () => {},
    getRestoreVisibleOverlayOnModalClose: () => restore,
    });

  const ok = runtime.sendToVisibleOverlay("runtime-options:open", undefined, {
    restoreOnModalClose: "runtime-options",
  });

  assert.equal(ok, true);
  assert.equal(visible, true);
  assert.equal(restore.has("runtime-options"), true);
  assert.deepEqual(sent, [["runtime-options:open"]]);
});

test("createFieldGroupingOverlayRuntimeService callback cancels when send fails", async () => {
  let resolver: ((choice: KikuFieldGroupingChoice) => void) | null = null;
  const runtime =
    createFieldGroupingOverlayRuntimeService<"runtime-options" | "subsync">({
      getMainWindow: () => null,
      getVisibleOverlayVisible: () => false,
      getInvisibleOverlayVisible: () => false,
      setVisibleOverlayVisible: () => {},
      setInvisibleOverlayVisible: () => {},
      getResolver: () => resolver,
      setResolver: (next) => {
        resolver = next;
      },
      getRestoreVisibleOverlayOnModalClose: () =>
        new Set<"runtime-options" | "subsync">(),
    });

  const callback = runtime.createFieldGroupingCallback();
  const result = await callback({
    original: {
      noteId: 1,
      expression: "a",
      sentencePreview: "a",
      hasAudio: false,
      hasImage: false,
      isOriginal: true,
    },
    duplicate: {
      noteId: 2,
      expression: "b",
      sentencePreview: "b",
      hasAudio: false,
      hasImage: false,
      isOriginal: false,
    },
  });

  assert.equal(result.cancelled, true);
  assert.equal(result.keepNoteId, 0);
  assert.equal(result.deleteNoteId, 0);
});
