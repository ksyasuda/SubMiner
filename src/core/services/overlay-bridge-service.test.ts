import test from "node:test";
import assert from "node:assert/strict";
import { KikuFieldGroupingChoice } from "../../types";
import {
  createFieldGroupingCallbackRuntimeService,
  sendToVisibleOverlayRuntimeService,
} from "./overlay-bridge-service";

test("sendToVisibleOverlayRuntimeService restores visibility flag when opening hidden overlay modal", () => {
  const sent: unknown[][] = [];
  const restoreSet = new Set<"runtime-options" | "subsync">();
  let visibleOverlayVisible = false;

  const ok = sendToVisibleOverlayRuntimeService({
    mainWindow: {
      isDestroyed: () => false,
      webContents: {
        send: (...args: unknown[]) => {
          sent.push(args);
        },
      },
    } as unknown as Electron.BrowserWindow,
    visibleOverlayVisible,
    setVisibleOverlayVisible: (visible) => {
      visibleOverlayVisible = visible;
    },
    channel: "runtime-options:open",
    restoreOnModalClose: "runtime-options",
    restoreVisibleOverlayOnModalClose: restoreSet,
  });

  assert.equal(ok, true);
  assert.equal(visibleOverlayVisible, true);
  assert.equal(restoreSet.has("runtime-options"), true);
  assert.deepEqual(sent, [["runtime-options:open"]]);
});

test("createFieldGroupingCallbackRuntimeService cancels when overlay request cannot be sent", async () => {
  let resolver: ((choice: KikuFieldGroupingChoice) => void) | null = null;
  const callback = createFieldGroupingCallbackRuntimeService<
    "runtime-options" | "subsync"
  >({
    getVisibleOverlayVisible: () => false,
    getInvisibleOverlayVisible: () => false,
    setVisibleOverlayVisible: () => {},
    setInvisibleOverlayVisible: () => {},
    getResolver: () => resolver,
    setResolver: (next) => {
      resolver = next;
    },
    sendToVisibleOverlay: () => false,
  });

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
