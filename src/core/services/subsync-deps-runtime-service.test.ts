import test from "node:test";
import assert from "node:assert/strict";
import { createSubsyncRuntimeDepsService } from "./subsync-deps-runtime-service";

test("createSubsyncRuntimeDepsService opens manual picker via visible overlay", () => {
  let inProgress = false;
  const calls: Array<{ channel: string; payload?: unknown; restore?: string }> = [];

  const deps = createSubsyncRuntimeDepsService({
    getMpvClient: () => null,
    getResolvedSubsyncConfig: () => ({
      defaultMode: "auto",
      ffsubsyncPath: "/usr/bin/ffsubsync",
      alassPath: "/usr/bin/alass",
      ffmpegPath: "/usr/bin/ffmpeg",
    }),
    isSubsyncInProgress: () => inProgress,
    setSubsyncInProgress: (next) => {
      inProgress = next;
    },
    showMpvOsd: () => {},
    sendToVisibleOverlay: (channel, payload, options) => {
      calls.push({ channel, payload, restore: options?.restoreOnModalClose });
      return true;
    },
  });

  deps.setSubsyncInProgress(true);
  deps.openManualPicker({
    sourceTracks: [{ id: 1, label: "Japanese Track" }],
  });

  assert.equal(deps.isSubsyncInProgress(), true);
  assert.equal(calls.length, 1);
  assert.equal(calls[0]?.channel, "subsync:open-manual");
  assert.equal(calls[0]?.restore, "subsync");
});
