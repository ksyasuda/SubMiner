import test from "node:test";
import assert from "node:assert/strict";
import {
  AnkiJimakuIpcDepsRuntimeOptions,
  createAnkiJimakuIpcDepsRuntimeService,
} from "./anki-jimaku-ipc-deps-runtime-service";

test("createAnkiJimakuIpcDepsRuntimeService returns passthrough runtime options", async () => {
  const calls: string[] = [];
  const options = {
    patchAnkiConnectEnabled: () => calls.push("patch"),
    getResolvedConfig: () => ({ ankiConnect: undefined }),
    getRuntimeOptionsManager: () => null,
    getSubtitleTimingTracker: () => null,
    getMpvClient: () => null,
    getAnkiIntegration: () => null,
    setAnkiIntegration: () => calls.push("set-integration"),
    showDesktopNotification: () => calls.push("notify"),
    createFieldGroupingCallback: () => async () => ({
      keepNoteId: 0,
      deleteNoteId: 0,
      deleteDuplicate: false,
      cancelled: true,
    }),
    broadcastRuntimeOptionsChanged: () => calls.push("broadcast"),
    getFieldGroupingResolver: () => null,
    setFieldGroupingResolver: () => calls.push("set-resolver"),
    parseMediaInfo: () => ({ mediaPath: null, baseName: null, episode: null }),
    getCurrentMediaPath: () => "/tmp/a.mp4",
    jimakuFetchJson: async () => ({ ok: true, data: [] }),
    getJimakuMaxEntryResults: () => 100,
    getJimakuLanguagePreference: () => "prefer-japanese",
    resolveJimakuApiKey: async () => "abc",
    isRemoteMediaPath: () => false,
    downloadToFile: async () => ({ ok: true, path: "/tmp/a.srt" }),
  } as unknown as AnkiJimakuIpcDepsRuntimeOptions;

  const runtime = createAnkiJimakuIpcDepsRuntimeService(options);

  runtime.patchAnkiConnectEnabled(true);
  runtime.broadcastRuntimeOptionsChanged();
  runtime.setFieldGroupingResolver(null);

  assert.deepEqual(calls, ["patch", "broadcast", "set-resolver"]);
  assert.equal(runtime.getCurrentMediaPath(), "/tmp/a.mp4");
  assert.equal(runtime.getJimakuMaxEntryResults(), 100);
  assert.equal(await runtime.resolveJimakuApiKey(), "abc");
});
