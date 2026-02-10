import test from "node:test";
import assert from "node:assert/strict";
import {
  createOverlayShortcutLifecycleDepsRuntimeService,
  createOverlayShortcutRuntimeDepsService,
} from "./overlay-shortcut-runtime-deps-service";

test("createOverlayShortcutRuntimeDepsService returns callable runtime deps", async () => {
  const calls: string[] = [];
  const deps = createOverlayShortcutRuntimeDepsService({
    showMpvOsd: () => calls.push("showMpvOsd"),
    openRuntimeOptions: () => calls.push("openRuntimeOptions"),
    openJimaku: () => calls.push("openJimaku"),
    markAudioCard: async () => {
      calls.push("markAudioCard");
    },
    copySubtitleMultiple: () => calls.push("copySubtitleMultiple"),
    copySubtitle: () => calls.push("copySubtitle"),
    toggleSecondarySub: () => calls.push("toggleSecondarySub"),
    updateLastCardFromClipboard: async () => {
      calls.push("updateLastCardFromClipboard");
    },
    triggerFieldGrouping: async () => {
      calls.push("triggerFieldGrouping");
    },
    triggerSubsync: async () => {
      calls.push("triggerSubsync");
    },
    mineSentence: async () => {
      calls.push("mineSentence");
    },
    mineSentenceMultiple: () => calls.push("mineSentenceMultiple"),
  });

  deps.copySubtitle();
  await deps.mineSentence();
  deps.mineSentenceMultiple(2);

  assert.deepEqual(calls, ["copySubtitle", "mineSentence", "mineSentenceMultiple"]);
});

test("createOverlayShortcutLifecycleDepsRuntimeService returns lifecycle passthrough", () => {
  const deps = createOverlayShortcutLifecycleDepsRuntimeService({
    getConfiguredShortcuts: () => ({ actions: [] } as never),
    getOverlayHandlers: () => ({} as never),
    cancelPendingMultiCopy: () => {},
    cancelPendingMineSentenceMultiple: () => {},
  });

  assert.ok(deps.getConfiguredShortcuts());
  assert.ok(deps.getOverlayHandlers());
});
