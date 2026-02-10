import test from "node:test";
import assert from "node:assert/strict";
import { runAppShutdownRuntimeService } from "./app-shutdown-runtime-service";

test("runAppShutdownRuntimeService runs teardown steps in order", () => {
  const calls: string[] = [];
  runAppShutdownRuntimeService({
    unregisterAllGlobalShortcuts: () => calls.push("unregisterAllGlobalShortcuts"),
    stopSubtitleWebsocket: () => calls.push("stopSubtitleWebsocket"),
    stopTexthookerService: () => calls.push("stopTexthookerService"),
    destroyYomitanParserWindow: () => calls.push("destroyYomitanParserWindow"),
    clearYomitanParserPromises: () => calls.push("clearYomitanParserPromises"),
    stopWindowTracker: () => calls.push("stopWindowTracker"),
    destroyMpvSocket: () => calls.push("destroyMpvSocket"),
    clearReconnectTimer: () => calls.push("clearReconnectTimer"),
    destroySubtitleTimingTracker: () => calls.push("destroySubtitleTimingTracker"),
    destroyAnkiIntegration: () => calls.push("destroyAnkiIntegration"),
  });

  assert.deepEqual(calls, [
    "unregisterAllGlobalShortcuts",
    "stopSubtitleWebsocket",
    "stopTexthookerService",
    "destroyYomitanParserWindow",
    "clearYomitanParserPromises",
    "stopWindowTracker",
    "destroyMpvSocket",
    "clearReconnectTimer",
    "destroySubtitleTimingTracker",
    "destroyAnkiIntegration",
  ]);
});
