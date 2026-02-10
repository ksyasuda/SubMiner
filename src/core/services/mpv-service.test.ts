import test from "node:test";
import assert from "node:assert/strict";
import { MpvIpcClient, MpvIpcClientDeps } from "./mpv-service";

function makeDeps(
  overrides: Partial<MpvIpcClientDeps> = {},
): MpvIpcClientDeps {
  return {
    getResolvedConfig: () => ({} as any),
    autoStartOverlay: false,
    setOverlayVisible: () => {},
    shouldBindVisibleOverlayToMpvSubVisibility: () => false,
    isVisibleOverlayVisible: () => false,
    getReconnectTimer: () => null,
    setReconnectTimer: () => {},
    getCurrentSubText: () => "",
    setCurrentSubText: () => {},
    setCurrentSubAssText: () => {},
    getSubtitleTimingTracker: () => null,
    subtitleWsBroadcast: () => {},
    getOverlayWindowsCount: () => 0,
    tokenizeSubtitle: async (text) => ({ text, tokens: null }),
    broadcastToOverlayWindows: () => {},
    updateCurrentMediaPath: () => {},
    updateMpvSubtitleRenderMetrics: () => {},
    getMpvSubtitleRenderMetrics: () => ({
      subPos: 100,
      subFontSize: 36,
      subScale: 1,
      subMarginY: 0,
      subMarginX: 0,
      subFont: "sans-serif",
      subSpacing: 0,
      subBold: false,
      subItalic: false,
      subBorderSize: 0,
      subShadowOffset: 0,
      subAssOverride: "yes",
      subScaleByWindow: true,
      subUseMargins: true,
      osdHeight: 720,
      osdDimensions: null,
    }),
    setPreviousSecondarySubVisibility: () => {},
    showMpvOsd: () => {},
    ...overrides,
  };
}

test("MpvIpcClient resolves pending request by request_id", async () => {
  const client = new MpvIpcClient("/tmp/mpv.sock", makeDeps());
  let resolved: unknown = null;
  (client as any).pendingRequests.set(1234, (msg: unknown) => {
    resolved = msg;
  });

  await (client as any).handleMessage({ request_id: 1234, data: "ok" });

  assert.deepEqual(resolved, { request_id: 1234, data: "ok" });
  assert.equal((client as any).pendingRequests.size, 0);
});

test("MpvIpcClient handles sub-text property change and broadcasts tokenized subtitle", async () => {
  const calls: string[] = [];
  const client = new MpvIpcClient(
    "/tmp/mpv.sock",
    makeDeps({
      setCurrentSubText: (text) => {
        calls.push(`setCurrentSubText:${text}`);
      },
      subtitleWsBroadcast: (text) => {
        calls.push(`subtitleWsBroadcast:${text}`);
      },
      getOverlayWindowsCount: () => 1,
      tokenizeSubtitle: async (text) => ({ text, tokens: null }),
      broadcastToOverlayWindows: (channel, payload) => {
        calls.push(`broadcast:${channel}:${String((payload as any).text ?? "")}`);
      },
    }),
  );

  await (client as any).handleMessage({
    event: "property-change",
    name: "sub-text",
    data: "字幕",
  });

  assert.ok(calls.includes("setCurrentSubText:字幕"));
  assert.ok(calls.includes("subtitleWsBroadcast:字幕"));
  assert.ok(calls.includes("broadcast:subtitle:set:字幕"));
});

test("MpvIpcClient parses JSON line protocol in processBuffer", () => {
  const client = new MpvIpcClient("/tmp/mpv.sock", makeDeps());
  const seen: Array<Record<string, unknown>> = [];
  (client as any).handleMessage = (msg: Record<string, unknown>) => {
    seen.push(msg);
  };
  (client as any).buffer =
    "{\"event\":\"property-change\",\"name\":\"path\",\"data\":\"a\"}\n{\"request_id\":1,\"data\":\"ok\"}\n{\"partial\":";

  (client as any).processBuffer();

  assert.equal(seen.length, 2);
  assert.equal(seen[0].name, "path");
  assert.equal(seen[1].request_id, 1);
  assert.equal((client as any).buffer, "{\"partial\":");
});

test("MpvIpcClient request rejects when disconnected", async () => {
  const client = new MpvIpcClient("/tmp/mpv.sock", makeDeps());
  await assert.rejects(
    async () => client.request(["get_property", "path"]),
    /MPV not connected/,
  );
});

test("MpvIpcClient requestProperty throws on mpv error response", async () => {
  const client = new MpvIpcClient("/tmp/mpv.sock", makeDeps());
  (client as any).request = async () => ({ error: "property unavailable" });
  await assert.rejects(
    async () => client.requestProperty("path"),
    /Failed to read MPV property 'path': property unavailable/,
  );
});

test("MpvIpcClient failPendingRequests resolves outstanding requests as disconnected", () => {
  const client = new MpvIpcClient("/tmp/mpv.sock", makeDeps());
  const resolved: unknown[] = [];
  (client as any).pendingRequests.set(10, (msg: unknown) => {
    resolved.push(msg);
  });
  (client as any).pendingRequests.set(11, (msg: unknown) => {
    resolved.push(msg);
  });

  (client as any).failPendingRequests();

  assert.deepEqual(resolved, [
    { request_id: 10, error: "disconnected" },
    { request_id: 11, error: "disconnected" },
  ]);
  assert.equal((client as any).pendingRequests.size, 0);
});

test("MpvIpcClient scheduleReconnect schedules timer and invokes connect", () => {
  const timers: Array<ReturnType<typeof setTimeout> | null> = [];
  const client = new MpvIpcClient(
    "/tmp/mpv.sock",
    makeDeps({
      getReconnectTimer: () => null,
      setReconnectTimer: (timer) => {
        timers.push(timer);
      },
    }),
  );

  let connectCalled = false;
  (client as any).connect = () => {
    connectCalled = true;
  };

  const originalSetTimeout = globalThis.setTimeout;
  (globalThis as any).setTimeout = (handler: () => void, _delay: number) => {
    handler();
    return 1 as unknown as ReturnType<typeof setTimeout>;
  };
  try {
    (client as any).scheduleReconnect();
  } finally {
    (globalThis as any).setTimeout = originalSetTimeout;
  }

  assert.equal(timers.length, 1);
  assert.equal(connectCalled, true);
});
