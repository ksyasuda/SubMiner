import test from "node:test";
import assert from "node:assert/strict";
import {
  MpvIpcClient,
  MpvIpcClientDeps,
  MpvIpcClientProtocolDeps,
  MPV_REQUEST_ID_SECONDARY_SUB_VISIBILITY,
} from "./mpv-service";

function makeDeps(
  overrides: Partial<MpvIpcClientProtocolDeps> = {},
): MpvIpcClientDeps {
  return {
    getResolvedConfig: () => ({} as any),
    autoStartOverlay: false,
    setOverlayVisible: () => {},
    shouldBindVisibleOverlayToMpvSubVisibility: () => false,
    isVisibleOverlayVisible: () => false,
    getReconnectTimer: () => null,
    setReconnectTimer: () => {},
    ...overrides,
  };
}

function invokeHandleMessage(client: MpvIpcClient, msg: unknown): Promise<void> {
  return (client as unknown as { handleMessage: (msg: unknown) => Promise<void> }).handleMessage(
    msg,
  );
}

test("MpvIpcClient resolves pending request by request_id", async () => {
  const client = new MpvIpcClient("/tmp/mpv.sock", makeDeps());
  let resolved: unknown = null;
  (client as any).pendingRequests.set(1234, (msg: unknown) => {
    resolved = msg;
  });

  await invokeHandleMessage(client, { request_id: 1234, data: "ok" });

  assert.deepEqual(resolved, { request_id: 1234, data: "ok" });
  assert.equal((client as any).pendingRequests.size, 0);
});

test("MpvIpcClient handles sub-text property change and broadcasts tokenized subtitle", async () => {
  const events: Array<{ text: string; isOverlayVisible: boolean }> = [];
  const client = new MpvIpcClient("/tmp/mpv.sock", makeDeps());
  client.on("subtitle-change", (payload) => {
    events.push(payload);
  });

  await invokeHandleMessage(client, {
    event: "property-change",
    name: "sub-text",
    data: "字幕",
  });

  assert.equal(events.length, 1);
  assert.equal(events[0].text, "字幕");
  assert.equal(events[0].isOverlayVisible, false);
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

test("MpvIpcClient captures and disables secondary subtitle visibility on request", async () => {
  const commands: unknown[] = [];
  const client = new MpvIpcClient("/tmp/mpv.sock", makeDeps());
  const previous: boolean[] = [];
  client.on("secondary-subtitle-visibility", ({ visible }) => {
    previous.push(visible);
  });

  (client as any).send = (payload: unknown) => {
    commands.push(payload);
    return true;
  };

  await invokeHandleMessage(client, {
    request_id: MPV_REQUEST_ID_SECONDARY_SUB_VISIBILITY,
    data: "yes",
  });

  assert.deepEqual(previous, [true]);
  assert.deepEqual(commands, [
    {
      command: ["set_property", "secondary-sub-visibility", "no"],
    },
  ]);
});

test("MpvIpcClient restorePreviousSecondarySubVisibility restores and clears tracked value", async () => {
  const commands: unknown[] = [];
  const client = new MpvIpcClient("/tmp/mpv.sock", makeDeps());
  const previous: boolean[] = [];
  client.on("secondary-subtitle-visibility", ({ visible }) => {
    previous.push(visible);
  });

  (client as any).send = (payload: unknown) => {
    commands.push(payload);
    return true;
  };

  await invokeHandleMessage(client, {
    request_id: MPV_REQUEST_ID_SECONDARY_SUB_VISIBILITY,
    data: "yes",
  });
  client.restorePreviousSecondarySubVisibility();

  assert.equal(previous[0], true);
  assert.equal(previous.length, 1);
  assert.deepEqual(commands, [
    {
      command: ["set_property", "secondary-sub-visibility", "no"],
    },
    {
      command: ["set_property", "secondary-sub-visibility", "no"],
    },
  ]);

  client.restorePreviousSecondarySubVisibility();
  assert.equal(commands.length, 2);
});
