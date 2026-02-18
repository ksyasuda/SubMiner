import test from "node:test";
import assert from "node:assert/strict";
import {
  buildJellyfinTimelinePayload,
  JellyfinRemoteSessionService,
} from "./jellyfin-remote";

class FakeWebSocket {
  private listeners: Record<string, Array<(...args: unknown[]) => void>> = {};

  on(event: string, listener: (...args: unknown[]) => void): this {
    if (!this.listeners[event]) {
      this.listeners[event] = [];
    }
    this.listeners[event].push(listener);
    return this;
  }

  close(): void {
    this.emit("close");
  }

  emit(event: string, ...args: unknown[]): void {
    for (const listener of this.listeners[event] ?? []) {
      listener(...args);
    }
  }
}

test("Jellyfin remote service has no traffic until started", async () => {
  let socketCreateCount = 0;
  const fetchCalls: Array<{ input: string; init: RequestInit }> = [];

  const service = new JellyfinRemoteSessionService({
    serverUrl: "http://jellyfin.local:8096",
    accessToken: "token-0",
    deviceId: "device-0",
    webSocketFactory: () => {
      socketCreateCount += 1;
      return new FakeWebSocket() as unknown as any;
    },
    fetchImpl: (async (input, init) => {
      fetchCalls.push({ input: String(input), init: init ?? {} });
      return new Response(null, { status: 200 });
    }) as typeof fetch,
  });

  await new Promise((resolve) => setTimeout(resolve, 0));

  assert.equal(socketCreateCount, 0);
  assert.equal(fetchCalls.length, 0);
  assert.equal(service.isConnected(), false);
});

test("start posts capabilities on socket connect", async () => {
  const sockets: FakeWebSocket[] = [];
  const fetchCalls: Array<{ input: string; init: RequestInit }> = [];

  const service = new JellyfinRemoteSessionService({
    serverUrl: "http://jellyfin.local:8096",
    accessToken: "token-1",
    deviceId: "device-1",
    webSocketFactory: (url) => {
      assert.equal(url, "ws://jellyfin.local:8096/socket?api_key=token-1&deviceId=device-1");
      const socket = new FakeWebSocket();
      sockets.push(socket);
      return socket as unknown as any;
    },
    fetchImpl: (async (input, init) => {
      fetchCalls.push({ input: String(input), init: init ?? {} });
      return new Response(null, { status: 200 });
    }) as typeof fetch,
  });

  service.start();
  sockets[0].emit("open");
  await new Promise((resolve) => setTimeout(resolve, 0));

  assert.equal(fetchCalls.length, 1);
  assert.equal(
    fetchCalls[0].input,
    "http://jellyfin.local:8096/Sessions/Capabilities/Full",
  );
  assert.equal(service.isConnected(), true);
});

test("socket headers include jellyfin authorization metadata", () => {
  const seenHeaders: Record<string, string>[] = [];

  const service = new JellyfinRemoteSessionService({
    serverUrl: "http://jellyfin.local:8096",
    accessToken: "token-auth",
    deviceId: "device-auth",
    clientName: "SubMiner",
    clientVersion: "0.1.0",
    deviceName: "SubMiner",
    socketHeadersFactory: (_url, headers) => {
      seenHeaders.push(headers);
      return new FakeWebSocket() as unknown as any;
    },
    fetchImpl: (async () => new Response(null, { status: 200 })) as typeof fetch,
  });

  service.start();
  assert.equal(seenHeaders.length, 1);
  assert.ok(seenHeaders[0].Authorization.includes('Client="SubMiner"'));
  assert.ok(seenHeaders[0].Authorization.includes('DeviceId="device-auth"'));
  assert.ok(seenHeaders[0]["X-Emby-Authorization"]);
});

test("dispatches inbound Play, Playstate, and GeneralCommand messages", () => {
  const sockets: FakeWebSocket[] = [];
  const playPayloads: unknown[] = [];
  const playstatePayloads: unknown[] = [];
  const commandPayloads: unknown[] = [];

  const service = new JellyfinRemoteSessionService({
    serverUrl: "http://jellyfin.local",
    accessToken: "token-2",
    deviceId: "device-2",
    webSocketFactory: () => {
      const socket = new FakeWebSocket();
      sockets.push(socket);
      return socket as unknown as any;
    },
    fetchImpl: (async () => new Response(null, { status: 200 })) as typeof fetch,
    onPlay: (payload) => playPayloads.push(payload),
    onPlaystate: (payload) => playstatePayloads.push(payload),
    onGeneralCommand: (payload) => commandPayloads.push(payload),
  });

  service.start();
  const socket = sockets[0];
  socket.emit(
    "message",
    JSON.stringify({ MessageType: "Play", Data: { ItemId: "movie-1" } }),
  );
  socket.emit(
    "message",
    JSON.stringify({ MessageType: "Playstate", Data: JSON.stringify({ Command: "Pause" }) }),
  );
  socket.emit(
    "message",
    Buffer.from(
      JSON.stringify({
        MessageType: "GeneralCommand",
        Data: { Name: "DisplayMessage" },
      }),
      "utf8",
    ),
  );

  assert.deepEqual(playPayloads, [{ ItemId: "movie-1" }]);
  assert.deepEqual(playstatePayloads, [{ Command: "Pause" }]);
  assert.deepEqual(commandPayloads, [{ Name: "DisplayMessage" }]);
});

test("schedules reconnect with bounded exponential backoff", () => {
  const sockets: FakeWebSocket[] = [];
  const delays: number[] = [];
  const pendingTimers: Array<() => void> = [];

  const service = new JellyfinRemoteSessionService({
    serverUrl: "http://jellyfin.local",
    accessToken: "token-3",
    deviceId: "device-3",
    webSocketFactory: () => {
      const socket = new FakeWebSocket();
      sockets.push(socket);
      return socket as unknown as any;
    },
    fetchImpl: (async () => new Response(null, { status: 200 })) as typeof fetch,
    reconnectBaseDelayMs: 100,
    reconnectMaxDelayMs: 400,
    setTimer: ((handler: () => void, delay?: number) => {
      pendingTimers.push(handler);
      delays.push(Number(delay));
      return pendingTimers.length as unknown as ReturnType<typeof setTimeout>;
    }) as typeof setTimeout,
    clearTimer: (() => {
      return;
    }) as typeof clearTimeout,
  });

  service.start();
  sockets[0].emit("close");
  pendingTimers.shift()?.();
  sockets[1].emit("close");
  pendingTimers.shift()?.();
  sockets[2].emit("close");
  pendingTimers.shift()?.();
  sockets[3].emit("close");

  assert.deepEqual(delays, [100, 200, 400, 400]);
  assert.equal(sockets.length, 4);
});

test("Jellyfin remote stop prevents further reconnect/network activity", () => {
  const sockets: FakeWebSocket[] = [];
  const fetchCalls: Array<{ input: string; init: RequestInit }> = [];
  const pendingTimers: Array<() => void> = [];
  const clearedTimers: unknown[] = [];

  const service = new JellyfinRemoteSessionService({
    serverUrl: "http://jellyfin.local",
    accessToken: "token-stop",
    deviceId: "device-stop",
    webSocketFactory: () => {
      const socket = new FakeWebSocket();
      sockets.push(socket);
      return socket as unknown as any;
    },
    fetchImpl: (async (input, init) => {
      fetchCalls.push({ input: String(input), init: init ?? {} });
      return new Response(null, { status: 200 });
    }) as typeof fetch,
    setTimer: ((handler: () => void) => {
      pendingTimers.push(handler);
      return pendingTimers.length as unknown as ReturnType<typeof setTimeout>;
    }) as typeof setTimeout,
    clearTimer: ((timer) => {
      clearedTimers.push(timer);
    }) as typeof clearTimeout,
  });

  service.start();
  assert.equal(sockets.length, 1);
  sockets[0].emit("close");
  assert.equal(pendingTimers.length, 1);

  service.stop();
  for (const reconnect of pendingTimers) reconnect();

  assert.ok(clearedTimers.length >= 1);
  assert.equal(sockets.length, 1);
  assert.equal(fetchCalls.length, 0);
  assert.equal(service.isConnected(), false);
});

test("reportProgress posts timeline payload and treats failure as non-fatal", async () => {
  const sockets: FakeWebSocket[] = [];
  const fetchCalls: Array<{ input: string; init: RequestInit }> = [];
  let shouldFailTimeline = false;

  const service = new JellyfinRemoteSessionService({
    serverUrl: "http://jellyfin.local",
    accessToken: "token-4",
    deviceId: "device-4",
    webSocketFactory: () => {
      const socket = new FakeWebSocket();
      sockets.push(socket);
      return socket as unknown as any;
    },
    fetchImpl: (async (input, init) => {
      fetchCalls.push({ input: String(input), init: init ?? {} });
      if (String(input).endsWith("/Sessions/Playing/Progress") && shouldFailTimeline) {
        return new Response("boom", { status: 500 });
      }
      return new Response(null, { status: 200 });
    }) as typeof fetch,
  });

  service.start();
  sockets[0].emit("open");
  await new Promise((resolve) => setTimeout(resolve, 0));

  const expectedPayload = buildJellyfinTimelinePayload({
    itemId: "movie-2",
    positionTicks: 123456,
    isPaused: true,
    volumeLevel: 33,
    audioStreamIndex: 1,
    subtitleStreamIndex: 2,
  });
  const expectedPostedPayload = JSON.parse(JSON.stringify(expectedPayload));

  const ok = await service.reportProgress({
    itemId: "movie-2",
    positionTicks: 123456,
    isPaused: true,
    volumeLevel: 33,
    audioStreamIndex: 1,
    subtitleStreamIndex: 2,
  });
  shouldFailTimeline = true;
  const failed = await service.reportProgress({
    itemId: "movie-2",
    positionTicks: 999,
  });

  const timelineCall = fetchCalls.find((call) =>
    call.input.endsWith("/Sessions/Playing/Progress"),
  );
  assert.ok(timelineCall);
  assert.equal(ok, true);
  assert.equal(failed, false);
  assert.ok(typeof timelineCall.init.body === "string");
  assert.deepEqual(
    JSON.parse(String(timelineCall.init.body)),
    expectedPostedPayload,
  );
});

test("advertiseNow validates server registration using Sessions endpoint", async () => {
  const sockets: FakeWebSocket[] = [];
  const calls: string[] = [];
  const service = new JellyfinRemoteSessionService({
    serverUrl: "http://jellyfin.local",
    accessToken: "token-5",
    deviceId: "device-5",
    webSocketFactory: () => {
      const socket = new FakeWebSocket();
      sockets.push(socket);
      return socket as unknown as any;
    },
    fetchImpl: (async (input) => {
      const url = String(input);
      calls.push(url);
      if (url.endsWith("/Sessions")) {
        return new Response(
          JSON.stringify([{ DeviceId: "device-5" }]),
          { status: 200 },
        );
      }
      return new Response(null, { status: 200 });
    }) as typeof fetch,
  });

  service.start();
  sockets[0].emit("open");
  const ok = await service.advertiseNow();
  assert.equal(ok, true);
  assert.ok(calls.some((url) => url.endsWith("/Sessions")));
});
