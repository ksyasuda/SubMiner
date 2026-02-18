import test from "node:test";
import assert from "node:assert/strict";
import * as net from "node:net";
import { EventEmitter } from "node:events";
import {
  getMpvReconnectDelay,
  MpvSocketMessagePayload,
  MpvSocketTransport,
  scheduleMpvReconnect,
} from "./mpv-transport";

class FakeSocket extends EventEmitter {
  public connectedPaths: string[] = [];
  public writePayloads: string[] = [];
  public destroyed = false;

  connect(path: string): void {
    this.connectedPaths.push(path);
    setTimeout(() => {
      this.emit("connect");
    }, 0);
  }

  write(payload: string): boolean {
    this.writePayloads.push(payload);
    return true;
  }

  destroy(): void {
    this.destroyed = true;
    this.emit("close");
  }
}

const wait = () => new Promise((resolve) => setTimeout(resolve, 0));

test("getMpvReconnectDelay follows existing reconnect ramp", () => {
  assert.equal(getMpvReconnectDelay(0, true), 1000);
  assert.equal(getMpvReconnectDelay(1, true), 1000);
  assert.equal(getMpvReconnectDelay(2, true), 2000);
  assert.equal(getMpvReconnectDelay(4, true), 5000);
  assert.equal(getMpvReconnectDelay(7, true), 10000);

  assert.equal(getMpvReconnectDelay(0, false), 200);
  assert.equal(getMpvReconnectDelay(2, false), 500);
  assert.equal(getMpvReconnectDelay(4, false), 1000);
  assert.equal(getMpvReconnectDelay(6, false), 2000);
});

test("scheduleMpvReconnect clears existing timer and increments attempt", () => {
  const existing = {} as ReturnType<typeof setTimeout>;
  const cleared: Array<ReturnType<typeof setTimeout> | null> = [];
  const setTimers: Array<ReturnType<typeof setTimeout> | null> = [];
  const calls: Array<{ attempt: number; delay: number }> = [];
  let connected = 0;

  const originalSetTimeout = globalThis.setTimeout;
  const originalClearTimeout = globalThis.clearTimeout;
  (globalThis as any).setTimeout = (handler: () => void, _delay: number) => {
    handler();
    return 1 as unknown as ReturnType<typeof setTimeout>;
  };
  (globalThis as any).clearTimeout = (
    timer: ReturnType<typeof setTimeout> | null,
  ) => {
    cleared.push(timer);
  };

  const nextAttempt = scheduleMpvReconnect({
    attempt: 3,
    hasConnectedOnce: true,
    getReconnectTimer: () => existing,
    setReconnectTimer: (timer) => {
      setTimers.push(timer);
    },
    onReconnectAttempt: (attempt, delay) => {
      calls.push({ attempt, delay });
    },
    connect: () => {
      connected += 1;
    },
  });

  (globalThis as any).setTimeout = originalSetTimeout;
  (globalThis as any).clearTimeout = originalClearTimeout;

  assert.equal(nextAttempt, 4);
  assert.equal(cleared.length, 1);
  assert.equal(cleared[0], existing);
  assert.equal(setTimers.length, 1);
  assert.equal(calls.length, 1);
  assert.equal(calls[0].attempt, 4);
  assert.equal(calls[0].delay, getMpvReconnectDelay(3, true));
  assert.equal(connected, 1);
});

test("MpvSocketTransport connects and sends payloads over a live socket", async () => {
  const events: string[] = [];
  const transport = new MpvSocketTransport({
    socketPath: "/tmp/mpv.sock",
    onConnect: () => {
      events.push("connect");
    },
    onData: () => {
      events.push("data");
    },
    onError: () => {
      events.push("error");
    },
    onClose: () => {
      events.push("close");
    },
    socketFactory: () => new FakeSocket() as unknown as net.Socket,
  });

  const payload: MpvSocketMessagePayload = {
    command: ["sub-seek", 1],
    request_id: 1,
  };

  assert.equal(transport.send(payload), false);

  transport.connect();
  await wait();

  assert.equal(events.includes("connect"), true);
  assert.equal(transport.send(payload), true);

  const fakeSocket = transport.getSocket() as unknown as FakeSocket;
  assert.equal(fakeSocket.connectedPaths.at(0), "/tmp/mpv.sock");
  assert.equal(fakeSocket.writePayloads.length, 1);
  assert.equal(fakeSocket.writePayloads.at(0), `${JSON.stringify(payload)}\n`);
});

test("MpvSocketTransport reports lifecycle transitions and callback order", async () => {
  const events: string[] = [];
  const fakeError = new Error("boom");
  const transport = new MpvSocketTransport({
    socketPath: "/tmp/mpv.sock",
    onConnect: () => {
      events.push("connect");
    },
    onData: () => {
      events.push("data");
    },
    onError: () => {
      events.push("error");
    },
    onClose: () => {
      events.push("close");
    },
    socketFactory: () => new FakeSocket() as unknown as net.Socket,
  });

  transport.connect();
  await wait();

  const socket = transport.getSocket() as unknown as FakeSocket;
  socket.emit("error", fakeError);
  socket.emit("data", Buffer.from("{}"));
  socket.destroy();
  await wait();

  assert.equal(events.includes("connect"), true);
  assert.equal(events.includes("data"), true);
  assert.equal(events.includes("error"), true);
  assert.equal(events.includes("close"), true);
  assert.equal(transport.isConnected, false);
  assert.equal(transport.isConnecting, false);
  assert.equal(socket.destroyed, true);
});

test("MpvSocketTransport ignores connect requests while already connecting or connected", async () => {
  const events: string[] = [];
  const transport = new MpvSocketTransport({
    socketPath: "/tmp/mpv.sock",
    onConnect: () => {
      events.push("connect");
    },
    onData: () => {
      events.push("data");
    },
    onError: () => {
      events.push("error");
    },
    onClose: () => {
      events.push("close");
    },
    socketFactory: () => new FakeSocket() as unknown as net.Socket,
  });

  transport.connect();
  transport.connect();
  await wait();

  assert.equal(events.includes("connect"), true);
  const socket = transport.getSocket() as unknown as FakeSocket;
  socket.emit("close");
  await wait();

  transport.connect();
  await wait();

  assert.equal(events.filter((entry) => entry === "connect").length, 2);
});

test("MpvSocketTransport.shutdown clears socket and lifecycle flags", async () => {
  const transport = new MpvSocketTransport({
    socketPath: "/tmp/mpv.sock",
    onConnect: () => {},
    onData: () => {},
    onError: () => {},
    onClose: () => {},
    socketFactory: () => new FakeSocket() as unknown as net.Socket,
  });

  transport.connect();
  await wait();
  assert.equal(transport.isConnected, true);

  transport.shutdown();
  assert.equal(transport.isConnected, false);
  assert.equal(transport.isConnecting, false);
  assert.equal(transport.getSocket(), null);
});
