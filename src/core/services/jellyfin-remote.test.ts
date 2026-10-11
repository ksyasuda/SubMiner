import test from 'node:test';
import assert from 'node:assert/strict';
import { JellyfinRemoteSessionService } from './jellyfin-remote';

class FakeWebSocket {
  private listeners: Record<string, Array<(...args: unknown[]) => void>> = {};
  sent: string[] = [];
  terminated = false;

  send(data: string): void {
    this.sent.push(data);
  }

  terminate(): void {
    this.terminated = true;
    this.emit('close');
  }

  on(event: string, listener: (...args: unknown[]) => void): this {
    if (!this.listeners[event]) {
      this.listeners[event] = [];
    }
    this.listeners[event].push(listener);
    return this;
  }

  close(): void {
    this.emit('close');
  }

  emit(event: string, ...args: unknown[]): void {
    for (const listener of this.listeners[event] ?? []) {
      listener(...args);
    }
  }
}

test('Jellyfin remote service has no traffic until started', async () => {
  let socketCreateCount = 0;
  const fetchCalls: Array<{ input: string; init: RequestInit }> = [];

  const service = new JellyfinRemoteSessionService({
    serverUrl: 'http://jellyfin.local:8096',
    accessToken: 'token-0',
    deviceId: 'device-0',
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

test('start posts capabilities on socket connect', async () => {
  const sockets: FakeWebSocket[] = [];
  const fetchCalls: Array<{ input: string; init: RequestInit }> = [];

  const service = new JellyfinRemoteSessionService({
    serverUrl: 'http://jellyfin.local:8096',
    accessToken: 'token-1',
    deviceId: 'device-1',
    webSocketFactory: (url) => {
      assert.equal(url, 'ws://jellyfin.local:8096/socket?ApiKey=token-1&deviceId=device-1');
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
  sockets[0]!.emit('open');
  await new Promise((resolve) => setTimeout(resolve, 0));

  assert.equal(fetchCalls.length, 1);
  assert.equal(fetchCalls[0]!.input, 'http://jellyfin.local:8096/Sessions/Capabilities/Full');
  assert.equal(service.isConnected(), true);
});

test('socket headers include jellyfin authorization metadata', () => {
  const seenHeaders: Record<string, string>[] = [];

  const service = new JellyfinRemoteSessionService({
    serverUrl: 'http://jellyfin.local:8096',
    accessToken: 'token-auth',
    deviceId: 'device-auth',
    clientName: 'SubMiner',
    clientVersion: '0.1.0',
    deviceName: 'SubMiner',
    socketHeadersFactory: (_url, headers) => {
      seenHeaders.push(headers);
      return new FakeWebSocket() as unknown as any;
    },
    fetchImpl: (async () => new Response(null, { status: 200 })) as typeof fetch,
  });

  service.start();
  assert.equal(seenHeaders.length, 1);
  assert.ok(seenHeaders[0]!['Authorization']!.includes('Client="SubMiner"'));
  assert.ok(seenHeaders[0]!['Authorization']!.includes('DeviceId="device-auth"'));
  assert.equal('X-Emby-Authorization' in seenHeaders[0]!, false);
  assert.equal('X-Emby-Token' in seenHeaders[0]!, false);
});

test('dispatches inbound Play, Playstate, and GeneralCommand messages', () => {
  const sockets: FakeWebSocket[] = [];
  const playPayloads: unknown[] = [];
  const playstatePayloads: unknown[] = [];
  const commandPayloads: unknown[] = [];

  const service = new JellyfinRemoteSessionService({
    serverUrl: 'http://jellyfin.local',
    accessToken: 'token-2',
    deviceId: 'device-2',
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
  const socket = sockets[0]!;
  socket.emit('message', JSON.stringify({ MessageType: 'Play', Data: { ItemId: 'movie-1' } }));
  socket.emit(
    'message',
    JSON.stringify({ MessageType: 'Playstate', Data: JSON.stringify({ Command: 'Pause' }) }),
  );
  socket.emit(
    'message',
    Buffer.from(
      JSON.stringify({
        MessageType: 'GeneralCommand',
        Data: { Name: 'DisplayMessage' },
      }),
      'utf8',
    ),
  );

  assert.deepEqual(playPayloads, [{ ItemId: 'movie-1' }]);
  assert.deepEqual(playstatePayloads, [{ Command: 'Pause' }]);
  assert.deepEqual(commandPayloads, [{ Name: 'DisplayMessage' }]);
});

test('schedules reconnect with bounded exponential backoff', () => {
  const sockets: FakeWebSocket[] = [];
  const delays: number[] = [];
  const pendingTimers: Array<() => void> = [];

  const service = new JellyfinRemoteSessionService({
    serverUrl: 'http://jellyfin.local',
    accessToken: 'token-3',
    deviceId: 'device-3',
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
  sockets[0]!.emit('close');
  pendingTimers.shift()?.();
  sockets[1]!.emit('close');
  pendingTimers.shift()?.();
  sockets[2]!.emit('close');
  pendingTimers.shift()?.();
  sockets[3]!.emit('close');

  assert.deepEqual(delays, [100, 200, 400, 400]);
  assert.equal(sockets.length, 4);
});

test('Jellyfin remote stop prevents further reconnect/network activity', () => {
  const sockets: FakeWebSocket[] = [];
  const fetchCalls: Array<{ input: string; init: RequestInit }> = [];
  const pendingTimers: Array<() => void> = [];
  const clearedTimers: unknown[] = [];

  const service = new JellyfinRemoteSessionService({
    serverUrl: 'http://jellyfin.local',
    accessToken: 'token-stop',
    deviceId: 'device-stop',
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
  sockets[0]!.emit('close');
  assert.equal(pendingTimers.length, 1);

  service.stop();
  for (const reconnect of pendingTimers) reconnect();

  assert.ok(clearedTimers.length >= 1);
  assert.equal(sockets.length, 1);
  assert.equal(fetchCalls.length, 0);
  assert.equal(service.isConnected(), false);
});

test('advertiseNow validates server registration using Sessions endpoint', async () => {
  const sockets: FakeWebSocket[] = [];
  const calls: string[] = [];
  const service = new JellyfinRemoteSessionService({
    serverUrl: 'http://jellyfin.local',
    accessToken: 'token-5',
    deviceId: 'device-5',
    webSocketFactory: () => {
      const socket = new FakeWebSocket();
      sockets.push(socket);
      return socket as unknown as any;
    },
    fetchImpl: (async (input) => {
      const url = String(input);
      calls.push(url);
      if (url.endsWith('/Sessions')) {
        return new Response(JSON.stringify([{ DeviceId: 'device-5' }]), { status: 200 });
      }
      return new Response(null, { status: 200 });
    }) as typeof fetch,
  });

  service.start();
  sockets[0]!.emit('open');
  const ok = await service.advertiseNow();
  assert.equal(ok, true);
  assert.ok(calls.some((url) => url.endsWith('/Sessions')));
});

test('answers ForceKeepAlive with KeepAlive messages on the advertised cadence', () => {
  const sockets: FakeWebSocket[] = [];
  const timers: Array<{ handler: () => void; delay: number }> = [];

  const service = new JellyfinRemoteSessionService({
    serverUrl: 'http://jellyfin.local',
    accessToken: 'token-ka',
    deviceId: 'device-ka',
    webSocketFactory: () => {
      const socket = new FakeWebSocket();
      sockets.push(socket);
      return socket as unknown as any;
    },
    fetchImpl: (async () => new Response(null, { status: 200 })) as typeof fetch,
    setTimer: ((handler: () => void, delay?: number) => {
      timers.push({ handler, delay: Number(delay) });
      return timers.length as unknown as ReturnType<typeof setTimeout>;
    }) as typeof setTimeout,
    clearTimer: (() => undefined) as typeof clearTimeout,
  });

  service.start();
  sockets[0]!.emit('open');
  assert.deepEqual(sockets[0]!.sent, ['{"MessageType":"KeepAlive"}']);
  assert.equal(timers[0]!.delay, 30_000);

  sockets[0]!.emit('message', JSON.stringify({ MessageType: 'ForceKeepAlive', Data: 20 }));
  assert.equal(sockets[0]!.sent.length, 2);
  assert.equal(timers.at(-1)!.delay, 10_000);

  timers.at(-1)!.handler();
  assert.equal(sockets[0]!.sent.length, 3);
});

test('reconnects when the server stops answering keep-alives', () => {
  let now = 1_000_000;
  const sockets: FakeWebSocket[] = [];
  const timers: Array<() => void> = [];
  const warnings: string[] = [];

  const service = new JellyfinRemoteSessionService({
    serverUrl: 'http://jellyfin.local',
    accessToken: 'token-lost',
    deviceId: 'device-lost',
    webSocketFactory: () => {
      const socket = new FakeWebSocket();
      sockets.push(socket);
      return socket as unknown as any;
    },
    fetchImpl: (async () => new Response(null, { status: 200 })) as typeof fetch,
    getNow: () => now,
    logWarn: (message) => {
      warnings.push(message);
    },
    reconnectBaseDelayMs: 100,
    setTimer: ((handler: () => void) => {
      timers.push(handler);
      return timers.length as unknown as ReturnType<typeof setTimeout>;
    }) as typeof setTimeout,
    clearTimer: (() => undefined) as typeof clearTimeout,
  });

  service.start();
  sockets[0]!.emit('open');

  // Two silent ticks are still within the 90s tolerance; the third marks the socket lost.
  now += 30_000;
  timers.shift()!();
  now += 30_000;
  timers.shift()!();
  assert.equal(sockets[0]!.sent.length, 3);
  assert.equal(sockets[0]!.terminated, false);

  now += 30_000;
  timers.shift()!();
  assert.equal(sockets[0]!.terminated, true);
  assert.equal(service.isConnected(), false);
  assert.equal(warnings.length, 1);

  timers.shift()!();
  assert.equal(sockets.length, 2);
});

test('ignores messages from a superseded socket', () => {
  const sockets: FakeWebSocket[] = [];
  const playPayloads: unknown[] = [];

  const service = new JellyfinRemoteSessionService({
    serverUrl: 'http://jellyfin.local',
    accessToken: 'token-stale',
    deviceId: 'device-stale',
    webSocketFactory: () => {
      const socket = new FakeWebSocket();
      sockets.push(socket);
      return socket as unknown as any;
    },
    fetchImpl: (async () => new Response(null, { status: 200 })) as typeof fetch,
    onPlay: (payload) => {
      playPayloads.push(payload);
    },
    setTimer: (() => 1 as unknown as ReturnType<typeof setTimeout>) as unknown as typeof setTimeout,
    clearTimer: (() => undefined) as typeof clearTimeout,
  });

  service.start();
  service.stop();
  service.start();
  sockets[1]!.emit('open');
  assert.equal(sockets.length, 2);

  sockets[0]!.emit('message', JSON.stringify({ MessageType: 'ForceKeepAlive', Data: 10 }));
  sockets[0]!.emit('message', JSON.stringify({ MessageType: 'Play', Data: { ItemIds: ['x'] } }));

  assert.deepEqual(sockets[0]!.sent, []);
  assert.deepEqual(playPayloads, []);
  assert.deepEqual(sockets[1]!.sent, ['{"MessageType":"KeepAlive"}']);
});
