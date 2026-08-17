import test from 'node:test';
import assert from 'node:assert/strict';
import { EventEmitter } from 'node:events';
import {
  MpvIpcClient,
  MpvIpcClientDeps,
  MpvIpcClientProtocolDeps,
  MPV_REQUEST_ID_SECONDARY_SUB_VISIBILITY,
} from './mpv';
import {
  MPV_REQUEST_ID_TRACK_LIST_AUDIO,
  MPV_REQUEST_ID_TRACK_LIST_SECONDARY,
} from './mpv-protocol';

function makeDeps(overrides: Partial<MpvIpcClientProtocolDeps> = {}): MpvIpcClientDeps {
  return {
    getResolvedConfig: () => ({}) as any,
    autoStartOverlay: false,
    setOverlayVisible: () => {},
    isVisibleOverlayVisible: () => false,
    getReconnectTimer: () => null,
    setReconnectTimer: () => {},
    ...overrides,
  };
}

function captureWarnLogs(run: () => void): string[] {
  const originalWarn = console.warn;
  const originalLogLevel = process.env.SUBMINER_LOG_LEVEL;
  const originalAppLog = process.env.SUBMINER_APP_LOG;
  const messages: string[] = [];

  console.warn = (...args: unknown[]) => {
    messages.push(args.map(String).join(' '));
  };
  process.env.SUBMINER_LOG_LEVEL = 'warn';
  process.env.SUBMINER_APP_LOG = process.platform === 'win32' ? 'NUL' : '/dev/null';

  try {
    run();
  } finally {
    console.warn = originalWarn;
    if (originalLogLevel === undefined) {
      delete process.env.SUBMINER_LOG_LEVEL;
    } else {
      process.env.SUBMINER_LOG_LEVEL = originalLogLevel;
    }
    if (originalAppLog === undefined) {
      delete process.env.SUBMINER_APP_LOG;
    } else {
      process.env.SUBMINER_APP_LOG = originalAppLog;
    }
  }

  return messages;
}

function invokeHandleMessage(client: MpvIpcClient, msg: unknown): Promise<void> {
  return (client as unknown as { handleMessage: (msg: unknown) => Promise<void> }).handleMessage(
    msg,
  );
}

test('MpvIpcClient resolves pending request by request_id', async () => {
  const client = new MpvIpcClient('/tmp/mpv.sock', makeDeps());
  let resolved: unknown = null;
  (client as any).pendingRequests.set(1234, (msg: unknown) => {
    resolved = msg;
  });

  await invokeHandleMessage(client, { request_id: 1234, data: 'ok' });

  assert.deepEqual(resolved, { request_id: 1234, data: 'ok' });
  assert.equal((client as any).pendingRequests.size, 0);
});

test('MpvIpcClient handles sub-text property change and broadcasts tokenized subtitle', async () => {
  const events: Array<{ text: string; isOverlayVisible: boolean }> = [];
  const client = new MpvIpcClient('/tmp/mpv.sock', makeDeps());
  client.on('subtitle-change', (payload) => {
    events.push(payload);
  });

  await invokeHandleMessage(client, {
    event: 'property-change',
    name: 'sub-text',
    data: '字幕',
  });

  assert.equal(events.length, 1);
  assert.equal(events[0]!.text, '字幕');
  assert.equal(events[0]!.isOverlayVisible, false);
});

test('MpvIpcClient emits fullscreen property changes', async () => {
  const events: Array<{ fullscreen: boolean }> = [];
  const client = new MpvIpcClient('/tmp/mpv.sock', makeDeps());
  client.on('fullscreen-change', (payload) => {
    events.push(payload);
  });

  await invokeHandleMessage(client, {
    event: 'property-change',
    name: 'fullscreen',
    data: true,
  });

  assert.deepEqual(events, [{ fullscreen: true }]);
});

test('MpvIpcClient clears cached media title when media path changes', async () => {
  const client = new MpvIpcClient('/tmp/mpv.sock', makeDeps());

  await invokeHandleMessage(client, {
    event: 'property-change',
    name: 'media-title',
    data: '[Jellyfin/direct] Episode 1',
  });
  assert.equal(client.currentMediaTitle, '[Jellyfin/direct] Episode 1');

  await invokeHandleMessage(client, {
    event: 'property-change',
    name: 'path',
    data: '/tmp/new-episode.mkv',
  });

  assert.equal(client.currentVideoPath, '/tmp/new-episode.mkv');
  assert.equal(client.currentMediaTitle, null);
});

test('MpvIpcClient skips secondary subtitle autoload when media path is managed', async () => {
  const commands: unknown[] = [];
  const originalSetTimeout = globalThis.setTimeout;
  const client = new MpvIpcClient(
    '/tmp/mpv.sock',
    makeDeps({
      getResolvedConfig: () =>
        ({
          secondarySub: {
            autoLoadSecondarySub: true,
            secondarySubLanguages: ['en'],
          },
        }) as any,
      shouldAutoLoadSecondarySubTrack: () => false,
    } as any),
  );
  (client as any).send = (command: unknown) => {
    commands.push(command);
    return true;
  };
  (globalThis as any).setTimeout = (callback: () => void) => {
    callback();
    return 0;
  };

  try {
    await invokeHandleMessage(client, {
      event: 'property-change',
      name: 'path',
      data: 'http://pve-main:8096/Videos/item/stream',
    });
  } finally {
    globalThis.setTimeout = originalSetTimeout;
  }

  assert.equal(
    commands.some(
      (command) =>
        Array.isArray((command as { command?: unknown[] }).command) &&
        (command as { command: unknown[] }).command[0] === 'get_property' &&
        (command as { command: unknown[] }).command[1] === 'track-list' &&
        (command as { request_id?: number }).request_id === MPV_REQUEST_ID_TRACK_LIST_SECONDARY,
    ),
    false,
  );
});

test('MpvIpcClient parses JSON line protocol in processBuffer', () => {
  const client = new MpvIpcClient('/tmp/mpv.sock', makeDeps());
  const seen: Array<Record<string, unknown>> = [];
  (client as any).handleMessage = (msg: Record<string, unknown>) => {
    seen.push(msg);
  };
  (client as any).buffer =
    '{"event":"property-change","name":"path","data":"a"}\n{"request_id":1,"data":"ok"}\n{"partial":';

  (client as any).processBuffer();

  assert.equal(seen.length, 2);
  assert.equal(seen[0]!.name, 'path');
  assert.equal(seen[1]!.request_id, 1);
  assert.equal((client as any).buffer, '{"partial":');
});

test('MpvIpcClient request rejects when disconnected', async () => {
  const client = new MpvIpcClient('/tmp/mpv.sock', makeDeps());
  await assert.rejects(async () => client.request(['get_property', 'path']), /MPV not connected/);
});

test('MpvIpcClient requestProperty throws on mpv error response', async () => {
  const client = new MpvIpcClient('/tmp/mpv.sock', makeDeps());
  (client as any).request = async () => ({ error: 'property unavailable' });
  await assert.rejects(
    async () => client.requestProperty('path'),
    /Failed to read MPV property 'path': property unavailable/,
  );
});

test('MpvIpcClient connect does not log connect-request at info level', () => {
  const originalLevel = process.env.SUBMINER_LOG_LEVEL;
  const originalInfo = console.info;
  const infoLines: string[] = [];
  process.env.SUBMINER_LOG_LEVEL = 'info';
  console.info = (message?: unknown) => {
    infoLines.push(String(message ?? ''));
  };

  try {
    const client = new MpvIpcClient('/tmp/mpv.sock', makeDeps());
    (client as any).transport.connect = () => {};
    client.connect();
  } finally {
    process.env.SUBMINER_LOG_LEVEL = originalLevel;
    console.info = originalInfo;
  }

  const requestLogs = infoLines.filter((line) => line.includes('MPV IPC connect requested.'));
  assert.equal(requestLogs.length, 0);
});

test('MpvIpcClient connect logs connect-request at debug level', () => {
  const originalLevel = process.env.SUBMINER_LOG_LEVEL;
  const originalDebug = console.debug;
  const debugLines: string[] = [];
  process.env.SUBMINER_LOG_LEVEL = 'debug';
  console.debug = (message?: unknown) => {
    debugLines.push(String(message ?? ''));
  };

  try {
    const client = new MpvIpcClient('/tmp/mpv.sock', makeDeps());
    (client as any).transport.connect = () => {};
    client.connect();
  } finally {
    process.env.SUBMINER_LOG_LEVEL = originalLevel;
    console.debug = originalDebug;
  }

  const requestLogs = debugLines.filter((line) => line.includes('MPV IPC connect requested.'));
  assert.equal(requestLogs.length, 1);
});

test('MpvIpcClient reconnect clears stale connected state and starts a fresh transport connect', () => {
  const client = new MpvIpcClient('/tmp/mpv.sock', makeDeps());
  const calls: string[] = [];
  const connectionChanges: boolean[] = [];
  const resolved: unknown[] = [];
  client.on('connection-change', ({ connected }) => {
    connectionChanges.push(connected);
  });
  (client as any).connected = true;
  (client as any).connecting = false;
  (client as any).socket = {};
  (client as any).pendingRequests.set(10, (message: unknown) => {
    resolved.push(message);
  });
  (client as any).transport.shutdown = () => {
    calls.push('shutdown');
  };
  (client as any).transport.connect = () => {
    calls.push('connect');
  };

  client.reconnect();

  assert.deepEqual(calls, ['shutdown', 'connect']);
  assert.equal(client.connected, false);
  assert.equal((client as any).connecting, true);
  assert.equal((client as any).socket, null);
  assert.deepEqual(connectionChanges, [false]);
  assert.deepEqual(resolved, [{ request_id: 10, error: 'disconnected' }]);
});

test('MpvIpcClient failPendingRequests resolves outstanding requests as disconnected', () => {
  const client = new MpvIpcClient('/tmp/mpv.sock', makeDeps());
  const resolved: unknown[] = [];
  (client as any).pendingRequests.set(10, (msg: unknown) => {
    resolved.push(msg);
  });
  (client as any).pendingRequests.set(11, (msg: unknown) => {
    resolved.push(msg);
  });

  (client as any).failPendingRequests();

  assert.deepEqual(resolved, [
    { request_id: 10, error: 'disconnected' },
    { request_id: 11, error: 'disconnected' },
  ]);
  assert.equal((client as any).pendingRequests.size, 0);
});

test('MpvIpcClient scheduleReconnect schedules timer and invokes connect', () => {
  const timers: Array<ReturnType<typeof setTimeout> | null> = [];
  const client = new MpvIpcClient(
    '/tmp/mpv.sock',
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

test('MpvIpcClient scheduleReconnect clears existing reconnect timer', () => {
  const timers: Array<ReturnType<typeof setTimeout> | null> = [];
  const cleared: Array<ReturnType<typeof setTimeout> | null> = [];
  const existingTimer = {} as ReturnType<typeof setTimeout>;
  const client = new MpvIpcClient(
    '/tmp/mpv.sock',
    makeDeps({
      getReconnectTimer: () => existingTimer,
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
  const originalClearTimeout = globalThis.clearTimeout;
  (globalThis as any).setTimeout = (handler: () => void, _delay: number) => {
    handler();
    return 1 as unknown as ReturnType<typeof setTimeout>;
  };
  (globalThis as any).clearTimeout = (timer: ReturnType<typeof setTimeout> | null) => {
    cleared.push(timer);
  };

  try {
    (client as any).scheduleReconnect();
  } finally {
    (globalThis as any).setTimeout = originalSetTimeout;
    (globalThis as any).clearTimeout = originalClearTimeout;
  }

  assert.equal(cleared.length, 1);
  assert.equal(cleared[0], existingTimer);
  assert.equal(timers.length, 1);
  assert.equal(connectCalled, true);
});

test('MpvIpcClient onClose resolves outstanding requests and schedules reconnect', () => {
  const timers: Array<ReturnType<typeof setTimeout> | null> = [];
  const client = new MpvIpcClient(
    '/tmp/mpv.sock',
    makeDeps({
      getReconnectTimer: () => null,
      setReconnectTimer: (timer) => {
        timers.push(timer);
      },
    }),
  );

  const resolved: Array<unknown> = [];
  (client as any).pendingRequests.set(1, (message: unknown) => {
    resolved.push(message);
  });

  let reconnectConnectCount = 0;
  (client as any).connect = () => {
    reconnectConnectCount += 1;
  };

  const originalSetTimeout = globalThis.setTimeout;
  (globalThis as any).setTimeout = (handler: () => void, _delay: number) => {
    handler();
    return 1 as unknown as ReturnType<typeof setTimeout>;
  };

  try {
    (client as any).transport.callbacks.onClose();
  } finally {
    (globalThis as any).setTimeout = originalSetTimeout;
  }

  assert.equal(resolved.length, 1);
  assert.deepEqual(resolved[0], { request_id: 1, error: 'disconnected' });
  assert.equal(reconnectConnectCount, 1);
  assert.equal(timers.length, 1);
});

test('MpvIpcClient onClose requests app quit for managed playback', () => {
  let quitRequests = 0;
  const client = new MpvIpcClient(
    '/tmp/mpv.sock',
    makeDeps({
      shouldQuitOnMpvShutdown: () => true,
      requestAppQuit: () => {
        quitRequests += 1;
      },
    }),
  );

  (client as any).scheduleReconnect = () => {};

  (client as any).transport.callbacks.onClose();

  assert.equal(quitRequests, 1);
});

test('MpvIpcClient only warns once for repeated post-disconnect socket failures', () => {
  const client = new MpvIpcClient('/tmp/mpv.sock', makeDeps());
  (client as any).send = () => true;
  (client as any).scheduleReconnect = () => {};

  const callbacks = (client as any).transport.callbacks;
  callbacks.onConnect();

  const messages = captureWarnLogs(() => {
    callbacks.onClose();
    for (let index = 0; index < 3; index += 1) {
      const error = Object.assign(new Error('connect ENOENT /tmp/mpv.sock'), {
        code: 'ENOENT',
      });
      callbacks.onError(error);
      callbacks.onClose();
    }
  });

  assert.equal(messages.filter((message) => message.includes('MPV IPC socket closed')).length, 1);
  assert.equal(messages.filter((message) => message.includes('MPV IPC socket error')).length, 1);
});

test('MpvIpcClient warns again after MPV reconnects and disconnects later', () => {
  const client = new MpvIpcClient('/tmp/mpv.sock', makeDeps());
  (client as any).send = () => true;
  (client as any).scheduleReconnect = () => {};

  const callbacks = (client as any).transport.callbacks;
  callbacks.onConnect();

  const messages = captureWarnLogs(() => {
    callbacks.onClose();
    callbacks.onError(Object.assign(new Error('connect ENOENT /tmp/mpv.sock'), { code: 'ENOENT' }));
    callbacks.onClose();
    callbacks.onConnect();
    callbacks.onClose();
    callbacks.onError(Object.assign(new Error('connect ENOENT /tmp/mpv.sock'), { code: 'ENOENT' }));
    callbacks.onClose();
  });

  assert.equal(messages.filter((message) => message.includes('MPV IPC socket closed')).length, 2);
  assert.equal(messages.filter((message) => message.includes('MPV IPC socket error')).length, 2);
});

test('MpvIpcClient reconnect replays property subscriptions and initial state requests', () => {
  const commands: unknown[] = [];
  const client = new MpvIpcClient('/tmp/mpv.sock', makeDeps());
  (client as any).send = (command: unknown) => {
    commands.push(command);
    return true;
  };

  const callbacks = (client as any).transport.callbacks;
  callbacks.onConnect();

  commands.length = 0;
  callbacks.onConnect();

  const hasSecondaryVisibilityReset = commands.some(
    (command) =>
      Array.isArray((command as { command: unknown[] }).command) &&
      (command as { command: unknown[] }).command[0] === 'set_property' &&
      (command as { command: unknown[] }).command[1] === 'secondary-sub-visibility' &&
      (command as { command: unknown[] }).command[2] === 'no',
  );
  const hasTrackSubscription = commands.some(
    (command) =>
      Array.isArray((command as { command: unknown[] }).command) &&
      (command as { command: unknown[] }).command[0] === 'observe_property' &&
      (command as { command: unknown[] }).command[1] === 1 &&
      (command as { command: unknown[] }).command[2] === 'sub-text',
  );
  const hasAssSubtitleSubscription = commands.some(
    (command) =>
      Array.isArray((command as { command: unknown[] }).command) &&
      (command as { command: unknown[] }).command[0] === 'observe_property' &&
      (command as { command: unknown[] }).command[2] === 'sub-text/ass',
  );
  const hasDeprecatedAssSubtitleProperty = commands.some(
    (command) =>
      Array.isArray((command as { command: unknown[] }).command) &&
      (command as { command: unknown[] }).command.includes('sub-text-ass'),
  );
  const hasPathRequest = commands.some(
    (command) =>
      Array.isArray((command as { command: unknown[] }).command) &&
      (command as { command: unknown[] }).command[0] === 'get_property' &&
      (command as { command: unknown[] }).command[1] === 'path',
  );

  assert.equal(hasSecondaryVisibilityReset, true);
  assert.equal(hasTrackSubscription, true);
  assert.equal(hasAssSubtitleSubscription, true);
  assert.equal(hasDeprecatedAssSubtitleProperty, false);
  assert.equal(hasPathRequest, true);
});

test('MpvIpcClient connect does not force primary subtitle visibility from binding path', () => {
  const commands: unknown[] = [];
  const client = new MpvIpcClient(
    '/tmp/mpv.sock',
    makeDeps({
      isVisibleOverlayVisible: () => true,
    }),
  );
  (client as any).send = (command: unknown) => {
    commands.push(command);
    return true;
  };

  const callbacks = (client as any).transport.callbacks;
  callbacks.onConnect();

  const hasPrimaryVisibilityMutation = commands.some(
    (command) =>
      Array.isArray((command as { command: unknown[] }).command) &&
      (command as { command: unknown[] }).command[0] === 'set_property' &&
      (command as { command: unknown[] }).command[1] === 'sub-visibility',
  );
  assert.equal(hasPrimaryVisibilityMutation, false);
});

test('MpvIpcClient snapshots current subtitles before connection side effects can hide them', () => {
  const commands: unknown[] = [];
  const client = new MpvIpcClient('/tmp/mpv.sock', makeDeps());
  (client as any).send = (command: unknown) => {
    commands.push(command);
    return true;
  };
  client.on('connection-change', ({ connected }) => {
    if (connected) {
      client.setSubVisibility(false);
    }
  });

  const callbacks = (client as any).transport.callbacks;
  callbacks.onConnect();

  const firstSubTextSnapshot = commands.findIndex((command) => {
    const args = (command as { command?: unknown[] }).command;
    return Array.isArray(args) && args[0] === 'get_property' && args[1] === 'sub-text';
  });
  const firstPrimaryHide = commands.findIndex((command) => {
    const args = (command as { command?: unknown[] }).command;
    return (
      Array.isArray(args) &&
      args[0] === 'set_property' &&
      args[1] === 'sub-visibility' &&
      (args[2] === false || args[2] === 'no')
    );
  });

  assert.notEqual(firstSubTextSnapshot, -1);
  assert.notEqual(firstPrimaryHide, -1);
  assert.ok(firstSubTextSnapshot < firstPrimaryHide);
});

test('MpvIpcClient setSubVisibility writes compatibility commands for visibility toggle', () => {
  const commands: unknown[] = [];
  const client = new MpvIpcClient('/tmp/mpv.sock', makeDeps());
  (client as any).send = (payload: unknown) => {
    commands.push(payload);
    return true;
  };

  client.setSubVisibility(false);

  assert.deepEqual(commands, [
    {
      command: ['set_property', 'sub-visibility', false],
    },
    {
      command: ['set_property', 'sub-visibility', 'no'],
    },
    {
      command: ['set', 'sub-visibility', 'no'],
    },
  ]);
});

test('MpvIpcClient captures and disables secondary subtitle visibility on request', async () => {
  const commands: unknown[] = [];
  const client = new MpvIpcClient('/tmp/mpv.sock', makeDeps());
  const previous: boolean[] = [];
  client.on('secondary-subtitle-visibility', ({ visible }) => {
    previous.push(visible);
  });

  (client as any).send = (payload: unknown) => {
    commands.push(payload);
    return true;
  };

  await invokeHandleMessage(client, {
    request_id: MPV_REQUEST_ID_SECONDARY_SUB_VISIBILITY,
    data: 'yes',
  });

  assert.deepEqual(previous, [true]);
  assert.deepEqual(commands, [
    {
      command: ['set_property', 'secondary-sub-visibility', 'no'],
    },
  ]);
});

test('MpvIpcClient restorePreviousSecondarySubVisibility restores and clears tracked value', async () => {
  const commands: unknown[] = [];
  const client = new MpvIpcClient('/tmp/mpv.sock', makeDeps());
  const previous: boolean[] = [];
  client.on('secondary-subtitle-visibility', ({ visible }) => {
    previous.push(visible);
  });

  (client as any).send = (payload: unknown) => {
    commands.push(payload);
    return true;
  };

  await invokeHandleMessage(client, {
    request_id: MPV_REQUEST_ID_SECONDARY_SUB_VISIBILITY,
    data: 'yes',
  });
  client.restorePreviousSecondarySubVisibility();

  assert.equal(previous[0], true);
  assert.equal(previous.length, 1);
  assert.deepEqual(commands, [
    {
      command: ['set_property', 'secondary-sub-visibility', 'no'],
    },
    {
      command: ['set_property', 'secondary-sub-visibility', 'yes'],
    },
  ]);

  client.restorePreviousSecondarySubVisibility();
  assert.equal(commands.length, 2);
});

test('MpvIpcClient updates current audio stream index from track list', async () => {
  const client = new MpvIpcClient('/tmp/mpv.sock', makeDeps());

  await invokeHandleMessage(client, {
    event: 'property-change',
    name: 'aid',
    data: 3,
  });
  await invokeHandleMessage(client, {
    request_id: MPV_REQUEST_ID_TRACK_LIST_AUDIO,
    data: [
      { type: 'sub', id: 5 },
      { type: 'audio', id: 1, selected: false, 'ff-index': 7 },
      { type: 'audio', id: 3, selected: false, 'ff-index': 11 },
      { type: 'audio', id: 4, selected: true, 'ff-index': 9 },
    ],
  });

  assert.equal(client.currentAudioStreamIndex, 11);
});

test('MpvIpcClient playNextSubtitle starts playback from paused state and auto-pauses at end', async () => {
  const commands: unknown[] = [];
  const client = new MpvIpcClient('/tmp/mpv.sock', makeDeps());
  (client as any).send = (payload: unknown) => {
    commands.push(payload);
    return true;
  };
  (client as any).pendingPauseAtSubEnd = true;
  (client as any).pauseAtTime = 42;

  await invokeHandleMessage(client, {
    event: 'property-change',
    name: 'pause',
    data: true,
  });

  client.playNextSubtitle();

  assert.equal((client as any).pendingPauseAtSubEnd, true);
  assert.equal((client as any).pauseAtTime, null);
  assert.deepEqual(commands, [
    { command: ['sub-seek', 1] },
    { command: ['set_property', 'pause', false] },
  ]);
});

test('MpvIpcClient playNextSubtitle starts playback when pause state is unknown', () => {
  const commands: unknown[] = [];
  const client = new MpvIpcClient('/tmp/mpv.sock', makeDeps());
  (client as any).send = (payload: unknown) => {
    commands.push(payload);
    return true;
  };

  client.playNextSubtitle();

  assert.equal((client as any).pendingPauseAtSubEnd, true);
  assert.deepEqual(commands, [
    { command: ['sub-seek', 1] },
    { command: ['set_property', 'pause', false] },
  ]);
});

test('MpvIpcClient playNextSubtitle still auto-pauses at end while already playing', async () => {
  const commands: unknown[] = [];
  const client = new MpvIpcClient('/tmp/mpv.sock', makeDeps());
  (client as any).send = (payload: unknown) => {
    commands.push(payload);
    return true;
  };

  await invokeHandleMessage(client, {
    event: 'property-change',
    name: 'pause',
    data: false,
  });

  client.playNextSubtitle();

  assert.equal((client as any).pendingPauseAtSubEnd, true);
  assert.deepEqual(commands, [{ command: ['sub-seek', 1] }]);
});

class HangingTestSocket extends EventEmitter {
  public connectedPaths: string[] = [];
  public destroyed = false;

  connect(path: string): void {
    this.connectedPaths.push(path);
    // Never resolves: models a stalled named-pipe dial.
  }

  write(): boolean {
    return true;
  }

  destroy(): void {
    this.destroyed = true;
  }
}

test('MpvIpcClient.setSocketPath aborts an in-flight connect so the next dial targets the new path', () => {
  const sockets: HangingTestSocket[] = [];
  const client = new MpvIpcClient(
    '/tmp/mpv-old.sock',
    makeDeps({
      socketFactory: () => {
        const socket = new HangingTestSocket();
        sockets.push(socket);
        return socket as unknown as import('node:net').Socket;
      },
    }),
  );

  client.connect();
  assert.equal(sockets.length, 1);
  assert.equal(sockets[0]!.connectedPaths.at(0), '/tmp/mpv-old.sock');
  assert.equal((client as any).connecting, true);

  client.setSocketPath('/tmp/mpv-new.sock');
  assert.equal((client as any).connecting, false);
  assert.equal(sockets[0]!.destroyed, true);

  client.connect();
  assert.equal(sockets.length, 2);
  assert.equal(sockets[1]!.connectedPaths.at(0), '/tmp/mpv-new.sock');

  (client as any).transport.shutdown();
});
