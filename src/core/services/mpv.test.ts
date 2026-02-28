import test from 'node:test';
import assert from 'node:assert/strict';
import {
  MpvIpcClient,
  MpvIpcClientDeps,
  MpvIpcClientProtocolDeps,
  MPV_REQUEST_ID_SECONDARY_SUB_VISIBILITY,
} from './mpv';
import { MPV_REQUEST_ID_TRACK_LIST_AUDIO } from './mpv-protocol';

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
  const hasPathRequest = commands.some(
    (command) =>
      Array.isArray((command as { command: unknown[] }).command) &&
      (command as { command: unknown[] }).command[0] === 'get_property' &&
      (command as { command: unknown[] }).command[1] === 'path',
  );

  assert.equal(hasSecondaryVisibilityReset, true);
  assert.equal(hasTrackSubscription, true);
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
