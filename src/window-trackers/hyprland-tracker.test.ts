import test from 'node:test';
import assert from 'node:assert/strict';
import {
  HyprlandWindowTracker,
  isHyprlandGeometryEvent,
  parseHyprctlClients,
  parseHyprctlMonitors,
  resolveHyprlandWindowGeometry,
  selectHyprlandMpvWindow,
  type HyprlandClient,
  type HyprlandMonitor,
} from './hyprland-tracker';

function makeClient(overrides: Partial<HyprlandClient> = {}): HyprlandClient {
  return {
    address: '0x1',
    class: 'mpv',
    initialClass: 'mpv',
    at: [0, 0],
    size: [1280, 720],
    mapped: true,
    hidden: false,
    ...overrides,
  };
}

function makeMonitor(overrides: Partial<HyprlandMonitor> = {}): HyprlandMonitor {
  return {
    id: 0,
    x: 0,
    y: 0,
    width: 1920,
    height: 1080,
    ...overrides,
  };
}

test('selectHyprlandMpvWindow ignores hidden and unmapped mpv clients', () => {
  const selected = selectHyprlandMpvWindow(
    [
      makeClient({
        address: '0xhidden',
        hidden: true,
      }),
      makeClient({
        address: '0xunmapped',
        mapped: false,
      }),
      makeClient({
        address: '0xvisible',
        at: [100, 200],
        size: [1920, 1080],
      }),
    ],
    {
      targetMpvSocketPath: null,
      activeWindowAddress: null,
      getWindowCommandLine: () => null,
    },
  );

  assert.equal(selected?.address, '0xvisible');
});

test('selectHyprlandMpvWindow prefers active visible window among socket matches', () => {
  const commandLines = new Map<string, string>([
    ['10', 'mpv --input-ipc-server=/tmp/subminer.sock first.mkv'],
    ['20', 'mpv --input-ipc-server=/tmp/subminer.sock second.mkv'],
  ]);

  const selected = selectHyprlandMpvWindow(
    [
      makeClient({
        address: '0xfirst',
        pid: 10,
      }),
      makeClient({
        address: '0xsecond',
        pid: 20,
      }),
    ],
    {
      targetMpvSocketPath: '/tmp/subminer.sock',
      activeWindowAddress: '0xsecond',
      getWindowCommandLine: (pid) => commandLines.get(String(pid)) ?? null,
    },
  );

  assert.equal(selected?.address, '0xsecond');
});

test('selectHyprlandMpvWindow matches mpv by initialClass when class is blank', () => {
  const selected = selectHyprlandMpvWindow(
    [
      makeClient({
        address: '0xinitial',
        class: '',
        initialClass: 'mpv',
      }),
    ],
    {
      targetMpvSocketPath: null,
      activeWindowAddress: null,
      getWindowCommandLine: () => null,
    },
  );

  assert.equal(selected?.address, '0xinitial');
});

test('parseHyprctlClients tolerates non-json prefix output', () => {
  const clients = parseHyprctlClients(`ok
[{"address":"0x1","class":"mpv","initialClass":"mpv","at":[1,2],"size":[3,4]}]`);

  assert.deepEqual(clients, [
    {
      address: '0x1',
      class: 'mpv',
      initialClass: 'mpv',
      at: [1, 2],
      size: [3, 4],
    },
  ]);
});

test('parseHyprctlMonitors returns null for malformed JSON output', () => {
  assert.equal(parseHyprctlMonitors('not-json'), null);
  assert.equal(parseHyprctlMonitors('[{"id":0,"x":0,"y":0,"width":1920'), null);
});

test('isHyprlandGeometryEvent treats geometry events as geometry-changing only', () => {
  assert.equal(isHyprlandGeometryEvent('fullscreenv2'), true);
  assert.equal(isHyprlandGeometryEvent('workspacev2'), true);
  assert.equal(isHyprlandGeometryEvent('windowtitle'), false);
  assert.equal(isHyprlandGeometryEvent('windowtitlev2'), false);
  assert.equal(isHyprlandGeometryEvent('activewindowv2'), false);
});

test('resolveHyprlandWindowGeometry uses monitor bounds for fullscreen clients', () => {
  const geometry = resolveHyprlandWindowGeometry(
    makeClient({
      at: [60, 80],
      size: [1280, 720],
      monitor: 1,
      fullscreen: 2,
      fullscreenClient: 2,
    }),
    [
      makeMonitor({ id: 0, x: 0, y: 0, width: 1920, height: 1080 }),
      makeMonitor({ id: 1, x: 1920, y: 0, width: 2560, height: 1440 }),
    ],
  );

  assert.deepEqual(geometry, {
    x: 1920,
    y: 0,
    width: 2560,
    height: 1440,
  });
});

test('resolveHyprlandWindowGeometry uses monitor bounds for client-requested fullscreen', () => {
  const geometry = resolveHyprlandWindowGeometry(
    makeClient({
      at: [0, 28],
      size: [1920, 1052],
      monitor: 0,
      fullscreen: 0,
      fullscreenClient: 2,
    }),
    [makeMonitor({ id: 0, x: 0, y: 0, width: 1920, height: 1080 })],
  );

  assert.deepEqual(geometry, {
    x: 0,
    y: 0,
    width: 1920,
    height: 1080,
  });
});

type PendingCommand = {
  command: string;
  args: string[];
  resolve: (output: string) => void;
};

// Runner whose commands stay pending until the test resolves them.
function makeDeferredRunner() {
  const pending: PendingCommand[] = [];
  const runCommand = (command: string, args: string[]) =>
    new Promise<string>((resolve) => {
      pending.push({ command, args, resolve });
    });
  return { pending, runCommand };
}

function flush(): Promise<void> {
  return new Promise((resolve) => setImmediate(resolve));
}

test('HyprlandWindowTracker re-emits focus callback on active window events for z-order refresh', async () => {
  const calls: string[] = [];
  const tracker = new HyprlandWindowTracker(undefined, async (command, args) => {
    calls.push(`${command} ${args.join(' ')}`);
    return '[]';
  });
  tracker.onWindowFocusChange = (focused) => {
    calls.push(`focus:${focused}`);
  };

  (tracker as unknown as { handleSocketEvent: (event: string) => void }).handleSocketEvent(
    'activewindowv2>>0xmpv',
  );
  await flush();

  assert.deepEqual(calls, ['hyprctl -j clients', 'focus:false']);
});

test('HyprlandWindowTracker coalesces polls requested while hyprctl is still running', async () => {
  const { pending, runCommand } = makeDeferredRunner();
  const tracker = new HyprlandWindowTracker(undefined, runCommand);

  const polls = [tracker.refreshNow(), tracker.refreshNow(), tracker.refreshNow()];
  await flush();
  assert.equal(pending.length, 1);

  pending[0]!.resolve(JSON.stringify([makeClient({ at: [10, 20] })]));
  await flush();
  assert.equal(pending.length, 2);

  pending[1]!.resolve(JSON.stringify([makeClient({ at: [30, 40] })]));
  await Promise.all(polls);
  assert.equal(pending.length, 2);
  assert.deepEqual(tracker.getGeometry(), { x: 30, y: 40, width: 1280, height: 720 });
});

test('HyprlandWindowTracker ignores a poll that resolves after stop', async () => {
  const { pending, runCommand } = makeDeferredRunner();
  const tracker = new HyprlandWindowTracker(undefined, runCommand);
  let found = 0;
  tracker.onWindowFound = () => {
    found += 1;
  };

  const poll = tracker.refreshNow();
  tracker.stop();
  await flush();
  pending[0]!.resolve(JSON.stringify([makeClient()]));
  await poll;

  assert.equal(found, 0);
  assert.equal(tracker.getGeometry(), null);
});

test('HyprlandWindowTracker reads each mpv command line once and re-reads after the pid leaves', async () => {
  let clients = [makeClient({ pid: 10 })];
  const psPids: string[] = [];
  const tracker = new HyprlandWindowTracker('/tmp/subminer.sock', async (command, args) => {
    if (command === 'hyprctl') {
      return JSON.stringify(clients);
    }
    psPids.push(args[1]!);
    return 'mpv --input-ipc-server=/tmp/subminer.sock video.mkv\n';
  });

  await tracker.refreshNow();
  await tracker.refreshNow();
  assert.equal(tracker.isTracking(), true);
  assert.deepEqual(psPids, ['10']);

  clients = [];
  await tracker.refreshNow();
  clients = [makeClient({ pid: 10 })];
  await tracker.refreshNow();
  assert.deepEqual(psPids, ['10', '10']);
});

test('HyprlandWindowTracker still tracks mpv when ps fails for another pid, and retries that pid', async () => {
  const psPids: string[] = [];
  const tracker = new HyprlandWindowTracker('/tmp/subminer.sock', async (command, args) => {
    if (command === 'hyprctl') {
      return JSON.stringify([makeClient({ pid: 10 }), makeClient({ pid: 20 })]);
    }
    psPids.push(args[1]!);
    if (args[1] === '20') {
      throw new Error('ps exited 1');
    }
    return 'mpv --input-ipc-server=/tmp/subminer.sock video.mkv\n';
  });

  await tracker.refreshNow();
  assert.equal(tracker.isTracking(), true);

  await tracker.refreshNow();
  assert.deepEqual(psPids.sort(), ['10', '20', '20']);
});
