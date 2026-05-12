import test from 'node:test';
import assert from 'node:assert/strict';
import {
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
