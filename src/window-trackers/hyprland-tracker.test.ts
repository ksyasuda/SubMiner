import test from 'node:test';
import assert from 'node:assert/strict';
import {
  selectHyprlandMpvWindow,
  type HyprlandClient,
} from './hyprland-tracker';

function makeClient(overrides: Partial<HyprlandClient> = {}): HyprlandClient {
  return {
    address: '0x1',
    class: 'mpv',
    at: [0, 0],
    size: [1280, 720],
    mapped: true,
    hidden: false,
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
