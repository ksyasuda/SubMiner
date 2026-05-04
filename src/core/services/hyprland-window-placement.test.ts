import assert from 'node:assert/strict';
import test from 'node:test';
import {
  buildHyprlandPlacementDispatches,
  ensureHyprlandWindowFloatingByTitle,
  findHyprlandWindowForPlacement,
  shouldAttemptHyprlandWindowPlacement,
} from './hyprland-window-placement';

test('shouldAttemptHyprlandWindowPlacement only enables on Hyprland Linux sessions', () => {
  assert.equal(
    shouldAttemptHyprlandWindowPlacement('linux', {
      HYPRLAND_INSTANCE_SIGNATURE: 'abc',
    }),
    true,
  );
  assert.equal(
    shouldAttemptHyprlandWindowPlacement('linux', {
      WAYLAND_DISPLAY: 'wayland-1',
    }),
    false,
  );
  assert.equal(
    shouldAttemptHyprlandWindowPlacement('darwin', {
      HYPRLAND_INSTANCE_SIGNATURE: 'abc',
    }),
    false,
  );
});

test('findHyprlandWindowForPlacement matches current process by title', () => {
  const client = findHyprlandWindowForPlacement(
    [
      {
        address: '0xother',
        pid: 123,
        title: 'SubMiner Stats',
        mapped: true,
      },
      {
        address: '0xmatch',
        pid: 456,
        title: 'SubMiner Stats',
        mapped: true,
      },
    ],
    {
      pid: 456,
      title: 'SubMiner Stats',
    },
  );

  assert.equal(client?.address, '0xmatch');
});

test('buildHyprlandPlacementDispatches floats and pins tiled overlay windows', () => {
  assert.deepEqual(
    buildHyprlandPlacementDispatches({
      address: '0xabc',
      floating: false,
      pinned: false,
    }),
    [
      ['dispatch', 'setfloating', 'address:0xabc'],
      ['dispatch', 'pin', 'address:0xabc'],
    ],
  );
});

test('buildHyprlandPlacementDispatches skips already floating and pinned windows', () => {
  assert.deepEqual(
    buildHyprlandPlacementDispatches({
      address: '0xabc',
      floating: true,
      pinned: true,
    }),
    [],
  );
});

test('ensureHyprlandWindowFloatingByTitle dispatches placement for matching tiled window', () => {
  const calls: unknown[][] = [];
  const placed = ensureHyprlandWindowFloatingByTitle({
    title: 'SubMiner Stats',
    platform: 'linux',
    env: {
      HYPRLAND_INSTANCE_SIGNATURE: 'abc',
    },
    pid: 456,
    execFileSync: ((command: string, args: string[], options: unknown) => {
      calls.push([command, args, options]);
      if (args.join(' ') === '-j clients') {
        return JSON.stringify([
          {
            address: '0xmatch',
            pid: 456,
            title: 'SubMiner Stats',
            mapped: true,
            floating: false,
            pinned: false,
          },
        ]);
      }
      return '';
    }) as never,
  });

  assert.equal(placed, true);
  assert.deepEqual(
    calls.map(([, args]) => args),
    [
      ['-j', 'clients'],
      ['dispatch', 'setfloating', 'address:0xmatch'],
      ['dispatch', 'pin', 'address:0xmatch'],
    ],
  );
});
