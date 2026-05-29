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

test('buildHyprlandPlacementDispatches floats tiled overlay windows without pinning them', () => {
  assert.deepEqual(
    buildHyprlandPlacementDispatches({
      address: '0xabc',
      floating: false,
      pinned: false,
    }),
    [
      ['dispatch', 'setfloating', 'address:0xabc'],
      ['dispatch', 'alterzorder', 'top,address:0xabc'],
    ],
  );
});

test('buildHyprlandPlacementDispatches force-aligns floating overlay windows to target bounds', () => {
  assert.deepEqual(
    buildHyprlandPlacementDispatches(
      {
        address: '0xabc',
        floating: true,
        pinned: false,
      },
      {
        x: 0,
        y: 0,
        width: 1920,
        height: 1080,
      },
    ),
    [
      ['dispatch', 'movewindowpixel', 'exact 0 0,address:0xabc'],
      ['dispatch', 'resizewindowpixel', 'exact 1920 1080,address:0xabc'],
      ['dispatch', 'setprop', 'address:0xabc rounding 0'],
      ['dispatch', 'setprop', 'address:0xabc border_size 0'],
      ['dispatch', 'setprop', 'address:0xabc no_shadow 1'],
      ['dispatch', 'setprop', 'address:0xabc no_blur 1'],
      ['dispatch', 'setprop', 'address:0xabc decorate 0'],
      ['dispatch', 'alterzorder', 'top,address:0xabc'],
    ],
  );
});

test('buildHyprlandPlacementDispatches emits Lua dispatchers for Lua-config Hyprland sessions', () => {
  assert.deepEqual(
    buildHyprlandPlacementDispatches(
      {
        address: '0xabc',
        floating: false,
        pinned: true,
      },
      {
        x: 0,
        y: 0,
        width: 1920,
        height: 1080,
      },
      {
        configProvider: 'lua',
      },
    ),
    [
      ['dispatch', 'hl.dsp.window.float({ action = "on", window = "address:0xabc" })'],
      ['dispatch', 'hl.dsp.window.pin({ action = "off", window = "address:0xabc" })'],
      ['dispatch', 'hl.dsp.window.move({ x = 0, y = 0, window = "address:0xabc" })'],
      ['dispatch', 'hl.dsp.window.resize({ x = 1920, y = 1080, window = "address:0xabc" })'],
      [
        'dispatch',
        'hl.dsp.window.set_prop({ prop = "rounding", value = "0", window = "address:0xabc" })',
      ],
      [
        'dispatch',
        'hl.dsp.window.set_prop({ prop = "border_size", value = "0", window = "address:0xabc" })',
      ],
      [
        'dispatch',
        'hl.dsp.window.set_prop({ prop = "no_shadow", value = "1", window = "address:0xabc" })',
      ],
      [
        'dispatch',
        'hl.dsp.window.set_prop({ prop = "no_blur", value = "1", window = "address:0xabc" })',
      ],
      [
        'dispatch',
        'hl.dsp.window.set_prop({ prop = "decorate", value = "0", window = "address:0xabc" })',
      ],
      ['dispatch', 'hl.dsp.window.alter_zorder({ mode = "top", window = "address:0xabc" })'],
    ],
  );
});

test('buildHyprlandPlacementDispatches does not pin already floating overlay windows', () => {
  assert.deepEqual(
    buildHyprlandPlacementDispatches({
      address: '0xabc',
      floating: true,
      pinned: false,
    }),
    [['dispatch', 'alterzorder', 'top,address:0xabc']],
  );
});

test('buildHyprlandPlacementDispatches can update placement without raising z-order', () => {
  const buildDispatches = buildHyprlandPlacementDispatches as (
    client: Parameters<typeof buildHyprlandPlacementDispatches>[0],
    bounds: Parameters<typeof buildHyprlandPlacementDispatches>[1],
    options: { promote: false },
  ) => string[][];

  assert.deepEqual(
    buildDispatches(
      {
        address: '0xabc',
        floating: true,
        pinned: false,
      },
      {
        x: 0,
        y: 0,
        width: 1920,
        height: 1080,
      },
      { promote: false },
    ),
    [
      ['dispatch', 'movewindowpixel', 'exact 0 0,address:0xabc'],
      ['dispatch', 'resizewindowpixel', 'exact 1920 1080,address:0xabc'],
      ['dispatch', 'setprop', 'address:0xabc rounding 0'],
      ['dispatch', 'setprop', 'address:0xabc border_size 0'],
      ['dispatch', 'setprop', 'address:0xabc no_shadow 1'],
      ['dispatch', 'setprop', 'address:0xabc no_blur 1'],
      ['dispatch', 'setprop', 'address:0xabc decorate 0'],
    ],
  );
});

test('buildHyprlandPlacementDispatches unpins previously pinned overlay windows', () => {
  assert.deepEqual(
    buildHyprlandPlacementDispatches({
      address: '0xabc',
      floating: true,
      pinned: true,
    }),
    [
      ['dispatch', 'pin', 'address:0xabc'],
      ['dispatch', 'alterzorder', 'top,address:0xabc'],
    ],
  );
});

test('ensureHyprlandWindowFloatingByTitle dispatches float-only placement for matching tiled window', () => {
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
      if (args.join(' ') === '-j status') {
        return JSON.stringify({ configProvider: 'hyprlang' });
      }
      return '';
    }) as never,
  });

  assert.equal(placed, true);
  assert.deepEqual(
    calls.map(([, args]) => args),
    [
      ['-j', 'clients'],
      ['-j', 'status'],
      ['dispatch', 'setfloating', 'address:0xmatch'],
      ['dispatch', 'alterzorder', 'top,address:0xmatch'],
    ],
  );
});

test('ensureHyprlandWindowFloatingByTitle dispatches exact Hyprland geometry when bounds are provided', () => {
  const calls: unknown[][] = [];
  const placed = ensureHyprlandWindowFloatingByTitle({
    title: 'SubMiner Stats',
    platform: 'linux',
    env: {
      HYPRLAND_INSTANCE_SIGNATURE: 'abc',
    },
    pid: 456,
    bounds: {
      x: 0,
      y: 0,
      width: 1920,
      height: 1080,
    },
    execFileSync: ((command: string, args: string[], options: unknown) => {
      calls.push([command, args, options]);
      if (args.join(' ') === '-j clients') {
        return JSON.stringify([
          {
            address: '0xmatch',
            pid: 456,
            title: 'SubMiner Stats',
            mapped: true,
            floating: true,
            pinned: false,
          },
        ]);
      }
      if (args.join(' ') === '-j status') {
        return JSON.stringify({ configProvider: 'hyprlang' });
      }
      return '';
    }) as never,
  });

  assert.equal(placed, true);
  assert.deepEqual(
    calls.map(([, args]) => args),
    [
      ['-j', 'clients'],
      ['-j', 'status'],
      ['dispatch', 'movewindowpixel', 'exact 0 0,address:0xmatch'],
      ['dispatch', 'resizewindowpixel', 'exact 1920 1080,address:0xmatch'],
      ['dispatch', 'setprop', 'address:0xmatch rounding 0'],
      ['dispatch', 'setprop', 'address:0xmatch border_size 0'],
      ['dispatch', 'setprop', 'address:0xmatch no_shadow 1'],
      ['dispatch', 'setprop', 'address:0xmatch no_blur 1'],
      ['dispatch', 'setprop', 'address:0xmatch decorate 0'],
      ['dispatch', 'alterzorder', 'top,address:0xmatch'],
    ],
  );
});

test('ensureHyprlandWindowFloatingByTitle dispatches Lua syntax for Lua-config Hyprland sessions', () => {
  const calls: unknown[][] = [];
  const placed = ensureHyprlandWindowFloatingByTitle({
    title: 'SubMiner Stats',
    platform: 'linux',
    env: {
      HYPRLAND_INSTANCE_SIGNATURE: 'abc',
    },
    pid: 456,
    bounds: {
      x: 0,
      y: 0,
      width: 1920,
      height: 1080,
    },
    execFileSync: ((command: string, args: string[], options: unknown) => {
      calls.push([command, args, options]);
      if (args.join(' ') === '-j clients') {
        return JSON.stringify([
          {
            address: '0xmatch',
            pid: 456,
            title: 'SubMiner Stats',
            mapped: true,
            floating: true,
            pinned: false,
          },
        ]);
      }
      if (args.join(' ') === '-j status') {
        return JSON.stringify({ configProvider: 'lua' });
      }
      return '';
    }) as never,
  });

  assert.equal(placed, true);
  assert.deepEqual(
    calls.map(([, args]) => args),
    [
      ['-j', 'clients'],
      ['-j', 'status'],
      ['dispatch', 'hl.dsp.window.move({ x = 0, y = 0, window = "address:0xmatch" })'],
      ['dispatch', 'hl.dsp.window.resize({ x = 1920, y = 1080, window = "address:0xmatch" })'],
      [
        'dispatch',
        'hl.dsp.window.set_prop({ prop = "rounding", value = "0", window = "address:0xmatch" })',
      ],
      [
        'dispatch',
        'hl.dsp.window.set_prop({ prop = "border_size", value = "0", window = "address:0xmatch" })',
      ],
      [
        'dispatch',
        'hl.dsp.window.set_prop({ prop = "no_shadow", value = "1", window = "address:0xmatch" })',
      ],
      [
        'dispatch',
        'hl.dsp.window.set_prop({ prop = "no_blur", value = "1", window = "address:0xmatch" })',
      ],
      [
        'dispatch',
        'hl.dsp.window.set_prop({ prop = "decorate", value = "0", window = "address:0xmatch" })',
      ],
      ['dispatch', 'hl.dsp.window.alter_zorder({ mode = "top", window = "address:0xmatch" })'],
    ],
  );
});
