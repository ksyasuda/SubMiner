import assert from 'node:assert/strict';
import test from 'node:test';
import {
  promoteSettingsWindowAboveOverlay,
  shouldPromoteSettingsWindowAboveOverlay,
} from './settings-window-z-order';

test('settings window promotion only applies to Hyprland sessions', () => {
  assert.equal(
    shouldPromoteSettingsWindowAboveOverlay('linux', { HYPRLAND_INSTANCE_SIGNATURE: 'hypr' }),
    true,
  );
  assert.equal(
    shouldPromoteSettingsWindowAboveOverlay('linux', { WAYLAND_DISPLAY: 'wayland-1' }),
    false,
  );
  assert.equal(
    shouldPromoteSettingsWindowAboveOverlay('darwin', { HYPRLAND_INSTANCE_SIGNATURE: 'hypr' }),
    false,
  );
});

test('promoteSettingsWindowAboveOverlay raises Hyprland settings windows above the overlay', () => {
  const calls: string[] = [];

  const promoted = promoteSettingsWindowAboveOverlay(
    {
      isDestroyed: () => false,
      getTitle: () => 'SubMiner Settings',
      setAlwaysOnTop: (flag: boolean) => calls.push(`always-on-top:${flag}`),
      moveTop: () => calls.push('move-top'),
    } as never,
    {
      platform: 'linux',
      env: { HYPRLAND_INSTANCE_SIGNATURE: 'hypr' },
      ensureHyprlandWindowFloatingByTitle: ({ title }) => {
        calls.push(`hyprland-top:${title}`);
        return true;
      },
    },
  );

  assert.equal(promoted, true);
  assert.deepEqual(calls, ['always-on-top:true', 'move-top', 'hyprland-top:SubMiner Settings']);
});

test('promoteSettingsWindowAboveOverlay skips destroyed windows', () => {
  const calls: string[] = [];

  const promoted = promoteSettingsWindowAboveOverlay(
    {
      isDestroyed: () => true,
      getTitle: () => 'SubMiner Settings',
      setAlwaysOnTop: () => calls.push('always-on-top'),
      moveTop: () => calls.push('move-top'),
    } as never,
    {
      platform: 'linux',
      env: { HYPRLAND_INSTANCE_SIGNATURE: 'hypr' },
    },
  );

  assert.equal(promoted, false);
  assert.deepEqual(calls, []);
});
