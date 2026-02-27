import test from 'node:test';
import assert from 'node:assert/strict';
import {
  isAutoUpdateEnabledRuntime,
  shouldAutoInitializeOverlayRuntimeFromConfig,
  shouldBindVisibleOverlayToMpvSubVisibility,
} from './startup';

const BASE_CONFIG = {
  auto_start_overlay: false,
  bind_visible_overlay_to_mpv_sub_visibility: true,
  ankiConnect: {
    behavior: {
      autoUpdateNewCards: true,
    },
  },
};

test('shouldAutoInitializeOverlayRuntimeFromConfig respects auto start', () => {
  assert.equal(shouldAutoInitializeOverlayRuntimeFromConfig(BASE_CONFIG), false);
  assert.equal(
    shouldAutoInitializeOverlayRuntimeFromConfig({
      ...BASE_CONFIG,
      auto_start_overlay: true,
    }),
    true,
  );
});

test('shouldBindVisibleOverlayToMpvSubVisibility returns config value', () => {
  assert.equal(shouldBindVisibleOverlayToMpvSubVisibility(BASE_CONFIG), true);
  assert.equal(
    shouldBindVisibleOverlayToMpvSubVisibility({
      ...BASE_CONFIG,
      bind_visible_overlay_to_mpv_sub_visibility: false,
    }),
    false,
  );
});

test('isAutoUpdateEnabledRuntime prefers runtime option and falls back to config', () => {
  assert.equal(
    isAutoUpdateEnabledRuntime(BASE_CONFIG, {
      getOptionValue: () => false,
    }),
    false,
  );
  assert.equal(
    isAutoUpdateEnabledRuntime(
      {
        ...BASE_CONFIG,
        ankiConnect: { behavior: { autoUpdateNewCards: false } },
      },
      null,
    ),
    false,
  );
  assert.equal(isAutoUpdateEnabledRuntime(BASE_CONFIG, null), true);
});
