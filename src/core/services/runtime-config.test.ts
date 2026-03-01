import test from 'node:test';
import assert from 'node:assert/strict';
import {
  isAutoUpdateEnabledRuntime,
  shouldAutoInitializeOverlayRuntimeFromConfig,
} from './startup';

const BASE_CONFIG = {
  auto_start_overlay: false,
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
