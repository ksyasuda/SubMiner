import test from "node:test";
import assert from "node:assert/strict";
import {
  getInitialInvisibleOverlayVisibility,
  isAutoUpdateEnabledRuntime,
  shouldAutoInitializeOverlayRuntimeFromConfig,
  shouldBindVisibleOverlayToMpvSubVisibility,
} from "./startup";

const BASE_CONFIG = {
  auto_start_overlay: false,
  bind_visible_overlay_to_mpv_sub_visibility: true,
  invisibleOverlay: {
    startupVisibility: "platform-default" as const,
  },
  ankiConnect: {
    behavior: {
      autoUpdateNewCards: true,
    },
  },
};

test("getInitialInvisibleOverlayVisibility handles visibility + platform", () => {
  assert.equal(
    getInitialInvisibleOverlayVisibility(
      { ...BASE_CONFIG, invisibleOverlay: { startupVisibility: "visible" } },
      "linux",
    ),
    true,
  );
  assert.equal(
    getInitialInvisibleOverlayVisibility(
      { ...BASE_CONFIG, invisibleOverlay: { startupVisibility: "hidden" } },
      "darwin",
    ),
    false,
  );
  assert.equal(
    getInitialInvisibleOverlayVisibility(BASE_CONFIG, "linux"),
    false,
  );
  assert.equal(
    getInitialInvisibleOverlayVisibility(BASE_CONFIG, "darwin"),
    true,
  );
});

test("shouldAutoInitializeOverlayRuntimeFromConfig respects auto start and visible startup", () => {
  assert.equal(
    shouldAutoInitializeOverlayRuntimeFromConfig(BASE_CONFIG),
    false,
  );
  assert.equal(
    shouldAutoInitializeOverlayRuntimeFromConfig({
      ...BASE_CONFIG,
      auto_start_overlay: true,
    }),
    true,
  );
  assert.equal(
    shouldAutoInitializeOverlayRuntimeFromConfig({
      ...BASE_CONFIG,
      invisibleOverlay: { startupVisibility: "visible" },
    }),
    true,
  );
});

test("shouldBindVisibleOverlayToMpvSubVisibility returns config value", () => {
  assert.equal(shouldBindVisibleOverlayToMpvSubVisibility(BASE_CONFIG), true);
  assert.equal(
    shouldBindVisibleOverlayToMpvSubVisibility({
      ...BASE_CONFIG,
      bind_visible_overlay_to_mpv_sub_visibility: false,
    }),
    false,
  );
});

test("isAutoUpdateEnabledRuntime prefers runtime option and falls back to config", () => {
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
