import test from "node:test";
import assert from "node:assert/strict";
import {
  getInitialInvisibleOverlayVisibilityService,
  isAutoUpdateEnabledRuntimeService,
  shouldAutoInitializeOverlayRuntimeFromConfigService,
  shouldBindVisibleOverlayToMpvSubVisibilityService,
} from "./runtime-config-service";

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

test("getInitialInvisibleOverlayVisibilityService handles visibility + platform", () => {
  assert.equal(
    getInitialInvisibleOverlayVisibilityService(
      { ...BASE_CONFIG, invisibleOverlay: { startupVisibility: "visible" } },
      "linux",
    ),
    true,
  );
  assert.equal(
    getInitialInvisibleOverlayVisibilityService(
      { ...BASE_CONFIG, invisibleOverlay: { startupVisibility: "hidden" } },
      "darwin",
    ),
    false,
  );
  assert.equal(
    getInitialInvisibleOverlayVisibilityService(BASE_CONFIG, "linux"),
    false,
  );
  assert.equal(
    getInitialInvisibleOverlayVisibilityService(BASE_CONFIG, "darwin"),
    true,
  );
});

test("shouldAutoInitializeOverlayRuntimeFromConfigService respects auto start and visible startup", () => {
  assert.equal(
    shouldAutoInitializeOverlayRuntimeFromConfigService(BASE_CONFIG),
    false,
  );
  assert.equal(
    shouldAutoInitializeOverlayRuntimeFromConfigService({
      ...BASE_CONFIG,
      auto_start_overlay: true,
    }),
    true,
  );
  assert.equal(
    shouldAutoInitializeOverlayRuntimeFromConfigService({
      ...BASE_CONFIG,
      invisibleOverlay: { startupVisibility: "visible" },
    }),
    true,
  );
});

test("shouldBindVisibleOverlayToMpvSubVisibilityService returns config value", () => {
  assert.equal(shouldBindVisibleOverlayToMpvSubVisibilityService(BASE_CONFIG), true);
  assert.equal(
    shouldBindVisibleOverlayToMpvSubVisibilityService({
      ...BASE_CONFIG,
      bind_visible_overlay_to_mpv_sub_visibility: false,
    }),
    false,
  );
});

test("isAutoUpdateEnabledRuntimeService prefers runtime option and falls back to config", () => {
  assert.equal(
    isAutoUpdateEnabledRuntimeService(BASE_CONFIG, {
      getOptionValue: () => false,
    }),
    false,
  );
  assert.equal(
    isAutoUpdateEnabledRuntimeService(
      {
        ...BASE_CONFIG,
        ankiConnect: { behavior: { autoUpdateNewCards: false } },
      },
      null,
    ),
    false,
  );
  assert.equal(isAutoUpdateEnabledRuntimeService(BASE_CONFIG, null), true);
});
