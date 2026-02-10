import test from "node:test";
import assert from "node:assert/strict";
import { createOverlayVisibilityFacadeDepsRuntimeService } from "./overlay-visibility-facade-deps-runtime-service";

test("createOverlayVisibilityFacadeDepsRuntimeService returns working deps object", () => {
  let visible = false;
  let invisible = true;
  let mpvSubVisible: boolean | null = null;
  let syncCalls = 0;

  const deps = createOverlayVisibilityFacadeDepsRuntimeService({
    getVisibleOverlayVisible: () => visible,
    getInvisibleOverlayVisible: () => invisible,
    setVisibleOverlayVisibleState: (nextVisible) => {
      visible = nextVisible;
    },
    setInvisibleOverlayVisibleState: (nextVisible) => {
      invisible = nextVisible;
    },
    updateVisibleOverlayVisibility: () => {},
    updateInvisibleOverlayVisibility: () => {},
    syncInvisibleOverlayMousePassthrough: () => {
      syncCalls += 1;
    },
    shouldBindVisibleOverlayToMpvSubVisibility: () => true,
    isMpvConnected: () => true,
    setMpvSubVisibility: (nextVisible) => {
      mpvSubVisible = nextVisible;
    },
  });

  assert.equal(deps.getVisibleOverlayVisible(), false);
  assert.equal(deps.getInvisibleOverlayVisible(), true);

  deps.setVisibleOverlayVisibleState(true);
  deps.setInvisibleOverlayVisibleState(false);
  deps.syncInvisibleOverlayMousePassthrough();
  deps.setMpvSubVisibility(false);

  assert.equal(visible, true);
  assert.equal(invisible, false);
  assert.equal(syncCalls, 1);
  assert.equal(mpvSubVisible, false);
  assert.equal(deps.shouldBindVisibleOverlayToMpvSubVisibility(), true);
  assert.equal(deps.isMpvConnected(), true);
});
