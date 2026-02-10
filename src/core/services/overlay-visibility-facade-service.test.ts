import test from "node:test";
import assert from "node:assert/strict";
import {
  OverlayVisibilityFacadeDeps,
  setVisibleOverlayVisibleRuntimeFacadeService,
  toggleInvisibleOverlayRuntimeFacadeService,
  toggleVisibleOverlayRuntimeFacadeService,
} from "./overlay-visibility-facade-service";

function makeDeps(initialVisible = false, initialInvisible = false): {
  deps: OverlayVisibilityFacadeDeps;
  getState: () => { visible: boolean; invisible: boolean; mpvSubVisible: boolean | null };
} {
  let visible = initialVisible;
  let invisible = initialInvisible;
  let mpvSubVisible: boolean | null = null;

  const deps: OverlayVisibilityFacadeDeps = {
    getVisibleOverlayVisible: () => visible,
    getInvisibleOverlayVisible: () => invisible,
    setVisibleOverlayVisibleState: (value) => {
      visible = value;
    },
    setInvisibleOverlayVisibleState: (value) => {
      invisible = value;
    },
    updateVisibleOverlayVisibility: () => {},
    updateInvisibleOverlayVisibility: () => {},
    syncInvisibleOverlayMousePassthrough: () => {},
    shouldBindVisibleOverlayToMpvSubVisibility: () => true,
    isMpvConnected: () => true,
    setMpvSubVisibility: (value) => {
      mpvSubVisible = value;
    },
  };

  return {
    deps,
    getState: () => ({ visible, invisible, mpvSubVisible }),
  };
}

test("setVisibleOverlayVisibleRuntimeFacadeService updates visible state and mpv subtitle visibility", () => {
  const { deps, getState } = makeDeps(false, true);
  setVisibleOverlayVisibleRuntimeFacadeService(true, deps);
  assert.deepEqual(getState(), {
    visible: true,
    invisible: true,
    mpvSubVisible: false,
  });
});

test("toggleVisibleOverlayRuntimeFacadeService flips visible overlay state", () => {
  const { deps, getState } = makeDeps(false, false);
  toggleVisibleOverlayRuntimeFacadeService(deps);
  assert.equal(getState().visible, true);
});

test("toggleInvisibleOverlayRuntimeFacadeService flips invisible overlay state", () => {
  const { deps, getState } = makeDeps(false, false);
  toggleInvisibleOverlayRuntimeFacadeService(deps);
  assert.equal(getState().invisible, true);
});
