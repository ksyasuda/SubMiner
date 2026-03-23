import assert from 'node:assert/strict';
import test from 'node:test';
import { syncOverlayMouseIgnoreState } from './overlay-mouse-ignore.js';

function createClassList() {
  const classes = new Set<string>();
  return {
    add: (...tokens: string[]) => {
      for (const token of tokens) classes.add(token);
    },
    remove: (...tokens: string[]) => {
      for (const token of tokens) classes.delete(token);
    },
    contains: (token: string) => classes.has(token),
  };
}

test('youtube picker keeps overlay interactive even when subtitle hover is inactive', () => {
  const classList = createClassList();
  const ignoreCalls: Array<{ ignore: boolean; forward?: boolean }> = [];
  const originalWindow = globalThis.window;

  Object.assign(globalThis, {
    window: {
      electronAPI: {
        setIgnoreMouseEvents: (ignore: boolean, options?: { forward?: boolean }) => {
          ignoreCalls.push({ ignore, forward: options?.forward });
        },
      },
    },
  });

  try {
    syncOverlayMouseIgnoreState({
      dom: {
        overlay: { classList },
      },
      platform: {
        shouldToggleMouseIgnore: true,
      },
      state: {
        isOverSubtitle: false,
        isOverSubtitleSidebar: false,
        yomitanPopupVisible: false,
        controllerSelectModalOpen: false,
        controllerDebugModalOpen: false,
        jimakuModalOpen: false,
        youtubePickerModalOpen: true,
        kikuModalOpen: false,
        runtimeOptionsModalOpen: false,
        subsyncModalOpen: false,
        sessionHelpModalOpen: false,
        subtitleSidebarModalOpen: false,
        subtitleSidebarConfig: null,
      },
    } as never);

    assert.equal(classList.contains('interactive'), true);
    assert.deepEqual(ignoreCalls, [{ ignore: false, forward: undefined }]);
  } finally {
    Object.assign(globalThis, { window: originalWindow });
  }
});
