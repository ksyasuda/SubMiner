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

test('idle visible overlay starts click-through on platforms that toggle mouse ignore', () => {
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
        youtubePickerModalOpen: false,
        kikuModalOpen: false,
        runtimeOptionsModalOpen: false,
        subsyncModalOpen: false,
        sessionHelpModalOpen: false,
        subtitleSidebarModalOpen: false,
        subtitleSidebarConfig: null,
      },
    } as never);

    assert.equal(classList.contains('interactive'), false);
    assert.deepEqual(ignoreCalls, [{ ignore: true, forward: true }]);
  } finally {
    Object.assign(globalThis, { window: originalWindow });
  }
});

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

test('visible yomitan popup host keeps overlay interactive even when cached popup state is false', () => {
  const classList = createClassList();
  const ignoreCalls: Array<{ ignore: boolean; forward?: boolean }> = [];
  const originalWindow = globalThis.window;
  const originalDocument = globalThis.document;

  Object.assign(globalThis, {
    window: {
      electronAPI: {
        setIgnoreMouseEvents: (ignore: boolean, options?: { forward?: boolean }) => {
          ignoreCalls.push({ ignore, forward: options?.forward });
        },
      },
      getComputedStyle: () => ({
        visibility: 'visible',
        display: 'block',
        opacity: '1',
      }),
    },
    document: {
      querySelectorAll: (selector: string) =>
        selector === '[data-subminer-yomitan-popup-host="true"][data-subminer-yomitan-popup-visible="true"]'
          ? [{ getAttribute: () => 'true' }]
          : [],
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
        youtubePickerModalOpen: false,
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
    Object.assign(globalThis, { window: originalWindow, document: originalDocument });
  }
});
