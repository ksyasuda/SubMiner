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

function replaceGlobalProperty(key: 'window' | 'document', value: unknown): () => void {
  const original = Object.getOwnPropertyDescriptor(globalThis, key);
  Object.defineProperty(globalThis, key, {
    configurable: true,
    writable: true,
    value,
  });
  return () => {
    if (original) {
      Object.defineProperty(globalThis, key, original);
      return;
    }
    delete (globalThis as Record<string, unknown>)[key];
  };
}

test('idle visible overlay starts click-through on platforms that toggle mouse ignore', () => {
  const classList = createClassList();
  const ignoreCalls: Array<{ ignore: boolean; forward?: boolean }> = [];

  const restoreWindow = replaceGlobalProperty('window', {
    electronAPI: {
      setIgnoreMouseEvents: (ignore: boolean, options?: { forward?: boolean }) => {
        ignoreCalls.push({ ignore, forward: options?.forward });
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
    restoreWindow();
  }
});

test('youtube picker keeps overlay interactive even when subtitle hover is inactive', () => {
  const classList = createClassList();
  const ignoreCalls: Array<{ ignore: boolean; forward?: boolean }> = [];

  const restoreWindow = replaceGlobalProperty('window', {
    electronAPI: {
      setIgnoreMouseEvents: (ignore: boolean, options?: { forward?: boolean }) => {
        ignoreCalls.push({ ignore, forward: options?.forward });
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
    restoreWindow();
  }
});

test('visible yomitan popup host keeps overlay interactive even when cached popup state is false', () => {
  const classList = createClassList();
  const ignoreCalls: Array<{ ignore: boolean; forward?: boolean }> = [];

  const restoreWindow = replaceGlobalProperty('window', {
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
  });
  const restoreDocument = replaceGlobalProperty('document', {
    querySelectorAll: (selector: string) =>
      selector ===
      '[data-subminer-yomitan-popup-host="true"][data-subminer-yomitan-popup-visible="true"]'
        ? [{ getAttribute: () => 'true' }]
        : [],
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
    restoreDocument();
    restoreWindow();
  }
});

test('visible yomitan popup host on macOS keeps overlay interactive so click-away reaches popup', () => {
  const classList = createClassList();
  const ignoreCalls: Array<{ ignore: boolean; forward?: boolean }> = [];

  const restoreWindow = replaceGlobalProperty('window', {
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
  });
  const restoreDocument = replaceGlobalProperty('document', {
    querySelectorAll: (selector: string) =>
      selector ===
      '[data-subminer-yomitan-popup-host="true"][data-subminer-yomitan-popup-visible="true"]'
        ? [{ getAttribute: () => 'true' }]
        : [],
  });

  try {
    syncOverlayMouseIgnoreState({
      dom: {
        overlay: { classList },
      },
      platform: {
        isMacOSPlatform: true,
        shouldToggleMouseIgnore: true,
      },
      state: {
        isOverSubtitle: false,
        isOverSubtitleSidebar: false,
        isOverYomitanPopup: false,
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
    restoreDocument();
    restoreWindow();
  }
});

test('macOS pointer over a visible yomitan popup keeps the overlay interactive', () => {
  const classList = createClassList();
  const ignoreCalls: Array<{ ignore: boolean; forward?: boolean }> = [];

  const restoreWindow = replaceGlobalProperty('window', {
    electronAPI: {
      setIgnoreMouseEvents: (ignore: boolean, options?: { forward?: boolean }) => {
        ignoreCalls.push({ ignore, forward: options?.forward });
      },
    },
  });

  try {
    syncOverlayMouseIgnoreState({
      dom: {
        overlay: { classList },
      },
      platform: {
        isMacOSPlatform: true,
        shouldToggleMouseIgnore: true,
      },
      state: {
        isOverSubtitle: false,
        isOverSubtitleSidebar: false,
        isOverYomitanPopup: true,
        yomitanPopupVisible: true,
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
    restoreWindow();
  }
});

test('Linux subtitle hover keeps root passive and does not report whole-window interactive hint', () => {
  const classList = createClassList();
  const interactiveHints: boolean[] = [];

  const restoreWindow = replaceGlobalProperty('window', {
    electronAPI: {
      reportOverlayInteractive: (interactive: boolean) => {
        interactiveHints.push(interactive);
      },
    },
  });

  try {
    syncOverlayMouseIgnoreState({
      dom: {
        overlay: { classList },
      },
      platform: {
        isLinuxPlatform: true,
        shouldToggleMouseIgnore: false,
      },
      state: {
        isOverSubtitle: true,
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
    assert.deepEqual(interactiveHints, [false]);
  } finally {
    restoreWindow();
  }
});

test('Linux modal state reports whole-window interactive hint', () => {
  const classList = createClassList();
  const interactiveHints: boolean[] = [];

  const restoreWindow = replaceGlobalProperty('window', {
    electronAPI: {
      reportOverlayInteractive: (interactive: boolean) => {
        interactiveHints.push(interactive);
      },
    },
  });

  try {
    syncOverlayMouseIgnoreState({
      dom: {
        overlay: { classList },
      },
      platform: {
        isLinuxPlatform: true,
        shouldToggleMouseIgnore: false,
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
        runtimeOptionsModalOpen: true,
        subsyncModalOpen: false,
        sessionHelpModalOpen: false,
        subtitleSidebarModalOpen: false,
        subtitleSidebarConfig: null,
      },
    } as never);

    assert.equal(classList.contains('interactive'), true);
    assert.deepEqual(interactiveHints, [true]);
  } finally {
    restoreWindow();
  }
});
