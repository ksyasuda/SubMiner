import assert from 'node:assert/strict';
import test from 'node:test';

import type { ElectronAPI, SubtitleSidebarSnapshot } from '../../types';
import { createRendererState } from '../state.js';
import { createSubtitleSidebarModal, findActiveSubtitleCueIndex } from './subtitle-sidebar.js';

function createClassList(initialTokens: string[] = []) {
  const tokens = new Set(initialTokens);
  return {
    add: (...entries: string[]) => {
      for (const entry of entries) tokens.add(entry);
    },
    remove: (...entries: string[]) => {
      for (const entry of entries) tokens.delete(entry);
    },
    contains: (entry: string) => tokens.has(entry),
    toggle: (entry: string, force?: boolean) => {
      if (force === true) tokens.add(entry);
      else if (force === false) tokens.delete(entry);
      else if (tokens.has(entry)) tokens.delete(entry);
      else tokens.add(entry);
    },
  };
}

function createCueRow() {
  const listeners = new Map<string, Array<(event: unknown) => void>>();
  return {
    className: '',
    classList: createClassList(),
    dataset: {} as Record<string, string>,
    textContent: '',
    tabIndex: -1,
    offsetTop: 0,
    clientHeight: 40,
    children: [] as unknown[],
    appendChild(child: unknown) {
      this.children.push(child);
    },
    attributes: {} as Record<string, string>,
    listeners,
    addEventListener(type: string, listener: (event: unknown) => void) {
      const bucket = listeners.get(type) ?? [];
      bucket.push(listener);
      listeners.set(type, bucket);
    },
    setAttribute(name: string, value: string) {
      this.attributes[name] = value;
    },
    scrollIntoViewCalls: [] as ScrollIntoViewOptions[],
    scrollIntoView(options?: ScrollIntoViewOptions) {
      this.scrollIntoViewCalls.push(options ?? {});
    },
  };
}

function createListStub() {
  return {
    innerHTML: '',
    children: [] as ReturnType<typeof createCueRow>[],
    appendChild(child: ReturnType<typeof createCueRow>) {
      child.offsetTop = this.children.length * child.clientHeight;
      this.children.push(child);
    },
    addEventListener: () => {},
    scrollTop: 0,
    clientHeight: 240,
    scrollHeight: 480,
    scrollToCalls: [] as ScrollToOptions[],
    scrollTo(options?: ScrollToOptions) {
      this.scrollToCalls.push(options ?? {});
    },
  };
}

test.afterEach(() => {
  if (
    Object.prototype.hasOwnProperty.call(globalThis, 'window') &&
    globalThis.window === undefined
  ) {
    Reflect.deleteProperty(globalThis, 'window');
  }
  if (
    Object.prototype.hasOwnProperty.call(globalThis, 'document') &&
    globalThis.document === undefined
  ) {
    Reflect.deleteProperty(globalThis, 'document');
  }
});

test('findActiveSubtitleCueIndex prefers timing match before text fallback', () => {
  const cues = [
    { startTime: 1, endTime: 2, text: 'same' },
    { startTime: 3, endTime: 4, text: 'same' },
  ];

  assert.equal(findActiveSubtitleCueIndex(cues, { text: 'same', startTime: 3.1 }), 1);
  assert.equal(findActiveSubtitleCueIndex(cues, { text: 'same', startTime: null }), 0);
});

test('findActiveSubtitleCueIndex prefers current subtitle timing over near-future clock lookahead', () => {
  const cues = [
    { startTime: 231, endTime: 233.2, text: 'previous' },
    { startTime: 233.05, endTime: 236, text: 'next' },
  ];

  assert.equal(findActiveSubtitleCueIndex(cues, { text: 'previous', startTime: 231 }, 233, 0), 0);
});

test('subtitle sidebar modal opens from snapshot and clicking cue seeks playback', async () => {
  const globals = globalThis as typeof globalThis & { window?: unknown; document?: unknown };
  const previousWindow = globals.window;
  const previousDocument = globals.document;
  const mpvCommands: Array<Array<string | number>> = [];

  const snapshot: SubtitleSidebarSnapshot = {
    cues: [
      { startTime: 1, endTime: 2, text: 'first' },
      { startTime: 3, endTime: 4, text: 'second' },
    ],
    currentSubtitle: {
      text: 'second',
      startTime: 3,
      endTime: 4,
    },
    config: {
      enabled: true,
      autoOpen: false,
      layout: 'overlay',
      toggleKey: 'Backslash',
      pauseVideoOnHover: false,
      autoScroll: true,
      maxWidth: 420,
      opacity: 0.92,
      backgroundColor: 'rgba(54, 58, 79, 0.88)',
      textColor: '#cad3f5',
      fontFamily: '"Iosevka Aile", sans-serif',
      fontSize: 17,
      timestampColor: '#a5adcb',
      activeLineColor: '#f5bde6',
      activeLineBackgroundColor: 'rgba(138, 173, 244, 0.22)',
      hoverLineBackgroundColor: 'rgba(54, 58, 79, 0.84)',
    },
  };

  Object.defineProperty(globalThis, 'window', {
    configurable: true,
    value: {
      electronAPI: {
        getSubtitleSidebarSnapshot: async () => snapshot,
        sendMpvCommand: (command: Array<string | number>) => {
          mpvCommands.push(command);
        },
      } as unknown as ElectronAPI,
    },
  });
  Object.defineProperty(globalThis, 'document', {
    configurable: true,
    value: {
      createElement: () => createCueRow(),
      body: {
        classList: createClassList(),
      },
      documentElement: {
        style: {
          setProperty: () => {},
        },
      },
    },
  });

  try {
    const state = createRendererState();
    const overlayClassList = createClassList();
    const modalClassList = createClassList(['hidden']);
    const cueList = createListStub();
    const ctx = {
      dom: {
        overlay: { classList: overlayClassList },
        subtitleSidebarModal: {
          classList: modalClassList,
          setAttribute: () => {},
          style: { setProperty: () => {} },
          addEventListener: () => {},
        },
        subtitleSidebarContent: {
          classList: createClassList(),
          getBoundingClientRect: () => ({ width: 420 }),
        },
        subtitleSidebarClose: { addEventListener: () => {} },
        subtitleSidebarStatus: { textContent: '' },
        subtitleSidebarList: cueList,
      },
      state,
    };

    const modal = createSubtitleSidebarModal(ctx as never, {
      modalStateReader: { isAnyModalOpen: () => false },
    });

    await modal.openSubtitleSidebarModal();

    assert.equal(state.subtitleSidebarModalOpen, true);
    assert.equal(modalClassList.contains('hidden'), false);
    assert.equal(state.subtitleSidebarActiveCueIndex, 1);
    assert.equal(cueList.children.length, 2);
    assert.equal(cueList.scrollTop, 0);
    assert.deepEqual(cueList.scrollToCalls, []);

    modal.seekToCue(snapshot.cues[0]!);
    assert.deepEqual(mpvCommands.at(-1), ['seek', 1.08, 'absolute+exact']);
  } finally {
    Object.defineProperty(globalThis, 'window', { configurable: true, value: previousWindow });
    Object.defineProperty(globalThis, 'document', { configurable: true, value: previousDocument });
  }
});

test('subtitle sidebar rows support keyboard activation', async () => {
  const globals = globalThis as typeof globalThis & { window?: unknown; document?: unknown };
  const previousWindow = globals.window;
  const previousDocument = globals.document;
  const mpvCommands: Array<Array<string | number>> = [];

  const snapshot: SubtitleSidebarSnapshot = {
    cues: [
      { startTime: 1, endTime: 2, text: 'first' },
      { startTime: 3, endTime: 4, text: 'second' },
    ],
    currentSubtitle: {
      text: 'second',
      startTime: 3,
      endTime: 4,
    },
    config: {
      enabled: true,
      autoOpen: false,
      layout: 'overlay',
      toggleKey: 'Backslash',
      pauseVideoOnHover: false,
      autoScroll: true,
      maxWidth: 420,
      opacity: 0.92,
      backgroundColor: 'rgba(54, 58, 79, 0.88)',
      textColor: '#cad3f5',
      fontFamily: '"Iosevka Aile", sans-serif',
      fontSize: 17,
      timestampColor: '#a5adcb',
      activeLineColor: '#f5bde6',
      activeLineBackgroundColor: 'rgba(138, 173, 244, 0.22)',
      hoverLineBackgroundColor: 'rgba(54, 58, 79, 0.84)',
    },
  };

  Object.defineProperty(globalThis, 'window', {
    configurable: true,
    value: {
      electronAPI: {
        getSubtitleSidebarSnapshot: async () => snapshot,
        sendMpvCommand: (command: Array<string | number>) => {
          mpvCommands.push(command);
        },
      } as unknown as ElectronAPI,
    },
  });
  Object.defineProperty(globalThis, 'document', {
    configurable: true,
    value: {
      createElement: () => createCueRow(),
      body: {
        classList: createClassList(),
      },
      documentElement: {
        style: {
          setProperty: () => {},
        },
      },
    },
  });

  try {
    const state = createRendererState();
    const cueList = createListStub();
    const ctx = {
      dom: {
        overlay: { classList: createClassList() },
        subtitleSidebarModal: {
          classList: createClassList(['hidden']),
          setAttribute: () => {},
          style: { setProperty: () => {} },
          addEventListener: () => {},
        },
        subtitleSidebarContent: {
          classList: createClassList(),
          getBoundingClientRect: () => ({ width: 420 }),
        },
        subtitleSidebarClose: { addEventListener: () => {} },
        subtitleSidebarStatus: { textContent: '' },
        subtitleSidebarList: cueList,
      },
      state,
    };

    const modal = createSubtitleSidebarModal(ctx as never, {
      modalStateReader: { isAnyModalOpen: () => false },
    });

    await modal.openSubtitleSidebarModal();

    const firstRow = cueList.children[0]!;
    const keydownListeners = firstRow.listeners.get('keydown') ?? [];
    assert.equal(keydownListeners.length > 0, true);

    keydownListeners[0]!({
      key: 'Enter',
      preventDefault: () => {},
    });

    assert.deepEqual(mpvCommands.at(-1), ['seek', 1.08, 'absolute+exact']);
  } finally {
    Object.defineProperty(globalThis, 'window', { configurable: true, value: previousWindow });
    Object.defineProperty(globalThis, 'document', { configurable: true, value: previousDocument });
  }
});

test('subtitle sidebar renders hour-long cue timestamps as HH:MM:SS', async () => {
  const globals = globalThis as typeof globalThis & { window?: unknown; document?: unknown };
  const previousWindow = globals.window;
  const previousDocument = globals.document;

  const snapshot: SubtitleSidebarSnapshot = {
    cues: [{ startTime: 3665, endTime: 3670, text: 'long cue' }],
    currentSubtitle: {
      text: 'long cue',
      startTime: 3665,
      endTime: 3670,
    },
    config: {
      enabled: true,
      autoOpen: false,
      layout: 'overlay',
      toggleKey: 'Backslash',
      pauseVideoOnHover: false,
      autoScroll: true,
      maxWidth: 420,
      opacity: 0.92,
      backgroundColor: 'rgba(54, 58, 79, 0.88)',
      textColor: '#cad3f5',
      fontFamily: '"Iosevka Aile", sans-serif',
      fontSize: 17,
      timestampColor: '#a5adcb',
      activeLineColor: '#f5bde6',
      activeLineBackgroundColor: 'rgba(138, 173, 244, 0.22)',
      hoverLineBackgroundColor: 'rgba(54, 58, 79, 0.84)',
    },
  };

  Object.defineProperty(globalThis, 'window', {
    configurable: true,
    value: {
      electronAPI: {
        getSubtitleSidebarSnapshot: async () => snapshot,
        sendMpvCommand: () => {},
      } as unknown as ElectronAPI,
    },
  });
  Object.defineProperty(globalThis, 'document', {
    configurable: true,
    value: {
      createElement: () => createCueRow(),
      body: {
        classList: createClassList(),
      },
      documentElement: {
        style: {
          setProperty: () => {},
        },
      },
    },
  });

  try {
    const state = createRendererState();
    const cueList = createListStub();
    const ctx = {
      dom: {
        overlay: { classList: createClassList() },
        subtitleSidebarModal: {
          classList: createClassList(['hidden']),
          setAttribute: () => {},
          style: { setProperty: () => {} },
          addEventListener: () => {},
        },
        subtitleSidebarContent: {
          classList: createClassList(),
          getBoundingClientRect: () => ({ width: 420 }),
        },
        subtitleSidebarClose: { addEventListener: () => {} },
        subtitleSidebarStatus: { textContent: '' },
        subtitleSidebarList: cueList,
      },
      state,
    };

    const modal = createSubtitleSidebarModal(ctx as never, {
      modalStateReader: { isAnyModalOpen: () => false },
    });

    await modal.openSubtitleSidebarModal();

    const firstRow = cueList.children[0]!;
    assert.equal(firstRow.attributes['aria-label'], 'Jump to subtitle at 01:01:05');
    assert.equal((firstRow.children[0] as { textContent: string }).textContent, '01:01:05');
  } finally {
    Object.defineProperty(globalThis, 'window', { configurable: true, value: previousWindow });
    Object.defineProperty(globalThis, 'document', { configurable: true, value: previousDocument });
  }
});

test('subtitle sidebar does not open when the feature is disabled', async () => {
  const globals = globalThis as typeof globalThis & { window?: unknown; document?: unknown };
  const previousWindow = globals.window;
  const previousDocument = globals.document;
  const snapshot: SubtitleSidebarSnapshot = {
    cues: [],
    currentSubtitle: {
      text: '',
      startTime: null,
      endTime: null,
    },
    config: {
      enabled: false,
      autoOpen: false,
      layout: 'overlay',
      toggleKey: 'Backslash',
      pauseVideoOnHover: false,
      autoScroll: true,
      maxWidth: 420,
      opacity: 0.92,
      backgroundColor: 'rgba(54, 58, 79, 0.88)',
      textColor: '#cad3f5',
      fontFamily: '"Iosevka Aile", sans-serif',
      fontSize: 17,
      timestampColor: '#a5adcb',
      activeLineColor: '#f5bde6',
      activeLineBackgroundColor: 'rgba(138, 173, 244, 0.22)',
      hoverLineBackgroundColor: 'rgba(54, 58, 79, 0.84)',
    },
  };

  Object.defineProperty(globalThis, 'window', {
    configurable: true,
    value: {
      electronAPI: {
        getSubtitleSidebarSnapshot: async () => snapshot,
        sendMpvCommand: () => {},
      } as unknown as ElectronAPI,
    },
  });
  Object.defineProperty(globalThis, 'document', {
    configurable: true,
    value: {
      createElement: () => createCueRow(),
      body: {
        classList: createClassList(),
      },
      documentElement: {
        style: {
          setProperty: () => {},
        },
      },
    },
  });

  try {
    const state = createRendererState();
    const modalClassList = createClassList(['hidden']);
    const cueList = createListStub();
    const ctx = {
      dom: {
        overlay: { classList: createClassList() },
        subtitleSidebarModal: {
          classList: modalClassList,
          setAttribute: () => {},
          style: { setProperty: () => {} },
          addEventListener: () => {},
        },
        subtitleSidebarContent: {
          classList: createClassList(),
          getBoundingClientRect: () => ({ width: 420 }),
        },
        subtitleSidebarClose: { addEventListener: () => {} },
        subtitleSidebarStatus: { textContent: '' },
        subtitleSidebarList: cueList,
      },
      state,
    };

    const modal = createSubtitleSidebarModal(ctx as never, {
      modalStateReader: { isAnyModalOpen: () => false },
    });

    await modal.openSubtitleSidebarModal();

    assert.equal(state.subtitleSidebarModalOpen, false);
    assert.equal(modalClassList.contains('hidden'), true);
    assert.equal(cueList.children.length, 0);
    assert.equal(ctx.dom.subtitleSidebarStatus.textContent, 'Subtitle sidebar disabled in config.');
  } finally {
    Object.defineProperty(globalThis, 'window', { configurable: true, value: previousWindow });
    Object.defineProperty(globalThis, 'document', { configurable: true, value: previousDocument });
  }
});

test('subtitle sidebar auto-open on startup only opens when enabled and configured', async () => {
  const globals = globalThis as typeof globalThis & { window?: unknown; document?: unknown };
  const previousWindow = globals.window;
  const previousDocument = globals.document;

  let snapshot: SubtitleSidebarSnapshot = {
    cues: [{ startTime: 1, endTime: 2, text: 'first' }],
    currentSubtitle: {
      text: 'first',
      startTime: 1,
      endTime: 2,
    },
    config: {
      enabled: true,
      autoOpen: true,
      layout: 'overlay',
      toggleKey: 'Backslash',
      pauseVideoOnHover: false,
      autoScroll: true,
      maxWidth: 420,
      opacity: 0.92,
      backgroundColor: 'rgba(54, 58, 79, 0.88)',
      textColor: '#cad3f5',
      fontFamily: '"Iosevka Aile", sans-serif',
      fontSize: 17,
      timestampColor: '#a5adcb',
      activeLineColor: '#f5bde6',
      activeLineBackgroundColor: 'rgba(138, 173, 244, 0.22)',
      hoverLineBackgroundColor: 'rgba(54, 58, 79, 0.84)',
    },
  };

  Object.defineProperty(globalThis, 'window', {
    configurable: true,
    value: {
      electronAPI: {
        getSubtitleSidebarSnapshot: async () => snapshot,
        sendMpvCommand: () => {},
      } as unknown as ElectronAPI,
      addEventListener: () => {},
      removeEventListener: () => {},
    },
  });
  Object.defineProperty(globalThis, 'document', {
    configurable: true,
    value: {
      createElement: () => createCueRow(),
      body: {
        classList: createClassList(),
      },
      documentElement: {
        style: {
          setProperty: () => {},
        },
      },
    },
  });

  try {
    const state = createRendererState();
    const modalClassList = createClassList(['hidden']);
    const cueList = createListStub();
    const ctx = {
      dom: {
        overlay: { classList: createClassList() },
        subtitleSidebarModal: {
          classList: modalClassList,
          setAttribute: () => {},
          style: { setProperty: () => {} },
          addEventListener: () => {},
        },
        subtitleSidebarContent: {
          classList: createClassList(),
          getBoundingClientRect: () => ({ width: 420 }),
        },
        subtitleSidebarClose: { addEventListener: () => {} },
        subtitleSidebarStatus: { textContent: '' },
        subtitleSidebarList: cueList,
      },
      state,
    };

    const modal = createSubtitleSidebarModal(ctx as never, {
      modalStateReader: { isAnyModalOpen: () => false },
    });

    await modal.autoOpenSubtitleSidebarOnStartup();

    assert.equal(state.subtitleSidebarModalOpen, true);
    assert.equal(modalClassList.contains('hidden'), false);
    assert.equal(cueList.children.length, 1);

    modal.closeSubtitleSidebarModal();
    snapshot = {
      ...snapshot,
      config: {
        ...snapshot.config,
        autoOpen: false,
      },
    };

    await modal.autoOpenSubtitleSidebarOnStartup();

    assert.equal(state.subtitleSidebarModalOpen, false);
  } finally {
    Object.defineProperty(globalThis, 'window', { configurable: true, value: previousWindow });
    Object.defineProperty(globalThis, 'document', { configurable: true, value: previousDocument });
  }
});

test('subtitle sidebar refresh closes and clears state when config becomes disabled', async () => {
  const globals = globalThis as typeof globalThis & { window?: unknown; document?: unknown };
  const previousWindow = globals.window;
  const previousDocument = globals.document;
  const bodyClassList = createClassList();
  let snapshot: SubtitleSidebarSnapshot = {
    cues: [{ startTime: 1, endTime: 2, text: 'first' }],
    currentSubtitle: {
      text: 'first',
      startTime: 1,
      endTime: 2,
    },
    currentTimeSec: 1.1,
    config: {
      enabled: true,
      autoOpen: false,
      layout: 'embedded',
      toggleKey: 'Backslash',
      pauseVideoOnHover: false,
      autoScroll: true,
      maxWidth: 360,
      opacity: 0.92,
      backgroundColor: 'rgba(54, 58, 79, 0.88)',
      textColor: '#cad3f5',
      fontFamily: '"Iosevka Aile", sans-serif',
      fontSize: 17,
      timestampColor: '#a5adcb',
      activeLineColor: '#f5bde6',
      activeLineBackgroundColor: 'rgba(138, 173, 244, 0.22)',
      hoverLineBackgroundColor: 'rgba(54, 58, 79, 0.84)',
    },
  };

  Object.defineProperty(globalThis, 'window', {
    configurable: true,
    value: {
      innerWidth: 1200,
      electronAPI: {
        getSubtitleSidebarSnapshot: async () => snapshot,
        sendMpvCommand: () => {},
        setIgnoreMouseEvents: () => {},
      } as unknown as ElectronAPI,
      addEventListener: () => {},
      removeEventListener: () => {},
    },
  });
  Object.defineProperty(globalThis, 'document', {
    configurable: true,
    value: {
      createElement: () => createCueRow(),
      body: {
        classList: bodyClassList,
      },
      documentElement: {
        style: {
          setProperty: () => {},
        },
      },
    },
  });

  try {
    const state = createRendererState();
    const modalClassList = createClassList(['hidden']);
    const cueList = createListStub();
    const ctx = {
      dom: {
        overlay: { classList: createClassList() },
        subtitleSidebarModal: {
          classList: modalClassList,
          setAttribute: () => {},
          style: { setProperty: () => {} },
          addEventListener: () => {},
        },
        subtitleSidebarContent: {
          classList: createClassList(),
          getBoundingClientRect: () => ({ width: 360 }),
        },
        subtitleSidebarClose: { addEventListener: () => {} },
        subtitleSidebarStatus: { textContent: '' },
        subtitleSidebarList: cueList,
      },
      platform: {
        shouldToggleMouseIgnore: false,
      },
      state,
    };

    const modal = createSubtitleSidebarModal(ctx as never, {
      modalStateReader: { isAnyModalOpen: () => false },
    });

    await modal.openSubtitleSidebarModal();
    assert.equal(state.subtitleSidebarModalOpen, true);
    assert.equal(bodyClassList.contains('subtitle-sidebar-embedded-open'), true);

    snapshot = {
      ...snapshot,
      cues: [],
      currentSubtitle: {
        text: '',
        startTime: null,
        endTime: null,
      },
      currentTimeSec: null,
      config: {
        ...snapshot.config,
        enabled: false,
      },
    };

    await modal.refreshSubtitleSidebarSnapshot();

    assert.equal(state.subtitleSidebarModalOpen, false);
    assert.equal(state.subtitleSidebarCues.length, 0);
    assert.equal(state.subtitleSidebarActiveCueIndex, -1);
    assert.equal(modalClassList.contains('hidden'), true);
    assert.equal(bodyClassList.contains('subtitle-sidebar-embedded-open'), false);
  } finally {
    Object.defineProperty(globalThis, 'window', { configurable: true, value: previousWindow });
    Object.defineProperty(globalThis, 'document', { configurable: true, value: previousDocument });
  }
});

test('subtitle sidebar keeps nearby repeated cue when subtitle update lacks timing', async () => {
  const globals = globalThis as typeof globalThis & { window?: unknown; document?: unknown };
  const previousWindow = globals.window;
  const previousDocument = globals.document;

  const snapshot: SubtitleSidebarSnapshot = {
    cues: [
      { startTime: 1, endTime: 2, text: 'same' },
      { startTime: 3, endTime: 4, text: 'other' },
      { startTime: 10, endTime: 11, text: 'same' },
    ],
    currentSubtitle: {
      text: 'same',
      startTime: 10,
      endTime: 11,
    },
    currentTimeSec: 10.1,
    config: {
      enabled: true,
      autoOpen: false,
      layout: 'overlay',
      toggleKey: 'Backslash',
      pauseVideoOnHover: false,
      autoScroll: true,
      maxWidth: 420,
      opacity: 0.92,
      backgroundColor: 'rgba(54, 58, 79, 0.88)',
      textColor: '#cad3f5',
      fontFamily: '"Iosevka Aile", sans-serif',
      fontSize: 17,
      timestampColor: '#a5adcb',
      activeLineColor: '#f5bde6',
      activeLineBackgroundColor: 'rgba(138, 173, 244, 0.22)',
      hoverLineBackgroundColor: 'rgba(54, 58, 79, 0.84)',
    },
  };

  Object.defineProperty(globalThis, 'window', {
    configurable: true,
    value: {
      electronAPI: {
        getSubtitleSidebarSnapshot: async () => snapshot,
        sendMpvCommand: () => {},
      } as unknown as ElectronAPI,
      addEventListener: () => {},
      removeEventListener: () => {},
    },
  });
  Object.defineProperty(globalThis, 'document', {
    configurable: true,
    value: {
      createElement: () => createCueRow(),
      body: {
        classList: createClassList(),
      },
      documentElement: {
        style: {
          setProperty: () => {},
        },
      },
    },
  });

  try {
    const state = createRendererState();
    const cueList = createListStub();
    const ctx = {
      dom: {
        overlay: { classList: createClassList() },
        subtitleSidebarModal: {
          classList: createClassList(['hidden']),
          setAttribute: () => {},
          style: { setProperty: () => {} },
          addEventListener: () => {},
        },
        subtitleSidebarContent: {
          classList: createClassList(),
          getBoundingClientRect: () => ({ width: 420 }),
        },
        subtitleSidebarClose: { addEventListener: () => {} },
        subtitleSidebarStatus: { textContent: '' },
        subtitleSidebarList: cueList,
      },
      state,
    };

    const modal = createSubtitleSidebarModal(ctx as never, {
      modalStateReader: { isAnyModalOpen: () => false },
    });

    await modal.openSubtitleSidebarModal();
    cueList.scrollToCalls.length = 0;

    modal.handleSubtitleUpdated({
      text: 'same',
      startTime: null,
      endTime: null,
      tokens: [],
    });

    assert.equal(state.subtitleSidebarActiveCueIndex, 2);
    assert.deepEqual(cueList.scrollToCalls, []);
  } finally {
    Object.defineProperty(globalThis, 'window', { configurable: true, value: previousWindow });
    Object.defineProperty(globalThis, 'document', { configurable: true, value: previousDocument });
  }
});

test('findActiveSubtitleCueIndex falls back to the latest matching cue when the preferred index is stale', () => {
  const cues = [
    { startTime: 1, endTime: 2, text: 'same' },
    { startTime: 3, endTime: 4, text: 'same' },
  ];

  assert.equal(findActiveSubtitleCueIndex(cues, { text: 'same', startTime: null }, null, 5), 1);
});

test('subtitle sidebar does not regress to previous cue on text-only transition update', async () => {
  const globals = globalThis as typeof globalThis & { window?: unknown; document?: unknown };
  const previousWindow = globals.window;
  const previousDocument = globals.document;

  const snapshot: SubtitleSidebarSnapshot = {
    cues: [
      { startTime: 1, endTime: 2, text: 'first' },
      { startTime: 3, endTime: 4, text: 'second' },
      { startTime: 5, endTime: 6, text: 'third' },
    ],
    currentSubtitle: {
      text: 'third',
      startTime: 5,
      endTime: 6,
    },
    currentTimeSec: 5.1,
    config: {
      enabled: true,
      autoOpen: false,
      layout: 'overlay',
      toggleKey: 'Backslash',
      pauseVideoOnHover: false,
      autoScroll: true,
      maxWidth: 420,
      opacity: 0.92,
      backgroundColor: 'rgba(54, 58, 79, 0.88)',
      textColor: '#cad3f5',
      fontFamily: '"Iosevka Aile", sans-serif',
      fontSize: 17,
      timestampColor: '#a5adcb',
      activeLineColor: '#f5bde6',
      activeLineBackgroundColor: 'rgba(138, 173, 244, 0.22)',
      hoverLineBackgroundColor: 'rgba(54, 58, 79, 0.84)',
    },
  };

  Object.defineProperty(globalThis, 'window', {
    configurable: true,
    value: {
      electronAPI: {
        getSubtitleSidebarSnapshot: async () => snapshot,
        sendMpvCommand: () => {},
      } as unknown as ElectronAPI,
      addEventListener: () => {},
      removeEventListener: () => {},
    },
  });
  Object.defineProperty(globalThis, 'document', {
    configurable: true,
    value: {
      createElement: () => createCueRow(),
      body: {
        classList: createClassList(),
      },
      documentElement: {
        style: {
          setProperty: () => {},
        },
      },
    },
  });

  try {
    const state = createRendererState();
    const cueList = createListStub();
    const ctx = {
      dom: {
        overlay: { classList: createClassList() },
        subtitleSidebarModal: {
          classList: createClassList(['hidden']),
          setAttribute: () => {},
          style: { setProperty: () => {} },
          addEventListener: () => {},
        },
        subtitleSidebarContent: {
          classList: createClassList(),
          getBoundingClientRect: () => ({ width: 420 }),
        },
        subtitleSidebarClose: { addEventListener: () => {} },
        subtitleSidebarStatus: { textContent: '' },
        subtitleSidebarList: cueList,
      },
      state,
    };

    const modal = createSubtitleSidebarModal(ctx as never, {
      modalStateReader: { isAnyModalOpen: () => false },
    });

    await modal.openSubtitleSidebarModal();
    cueList.scrollToCalls.length = 0;

    modal.handleSubtitleUpdated({
      text: 'second',
      startTime: null,
      endTime: null,
      tokens: [],
    });

    assert.equal(state.subtitleSidebarActiveCueIndex, 2);
    assert.deepEqual(cueList.scrollToCalls, []);
  } finally {
    Object.defineProperty(globalThis, 'window', { configurable: true, value: previousWindow });
    Object.defineProperty(globalThis, 'document', { configurable: true, value: previousDocument });
  }
});

test('subtitle sidebar jumps to first resolved active cue, then resumes smooth auto-follow', async () => {
  const globals = globalThis as typeof globalThis & { window?: unknown; document?: unknown };
  const previousWindow = globals.window;
  const previousDocument = globals.document;

  let snapshot: SubtitleSidebarSnapshot = {
    cues: Array.from({ length: 12 }, (_, index) => ({
      startTime: index * 2,
      endTime: index * 2 + 1.5,
      text: `line-${index}`,
    })),
    currentSubtitle: {
      text: '',
      startTime: null,
      endTime: null,
    },
    currentTimeSec: null,
    config: {
      enabled: true,
      autoOpen: false,
      layout: 'overlay',
      toggleKey: 'Backslash',
      pauseVideoOnHover: false,
      autoScroll: true,
      maxWidth: 420,
      opacity: 0.92,
      backgroundColor: 'rgba(54, 58, 79, 0.88)',
      textColor: '#cad3f5',
      fontFamily: '"Iosevka Aile", sans-serif',
      fontSize: 17,
      timestampColor: '#a5adcb',
      activeLineColor: '#f5bde6',
      activeLineBackgroundColor: 'rgba(138, 173, 244, 0.22)',
      hoverLineBackgroundColor: 'rgba(54, 58, 79, 0.84)',
    },
  };

  Object.defineProperty(globalThis, 'window', {
    configurable: true,
    value: {
      electronAPI: {
        getSubtitleSidebarSnapshot: async () => snapshot,
        sendMpvCommand: () => {},
      } as unknown as ElectronAPI,
      addEventListener: () => {},
      removeEventListener: () => {},
    },
  });
  Object.defineProperty(globalThis, 'document', {
    configurable: true,
    value: {
      createElement: () => createCueRow(),
      body: {
        classList: createClassList(),
      },
      documentElement: {
        style: {
          setProperty: () => {},
        },
      },
    },
  });

  try {
    const state = createRendererState();
    const cueList = createListStub();
    const ctx = {
      dom: {
        overlay: { classList: createClassList() },
        subtitleSidebarModal: {
          classList: createClassList(['hidden']),
          setAttribute: () => {},
          style: { setProperty: () => {} },
          addEventListener: () => {},
        },
        subtitleSidebarContent: {
          classList: createClassList(),
          getBoundingClientRect: () => ({ width: 420 }),
        },
        subtitleSidebarClose: { addEventListener: () => {} },
        subtitleSidebarStatus: { textContent: '' },
        subtitleSidebarList: cueList,
      },
      state,
    };

    const modal = createSubtitleSidebarModal(ctx as never, {
      modalStateReader: { isAnyModalOpen: () => false },
    });

    await modal.openSubtitleSidebarModal();
    assert.equal(state.subtitleSidebarActiveCueIndex, -1);
    cueList.scrollToCalls.length = 0;

    snapshot = {
      ...snapshot,
      currentSubtitle: {
        text: 'line-9',
        startTime: 18,
        endTime: 19.5,
      },
      currentTimeSec: 18.1,
    };

    await modal.refreshSubtitleSidebarSnapshot();

    assert.equal(state.subtitleSidebarActiveCueIndex, 9);
    assert.equal(cueList.scrollTop, 260);
    assert.deepEqual(cueList.scrollToCalls, []);

    cueList.scrollToCalls.length = 0;
    snapshot = {
      ...snapshot,
      currentSubtitle: {
        text: 'line-10',
        startTime: 20,
        endTime: 21.5,
      },
      currentTimeSec: 20.1,
    };

    await modal.refreshSubtitleSidebarSnapshot();

    assert.equal(state.subtitleSidebarActiveCueIndex, 10);
    assert.deepEqual(cueList.scrollToCalls.at(-1), {
      top: 300,
      behavior: 'smooth',
    });
  } finally {
    Object.defineProperty(globalThis, 'window', { configurable: true, value: previousWindow });
    Object.defineProperty(globalThis, 'document', { configurable: true, value: previousDocument });
  }
});

test('subtitle sidebar polling schedules serialized timeouts instead of intervals', async () => {
  const globals = globalThis as typeof globalThis & {
    window?: unknown;
    document?: unknown;
    setTimeout?: typeof globalThis.setTimeout;
    clearTimeout?: typeof globalThis.clearTimeout;
    setInterval?: typeof globalThis.setInterval;
    clearInterval?: typeof globalThis.clearInterval;
  };
  const previousWindow = globals.window;
  const previousDocument = globals.document;
  const previousSetTimeout = globals.setTimeout;
  const previousClearTimeout = globals.clearTimeout;
  const previousSetInterval = globals.setInterval;
  const previousClearInterval = globals.clearInterval;
  let timeoutCount = 0;
  let intervalCount = 0;

  Object.defineProperty(globalThis, 'setTimeout', {
    configurable: true,
    value: (callback: (...args: never[]) => void) => {
      timeoutCount += 1;
      return timeoutCount as unknown as ReturnType<typeof setTimeout>;
    },
  });
  Object.defineProperty(globalThis, 'clearTimeout', {
    configurable: true,
    value: () => {},
  });
  Object.defineProperty(globalThis, 'setInterval', {
    configurable: true,
    value: () => {
      intervalCount += 1;
      return intervalCount as unknown as ReturnType<typeof setInterval>;
    },
  });
  Object.defineProperty(globalThis, 'clearInterval', {
    configurable: true,
    value: () => {},
  });

  const snapshot: SubtitleSidebarSnapshot = {
    cues: [{ startTime: 1, endTime: 2, text: 'first' }],
    currentSubtitle: {
      text: 'first',
      startTime: 1,
      endTime: 2,
    },
    currentTimeSec: 1.1,
    config: {
      enabled: true,
      autoOpen: false,
      layout: 'overlay',
      toggleKey: 'Backslash',
      pauseVideoOnHover: false,
      autoScroll: true,
      maxWidth: 420,
      opacity: 0.92,
      backgroundColor: 'rgba(54, 58, 79, 0.88)',
      textColor: '#cad3f5',
      fontFamily: '"Iosevka Aile", sans-serif',
      fontSize: 17,
      timestampColor: '#a5adcb',
      activeLineColor: '#f5bde6',
      activeLineBackgroundColor: 'rgba(138, 173, 244, 0.22)',
      hoverLineBackgroundColor: 'rgba(54, 58, 79, 0.84)',
    },
  };

  Object.defineProperty(globalThis, 'window', {
    configurable: true,
    value: {
      electronAPI: {
        getSubtitleSidebarSnapshot: async () => snapshot,
        sendMpvCommand: () => {},
      } as unknown as ElectronAPI,
    },
  });
  Object.defineProperty(globalThis, 'document', {
    configurable: true,
    value: {
      createElement: () => createCueRow(),
      body: {
        classList: createClassList(),
      },
      documentElement: {
        style: {
          setProperty: () => {},
        },
      },
    },
  });

  try {
    const state = createRendererState();
    const ctx = {
      dom: {
        overlay: { classList: createClassList() },
        subtitleSidebarModal: {
          classList: createClassList(['hidden']),
          setAttribute: () => {},
          style: { setProperty: () => {} },
          addEventListener: () => {},
        },
        subtitleSidebarContent: {
          classList: createClassList(),
          getBoundingClientRect: () => ({ width: 420 }),
        },
        subtitleSidebarClose: { addEventListener: () => {} },
        subtitleSidebarStatus: { textContent: '' },
        subtitleSidebarList: createListStub(),
      },
      state,
    };

    const modal = createSubtitleSidebarModal(ctx as never, {
      modalStateReader: { isAnyModalOpen: () => false },
    });

    await modal.openSubtitleSidebarModal();

    assert.equal(timeoutCount > 0, true);
    assert.equal(intervalCount, 0);
  } finally {
    Object.defineProperty(globalThis, 'setTimeout', {
      configurable: true,
      value: previousSetTimeout,
    });
    Object.defineProperty(globalThis, 'clearTimeout', {
      configurable: true,
      value: previousClearTimeout,
    });
    Object.defineProperty(globalThis, 'setInterval', {
      configurable: true,
      value: previousSetInterval,
    });
    Object.defineProperty(globalThis, 'clearInterval', {
      configurable: true,
      value: previousClearInterval,
    });
    Object.defineProperty(globalThis, 'window', { configurable: true, value: previousWindow });
    Object.defineProperty(globalThis, 'document', { configurable: true, value: previousDocument });
  }
});

test('subtitle sidebar closes and resumes a hover pause', async () => {
  const globals = globalThis as typeof globalThis & { window?: unknown; document?: unknown };
  const previousWindow = globals.window;
  const previousDocument = globals.document;
  const mpvCommands: Array<Array<string | number>> = [];
  const modalListeners = new Map<string, Array<() => void>>();
  const contentListeners = new Map<string, Array<() => void>>();

  const snapshot: SubtitleSidebarSnapshot = {
    cues: [{ startTime: 1, endTime: 2, text: 'first' }],
    currentSubtitle: {
      text: 'first',
      startTime: 1,
      endTime: 2,
    },
    currentTimeSec: 1.1,
    config: {
      enabled: true,
      autoOpen: false,
      layout: 'overlay',
      toggleKey: 'Backslash',
      pauseVideoOnHover: true,
      autoScroll: true,
      maxWidth: 420,
      opacity: 0.92,
      backgroundColor: 'rgba(54, 58, 79, 0.88)',
      textColor: '#cad3f5',
      fontFamily: '"Iosevka Aile", sans-serif',
      fontSize: 17,
      timestampColor: '#a5adcb',
      activeLineColor: '#f5bde6',
      activeLineBackgroundColor: 'rgba(138, 173, 244, 0.22)',
      hoverLineBackgroundColor: 'rgba(54, 58, 79, 0.84)',
    },
  };

  Object.defineProperty(globalThis, 'window', {
    configurable: true,
    value: {
      addEventListener: () => {},
      removeEventListener: () => {},
      electronAPI: {
        getSubtitleSidebarSnapshot: async () => snapshot,
        getPlaybackPaused: async () => false,
        sendMpvCommand: (command: Array<string | number>) => {
          mpvCommands.push(command);
        },
      } as unknown as ElectronAPI,
    },
  });
  Object.defineProperty(globalThis, 'document', {
    configurable: true,
    value: {
      createElement: () => createCueRow(),
      body: {
        classList: createClassList(),
      },
      documentElement: {
        style: {
          setProperty: () => {},
        },
      },
    },
  });

  try {
    const state = createRendererState();
    const ctx = {
      dom: {
        overlay: { classList: createClassList() },
        subtitleSidebarModal: {
          classList: createClassList(['hidden']),
          setAttribute: () => {},
          style: { setProperty: () => {} },
          addEventListener: (type: string, listener: () => void) => {
            const bucket = modalListeners.get(type) ?? [];
            bucket.push(listener);
            modalListeners.set(type, bucket);
          },
        },
        subtitleSidebarContent: {
          classList: createClassList(),
          getBoundingClientRect: () => ({ width: 420 }),
          addEventListener: (type: string, listener: () => void) => {
            const bucket = contentListeners.get(type) ?? [];
            bucket.push(listener);
            contentListeners.set(type, bucket);
          },
        },
        subtitleSidebarClose: { addEventListener: () => {} },
        subtitleSidebarStatus: { textContent: '' },
        subtitleSidebarList: createListStub(),
      },
      state,
    };

    const modal = createSubtitleSidebarModal(ctx as never, {
      modalStateReader: { isAnyModalOpen: () => false },
    });
    modal.wireDomEvents();

    await modal.openSubtitleSidebarModal();
    await modal.refreshSubtitleSidebarSnapshot();
    mpvCommands.length = 0;
    await contentListeners.get('mouseenter')?.[0]?.();

    assert.deepEqual(mpvCommands.at(-1), ['set_property', 'pause', 'yes']);

    modal.closeSubtitleSidebarModal();

    assert.deepEqual(mpvCommands.at(-1), ['set_property', 'pause', 'no']);
    assert.equal(state.subtitleSidebarPausedByHover, false);
  } finally {
    Object.defineProperty(globalThis, 'window', { configurable: true, value: previousWindow });
    Object.defineProperty(globalThis, 'document', { configurable: true, value: previousDocument });
  }
});

test('subtitle sidebar hover pause ignores playback-state IPC failures', async () => {
  const globals = globalThis as typeof globalThis & { window?: unknown; document?: unknown };
  const previousWindow = globals.window;
  const previousDocument = globals.document;
  const mpvCommands: Array<Array<string | number>> = [];
  const modalListeners = new Map<string, Array<() => Promise<void> | void>>();
  const contentListeners = new Map<string, Array<() => Promise<void> | void>>();

  const snapshot: SubtitleSidebarSnapshot = {
    cues: [{ startTime: 1, endTime: 2, text: 'first' }],
    currentSubtitle: {
      text: 'first',
      startTime: 1,
      endTime: 2,
    },
    currentTimeSec: 1.1,
    config: {
      enabled: true,
      autoOpen: false,
      layout: 'overlay',
      toggleKey: 'Backslash',
      pauseVideoOnHover: true,
      autoScroll: true,
      maxWidth: 420,
      opacity: 0.92,
      backgroundColor: 'rgba(54, 58, 79, 0.88)',
      textColor: '#cad3f5',
      fontFamily: '"Iosevka Aile", sans-serif',
      fontSize: 17,
      timestampColor: '#a5adcb',
      activeLineColor: '#f5bde6',
      activeLineBackgroundColor: 'rgba(138, 173, 244, 0.22)',
      hoverLineBackgroundColor: 'rgba(54, 58, 79, 0.84)',
    },
  };

  Object.defineProperty(globalThis, 'window', {
    configurable: true,
    value: {
      addEventListener: () => {},
      removeEventListener: () => {},
      electronAPI: {
        getSubtitleSidebarSnapshot: async () => snapshot,
        getPlaybackPaused: async () => {
          throw new Error('ipc failed');
        },
        sendMpvCommand: (command: Array<string | number>) => {
          mpvCommands.push(command);
        },
      } as unknown as ElectronAPI,
    },
  });
  Object.defineProperty(globalThis, 'document', {
    configurable: true,
    value: {
      createElement: () => createCueRow(),
      body: {
        classList: createClassList(),
      },
      documentElement: {
        style: {
          setProperty: () => {},
        },
      },
    },
  });

  try {
    const state = createRendererState();
    const ctx = {
      dom: {
        overlay: { classList: createClassList() },
        subtitleSidebarModal: {
          classList: createClassList(['hidden']),
          setAttribute: () => {},
          style: { setProperty: () => {} },
          addEventListener: (type: string, listener: () => Promise<void> | void) => {
            const bucket = modalListeners.get(type) ?? [];
            bucket.push(listener);
            modalListeners.set(type, bucket);
          },
        },
        subtitleSidebarContent: {
          classList: createClassList(),
          getBoundingClientRect: () => ({ width: 420 }),
          addEventListener: (type: string, listener: () => Promise<void> | void) => {
            const bucket = contentListeners.get(type) ?? [];
            bucket.push(listener);
            contentListeners.set(type, bucket);
          },
        },
        subtitleSidebarClose: { addEventListener: () => {} },
        subtitleSidebarStatus: { textContent: '' },
        subtitleSidebarList: createListStub(),
      },
      state,
    };

    const modal = createSubtitleSidebarModal(ctx as never, {
      modalStateReader: { isAnyModalOpen: () => false },
    });
    modal.wireDomEvents();

    await modal.openSubtitleSidebarModal();
    await assert.doesNotReject(async () => {
      await contentListeners.get('mouseenter')?.[0]?.();
    });

    assert.equal(state.subtitleSidebarPausedByHover, false);
    assert.equal(
      mpvCommands.some((command) => command[0] === 'set_property' && command[2] === 'yes'),
      false,
    );
  } finally {
    Object.defineProperty(globalThis, 'window', { configurable: true, value: previousWindow });
    Object.defineProperty(globalThis, 'document', { configurable: true, value: previousDocument });
  }
});

test('subtitle sidebar embedded layout reserves and releases mpv right margin', async () => {
  const globals = globalThis as typeof globalThis & { window?: unknown; document?: unknown };
  const previousWindow = globals.window;
  const previousDocument = globals.document;
  const mpvCommands: Array<Array<string | number>> = [];

  const snapshot: SubtitleSidebarSnapshot = {
    cues: [{ startTime: 1, endTime: 2, text: 'first' }],
    currentSubtitle: {
      text: 'first',
      startTime: 1,
      endTime: 2,
    },
    currentTimeSec: 1.1,
    config: {
      enabled: true,
      autoOpen: false,
      layout: 'embedded',
      toggleKey: 'Backslash',
      pauseVideoOnHover: false,
      autoScroll: true,
      maxWidth: 360,
      opacity: 0.92,
      backgroundColor: 'rgba(54, 58, 79, 0.88)',
      textColor: '#cad3f5',
      fontFamily: '"Iosevka Aile", sans-serif',
      fontSize: 17,
      timestampColor: '#a5adcb',
      activeLineColor: '#f5bde6',
      activeLineBackgroundColor: 'rgba(138, 173, 244, 0.22)',
      hoverLineBackgroundColor: 'rgba(54, 58, 79, 0.84)',
    },
  };

  const rootStyleCalls: Array<[string, string]> = [];
  const bodyClassList = createClassList();

  Object.defineProperty(globalThis, 'window', {
    configurable: true,
    value: {
      innerWidth: 1200,
      electronAPI: {
        getSubtitleSidebarSnapshot: async () => snapshot,
        sendMpvCommand: (command: Array<string | number>) => {
          mpvCommands.push(command);
        },
      } as unknown as ElectronAPI,
      addEventListener: () => {},
      removeEventListener: () => {},
    },
  });
  Object.defineProperty(globalThis, 'document', {
    configurable: true,
    value: {
      createElement: () => createCueRow(),
      body: {
        classList: bodyClassList,
      },
      documentElement: {
        style: {
          setProperty: (name: string, value: string) => {
            rootStyleCalls.push([name, value]);
          },
        },
      },
    },
  });

  try {
    const state = createRendererState();
    const cueList = createListStub();
    const modalClassList = createClassList(['hidden']);
    const contentClassList = createClassList();
    const ctx = {
      dom: {
        overlay: { classList: createClassList() },
        subtitleSidebarModal: {
          classList: modalClassList,
          setAttribute: () => {},
          style: { setProperty: () => {} },
          addEventListener: () => {},
        },
        subtitleSidebarContent: {
          classList: contentClassList,
          getBoundingClientRect: () => ({ width: 360 }),
        },
        subtitleSidebarClose: { addEventListener: () => {} },
        subtitleSidebarStatus: { textContent: '' },
        subtitleSidebarList: cueList,
      },
      platform: {
        shouldToggleMouseIgnore: false,
      },
      state,
    };

    const modal = createSubtitleSidebarModal(ctx as never, {
      modalStateReader: { isAnyModalOpen: () => false },
    });

    await modal.openSubtitleSidebarModal();

    assert.ok(
      mpvCommands.some(
        (command) =>
          command[0] === 'set_property' &&
          command[1] === 'video-margin-ratio-right' &&
          command[2] === 0.3,
      ),
    );
    assert.ok(
      mpvCommands.some(
        (command) =>
          command[0] === 'set_property' && command[1] === 'osd-align-x' && command[2] === 'left',
      ),
    );
    assert.ok(
      mpvCommands.some(
        (command) =>
          command[0] === 'set_property' && command[1] === 'osd-align-y' && command[2] === 'top',
      ),
    );
    assert.ok(
      mpvCommands.some(
        (command) =>
          command[0] === 'set_property' &&
          command[1] === 'user-data/osc/margins' &&
          command[2] === '{"l":0,"r":0.3,"t":0,"b":0}',
      ),
    );
    assert.ok(bodyClassList.contains('subtitle-sidebar-embedded-open'));
    assert.ok(
      rootStyleCalls.some(
        ([name, value]) => name === '--subtitle-sidebar-reserved-width' && value === '360px',
      ),
    );

    modal.closeSubtitleSidebarModal();

    assert.deepEqual(mpvCommands.at(-5), ['set_property', 'video-margin-ratio-right', 0]);
    assert.deepEqual(mpvCommands.at(-4), ['set_property', 'osd-align-x', 'left']);
    assert.deepEqual(mpvCommands.at(-3), ['set_property', 'osd-align-y', 'top']);
    assert.deepEqual(mpvCommands.at(-2), [
      'set_property',
      'user-data/osc/margins',
      '{"l":0,"r":0,"t":0,"b":0}',
    ]);
    assert.deepEqual(mpvCommands.at(-1), ['set_property', 'video-pan-x', 0]);
    assert.equal(bodyClassList.contains('subtitle-sidebar-embedded-open'), false);
    assert.deepEqual(rootStyleCalls.at(-1), ['--subtitle-sidebar-reserved-width', '0px']);
  } finally {
    Object.defineProperty(globalThis, 'window', { configurable: true, value: previousWindow });
    Object.defineProperty(globalThis, 'document', { configurable: true, value: previousDocument });
  }
});

test('subtitle sidebar embedded layout measures reserved width after embedded classes apply', async () => {
  const globals = globalThis as typeof globalThis & { window?: unknown; document?: unknown };
  const previousWindow = globals.window;
  const previousDocument = globals.document;
  const mpvCommands: Array<Array<string | number>> = [];
  const rootStyleCalls: Array<[string, string]> = [];
  const bodyClassList = createClassList();
  const contentClassList = createClassList();

  const snapshot: SubtitleSidebarSnapshot = {
    cues: [{ startTime: 1, endTime: 2, text: 'first' }],
    currentSubtitle: {
      text: 'first',
      startTime: 1,
      endTime: 2,
    },
    currentTimeSec: 1.1,
    config: {
      enabled: true,
      autoOpen: false,
      layout: 'embedded',
      toggleKey: 'Backslash',
      pauseVideoOnHover: false,
      autoScroll: true,
      maxWidth: 420,
      opacity: 0.92,
      backgroundColor: 'rgba(54, 58, 79, 0.88)',
      textColor: '#cad3f5',
      fontFamily: '"Iosevka Aile", sans-serif',
      fontSize: 17,
      timestampColor: '#a5adcb',
      activeLineColor: '#f5bde6',
      activeLineBackgroundColor: 'rgba(138, 173, 244, 0.22)',
      hoverLineBackgroundColor: 'rgba(54, 58, 79, 0.84)',
    },
  };

  Object.defineProperty(globalThis, 'window', {
    configurable: true,
    value: {
      innerWidth: 1200,
      electronAPI: {
        getSubtitleSidebarSnapshot: async () => snapshot,
        sendMpvCommand: (command: Array<string | number>) => {
          mpvCommands.push(command);
        },
      } as unknown as ElectronAPI,
      addEventListener: () => {},
      removeEventListener: () => {},
    },
  });
  Object.defineProperty(globalThis, 'document', {
    configurable: true,
    value: {
      createElement: () => createCueRow(),
      body: {
        classList: bodyClassList,
      },
      documentElement: {
        style: {
          setProperty: (name: string, value: string) => {
            rootStyleCalls.push([name, value]);
          },
        },
      },
    },
  });

  try {
    const state = createRendererState();
    const ctx = {
      dom: {
        overlay: { classList: createClassList() },
        subtitleSidebarModal: {
          classList: createClassList(['hidden']),
          setAttribute: () => {},
          style: { setProperty: () => {} },
          addEventListener: () => {},
        },
        subtitleSidebarContent: {
          classList: contentClassList,
          getBoundingClientRect: () => ({
            width: contentClassList.contains('subtitle-sidebar-content-embedded') ? 300 : 0,
          }),
        },
        subtitleSidebarClose: { addEventListener: () => {} },
        subtitleSidebarStatus: { textContent: '' },
        subtitleSidebarList: createListStub(),
      },
      platform: {
        shouldToggleMouseIgnore: false,
      },
      state,
    };

    const modal = createSubtitleSidebarModal(ctx as never, {
      modalStateReader: { isAnyModalOpen: () => false },
    });

    await modal.openSubtitleSidebarModal();

    assert.ok(bodyClassList.contains('subtitle-sidebar-embedded-open'));
    assert.ok(
      rootStyleCalls.some(
        ([name, value]) => name === '--subtitle-sidebar-reserved-width' && value === '300px',
      ),
    );
    assert.ok(
      mpvCommands.some(
        (command) =>
          command[0] === 'set_property' &&
          command[1] === 'video-margin-ratio-right' &&
          command[2] === 0.25,
      ),
    );
  } finally {
    Object.defineProperty(globalThis, 'window', { configurable: true, value: previousWindow });
    Object.defineProperty(globalThis, 'document', { configurable: true, value: previousDocument });
  }
});

test('subtitle sidebar embedded layout restores macOS and Windows passthrough outside sidebar hover', async () => {
  const globals = globalThis as typeof globalThis & { window?: unknown; document?: unknown };
  const previousWindow = globals.window;
  const previousDocument = globals.document;
  const mpvCommands: Array<Array<string | number>> = [];
  const ignoreMouseCalls: Array<[boolean, { forward?: boolean } | undefined]> = [];
  const modalListeners = new Map<string, Array<() => void>>();
  const contentListeners = new Map<string, Array<() => void>>();

  const snapshot: SubtitleSidebarSnapshot = {
    cues: [{ startTime: 1, endTime: 2, text: 'first' }],
    currentSubtitle: {
      text: 'first',
      startTime: 1,
      endTime: 2,
    },
    currentTimeSec: 1.1,
    config: {
      enabled: true,
      autoOpen: false,
      layout: 'embedded',
      toggleKey: 'Backslash',
      pauseVideoOnHover: false,
      autoScroll: true,
      maxWidth: 360,
      opacity: 0.92,
      backgroundColor: 'rgba(54, 58, 79, 0.88)',
      textColor: '#cad3f5',
      fontFamily: '"Iosevka Aile", sans-serif',
      fontSize: 17,
      timestampColor: '#a5adcb',
      activeLineColor: '#f5bde6',
      activeLineBackgroundColor: 'rgba(138, 173, 244, 0.22)',
      hoverLineBackgroundColor: 'rgba(54, 58, 79, 0.84)',
    },
  };

  Object.defineProperty(globalThis, 'window', {
    configurable: true,
    value: {
      innerWidth: 1200,
      electronAPI: {
        getSubtitleSidebarSnapshot: async () => snapshot,
        sendMpvCommand: (command: Array<string | number>) => {
          mpvCommands.push(command);
        },
        setIgnoreMouseEvents: (ignore: boolean, options?: { forward?: boolean }) => {
          ignoreMouseCalls.push([ignore, options]);
        },
      } as unknown as ElectronAPI,
      addEventListener: () => {},
      removeEventListener: () => {},
    },
  });
  Object.defineProperty(globalThis, 'document', {
    configurable: true,
    value: {
      createElement: () => createCueRow(),
      body: {
        classList: createClassList(),
      },
      documentElement: {
        style: {
          setProperty: () => {},
        },
      },
    },
  });

  try {
    const state = createRendererState();
    const ctx = {
      dom: {
        overlay: { classList: createClassList() },
        subtitleSidebarModal: {
          classList: createClassList(['hidden']),
          setAttribute: () => {},
          style: { setProperty: () => {} },
          addEventListener: (type: string, listener: () => void) => {
            const bucket = modalListeners.get(type) ?? [];
            bucket.push(listener);
            modalListeners.set(type, bucket);
          },
        },
        subtitleSidebarContent: {
          classList: createClassList(),
          getBoundingClientRect: () => ({ width: 360 }),
          addEventListener: (type: string, listener: () => void) => {
            const bucket = contentListeners.get(type) ?? [];
            bucket.push(listener);
            contentListeners.set(type, bucket);
          },
        },
        subtitleSidebarClose: { addEventListener: () => {} },
        subtitleSidebarStatus: { textContent: '' },
        subtitleSidebarList: createListStub(),
      },
      platform: {
        shouldToggleMouseIgnore: true,
      },
      state,
    };

    const modal = createSubtitleSidebarModal(ctx as never, {
      modalStateReader: { isAnyModalOpen: () => false },
    });
    modal.wireDomEvents();

    await modal.openSubtitleSidebarModal();
    assert.deepEqual(ignoreMouseCalls.at(-1), [true, { forward: true }]);

    contentListeners.get('mouseenter')?.[0]?.();
    assert.deepEqual(ignoreMouseCalls.at(-1), [false, undefined]);

    contentListeners.get('mouseleave')?.[0]?.();
    assert.deepEqual(ignoreMouseCalls.at(-1), [true, { forward: true }]);

    state.isOverSubtitle = true;
    contentListeners.get('mouseenter')?.[0]?.();
    contentListeners.get('mouseleave')?.[0]?.();
    assert.deepEqual(ignoreMouseCalls.at(-1), [false, undefined]);

    void mpvCommands;
  } finally {
    Object.defineProperty(globalThis, 'window', { configurable: true, value: previousWindow });
    Object.defineProperty(globalThis, 'document', { configurable: true, value: previousDocument });
  }
});

test('subtitle sidebar overlay layout restores macOS and Windows passthrough outside sidebar hover', async () => {
  const globals = globalThis as typeof globalThis & { window?: unknown; document?: unknown };
  const previousWindow = globals.window;
  const previousDocument = globals.document;
  const mpvCommands: Array<Array<string | number>> = [];
  const ignoreMouseCalls: Array<[boolean, { forward?: boolean } | undefined]> = [];
  const modalListeners = new Map<string, Array<() => void>>();
  const contentListeners = new Map<string, Array<() => void>>();

  const snapshot: SubtitleSidebarSnapshot = {
    cues: [{ startTime: 1, endTime: 2, text: 'first' }],
    currentSubtitle: {
      text: 'first',
      startTime: 1,
      endTime: 2,
    },
    currentTimeSec: 1.1,
    config: {
      enabled: true,
      autoOpen: false,
      layout: 'overlay',
      toggleKey: 'Backslash',
      pauseVideoOnHover: false,
      autoScroll: true,
      maxWidth: 360,
      opacity: 0.92,
      backgroundColor: 'rgba(54, 58, 79, 0.88)',
      textColor: '#cad3f5',
      fontFamily: '"Iosevka Aile", sans-serif',
      fontSize: 17,
      timestampColor: '#a5adcb',
      activeLineColor: '#f5bde6',
      activeLineBackgroundColor: 'rgba(138, 173, 244, 0.22)',
      hoverLineBackgroundColor: 'rgba(54, 58, 79, 0.84)',
    },
  };

  Object.defineProperty(globalThis, 'window', {
    configurable: true,
    value: {
      innerWidth: 1200,
      electronAPI: {
        getSubtitleSidebarSnapshot: async () => snapshot,
        sendMpvCommand: (command: Array<string | number>) => {
          mpvCommands.push(command);
        },
        setIgnoreMouseEvents: (ignore: boolean, options?: { forward?: boolean }) => {
          ignoreMouseCalls.push([ignore, options]);
        },
      } as unknown as ElectronAPI,
      addEventListener: () => {},
      removeEventListener: () => {},
    },
  });
  Object.defineProperty(globalThis, 'document', {
    configurable: true,
    value: {
      createElement: () => createCueRow(),
      body: {
        classList: createClassList(),
      },
      documentElement: {
        style: {
          setProperty: () => {},
        },
      },
    },
  });

  try {
    const state = createRendererState();
    const ctx = {
      dom: {
        overlay: { classList: createClassList() },
        subtitleSidebarModal: {
          classList: createClassList(['hidden']),
          setAttribute: () => {},
          style: { setProperty: () => {} },
          addEventListener: (type: string, listener: () => void) => {
            const bucket = modalListeners.get(type) ?? [];
            bucket.push(listener);
            modalListeners.set(type, bucket);
          },
        },
        subtitleSidebarContent: {
          classList: createClassList(),
          getBoundingClientRect: () => ({ width: 360 }),
          addEventListener: (type: string, listener: () => void) => {
            const bucket = contentListeners.get(type) ?? [];
            bucket.push(listener);
            contentListeners.set(type, bucket);
          },
        },
        subtitleSidebarClose: { addEventListener: () => {} },
        subtitleSidebarStatus: { textContent: '' },
        subtitleSidebarList: createListStub(),
      },
      platform: {
        shouldToggleMouseIgnore: true,
      },
      state,
    };

    const modal = createSubtitleSidebarModal(ctx as never, {
      modalStateReader: { isAnyModalOpen: () => false },
    });
    modal.wireDomEvents();

    assert.equal(modalListeners.get('mouseenter')?.length ?? 0, 0);
    assert.equal(modalListeners.get('mouseleave')?.length ?? 0, 0);
    assert.equal(contentListeners.get('mouseenter')?.length ?? 0, 1);
    assert.equal(contentListeners.get('mouseleave')?.length ?? 0, 1);

    await modal.openSubtitleSidebarModal();
    assert.deepEqual(ignoreMouseCalls.at(-1), [true, { forward: true }]);

    contentListeners.get('mouseenter')?.[0]?.();
    assert.deepEqual(ignoreMouseCalls.at(-1), [false, undefined]);

    contentListeners.get('mouseleave')?.[0]?.();
    assert.deepEqual(ignoreMouseCalls.at(-1), [true, { forward: true }]);

    void mpvCommands;
  } finally {
    Object.defineProperty(globalThis, 'window', { configurable: true, value: previousWindow });
    Object.defineProperty(globalThis, 'document', { configurable: true, value: previousDocument });
  }
});

test('subtitle sidebar overlay layout only stays interactive while focus remains inside the sidebar panel', async () => {
  const globals = globalThis as typeof globalThis & { window?: unknown; document?: unknown };
  const previousWindow = globals.window;
  const previousDocument = globals.document;
  const ignoreMouseCalls: Array<[boolean, { forward?: boolean } | undefined]> = [];
  const contentListeners = new Map<string, Array<(event?: FocusEvent) => void>>();

  const snapshot: SubtitleSidebarSnapshot = {
    cues: [{ startTime: 1, endTime: 2, text: 'first' }],
    currentSubtitle: {
      text: 'first',
      startTime: 1,
      endTime: 2,
    },
    currentTimeSec: 1.1,
    config: {
      enabled: true,
      autoOpen: false,
      layout: 'overlay',
      toggleKey: 'Backslash',
      pauseVideoOnHover: false,
      autoScroll: true,
      maxWidth: 360,
      opacity: 0.92,
      backgroundColor: 'rgba(54, 58, 79, 0.88)',
      textColor: '#cad3f5',
      fontFamily: '"Iosevka Aile", sans-serif',
      fontSize: 17,
      timestampColor: '#a5adcb',
      activeLineColor: '#f5bde6',
      activeLineBackgroundColor: 'rgba(138, 173, 244, 0.22)',
      hoverLineBackgroundColor: 'rgba(54, 58, 79, 0.84)',
    },
  };

  Object.defineProperty(globalThis, 'window', {
    configurable: true,
    value: {
      innerWidth: 1200,
      electronAPI: {
        getSubtitleSidebarSnapshot: async () => snapshot,
        sendMpvCommand: () => {},
        setIgnoreMouseEvents: (ignore: boolean, options?: { forward?: boolean }) => {
          ignoreMouseCalls.push([ignore, options]);
        },
      } as unknown as ElectronAPI,
      addEventListener: () => {},
      removeEventListener: () => {},
    },
  });
  Object.defineProperty(globalThis, 'document', {
    configurable: true,
    value: {
      createElement: () => createCueRow(),
      body: {
        classList: createClassList(),
      },
      documentElement: {
        style: {
          setProperty: () => {},
        },
      },
    },
  });

  try {
    const state = createRendererState();
    const sidebarContent = {
      classList: createClassList(),
      getBoundingClientRect: () => ({ width: 360 }),
      addEventListener: (type: string, listener: (event?: FocusEvent) => void) => {
        const bucket = contentListeners.get(type) ?? [];
        bucket.push(listener);
        contentListeners.set(type, bucket);
      },
      contains: () => false,
    };
    const ctx = {
      dom: {
        overlay: { classList: createClassList() },
        subtitleSidebarModal: {
          classList: createClassList(['hidden']),
          setAttribute: () => {},
          style: { setProperty: () => {} },
          addEventListener: () => {},
        },
        subtitleSidebarContent: sidebarContent,
        subtitleSidebarClose: { addEventListener: () => {} },
        subtitleSidebarStatus: { textContent: '' },
        subtitleSidebarList: createListStub(),
      },
      platform: {
        shouldToggleMouseIgnore: true,
      },
      state,
    };

    const modal = createSubtitleSidebarModal(ctx as never, {
      modalStateReader: { isAnyModalOpen: () => false },
    });
    modal.wireDomEvents();

    await modal.openSubtitleSidebarModal();
    assert.deepEqual(ignoreMouseCalls.at(-1), [true, { forward: true }]);

    contentListeners.get('focusin')?.[0]?.();
    assert.deepEqual(ignoreMouseCalls.at(-1), [false, undefined]);

    contentListeners.get('focusout')?.[0]?.({ relatedTarget: null } as FocusEvent);
    assert.deepEqual(ignoreMouseCalls.at(-1), [true, { forward: true }]);
  } finally {
    Object.defineProperty(globalThis, 'window', { configurable: true, value: previousWindow });
    Object.defineProperty(globalThis, 'document', { configurable: true, value: previousDocument });
  }
});

test('closing embedded subtitle sidebar recomputes passthrough from remaining subtitle hover state', async () => {
  const globals = globalThis as typeof globalThis & { window?: unknown; document?: unknown };
  const previousWindow = globals.window;
  const previousDocument = globals.document;
  const ignoreMouseCalls: Array<[boolean, { forward?: boolean } | undefined]> = [];

  const snapshot: SubtitleSidebarSnapshot = {
    cues: [{ startTime: 1, endTime: 2, text: 'first' }],
    currentSubtitle: {
      text: 'first',
      startTime: 1,
      endTime: 2,
    },
    currentTimeSec: 1.1,
    config: {
      enabled: true,
      autoOpen: false,
      layout: 'embedded',
      toggleKey: 'Backslash',
      pauseVideoOnHover: false,
      autoScroll: true,
      maxWidth: 360,
      opacity: 0.92,
      backgroundColor: 'rgba(54, 58, 79, 0.88)',
      textColor: '#cad3f5',
      fontFamily: '"Iosevka Aile", sans-serif',
      fontSize: 17,
      timestampColor: '#a5adcb',
      activeLineColor: '#f5bde6',
      activeLineBackgroundColor: 'rgba(138, 173, 244, 0.22)',
      hoverLineBackgroundColor: 'rgba(54, 58, 79, 0.84)',
    },
  };

  Object.defineProperty(globalThis, 'window', {
    configurable: true,
    value: {
      innerWidth: 1200,
      electronAPI: {
        getSubtitleSidebarSnapshot: async () => snapshot,
        sendMpvCommand: () => {},
        setIgnoreMouseEvents: (ignore: boolean, options?: { forward?: boolean }) => {
          ignoreMouseCalls.push([ignore, options]);
        },
      } as unknown as ElectronAPI,
      addEventListener: () => {},
      removeEventListener: () => {},
    },
  });
  Object.defineProperty(globalThis, 'document', {
    configurable: true,
    value: {
      createElement: () => createCueRow(),
      body: {
        classList: createClassList(),
      },
      documentElement: {
        style: {
          setProperty: () => {},
        },
      },
    },
  });

  try {
    const state = createRendererState();
    const ctx = {
      dom: {
        overlay: { classList: createClassList() },
        subtitleSidebarModal: {
          classList: createClassList(['hidden']),
          setAttribute: () => {},
          style: { setProperty: () => {} },
          addEventListener: () => {},
        },
        subtitleSidebarContent: {
          classList: createClassList(),
          getBoundingClientRect: () => ({ width: 360 }),
        },
        subtitleSidebarClose: { addEventListener: () => {} },
        subtitleSidebarStatus: { textContent: '' },
        subtitleSidebarList: createListStub(),
      },
      platform: {
        shouldToggleMouseIgnore: true,
      },
      state,
    };

    const modal = createSubtitleSidebarModal(ctx as never, {
      modalStateReader: { isAnyModalOpen: () => false },
    });

    await modal.openSubtitleSidebarModal();
    state.isOverSubtitle = true;
    modal.closeSubtitleSidebarModal();
    assert.deepEqual(ignoreMouseCalls.at(-1), [false, undefined]);

    await modal.openSubtitleSidebarModal();
    state.isOverSubtitle = false;
    modal.closeSubtitleSidebarModal();
    assert.deepEqual(ignoreMouseCalls.at(-1), [true, { forward: true }]);
  } finally {
    Object.defineProperty(globalThis, 'window', { configurable: true, value: previousWindow });
    Object.defineProperty(globalThis, 'document', { configurable: true, value: previousDocument });
  }
});

test('subtitle sidebar resets embedded mpv margin on startup while closed', async () => {
  const globals = globalThis as typeof globalThis & { window?: unknown; document?: unknown };
  const previousWindow = globals.window;
  const previousDocument = globals.document;
  const mpvCommands: Array<Array<string | number>> = [];

  const snapshot: SubtitleSidebarSnapshot = {
    cues: [{ startTime: 1, endTime: 2, text: 'first' }],
    currentSubtitle: {
      text: 'first',
      startTime: 1,
      endTime: 2,
    },
    currentTimeSec: 1.1,
    config: {
      enabled: true,
      autoOpen: false,
      layout: 'embedded',
      toggleKey: 'Backslash',
      pauseVideoOnHover: false,
      autoScroll: true,
      maxWidth: 360,
      opacity: 0.92,
      backgroundColor: 'rgba(54, 58, 79, 0.88)',
      textColor: '#cad3f5',
      fontFamily: '"Iosevka Aile", sans-serif',
      fontSize: 17,
      timestampColor: '#a5adcb',
      activeLineColor: '#f5bde6',
      activeLineBackgroundColor: 'rgba(138, 173, 244, 0.22)',
      hoverLineBackgroundColor: 'rgba(54, 58, 79, 0.84)',
    },
  };

  Object.defineProperty(globalThis, 'window', {
    configurable: true,
    value: {
      innerWidth: 1200,
      electronAPI: {
        getSubtitleSidebarSnapshot: async () => snapshot,
        sendMpvCommand: (command: Array<string | number>) => {
          mpvCommands.push(command);
        },
      } as unknown as ElectronAPI,
      addEventListener: () => {},
      removeEventListener: () => {},
    },
  });
  Object.defineProperty(globalThis, 'document', {
    configurable: true,
    value: {
      createElement: () => createCueRow(),
      body: {
        classList: createClassList(),
      },
      documentElement: {
        style: {
          setProperty: () => {},
        },
      },
    },
  });

  try {
    const state = createRendererState();
    const ctx = {
      dom: {
        overlay: { classList: createClassList() },
        subtitleSidebarModal: {
          classList: createClassList(['hidden']),
          setAttribute: () => {},
          style: { setProperty: () => {} },
          addEventListener: () => {},
        },
        subtitleSidebarContent: {
          classList: createClassList(),
          getBoundingClientRect: () => ({ width: 360 }),
        },
        subtitleSidebarClose: { addEventListener: () => {} },
        subtitleSidebarStatus: { textContent: '' },
        subtitleSidebarList: createListStub(),
      },
      state,
    };

    const modal = createSubtitleSidebarModal(ctx as never, {
      modalStateReader: { isAnyModalOpen: () => false },
    });

    await modal.refreshSubtitleSidebarSnapshot();

    assert.deepEqual(mpvCommands, [
      ['set_property', 'video-margin-ratio-right', 0],
      ['set_property', 'osd-align-x', 'left'],
      ['set_property', 'osd-align-y', 'top'],
      ['set_property', 'user-data/osc/margins', '{"l":0,"r":0,"t":0,"b":0}'],
      ['set_property', 'video-pan-x', 0],
    ]);
  } finally {
    Object.defineProperty(globalThis, 'window', { configurable: true, value: previousWindow });
    Object.defineProperty(globalThis, 'document', { configurable: true, value: previousDocument });
  }
});
