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
  return {
    className: '',
    classList: createClassList(),
    dataset: {} as Record<string, string>,
    textContent: '',
    offsetTop: 0,
    clientHeight: 40,
    children: [] as unknown[],
    appendChild(child: unknown) {
      this.children.push(child);
    },
    addEventListener: () => {},
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

test('findActiveSubtitleCueIndex prefers timing match before text fallback', () => {
  const cues = [
    { startTime: 1, endTime: 2, text: 'same' },
    { startTime: 3, endTime: 4, text: 'same' },
  ];

  assert.equal(findActiveSubtitleCueIndex(cues, { text: 'same', startTime: 3.1 }), 1);
  assert.equal(findActiveSubtitleCueIndex(cues, { text: 'same', startTime: null }), 0);
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
    assert.deepEqual(cueList.scrollToCalls[0], {
      top: 0,
      behavior: 'auto',
    });

    modal.seekToCue(snapshot.cues[0]!);
    assert.deepEqual(mpvCommands.at(-1), ['seek', 1.08, 'absolute+exact']);
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
    assert.ok(bodyClassList.contains('subtitle-sidebar-embedded-open'));
    assert.ok(
      rootStyleCalls.some(
        ([name, value]) => name === '--subtitle-sidebar-reserved-width' && value === '360px',
      ),
    );

    modal.closeSubtitleSidebarModal();

    assert.deepEqual(mpvCommands.at(-2), ['set_property', 'video-margin-ratio-right', 0]);
    assert.deepEqual(mpvCommands.at(-1), ['set_property', 'video-pan-x', 0]);
    assert.equal(bodyClassList.contains('subtitle-sidebar-embedded-open'), false);
    assert.deepEqual(rootStyleCalls.at(-1), ['--subtitle-sidebar-reserved-width', '0px']);
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
      ['set_property', 'video-pan-x', 0],
    ]);
  } finally {
    Object.defineProperty(globalThis, 'window', { configurable: true, value: previousWindow });
    Object.defineProperty(globalThis, 'document', { configurable: true, value: previousDocument });
  }
});
