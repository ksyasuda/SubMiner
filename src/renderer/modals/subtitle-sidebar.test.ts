import assert from 'node:assert/strict';
import test from 'node:test';

import type { ElectronAPI, SubtitleSidebarSnapshot } from '../../types';
import { createRendererState } from '../state.js';
import {
  applySidebarCssDeclarations,
  createSubtitleSidebarModal,
  findActiveSubtitleCueIndex,
} from './subtitle-sidebar.js';
import { YOMITAN_POPUP_HIDDEN_EVENT, YOMITAN_POPUP_SHOWN_EVENT } from '../yomitan-popup.js';

type SidebarConfig = SubtitleSidebarSnapshot['config'];
type SidebarModalOptions = Parameters<typeof createSubtitleSidebarModal>[1];
type MpvCommand = Array<string | number>;
type Listener = (event?: unknown) => unknown;
type ClassList = ReturnType<typeof createClassList>;

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

function createListenerTarget() {
  const listeners = new Map<string, Listener[]>();
  return {
    addEventListener: (type: string, listener: Listener) => {
      listeners.set(type, [...(listeners.get(type) ?? []), listener]);
    },
    removeEventListener: () => {},
    count: (type: string) => listeners.get(type)?.length ?? 0,
    async dispatch(type: string, event?: unknown) {
      for (const listener of listeners.get(type) ?? []) {
        await listener(event);
      }
    },
  };
}

/** Replaces globals and returns a function that puts the previous descriptors back. */
function installGlobals(values: Record<string, unknown>): () => void {
  const previous = Object.keys(values).map(
    (key) => [key, Object.getOwnPropertyDescriptor(globalThis, key)] as const,
  );
  for (const [key, value] of Object.entries(values)) {
    Object.defineProperty(globalThis, key, { configurable: true, writable: true, value });
  }
  return () => {
    for (const [key, descriptor] of previous) {
      if (descriptor) Object.defineProperty(globalThis, key, descriptor);
      else Reflect.deleteProperty(globalThis, key);
    }
  };
}

/** Folds `set_property` commands into the final value mpv would hold for each property. */
function mpvPropertyState(commands: MpvCommand[]): Map<string, string | number> {
  const properties = new Map<string, string | number>();
  for (const [command, name, value] of commands) {
    if (command === 'set_property' && typeof name === 'string' && value !== undefined) {
      properties.set(name, value);
    }
  }
  return properties;
}

const RELEASED_EMBEDDED_MARGIN = new Map<string, string | number>([
  ['video-margin-ratio-right', 0],
  ['osd-align-x', 'left'],
  ['osd-align-y', 'top'],
  ['user-data/osc/margins', '{"l":0,"r":0,"t":0,"b":0}'],
  ['video-pan-x', 0],
]);

const DEFAULT_SIDEBAR_CONFIG: SidebarConfig = {
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
};

interface SidebarHarnessOptions {
  config?: Partial<SidebarConfig>;
  cues?: SubtitleSidebarSnapshot['cues'];
  currentSubtitle?: SubtitleSidebarSnapshot['currentSubtitle'];
  currentTimeSec?: number | null;
  /** `ctx.platform.shouldToggleMouseIgnore`, i.e. macOS/Windows window passthrough. */
  toggleMouseIgnore?: boolean;
  /** Panel width from getBoundingClientRect; a function sees the panel class list. */
  contentWidth?: number | ((classList: ClassList) => number);
  electronAPI?: Partial<ElectronAPI>;
  modalOptions?: Partial<SidebarModalOptions>;
}

type SnapshotPatch = Partial<Omit<SubtitleSidebarSnapshot, 'config'>> & {
  config?: Partial<SidebarConfig>;
};

const harnessCleanups: Array<() => void> = [];

test.afterEach(() => {
  for (const cleanup of harnessCleanups.splice(0).reverse()) cleanup();
});

/**
 * Builds a sidebar modal over fake window/document/dom globals. The snapshot served to the
 * modal can be changed with `setSnapshot`. Globals are restored (and polling stopped) after
 * each test.
 */
function createSidebarHarness(options: SidebarHarnessOptions = {}) {
  let snapshot: SubtitleSidebarSnapshot = {
    sourceKey: 'test-subtitles',
    cues: options.cues ?? [{ startTime: 1, endTime: 2, text: 'first' }],
    currentSubtitle: options.currentSubtitle ?? { text: 'first', startTime: 1, endTime: 2 },
    currentTimeSec: 'currentTimeSec' in options ? options.currentTimeSec : 1.1,
    config: { ...DEFAULT_SIDEBAR_CONFIG, ...options.config },
  };

  const mpvCommands: MpvCommand[] = [];
  const ignoreMouseCalls: Array<[boolean, { forward?: boolean } | undefined]> = [];
  const modalNotifications: string[] = [];
  const visibilityChanges: boolean[] = [];
  const rootStyle = new Map<string, string>();
  const bodyClassList = createClassList();
  const windowEvents = createListenerTarget();
  const contentEvents = createListenerTarget();
  const modalEvents = createListenerTarget();

  const restoreGlobals = installGlobals({
    window: {
      innerWidth: 1200,
      addEventListener: windowEvents.addEventListener,
      removeEventListener: windowEvents.removeEventListener,
      electronAPI: {
        getSubtitleSidebarSnapshot: async () => snapshot,
        getPlaybackPaused: async () => false,
        sendMpvCommand: (command: MpvCommand) => {
          mpvCommands.push(command);
        },
        setIgnoreMouseEvents: (ignore: boolean, forward?: { forward?: boolean }) => {
          ignoreMouseCalls.push([ignore, forward]);
        },
        notifyOverlayModalOpened: (modal: string) => {
          modalNotifications.push(`open:${modal}`);
        },
        notifyOverlayModalClosed: (modal: string) => {
          modalNotifications.push(`close:${modal}`);
        },
        ...options.electronAPI,
      },
    },
    document: {
      createElement: () => createCueRow(),
      body: { classList: bodyClassList },
      documentElement: {
        style: {
          setProperty: (name: string, value: string) => {
            rootStyle.set(name, value);
          },
        },
      },
    },
  });

  const contentClassList = createClassList();
  const contentWidth = options.contentWidth ?? 420;
  const contentStyle: Record<string, string> = {};
  const dom = {
    overlay: { classList: createClassList() },
    subtitleSidebarModal: {
      classList: createClassList(['hidden']),
      setAttribute: () => {},
      style: { setProperty: () => {} },
      addEventListener: modalEvents.addEventListener,
    },
    subtitleSidebarContent: {
      classList: contentClassList,
      style: {
        setProperty: (name: string, value: string) => {
          contentStyle[name] = value;
        },
        removeProperty: (name: string) => {
          delete contentStyle[name];
        },
      },
      getBoundingClientRect: () => ({
        width: typeof contentWidth === 'function' ? contentWidth(contentClassList) : contentWidth,
      }),
      addEventListener: contentEvents.addEventListener,
      contains: () => false,
    },
    subtitleSidebarClose: { addEventListener: () => {} },
    subtitleSidebarStatus: { textContent: '' },
    subtitleSidebarList: createListStub(),
  };
  const state = createRendererState();
  const ctx = {
    dom,
    platform: { shouldToggleMouseIgnore: options.toggleMouseIgnore ?? false },
    state,
  };

  const modal = createSubtitleSidebarModal(ctx as never, {
    modalStateReader: { isAnyModalOpen: () => false },
    onVisibilityChanged: (visible) => {
      visibilityChanges.push(visible);
    },
    ...options.modalOptions,
  });

  harnessCleanups.push(() => {
    // Closing stops the snapshot poll so it cannot fire against the next test's globals.
    modal.closeSubtitleSidebarModal();
    modal.disposeDomEvents();
    restoreGlobals();
  });

  return {
    modal,
    state,
    dom,
    cueList: dom.subtitleSidebarList,
    modalClassList: dom.subtitleSidebarModal.classList,
    contentStyle,
    mpvCommands,
    ignoreMouseCalls,
    modalNotifications,
    visibilityChanges,
    rootStyle,
    bodyClassList,
    windowEvents,
    contentEvents,
    modalEvents,
    setSnapshot(patch: SnapshotPatch) {
      snapshot = { ...snapshot, ...patch, config: { ...snapshot.config, ...patch.config } };
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

test('findActiveSubtitleCueIndex prefers current subtitle timing over near-future clock lookahead', () => {
  const cues = [
    { startTime: 231, endTime: 233.2, text: 'previous' },
    { startTime: 233.05, endTime: 236, text: 'next' },
  ];

  assert.equal(findActiveSubtitleCueIndex(cues, { text: 'previous', startTime: 231 }, 233, 0), 0);
});

test('findActiveSubtitleCueIndex follows playback through empty subtitle gaps', () => {
  const cues = [
    { startTime: 0, endTime: 2, text: 'first' },
    { startTime: 100, endTime: 102, text: 'later' },
    { startTime: 105, endTime: 107, text: 'next' },
  ];

  assert.equal(findActiveSubtitleCueIndex(cues, { text: 'later', startTime: 100 }, 101, 1), 1);
  assert.equal(findActiveSubtitleCueIndex(cues, { text: '', startTime: 0 }, 103, 1), 2);
  assert.equal(findActiveSubtitleCueIndex(cues, { text: 'next', startTime: 105 }, 105, 2), 2);
  assert.equal(findActiveSubtitleCueIndex(cues, { text: '', startTime: 0 }, 108, 2), -1);
  assert.equal(findActiveSubtitleCueIndex(cues, { text: 'first', startTime: 0 }, 0, 2), 0);
});

test('findActiveSubtitleCueIndex falls back to the latest matching cue when the preferred index is stale', () => {
  const cues = [
    { startTime: 1, endTime: 2, text: 'same' },
    { startTime: 3, endTime: 4, text: 'same' },
  ];

  assert.equal(findActiveSubtitleCueIndex(cues, { text: 'same', startTime: null }, null, 5), 1);
});

test('subtitle sidebar mining context resolves selected row cue timing', () => {
  class FakeNode {
    parentElement: FakeElement | null = null;
  }
  class FakeElement extends FakeNode {
    dataset: Record<string, string> = {};

    closest(selector: string) {
      return selector === '.subtitle-sidebar-item' ? this : null;
    }
  }

  const row = new FakeElement();
  row.dataset.index = '1';
  const textNode = new FakeNode();
  textNode.parentElement = row;

  const restoreGlobals = installGlobals({
    Node: FakeNode,
    Element: FakeElement,
    window: { getSelection: () => ({ anchorNode: textNode, focusNode: null }) },
  });

  try {
    const state = createRendererState();
    state.subtitleSidebarModalOpen = true;
    state.subtitleSidebarCues = [
      { startTime: 1, endTime: 2, text: 'current line' },
      { startTime: 3, endTime: 5, text: 'sidebar previous line' },
    ];
    const modal = createSubtitleSidebarModal(
      {
        dom: {
          overlay: { classList: createClassList() },
          subtitleSidebarModal: {
            classList: createClassList(),
            setAttribute: () => {},
            style: { setProperty: () => {} },
            addEventListener: () => {},
          },
          subtitleSidebarContent: {
            classList: createClassList(),
            getBoundingClientRect: () => ({ width: 420 }),
            style: { setProperty: () => {} },
          },
          subtitleSidebarClose: { addEventListener: () => {} },
          subtitleSidebarStatus: { textContent: '' },
          subtitleSidebarList: createListStub(),
        },
        state,
      } as never,
      {
        modalStateReader: { isAnyModalOpen: () => false },
      },
    );

    const context = modal.getSubtitleSidebarMiningContext();

    assert.equal(context?.source, 'subtitle-sidebar');
    assert.equal(context?.text, 'sidebar previous line');
    assert.equal(context?.startTime, 3);
    assert.equal(context?.endTime, 5);
    assert.equal(typeof context?.capturedAtMs, 'number');
  } finally {
    restoreGlobals();
  }
});

test('applySidebarCssDeclarations clears declarations removed by config reload', () => {
  const removed: string[] = [];
  const style = {
    color: '',
    backgroundColor: '',
    setProperty(property: string, value: string) {
      (this as unknown as Record<string, string>)[property] = value;
    },
    removeProperty(property: string) {
      removed.push(property);
      delete (this as unknown as Record<string, string>)[property];
    },
  };
  const target = { style } as unknown as HTMLElement;

  applySidebarCssDeclarations(target, {
    color: '#cad3f5',
    'background-color': '#181926',
  });
  applySidebarCssDeclarations(target, {
    color: '#ffffff',
  });

  assert.equal(style.color, '#ffffff');
  assert.equal(style.backgroundColor, '');
  assert.deepEqual(removed, ['background-color']);

  applySidebarCssDeclarations(target, {
    color: '',
    'background-color': '',
  });

  assert.equal(style.color, '');
  assert.deepEqual(removed, ['background-color', 'background-color']);
});

const overlappingCues = [
  { startTime: 1, endTime: 3.4, text: 'first' },
  { startTime: 3, endTime: 4, text: 'second' },
];

test('subtitle sidebar opens from snapshot with the current cue active and config css applied', async () => {
  const h = createSidebarHarness({
    cues: overlappingCues,
    currentSubtitle: { text: 'second', startTime: 3.5, endTime: 4 },
    currentTimeSec: 3.5,
    config: { css: { 'font-size': '22px' } },
  });

  await h.modal.openSubtitleSidebarModal();

  assert.equal(h.state.subtitleSidebarModalOpen, true);
  assert.equal(h.modalClassList.contains('hidden'), false);
  assert.equal(h.state.subtitleSidebarActiveCueIndex, 1);
  assert.equal(h.cueList.children.length, 2);
  assert.equal(h.cueList.scrollTop, 0);
  assert.deepEqual(h.cueList.scrollToCalls, []);
  assert.equal(h.contentStyle['font-size'], '22px');
  assert.deepEqual(h.visibilityChanges, [true]);
  assert.deepEqual(h.modalNotifications, ['open:subtitle-sidebar']);
});

test('subtitle sidebar seeks to the selected cue start using the overlap-aware seek time', async () => {
  const h = createSidebarHarness({ cues: overlappingCues });
  await h.modal.openSubtitleSidebarModal();

  h.modal.seekToCue(overlappingCues[0]!);
  assert.deepEqual(h.mpvCommands.at(-1), ['seek', 1.08, 'absolute+exact']);

  h.modal.seekToCue(overlappingCues[1]!);
  assert.deepEqual(h.mpvCommands.at(-1), ['seek', 3.48, 'absolute+exact']);
});

test('subtitle sidebar close reports hidden visibility and notifies the modal close', async () => {
  const h = createSidebarHarness();
  await h.modal.openSubtitleSidebarModal();

  h.modal.closeSubtitleSidebarModal();

  assert.equal(h.state.subtitleSidebarModalOpen, false);
  assert.equal(h.modalClassList.contains('hidden'), true);
  assert.deepEqual(h.visibilityChanges, [true, false]);
  assert.deepEqual(h.modalNotifications, ['open:subtitle-sidebar', 'close:subtitle-sidebar']);
});

test('subtitle sidebar rows seek with Enter and leave Space to playback shortcuts', async () => {
  const h = createSidebarHarness({
    cues: [
      { startTime: 1, endTime: 2, text: 'first' },
      { startTime: 3, endTime: 4, text: 'second' },
    ],
    currentSubtitle: { text: 'second', startTime: 3, endTime: 4 },
    currentTimeSec: 3.1,
  });
  await h.modal.openSubtitleSidebarModal();

  const keydown = h.cueList.children[0]!.listeners.get('keydown')?.[0];
  assert.ok(keydown);

  h.mpvCommands.length = 0;
  let spacePrevented = false;
  keydown({
    key: ' ',
    preventDefault: () => {
      spacePrevented = true;
    },
  });

  assert.deepEqual(h.mpvCommands, []);
  assert.equal(spacePrevented, false);

  keydown({ key: 'Enter', preventDefault: () => {} });

  assert.deepEqual(h.mpvCommands.at(-1), ['seek', 1.08, 'absolute+exact']);
});

test('subtitle sidebar renders hour-long cue timestamps as HH:MM:SS', async () => {
  const h = createSidebarHarness({
    cues: [{ startTime: 3665, endTime: 3670, text: 'long cue' }],
    currentSubtitle: { text: 'long cue', startTime: 3665, endTime: 3670 },
    currentTimeSec: 3665,
  });

  await h.modal.openSubtitleSidebarModal();

  const firstRow = h.cueList.children[0]!;
  assert.equal(firstRow.attributes['aria-label'], 'Jump to subtitle at 01:01:05');
  assert.equal((firstRow.children[0] as { textContent: string }).textContent, '01:01:05');
});

test('subtitle sidebar does not open when the feature is disabled', async () => {
  const h = createSidebarHarness({ config: { enabled: false } });

  await h.modal.openSubtitleSidebarModal();

  assert.equal(h.state.subtitleSidebarModalOpen, false);
  assert.equal(h.modalClassList.contains('hidden'), true);
  assert.equal(h.cueList.children.length, 0);
  assert.equal(h.dom.subtitleSidebarStatus.textContent, 'Subtitle sidebar disabled in config.');
});

test('subtitle sidebar auto-open on startup only opens when enabled and configured', async () => {
  const h = createSidebarHarness({ config: { autoOpen: true } });

  await h.modal.autoOpenSubtitleSidebarOnStartup();

  assert.equal(h.state.subtitleSidebarModalOpen, true);
  assert.equal(h.modalClassList.contains('hidden'), false);
  assert.equal(h.cueList.children.length, 1);

  h.modal.closeSubtitleSidebarModal();
  h.setSnapshot({ config: { autoOpen: false } });

  await h.modal.autoOpenSubtitleSidebarOnStartup();

  assert.equal(h.state.subtitleSidebarModalOpen, false);
});

test('subtitle sidebar auto-open restores previously open sidebar after renderer replacement', async () => {
  const h = createSidebarHarness({
    modalOptions: { shouldRestoreOpenOnStartup: async () => true },
  });

  await h.modal.autoOpenSubtitleSidebarOnStartup();

  assert.equal(h.state.subtitleSidebarModalOpen, true);
  assert.equal(h.modalClassList.contains('hidden'), false);
  assert.equal(h.cueList.children.length, 1);
});

test('subtitle sidebar refresh closes and clears state when config becomes disabled', async () => {
  const h = createSidebarHarness({
    config: { layout: 'embedded', maxWidth: 360 },
    contentWidth: 360,
  });

  await h.modal.openSubtitleSidebarModal();
  assert.equal(h.state.subtitleSidebarModalOpen, true);
  assert.equal(h.bodyClassList.contains('subtitle-sidebar-embedded-open'), true);

  h.setSnapshot({
    cues: [],
    currentSubtitle: { text: '', startTime: null, endTime: null },
    currentTimeSec: null,
    config: { enabled: false },
  });
  await h.modal.refreshSubtitleSidebarSnapshot();

  assert.equal(h.state.subtitleSidebarModalOpen, false);
  assert.equal(h.state.subtitleSidebarCues.length, 0);
  assert.equal(h.state.subtitleSidebarActiveCueIndex, -1);
  assert.equal(h.modalClassList.contains('hidden'), true);
  assert.equal(h.bodyClassList.contains('subtitle-sidebar-embedded-open'), false);
});

test('subtitle sidebar keeps nearby repeated cue when subtitle update lacks timing', async () => {
  const h = createSidebarHarness({
    cues: [
      { startTime: 1, endTime: 2, text: 'same' },
      { startTime: 3, endTime: 4, text: 'other' },
      { startTime: 10, endTime: 11, text: 'same' },
    ],
    currentSubtitle: { text: 'same', startTime: 10, endTime: 11 },
    currentTimeSec: 10.1,
  });
  await h.modal.openSubtitleSidebarModal();
  h.cueList.scrollToCalls.length = 0;

  h.modal.handleSubtitleUpdated({ text: 'same', startTime: null, endTime: null, tokens: [] });

  assert.equal(h.state.subtitleSidebarActiveCueIndex, 2);
  assert.deepEqual(h.cueList.scrollToCalls, []);
});

test('subtitle sidebar does not regress to previous cue on text-only transition update', async () => {
  const h = createSidebarHarness({
    cues: [
      { startTime: 1, endTime: 2, text: 'first' },
      { startTime: 3, endTime: 4, text: 'second' },
      { startTime: 5, endTime: 6, text: 'third' },
    ],
    currentSubtitle: { text: 'third', startTime: 5, endTime: 6 },
    currentTimeSec: 5.1,
  });
  await h.modal.openSubtitleSidebarModal();
  h.cueList.scrollToCalls.length = 0;

  h.modal.handleSubtitleUpdated({ text: 'second', startTime: null, endTime: null, tokens: [] });

  assert.equal(h.state.subtitleSidebarActiveCueIndex, 2);
  assert.deepEqual(h.cueList.scrollToCalls, []);
});

test('subtitle sidebar jumps to first resolved active cue, then resumes smooth auto-follow', async () => {
  const h = createSidebarHarness({
    cues: Array.from({ length: 12 }, (_, index) => ({
      startTime: index * 2,
      endTime: index * 2 + 1.5,
      text: `line-${index}`,
    })),
    currentSubtitle: { text: '', startTime: null, endTime: null },
    currentTimeSec: null,
  });

  await h.modal.openSubtitleSidebarModal();
  assert.equal(h.state.subtitleSidebarActiveCueIndex, -1);
  h.cueList.scrollToCalls.length = 0;

  h.setSnapshot({
    currentSubtitle: { text: 'line-9', startTime: 18, endTime: 19.5 },
    currentTimeSec: 18.1,
  });
  await h.modal.refreshSubtitleSidebarSnapshot();

  assert.equal(h.state.subtitleSidebarActiveCueIndex, 9);
  assert.equal(h.cueList.scrollTop, 260);
  assert.deepEqual(h.cueList.scrollToCalls, []);

  h.setSnapshot({
    currentSubtitle: { text: 'line-10', startTime: 20, endTime: 21.5 },
    currentTimeSec: 20.1,
  });
  await h.modal.refreshSubtitleSidebarSnapshot();

  assert.equal(h.state.subtitleSidebarActiveCueIndex, 10);
  assert.deepEqual(h.cueList.scrollToCalls.at(-1), { top: 300, behavior: 'smooth' });
});

test('subtitle sidebar closes and resumes a hover pause', async () => {
  const h = createSidebarHarness({ config: { pauseVideoOnHover: true } });
  h.modal.wireDomEvents();
  await h.modal.openSubtitleSidebarModal();

  await h.contentEvents.dispatch('mouseenter');

  assert.equal(mpvPropertyState(h.mpvCommands).get('pause'), 'yes');

  h.modal.closeSubtitleSidebarModal();

  assert.equal(mpvPropertyState(h.mpvCommands).get('pause'), 'no');
  assert.equal(h.state.subtitleSidebarPausedByHover, false);
});

test('subtitle sidebar hover pause ignores playback-state IPC failures', async () => {
  const h = createSidebarHarness({
    config: { pauseVideoOnHover: true },
    electronAPI: {
      getPlaybackPaused: async () => {
        throw new Error('ipc failed');
      },
    },
  });
  h.modal.wireDomEvents();
  await h.modal.openSubtitleSidebarModal();

  await assert.doesNotReject(() => h.contentEvents.dispatch('mouseenter'));

  assert.equal(h.state.subtitleSidebarPausedByHover, false);
  assert.equal(mpvPropertyState(h.mpvCommands).has('pause'), false);
});

test('subtitle sidebar keeps hover pause while a Yomitan lookup popup remains open', async () => {
  const h = createSidebarHarness({ config: { pauseVideoOnHover: true } });
  h.state.autoPauseVideoOnYomitanPopup = true;
  h.modal.wireDomEvents();
  await h.modal.openSubtitleSidebarModal();
  h.mpvCommands.length = 0;

  await h.contentEvents.dispatch('mouseenter');

  assert.deepEqual(h.mpvCommands, [['set_property', 'pause', 'yes']]);

  await h.windowEvents.dispatch(YOMITAN_POPUP_SHOWN_EVENT);
  await h.contentEvents.dispatch('mouseleave');

  assert.deepEqual(h.mpvCommands, [['set_property', 'pause', 'yes']]);
  assert.equal(h.state.subtitleSidebarPausedByHover, true);

  await h.windowEvents.dispatch(YOMITAN_POPUP_HIDDEN_EVENT);

  assert.deepEqual(h.mpvCommands, [
    ['set_property', 'pause', 'yes'],
    ['set_property', 'pause', 'no'],
  ]);
  assert.equal(h.state.subtitleSidebarPausedByHover, false);
});

test('subtitle sidebar embedded layout reserves and releases mpv right margin', async () => {
  const h = createSidebarHarness({
    config: { layout: 'embedded', maxWidth: 360 },
    contentWidth: 360,
  });

  await h.modal.openSubtitleSidebarModal();

  assert.deepEqual(
    mpvPropertyState(h.mpvCommands),
    new Map<string, string | number>([
      ['video-margin-ratio-right', 0.3],
      ['osd-align-x', 'left'],
      ['osd-align-y', 'top'],
      ['user-data/osc/margins', '{"l":0,"r":0.3,"t":0,"b":0}'],
      ['video-pan-x', 0],
    ]),
  );
  assert.ok(h.bodyClassList.contains('subtitle-sidebar-embedded-open'));
  assert.equal(h.rootStyle.get('--subtitle-sidebar-reserved-width'), '360px');

  h.mpvCommands.length = 0;
  h.modal.closeSubtitleSidebarModal();

  assert.deepEqual(mpvPropertyState(h.mpvCommands), RELEASED_EMBEDDED_MARGIN);
  assert.equal(h.bodyClassList.contains('subtitle-sidebar-embedded-open'), false);
  assert.equal(h.rootStyle.get('--subtitle-sidebar-reserved-width'), '0px');
});

test('subtitle sidebar embedded layout measures reserved width after embedded classes apply', async () => {
  const h = createSidebarHarness({
    config: { layout: 'embedded' },
    contentWidth: (classList) =>
      classList.contains('subtitle-sidebar-content-embedded') ? 300 : 0,
  });

  await h.modal.openSubtitleSidebarModal();

  assert.ok(h.bodyClassList.contains('subtitle-sidebar-embedded-open'));
  assert.equal(h.rootStyle.get('--subtitle-sidebar-reserved-width'), '300px');
  assert.equal(mpvPropertyState(h.mpvCommands).get('video-margin-ratio-right'), 0.25);
});

test('subtitle sidebar resets embedded mpv margin on startup while closed', async () => {
  const h = createSidebarHarness({
    config: { layout: 'embedded', maxWidth: 360 },
    contentWidth: 360,
  });

  await h.modal.refreshSubtitleSidebarSnapshot();

  assert.deepEqual(mpvPropertyState(h.mpvCommands), RELEASED_EMBEDDED_MARGIN);
});

for (const layout of ['overlay', 'embedded'] as const) {
  test(`subtitle sidebar ${layout} layout restores macOS and Windows passthrough outside sidebar hover`, async () => {
    const h = createSidebarHarness({
      config: { layout, maxWidth: 360 },
      contentWidth: 360,
      toggleMouseIgnore: true,
    });
    h.modal.wireDomEvents();

    // Hover is tracked on the panel, not on the full-window modal backdrop.
    assert.equal(h.modalEvents.count('mouseenter'), 0);
    assert.equal(h.modalEvents.count('mouseleave'), 0);
    assert.equal(h.contentEvents.count('mouseenter'), 1);
    assert.equal(h.contentEvents.count('mouseleave'), 1);

    await h.modal.openSubtitleSidebarModal();
    assert.deepEqual(h.ignoreMouseCalls.at(-1), [true, { forward: true }]);

    await h.contentEvents.dispatch('mouseenter');
    assert.deepEqual(h.ignoreMouseCalls.at(-1), [false, undefined]);

    await h.contentEvents.dispatch('mouseleave');
    assert.deepEqual(h.ignoreMouseCalls.at(-1), [true, { forward: true }]);

    h.state.isOverSubtitle = true;
    await h.contentEvents.dispatch('mouseenter');
    await h.contentEvents.dispatch('mouseleave');
    assert.deepEqual(h.ignoreMouseCalls.at(-1), [false, undefined]);
  });
}

test('subtitle sidebar overlay layout only stays interactive while focus remains inside the sidebar panel', async () => {
  const h = createSidebarHarness({ toggleMouseIgnore: true });
  h.modal.wireDomEvents();

  await h.modal.openSubtitleSidebarModal();
  assert.deepEqual(h.ignoreMouseCalls.at(-1), [true, { forward: true }]);

  await h.contentEvents.dispatch('focusin');
  assert.deepEqual(h.ignoreMouseCalls.at(-1), [false, undefined]);

  await h.contentEvents.dispatch('focusout', { relatedTarget: null });
  assert.deepEqual(h.ignoreMouseCalls.at(-1), [true, { forward: true }]);
});

test('closing embedded subtitle sidebar recomputes passthrough from remaining subtitle hover state', async () => {
  const h = createSidebarHarness({
    config: { layout: 'embedded', maxWidth: 360 },
    contentWidth: 360,
    toggleMouseIgnore: true,
  });

  await h.modal.openSubtitleSidebarModal();
  h.state.isOverSubtitle = true;
  h.modal.closeSubtitleSidebarModal();
  assert.deepEqual(h.ignoreMouseCalls.at(-1), [false, undefined]);

  await h.modal.openSubtitleSidebarModal();
  h.state.isOverSubtitle = false;
  h.modal.closeSubtitleSidebarModal();
  assert.deepEqual(h.ignoreMouseCalls.at(-1), [true, { forward: true }]);
});
