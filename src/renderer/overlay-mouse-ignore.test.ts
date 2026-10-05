import assert from 'node:assert/strict';
import test from 'node:test';
import { syncOverlayMouseIgnoreState } from './overlay-mouse-ignore.js';
import { createRendererState, type RendererState } from './state.js';
import { YOMITAN_POPUP_VISIBLE_HOST_SELECTOR } from './yomitan-popup.js';

type IgnoreCall = { ignore: boolean; forward?: boolean };

type RunOptions = {
  // `toggle` matches Windows/macOS (renderer toggles ignore), `linux` matches the cursor-poll path.
  platform: 'toggle' | 'linux';
  state?: Partial<RendererState>;
  // Popup host marked visible in the DOM, independent of the cached `yomitanPopupVisible` flag.
  popupHostVisible?: boolean;
  sidebarSelecting?: boolean;
};

function makeState(overrides: Partial<RendererState> = {}): RendererState {
  return { ...createRendererState(), ...overrides };
}

function replaceGlobalProperty(key: 'window' | 'document', value: unknown): () => void {
  const original = Object.getOwnPropertyDescriptor(globalThis, key);
  Object.defineProperty(globalThis, key, { configurable: true, writable: true, value });
  return () => {
    if (original) {
      Object.defineProperty(globalThis, key, original);
      return;
    }
    delete (globalThis as Record<string, unknown>)[key];
  };
}

function run({ platform, state, popupHostVisible = false, sidebarSelecting = false }: RunOptions) {
  const classes = new Set<string>();
  const ignoreCalls: IgnoreCall[] = [];
  const interactiveHints: boolean[] = [];

  const restoreWindow = replaceGlobalProperty('window', {
    electronAPI: {
      setIgnoreMouseEvents: (ignore: boolean, options?: { forward?: boolean }) => {
        ignoreCalls.push({ ignore, forward: options?.forward });
      },
      reportOverlayInteractive: (interactive: boolean) => {
        interactiveHints.push(interactive);
      },
    },
  });
  const restoreDocument = replaceGlobalProperty('document', {
    querySelectorAll: (selector: string) =>
      popupHostVisible && selector === YOMITAN_POPUP_VISIBLE_HOST_SELECTOR ? [{}] : [],
  });

  try {
    syncOverlayMouseIgnoreState({
      dom: {
        overlay: {
          classList: {
            add: (token: string) => classes.add(token),
            remove: (token: string) => classes.delete(token),
          },
        },
        subtitleSidebarList: { dataset: sidebarSelecting ? { selecting: 'true' } : {} },
      },
      platform:
        platform === 'linux'
          ? { isLinuxPlatform: true, shouldToggleMouseIgnore: false }
          : { isLinuxPlatform: false, shouldToggleMouseIgnore: true },
      state: makeState(state),
    } as never);
  } finally {
    restoreDocument();
    restoreWindow();
  }

  return { interactive: classes.has('interactive'), ignoreCalls, interactiveHints };
}

const CLICK_THROUGH: IgnoreCall[] = [{ ignore: true, forward: true }];
const CAPTURE: IgnoreCall[] = [{ ignore: false, forward: undefined }];

const toggleCases: Array<{
  name: string;
  options: Omit<RunOptions, 'platform'>;
  capture: boolean;
}> = [
  { name: 'idle overlay starts click-through', options: {}, capture: false },
  {
    name: 'youtube picker keeps overlay interactive even when subtitle hover is inactive',
    options: { state: { youtubePickerModalOpen: true } },
    capture: true,
  },
  {
    name: 'visible yomitan popup host keeps overlay interactive even when cached popup state is false',
    options: { popupHostVisible: true, state: { yomitanPopupVisible: false } },
    capture: true,
  },
  {
    name: 'pointer over the yomitan popup keeps overlay interactive',
    options: { state: { isOverYomitanPopup: true } },
    capture: true,
  },
  {
    name: 'pointer over an overlay notification keeps overlay interactive',
    options: { state: { isOverOverlayNotification: true } },
    capture: true,
  },
  {
    name: 'pointer over notification history keeps overlay interactive',
    options: { state: { isOverNotificationHistory: true } },
    capture: true,
  },
  {
    name: 'subtitle sidebar drag selection keeps overlay interactive',
    options: { sidebarSelecting: true },
    capture: true,
  },
];

for (const c of toggleCases) {
  test(c.name, () => {
    const result = run({ platform: 'toggle', ...c.options });
    assert.equal(result.interactive, c.capture);
    assert.deepEqual(result.ignoreCalls, c.capture ? CAPTURE : CLICK_THROUGH);
    assert.deepEqual(result.interactiveHints, []);
  });
}

// Linux never toggles ignore itself; it only reports a whole-window hint for popups/modals.
const linuxCases: Array<{ name: string; state: Partial<RendererState>; interactive: boolean }> = [
  {
    name: 'Linux subtitle hover keeps root passive and does not report whole-window interactive hint',
    state: { isOverSubtitle: true },
    interactive: false,
  },
  {
    name: 'Linux modal state reports whole-window interactive hint',
    state: { runtimeOptionsModalOpen: true },
    interactive: true,
  },
  {
    name: 'Linux visible yomitan popup reports whole-window interactive hint',
    state: { yomitanPopupVisible: true },
    interactive: true,
  },
];

for (const c of linuxCases) {
  test(c.name, () => {
    const result = run({ platform: 'linux', state: c.state });
    assert.equal(result.interactive, c.interactive);
    assert.deepEqual(result.interactiveHints, [c.interactive]);
    assert.deepEqual(result.ignoreCalls, []);
  });
}
