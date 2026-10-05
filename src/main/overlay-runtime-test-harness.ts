import type { WindowGeometry } from '../types';
import {
  OVERLAY_WINDOW_CONTENT_READY_FLAG,
  OVERLAY_WINDOW_DOCUMENT_LOADED_FLAG,
} from '../core/services/overlay-window-flags';
import {
  createOverlayModalRuntimeService,
  type OverlayModalRuntimeOptions,
  type OverlayWindowResolver,
} from './overlay-runtime';

export const MODAL_URL = 'file:///overlay/index.html?layer=modal';
export const MODAL_GEOMETRY: WindowGeometry = { x: 0, y: 0, width: 400, height: 300 };

export interface FakeOverlayWindowInit {
  // Prefix for show/hide entries pushed to `log`, e.g. `modal:show`.
  name?: string;
  log?: string[];
  visible?: boolean;
  focused?: boolean;
  webContentsFocused?: boolean;
  ignoreMouseEvents?: boolean;
  loading?: boolean;
  url?: string;
  contentReady?: boolean;
  documentLoaded?: boolean;
}

// A still-loading window whose document has not been committed yet.
export const LOADING_BLANK_WINDOW: FakeOverlayWindowInit = {
  loading: true,
  url: '',
  contentReady: false,
};

/**
 * BrowserWindow stand-in covering the surface overlay-runtime touches. Defaults to a hidden,
 * fully loaded, content-ready modal window.
 */
export class FakeOverlayWindow {
  destroyed = false;
  visible = false;
  focused = false;
  webContentsFocused = false;
  ignoreMouseEvents = false;
  forwardedIgnoreMouseEvents = false;
  loading = false;
  url = MODAL_URL;
  alwaysOnTopCalls: string[] = [];
  showCount = 0;
  hideCount = 0;
  sent: unknown[][] = [];
  [OVERLAY_WINDOW_CONTENT_READY_FLAG] = true;
  [OVERLAY_WINDOW_DOCUMENT_LOADED_FLAG] = true;

  private readonly name: string;
  private readonly log: string[] | undefined;
  private loadListeners: Array<() => void> = [];
  private stopLoadingListeners: Array<() => void> = [];
  private readyToShowListeners: Array<() => void> = [];

  readonly webContents = {
    isLoading: (): boolean => this.loading,
    getURL: (): string => this.url,
    send: (channel: string, ...payload: unknown[]): void => {
      this.sent.push([channel, ...payload]);
    },
    isFocused: (): boolean => this.webContentsFocused,
    focus: (): void => {
      this.webContentsFocused = true;
    },
    once: (event: 'did-finish-load' | 'did-stop-loading', listener: () => void): void => {
      (event === 'did-stop-loading' ? this.stopLoadingListeners : this.loadListeners).push(
        listener,
      );
    },
  };

  constructor(init: FakeOverlayWindowInit = {}) {
    const { name = 'window', log, contentReady = true, documentLoaded = true, ...fields } = init;
    this.name = name;
    this.log = log;
    Object.assign(this, fields);
    this[OVERLAY_WINDOW_CONTENT_READY_FLAG] = contentReady;
    this[OVERLAY_WINDOW_DOCUMENT_LOADED_FLAG] = documentLoaded;
  }

  isDestroyed(): boolean {
    return this.destroyed;
  }

  isVisible(): boolean {
    return this.visible;
  }

  isFocused(): boolean {
    return this.focused;
  }

  setIgnoreMouseEvents(ignore: boolean, options?: { forward?: boolean }): void {
    this.ignoreMouseEvents = ignore;
    this.forwardedIgnoreMouseEvents = options?.forward === true;
  }

  setAlwaysOnTop(flag: boolean, level?: string, relativeLevel?: number): void {
    this.alwaysOnTopCalls.push(`top:${flag}:${level ?? ''}:${relativeLevel ?? ''}`);
  }

  moveTop(): void {}

  show(): void {
    this.log?.push(`${this.name}:show`);
    this.visible = true;
    this.showCount += 1;
  }

  showInactive(): void {
    this.log?.push(`${this.name}:show-inactive`);
    this.visible = true;
    this.showCount += 1;
  }

  hide(): void {
    this.log?.push(`${this.name}:hide`);
    this.visible = false;
    this.hideCount += 1;
  }

  destroy(): void {
    this.destroyed = true;
    this.visible = false;
  }

  focus(): void {
    this.focused = true;
  }

  once(_event: 'ready-to-show', listener: () => void): void {
    this.readyToShowListeners.push(listener);
  }

  // Commits the modal document like Electron does: isLoading() stays true inside
  // did-finish-load and only settles before did-stop-loading.
  finishLoad(): void {
    this.url = MODAL_URL;
    this[OVERLAY_WINDOW_DOCUMENT_LOADED_FLAG] = true;
    for (const listener of this.loadListeners.splice(0)) listener();
    this.loading = false;
    for (const listener of this.stopLoadingListeners.splice(0)) listener();
  }

  // Marks renderer content ready and fires the queued ready-to-show listeners.
  emitReadyToShow(): void {
    this[OVERLAY_WINDOW_CONTENT_READY_FLAG] = true;
    for (const listener of this.readyToShowListeners.splice(0)) listener();
  }
}

export interface OverlayRuntimeHarnessInit {
  main?: FakeOverlayWindow | null;
  modal?: FakeOverlayWindow | null;
  // Called by the runtime's createModalWindow; its result becomes the current modal window.
  createModal?: () => FakeOverlayWindow | null;
  // Defaults to linux so behavior does not depend on the host running the tests.
  platform?: NodeJS.Platform;
  options?: Omit<OverlayModalRuntimeOptions, 'platform'>;
  setModalWindowBounds?: OverlayWindowResolver['setModalWindowBounds'];
}

/** Builds the modal runtime over fake windows and records resolver traffic. */
export function createHarness(init: OverlayRuntimeHarnessInit = {}) {
  const harness = {
    modal: init.modal ?? null,
    createCalls: 0,
    boundsCalls: [] as WindowGeometry[],
  };
  const runtime = createOverlayModalRuntimeService(
    {
      getMainWindow: () => (init.main ?? null) as never,
      getModalWindow: () => harness.modal as never,
      createModalWindow: () => {
        harness.createCalls += 1;
        harness.modal = init.createModal?.() ?? null;
        return harness.modal as never;
      },
      getModalGeometry: () => MODAL_GEOMETRY,
      setModalWindowBounds: (geometry) => {
        harness.boundsCalls.push(geometry);
        return init.setModalWindowBounds?.(geometry);
      },
    },
    { ...init.options, platform: init.platform ?? 'linux' },
  );
  return Object.assign(harness, { runtime });
}
