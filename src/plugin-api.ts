/** Public plugin API v1. This file can also be copied into a plugin as subminer.d.ts. */
import type { BrowserWindow } from 'electron';

export type SessionActionId =
  | 'toggleStatsOverlay'
  | 'markWatched'
  | 'toggleVisibleOverlay'
  | 'copySubtitle'
  | 'copySubtitleMultiple'
  | 'updateLastCardFromClipboard'
  | 'triggerFieldGrouping'
  | 'triggerSubsync'
  | 'mineSentence'
  | 'mineSentenceMultiple'
  | 'toggleSecondarySub'
  | 'toggleSubtitleSidebar'
  | 'toggleNotificationHistory'
  | 'appendClipboardVideoToQueue'
  | 'markAudioCard'
  | 'openRuntimeOptions'
  | 'openSessionHelp'
  | 'openCharacterDictionaryManager'
  | 'openControllerSelect'
  | 'openControllerDebug'
  | 'openJimaku'
  | 'openTsukihime'
  | 'openSubtitleSelection'
  | 'openSubtitleGeneration'
  | 'openYoutubePicker'
  | 'openPlaylistBrowser'
  | 'replayCurrentSubtitle'
  | 'playNextSubtitle'
  | 'cycleRuntimeOption'
  | 'runPluginAction';

export interface SessionActionPayload {
  count?: number;
  runtimeOptionId?: string;
  direction?: 1 | -1;
  /** `<plugin id>/<action name>` for runPluginAction. */
  pluginActionId?: string;
}

export interface PluginMpvEventMap {
  'connection-change': { connected: boolean };
  'subtitle-change': { text: string; isOverlayVisible: boolean };
  'subtitle-timing': { text: string; start: number; end: number };
  'time-pos-change': { time: number };
  'duration-change': { duration: number };
  'pause-change': { paused: boolean };
  'fullscreen-change': { fullscreen: boolean };
  'secondary-subtitle-change': { text: string };
  'subtitle-track-change': { sid: number | null };
  'secondary-subtitle-track-change': { sid: number | null };
  'secondary-subtitle-delay-change': { delay: number };
  'subtitle-track-list-change': { trackList: unknown[] | null };
  'media-path-change': { path: string };
  'media-title-change': { title: string | null };
  'secondary-subtitle-visibility': { visible: boolean };
  'client-message': { args: string[] };
  'file-loaded': undefined;
  'end-file': {
    reason: 'eof' | 'stop' | 'quit' | 'error' | 'redirect' | 'unknown';
    playlistEntryId: number | null;
    error: string | null;
  };
}

export type PluginMpvEventName = keyof PluginMpvEventMap;

/** A cached playback snapshot. A new load gets a new id, including reopening the same path. */
export interface PluginMediaSnapshot {
  readonly playbackId: number;
  readonly path: string;
  readonly title: string | null;
}

/** Captured when mining starts. Null provenance means the source could not be established. */
export interface PluginMiningProvenance {
  readonly minedAt: number;
  readonly media: PluginMediaSnapshot | null;
  readonly subtitle: PluginSubtitleLine | null;
}

export interface PluginMinedNote {
  readonly originalNoteId: number;
  /** The note that kept the content after any field grouping. */
  readonly noteId: number;
  readonly provenance: PluginMiningProvenance | null;
}

/** Anki events plugins may observe. */
export interface PluginAnkiEventMap {
  /**
   * SubMiner finished writing a mined card, or polling saw it while auto-update is off.
   * Field grouping may have merged the note into another one by the time this fires.
   */
  'card-mined': { noteIds: number[]; notes: PluginMinedNote[] };
}
export type PluginAnkiEventName = keyof PluginAnkiEventMap;

export interface PluginAnkiInvokeOptions {
  /** Request timeout; AnkiConnect's default is 10 seconds. */
  timeoutMs?: number;
}

export interface PluginSubtitleLine {
  text: string;
  /** Seconds on the subtitle file's timeline; null when mpv did not report timing. */
  startTime: number | null;
  endTime: number | null;
}

export interface PluginSubtitleContext {
  current: PluginSubtitleLine | null;
  /** Oldest first. */
  before: PluginSubtitleLine[];
  after: PluginSubtitleLine[];
}

export interface PluginPanelContent {
  title: string;
  /** Rendered as a safe markdown subset; raw HTML is shown as text. */
  markdown: string;
}

export interface PluginPanelHandle {
  /** Aborts when the user closes the panel or another panel replaces it. */
  readonly signal: AbortSignal;
  readonly closed: boolean;
  update: (content: Partial<PluginPanelContent>) => void;
  close: () => void;
}

export interface PluginWindowOptions {
  /** HTML file to load, relative to the plugin directory. */
  html: string;
  /** Preload script, relative to the plugin directory. Runs with context isolation. */
  preload?: string;
  /** Window title; Hyprland matches it to float the window above the overlay. */
  title?: string;
  width?: number;
  height?: number;
}

export interface PluginExecOptions {
  /** Written to the child's stdin, which is then closed. */
  stdin?: string;
  cwd?: string;
  /** Merged over the parent environment. */
  env?: Record<string, string>;
  /** Kills the child and rejects after this many milliseconds. */
  timeoutMs?: number;
  /** Kills the child and rejects when aborted. */
  signal?: AbortSignal;
}

export interface PluginExecResult {
  /** Null when the child was terminated by a signal. */
  exitCode: number | null;
  stdout: string;
  stderr: string;
}

export type PluginNotifyVariant = 'info' | 'warning' | 'error';

export interface PluginLogger {
  debug: (message: string, ...meta: unknown[]) => void;
  info: (message: string, ...meta: unknown[]) => void;
  warn: (message: string, ...meta: unknown[]) => void;
  error: (message: string, ...meta: unknown[]) => void;
}

type Disposer = () => void;

/** The API object passed to a plugin's `activate(ctx)`. */
export interface PluginContext<TUnstable = unknown> {
  readonly apiVersion: 1;
  readonly id: string;
  /** Aborts as unloading starts, before deactivate runs. */
  readonly signal: AbortSignal;
  /** Registers cleanup; runs immediately if unloading has started. The disposer runs it once. */
  onDispose: (dispose: () => void) => Disposer;
  /** Absolute path of the plugin directory. */
  readonly dir: string;
  /** `plugins.<id>.settings` from config. */
  readonly settings: Readonly<Record<string, unknown>>;
  /**
   * `<config dir>/plugin-data/<id>/` for caches and state. Kept apart from `dir`, so replacing the
   * plugin folder keeps its data. Created before `activate`.
   */
  readonly dataDir: string;
  readonly log: PluginLogger;
  /** Repeating timer; cleared automatically when the plugin unloads. */
  every: (intervalMs: number, callback: () => unknown) => Disposer;
  /** One-shot timer; cleared automatically when the plugin unloads. */
  after: (delayMs: number, callback: () => unknown) => Disposer;
  readonly mpv: {
    readonly connected: boolean;
    on: <E extends PluginMpvEventName>(
      event: E,
      listener: (payload: PluginMpvEventMap[E]) => unknown,
    ) => Disposer;
    command: (command: (string | number)[]) => void;
    /** Resolves the mpv result; rejects on command error, disconnect, or a 4-second timeout. */
    request: (command: (string | number | boolean)[]) => Promise<unknown>;
    getProperty: (name: string) => Promise<unknown>;
    /** Connects to mpv, launching an idle mpv if none is running. Resolves false when that fails. */
    ensureRunning: () => Promise<boolean>;
  };
  readonly media: {
    current: () => PluginMediaSnapshot | null;
    /** Immediately delivers the current snapshot unless emitCurrent is false. */
    on: (
      event: 'change',
      listener: (media: PluginMediaSnapshot | null) => unknown,
      options?: { emitCurrent?: boolean },
    ) => Disposer;
  };
  readonly anki: {
    on: <E extends PluginAnkiEventName>(
      event: E,
      listener: (payload: PluginAnkiEventMap[E]) => unknown,
    ) => Disposer;
    /** Sends one AnkiConnect action to the configured `ankiConnect.url`. Rejects on an Anki error. */
    invoke: (
      action: string,
      params?: Record<string, unknown>,
      options?: PluginAnkiInvokeOptions,
    ) => Promise<unknown>;
  };
  readonly subtitles: {
    current: () => PluginSubtitleLine | null;
    context: (options?: { before?: number; after?: number }) => PluginSubtitleContext;
  };
  readonly actions: {
    /** Bind with `["__plugin-action:<plugin id>/<name>"]` in `keybindings`. */
    register: (name: string, handler: () => unknown) => Disposer;
    /** Runs a built-in session action, the same ones keybindings use (e.g. `openJimaku`). */
    run: (actionId: SessionActionId, payload?: SessionActionPayload) => Promise<void>;
  };
  notify: (message: string, options?: { variant?: PluginNotifyVariant }) => void;
  readonly ui: {
    /** Opens the shared plugin panel, replacing any panel that is already open. */
    panel: (content: Partial<PluginPanelContent> & { markdown: string }) => PluginPanelHandle;
    /**
     * Opens a window for this plugin, or focuses it if the same `html` is already open.
     * The window is destroyed when the plugin unloads.
     */
    openWindow: (options: PluginWindowOptions) => BrowserWindow;
  };
  /** Runs a binary without a shell, with the AppImage library path removed from its environment. */
  exec: (command: string, args: string[], options?: PluginExecOptions) => Promise<PluginExecResult>;
  /** Raw app internals. No stability guarantee across releases. */
  readonly unstable: TUnstable;
}

export interface PluginModule {
  activate?: (ctx: PluginContext) => unknown;
  deactivate?: () => unknown;
  default?: { activate?: (ctx: PluginContext) => unknown; deactivate?: () => unknown };
}
