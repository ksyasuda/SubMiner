import type {
  KikuFieldGroupingChoice,
  KikuFieldGroupingRequestData,
  KikuMergePreviewRequest,
  KikuMergePreviewResponse,
} from './anki';
import type { ResolvedConfig, ShortcutsConfig } from './config';
import type {
  CompiledSessionBinding,
  SessionActionId,
  SessionActionPayload,
  SessionBindingWarning,
} from './session-bindings';
import type {
  AnimetoshoApiResponse,
  AnimetoshoDownloadQuery,
  AnimetoshoDownloadResult,
  AnimetoshoEntry,
  AnimetoshoFilesQuery,
  AnimetoshoSearchQuery,
  AnimetoshoSubtitleFile,
  JimakuApiResponse,
  JimakuDownloadQuery,
  JimakuDownloadResult,
  JimakuEntry,
  JimakuFileEntry,
  JimakuFilesQuery,
  JimakuMediaInfo,
  JimakuSearchQuery,
  YoutubePickerOpenPayload,
  YoutubePickerResolveRequest,
  YoutubePickerResolveResult,
} from './integrations';
import type {
  PrimarySubMode,
  ResolvedSubtitleSidebarConfig,
  SecondarySubMode,
  SubtitleData,
  SubtitleMiningContext,
  SubtitlePosition,
  SubtitleSidebarSnapshot,
  SubtitleRendererStyleConfig,
  SubtitleStyleConfig,
} from './subtitle';
import type {
  RuntimeOptionApplyResult,
  RuntimeOptionId,
  RuntimeOptionState,
  RuntimeOptionValue,
} from './runtime-options';
import type {
  OverlayNotificationAction,
  OverlayNotificationEventPayload,
  OverlayNotificationPosition,
} from './notification';

export interface WindowGeometry {
  x: number;
  y: number;
  width: number;
  height: number;
}

export interface Keybinding {
  key: string;
  command: (string | number)[] | null;
}

export interface MpvClient {
  currentSubText: string;
  currentVideoPath: string;
  currentMediaTitle?: string | null;
  currentTimePos: number;
  currentSubStart: number;
  currentSubEnd: number;
  currentAudioStreamIndex: number | null;
  requestProperty?: (name: string) => Promise<unknown>;
  send(command: { command: unknown[]; request_id?: number }): boolean;
}

export interface SubsyncSourceTrack {
  id: number;
  label: string;
}

export interface SubsyncManualPayload {
  sourceTracks: SubsyncSourceTrack[];
  ffsubsyncAvailable: boolean;
}

export interface SubsyncManualRunRequest {
  engine: 'alass' | 'ffsubsync';
  sourceTrackId?: number | null;
}

export interface SubsyncResult {
  ok: boolean;
  message: string;
}

export interface PlaylistBrowserDirectoryItem {
  path: string;
  basename: string;
  episodeLabel?: string | null;
  isCurrentFile: boolean;
}

export interface PlaylistBrowserQueueItem {
  index: number;
  id: number | null;
  filename: string;
  title: string | null;
  displayLabel: string;
  current: boolean;
  playing: boolean;
  path: string | null;
}

export interface PlaylistBrowserSnapshot {
  directoryPath: string | null;
  directoryAvailable: boolean;
  directoryStatus: string;
  directoryItems: PlaylistBrowserDirectoryItem[];
  playlistItems: PlaylistBrowserQueueItem[];
  playingIndex: number | null;
  currentFilePath: string | null;
}

export interface PlaylistBrowserMutationResult {
  ok: boolean;
  message: string;
  snapshot?: PlaylistBrowserSnapshot;
}

export type ControllerButtonBinding =
  | 'none'
  | 'select'
  | 'buttonSouth'
  | 'buttonEast'
  | 'buttonNorth'
  | 'buttonWest'
  | 'leftShoulder'
  | 'rightShoulder'
  | 'leftStickPress'
  | 'rightStickPress'
  | 'leftTrigger'
  | 'rightTrigger';

export type ControllerAxisBinding = 'leftStickX' | 'leftStickY' | 'rightStickX' | 'rightStickY';
export type ControllerTriggerInputMode = 'auto' | 'digital' | 'analog';
export type ControllerAxisDirection = 'negative' | 'positive';
export type ControllerDpadFallback = 'none' | 'horizontal' | 'vertical';

export interface ControllerNoneBinding {
  kind: 'none';
}

export interface ControllerButtonInputBinding {
  kind: 'button';
  buttonIndex: number;
}

export interface ControllerAxisDirectionInputBinding {
  kind: 'axis';
  axisIndex: number;
  direction: ControllerAxisDirection;
}

export interface ControllerAxisInputBinding {
  kind: 'axis';
  axisIndex: number;
  dpadFallback?: ControllerDpadFallback;
}

export type ControllerDiscreteBindingConfig =
  | ControllerButtonBinding
  | ControllerNoneBinding
  | ControllerButtonInputBinding
  | ControllerAxisDirectionInputBinding;

export type ResolvedControllerDiscreteBinding =
  | ControllerNoneBinding
  | ControllerButtonInputBinding
  | ControllerAxisDirectionInputBinding;

export type ControllerAxisBindingConfig =
  | ControllerAxisBinding
  | ControllerNoneBinding
  | ControllerAxisInputBinding;

export type ResolvedControllerAxisBinding =
  | ControllerNoneBinding
  | {
      kind: 'axis';
      axisIndex: number;
      dpadFallback: ControllerDpadFallback;
    };

export interface ControllerBindingsConfig {
  toggleLookup?: ControllerDiscreteBindingConfig;
  closeLookup?: ControllerDiscreteBindingConfig;
  toggleKeyboardOnlyMode?: ControllerDiscreteBindingConfig;
  mineCard?: ControllerDiscreteBindingConfig;
  quitMpv?: ControllerDiscreteBindingConfig;
  previousAudio?: ControllerDiscreteBindingConfig;
  nextAudio?: ControllerDiscreteBindingConfig;
  playCurrentAudio?: ControllerDiscreteBindingConfig;
  toggleMpvPause?: ControllerDiscreteBindingConfig;
  leftStickHorizontal?: ControllerAxisBindingConfig;
  leftStickVertical?: ControllerAxisBindingConfig;
  rightStickHorizontal?: ControllerAxisBindingConfig;
  rightStickVertical?: ControllerAxisBindingConfig;
}

export interface ResolvedControllerBindingsConfig {
  toggleLookup?: ResolvedControllerDiscreteBinding;
  closeLookup?: ResolvedControllerDiscreteBinding;
  toggleKeyboardOnlyMode?: ResolvedControllerDiscreteBinding;
  mineCard?: ResolvedControllerDiscreteBinding;
  quitMpv?: ResolvedControllerDiscreteBinding;
  previousAudio?: ResolvedControllerDiscreteBinding;
  nextAudio?: ResolvedControllerDiscreteBinding;
  playCurrentAudio?: ResolvedControllerDiscreteBinding;
  toggleMpvPause?: ResolvedControllerDiscreteBinding;
  leftStickHorizontal?: ResolvedControllerAxisBinding;
  leftStickVertical?: ResolvedControllerAxisBinding;
  rightStickHorizontal?: ResolvedControllerAxisBinding;
  rightStickVertical?: ResolvedControllerAxisBinding;
}

export interface ControllerButtonIndicesConfig {
  select?: number;
  buttonSouth?: number;
  buttonEast?: number;
  buttonNorth?: number;
  buttonWest?: number;
  leftShoulder?: number;
  rightShoulder?: number;
  leftStickPress?: number;
  rightStickPress?: number;
  leftTrigger?: number;
  rightTrigger?: number;
}

export interface ControllerProfileConfig {
  label?: string;
  buttonIndices?: ControllerButtonIndicesConfig;
  bindings?: ControllerBindingsConfig;
}

export interface ResolvedControllerProfileConfig {
  label: string;
  buttonIndices: Required<ControllerButtonIndicesConfig>;
  bindings: Required<ResolvedControllerBindingsConfig>;
}

export interface ControllerConfig {
  enabled?: boolean;
  preferredGamepadId?: string;
  preferredGamepadLabel?: string;
  smoothScroll?: boolean;
  scrollPixelsPerSecond?: number;
  horizontalJumpPixels?: number;
  stickDeadzone?: number;
  triggerInputMode?: ControllerTriggerInputMode;
  triggerDeadzone?: number;
  repeatDelayMs?: number;
  repeatIntervalMs?: number;
  buttonIndices?: ControllerButtonIndicesConfig;
  bindings?: ControllerBindingsConfig;
  profiles?: Record<string, ControllerProfileConfig>;
}

export interface ControllerPreferenceUpdate {
  preferredGamepadId: string;
  preferredGamepadLabel: string;
}

export type ControllerConfigUpdate = ControllerConfig;

export interface ControllerDeviceInfo {
  id: string;
  index: number;
  mapping: string;
  connected: boolean;
}

export interface ControllerButtonSnapshot {
  value: number;
  pressed: boolean;
  touched?: boolean;
}

export interface ControllerRuntimeSnapshot {
  connectedGamepads: ControllerDeviceInfo[];
  activeGamepadId: string | null;
  rawAxes: number[];
  rawButtons: ControllerButtonSnapshot[];
}

export interface MpvSubtitleRenderMetrics {
  subPos: number;
  subFontSize: number;
  subScale: number;
  subMarginY: number;
  subMarginX: number;
  subFont: string;
  subSpacing: number;
  subBold: boolean;
  subItalic: boolean;
  subBorderSize: number;
  subShadowOffset: number;
  subAssOverride: string;
  subScaleByWindow: boolean;
  subUseMargins: boolean;
  osdHeight: number;
  osdDimensions: {
    w: number;
    h: number;
    ml: number;
    mr: number;
    mt: number;
    mb: number;
  } | null;
}

export type OverlayLayer = 'visible';

export interface OverlayContentRect {
  x: number;
  y: number;
  width: number;
  height: number;
}

export interface OverlayContentMeasurement {
  layer: OverlayLayer;
  measuredAtMs: number;
  viewport: {
    width: number;
    height: number;
  };
  contentRect: OverlayContentRect | null;
  interactiveRects?: OverlayContentRect[];
}

export interface MecabStatus {
  available: boolean;
  enabled: boolean;
  path: string | null;
}

export interface ClipboardAppendResult {
  ok: boolean;
  message: string;
}

export interface ConfigHotReloadPayload {
  keybindings: Keybinding[];
  sessionBindings: CompiledSessionBinding[];
  sessionBindingWarnings: SessionBindingWarning[];
  subtitleStyle: SubtitleRendererStyleConfig | null;
  subtitleSidebar: ResolvedSubtitleSidebarConfig;
  primarySubMode: PrimarySubMode;
  secondarySubMode: SecondarySubMode;
}

export interface SessionActionDispatchRequest {
  actionId: SessionActionId;
  payload?: SessionActionPayload;
}

export type ResolvedControllerConfig = ResolvedConfig['controller'];

export interface CharacterDictionaryCandidate {
  id: number;
  title: string;
  episodes: number | null;
}

export interface CharacterDictionarySelectionSnapshot {
  seriesKey: string;
  guessTitle: string | null;
  current: CharacterDictionaryCandidate | null;
  override: CharacterDictionaryCandidate | null;
  candidates: CharacterDictionaryCandidate[];
}

export interface CharacterDictionarySelectionResult {
  ok: boolean;
  seriesKey: string;
  selected: CharacterDictionaryCandidate;
  staleMediaIds: number[];
}

export interface CharacterDictionaryManagerEntry {
  mediaId: number;
  label: string;
  title: string;
  current: boolean;
}

export interface CharacterDictionaryManagerSnapshot {
  entries: CharacterDictionaryManagerEntry[];
}

export type CharacterDictionaryManagerMutationResult =
  | (CharacterDictionaryManagerSnapshot & { ok: true; rebuildRequired?: boolean })
  | { ok: false; message: string; entries: CharacterDictionaryManagerEntry[] };

export interface SessionNumericSelectionStartPayload {
  actionId: Extract<SessionActionId, 'copySubtitleMultiple' | 'mineSentenceMultiple'>;
  timeoutMs: number;
}

export interface ElectronAPI {
  getOverlayLayer: () => 'visible' | 'modal' | null;
  onSubtitle: (callback: (data: SubtitleData) => void) => void;
  onOverlayPointerRecoveryRequested: (callback: () => void) => void;
  onOverlayNotification: (callback: (payload: OverlayNotificationEventPayload) => void) => void;
  sendOverlayNotificationAction?: (
    notificationId: string,
    actionId: string,
    options?: Pick<OverlayNotificationAction, 'noteId'>,
  ) => void;
  onNotificationHistoryToggle: (callback: () => void) => void;
  onVisibility: (callback: (visible: boolean) => void) => void;
  onSubtitlePosition: (callback: (position: SubtitlePosition | null) => void) => void;
  getOverlayVisibility: () => Promise<boolean>;
  getCurrentSubtitle: () => Promise<SubtitleData>;
  getCurrentSubtitleRaw: () => Promise<string>;
  getCurrentSubtitleAss: () => Promise<string>;
  getSubtitleSidebarSnapshot: () => Promise<SubtitleSidebarSnapshot>;
  getSubtitleSidebarOpen: () => Promise<boolean>;
  getPlaybackPaused: () => Promise<boolean | null>;
  onSubtitleAss: (callback: (assText: string) => void) => void;
  setIgnoreMouseEvents: (ignore: boolean, options?: { forward?: boolean }) => void;
  reportOverlayInteractive: (interactive: boolean) => void;
  openYomitanSettings: () => void;
  recordYomitanLookup: (context?: SubtitleMiningContext | null) => void;
  getSubtitlePosition: () => Promise<SubtitlePosition | null>;
  saveSubtitlePosition: (position: SubtitlePosition) => void;
  getMecabStatus: () => Promise<MecabStatus>;
  setMecabEnabled: (enabled: boolean) => void;
  sendMpvCommand: (command: (string | number)[]) => void;
  getKeybindings: () => Promise<Keybinding[]>;
  getSessionBindings: () => Promise<CompiledSessionBinding[]>;
  getConfiguredShortcuts: () => Promise<Required<ShortcutsConfig>>;
  dispatchSessionAction: (
    actionId: SessionActionId,
    payload?: SessionActionPayload,
  ) => Promise<void>;
  getStatsToggleKey: () => Promise<string>;
  getMarkWatchedKey: () => Promise<string>;
  getOverlayNotificationPosition: () => Promise<OverlayNotificationPosition>;
  markActiveVideoWatched: () => Promise<boolean>;
  getControllerConfig: () => Promise<ResolvedControllerConfig>;
  saveControllerConfig: (update: ControllerConfigUpdate) => Promise<void>;
  saveControllerPreference: (update: ControllerPreferenceUpdate) => Promise<void>;
  getJimakuMediaInfo: () => Promise<JimakuMediaInfo>;
  jimakuSearchEntries: (query: JimakuSearchQuery) => Promise<JimakuApiResponse<JimakuEntry[]>>;
  jimakuListFiles: (query: JimakuFilesQuery) => Promise<JimakuApiResponse<JimakuFileEntry[]>>;
  jimakuDownloadFile: (query: JimakuDownloadQuery) => Promise<JimakuDownloadResult>;
  animetoshoSearchEntries: (
    query: AnimetoshoSearchQuery,
  ) => Promise<AnimetoshoApiResponse<AnimetoshoEntry[]>>;
  animetoshoListFiles: (
    query: AnimetoshoFilesQuery,
  ) => Promise<AnimetoshoApiResponse<AnimetoshoSubtitleFile[]>>;
  animetoshoDownloadFile: (query: AnimetoshoDownloadQuery) => Promise<AnimetoshoDownloadResult>;
  animetoshoGetSecondaryLanguages: () => Promise<string[]>;
  quitApp: () => void;
  toggleDevTools: () => void;
  toggleOverlay: () => void;
  toggleStatsOverlay: () => void;
  getAnkiConnectStatus: () => Promise<boolean>;
  setAnkiConnectEnabled: (enabled: boolean) => void;
  clearAnkiConnectHistory: () => void;
  onSecondarySub: (callback: (text: string) => void) => void;
  onSecondarySubMode: (callback: (mode: SecondarySubMode) => void) => void;
  getSecondarySubMode: () => Promise<SecondarySubMode>;
  getCurrentSecondarySub: () => Promise<string>;
  focusMainWindow: () => Promise<void>;
  activatePlaybackWindowForOverlayInteraction: () => Promise<boolean>;
  getSubtitleStyle: () => Promise<SubtitleRendererStyleConfig | null>;
  onSubsyncManualOpen: (callback: (payload: SubsyncManualPayload) => void) => void;
  runSubsyncManual: (request: SubsyncManualRunRequest) => Promise<SubsyncResult>;
  onKikuFieldGroupingRequest: (callback: (data: KikuFieldGroupingRequestData) => void) => void;
  onKikuFieldGroupingCancel: (callback: () => void) => void;
  kikuBuildMergePreview: (request: KikuMergePreviewRequest) => Promise<KikuMergePreviewResponse>;
  kikuFieldGroupingRespond: (choice: KikuFieldGroupingChoice) => void;
  getRuntimeOptions: () => Promise<RuntimeOptionState[]>;
  setRuntimeOptionValue: (
    id: RuntimeOptionId,
    value: RuntimeOptionValue,
  ) => Promise<RuntimeOptionApplyResult>;
  cycleRuntimeOption: (id: RuntimeOptionId, direction: 1 | -1) => Promise<RuntimeOptionApplyResult>;
  onRuntimeOptionsChanged: (callback: (options: RuntimeOptionState[]) => void) => void;
  onOpenRuntimeOptions: (callback: () => void) => void;
  onOpenSessionHelp: (callback: () => void) => void;
  onOpenControllerSelect: (callback: () => void) => void;
  onOpenControllerDebug: (callback: () => void) => void;
  onOpenJimaku: (callback: () => void) => void;
  onOpenAnimetosho: (callback: () => void) => void;
  onOpenYoutubeTrackPicker: (callback: (payload: YoutubePickerOpenPayload) => void) => void;
  onOpenPlaylistBrowser: (callback: () => void) => void;
  onOpenCharacterDictionaryManager: (callback: () => void) => void;
  onSubtitleSidebarToggle: (callback: () => void) => void;
  onPrimarySubtitleBarToggle: (callback: () => void) => void;
  onCancelYoutubeTrackPicker: (callback: () => void) => void;
  onSessionNumericSelectionStart: (
    callback: (payload: SessionNumericSelectionStartPayload) => void,
  ) => void;
  onKeyboardModeToggleRequested: (callback: () => void) => void;
  onLookupWindowToggleRequested: (callback: () => void) => void;
  appendClipboardVideoToQueue: () => Promise<ClipboardAppendResult>;
  getPlaylistBrowserSnapshot: () => Promise<PlaylistBrowserSnapshot>;
  appendPlaylistBrowserFile: (path: string) => Promise<PlaylistBrowserMutationResult>;
  playPlaylistBrowserIndex: (index: number) => Promise<PlaylistBrowserMutationResult>;
  removePlaylistBrowserIndex: (index: number) => Promise<PlaylistBrowserMutationResult>;
  movePlaylistBrowserIndex: (
    index: number,
    direction: 1 | -1,
  ) => Promise<PlaylistBrowserMutationResult>;
  youtubePickerResolve: (
    request: YoutubePickerResolveRequest,
  ) => Promise<YoutubePickerResolveResult>;
  getCharacterDictionarySelection: (
    searchTitle?: string,
  ) => Promise<CharacterDictionarySelectionSnapshot>;
  setCharacterDictionarySelection: (
    mediaId: number,
    replaceManagedMediaId?: number,
    mediaTitle?: string,
  ) => Promise<CharacterDictionarySelectionResult | CharacterDictionaryManagerMutationResult>;
  getCharacterDictionaryManagerSnapshot: () => Promise<CharacterDictionaryManagerSnapshot>;
  removeCharacterDictionaryManagedEntry: (
    mediaId: number,
  ) => Promise<CharacterDictionaryManagerMutationResult>;
  moveCharacterDictionaryManagedEntry: (
    mediaId: number,
    direction: 1 | -1,
  ) => Promise<CharacterDictionaryManagerMutationResult>;
  notifyOverlayModalClosed: (
    modal:
      | 'runtime-options'
      | 'subsync'
      | 'jimaku'
      | 'animetosho'
      | 'youtube-track-picker'
      | 'playlist-browser'
      | 'kiku'
      | 'controller-select'
      | 'controller-debug'
      | 'subtitle-sidebar'
      | 'session-help'
      | 'character-dictionary',
  ) => void;
  notifyOverlayModalOpened: (
    modal:
      | 'runtime-options'
      | 'subsync'
      | 'jimaku'
      | 'animetosho'
      | 'youtube-track-picker'
      | 'playlist-browser'
      | 'kiku'
      | 'controller-select'
      | 'controller-debug'
      | 'subtitle-sidebar'
      | 'session-help'
      | 'character-dictionary',
  ) => void;
  reportOverlayContentBounds: (measurement: OverlayContentMeasurement) => void;
  onConfigHotReload: (callback: (payload: ConfigHotReloadPayload) => void) => void;
}

declare global {
  interface Window {
    electronAPI: ElectronAPI;
  }
}
