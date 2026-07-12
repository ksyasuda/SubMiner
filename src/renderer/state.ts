import type {
  CompiledSessionBinding,
  PlaylistBrowserSnapshot,
  ControllerButtonSnapshot,
  ControllerDeviceInfo,
  ResolvedControllerConfig,
  AnimetoshoEntry,
  AnimetoshoSubtitleFile,
  JimakuEntry,
  JimakuFileEntry,
  KikuDuplicateCardInfo,
  KikuFieldGroupingChoice,
  RuntimeOptionId,
  RuntimeOptionState,
  RuntimeOptionValue,
  CharacterDictionarySelectionSnapshot,
  PrimarySubMode,
  SubtitlePosition,
  SubtitleSidebarSnapshotConfig,
  SubtitleCue,
  SubsyncSourceTrack,
  YoutubePickerOpenPayload,
} from '../types';

export type KikuModalStep = 'select' | 'preview';
export type KikuPreviewMode = 'compact' | 'full';

export type ChordAction =
  | { type: 'mpv'; command: string[] }
  | { type: 'electron'; action: () => void }
  | { type: 'noop' };

export type RendererState = {
  isOverSubtitle: boolean;
  isOverSubtitleSidebar: boolean;
  isOverOverlayNotification: boolean;
  isOverNotificationHistory: boolean;
  notificationHistoryOpen: boolean;
  isDragging: boolean;
  dragStartY: number;
  startYPercent: number;
  currentYPercent: number | null;
  persistedSubtitlePosition: SubtitlePosition;

  jimakuModalOpen: boolean;
  jimakuEntries: JimakuEntry[];
  jimakuFiles: JimakuFileEntry[];
  selectedEntryIndex: number;
  selectedFileIndex: number;
  currentEpisodeFilter: number | null;
  currentEntryId: number | null;

  animetoshoModalOpen: boolean;
  animetoshoActiveTab: 'en' | 'ja';
  animetoshoEntries: AnimetoshoEntry[];
  animetoshoFiles: AnimetoshoSubtitleFile[];
  selectedAnimetoshoEntryIndex: number;
  selectedAnimetoshoFileIndex: number;
  currentAnimetoshoEntryId: number | null;

  youtubePickerModalOpen: boolean;
  youtubePickerPayload: YoutubePickerOpenPayload | null;
  youtubePickerPrimaryTrackId: string | null;
  youtubePickerSecondaryTrackId: string | null;
  youtubePickerStatus: string;

  kikuModalOpen: boolean;
  kikuSelectedCard: 1 | 2;
  kikuOriginalData: KikuDuplicateCardInfo | null;
  kikuDuplicateData: KikuDuplicateCardInfo | null;
  kikuModalStep: KikuModalStep;
  kikuPreviewMode: KikuPreviewMode;
  kikuPendingChoice: KikuFieldGroupingChoice | null;
  kikuPreviewCompactData: Record<string, unknown> | null;
  kikuPreviewFullData: Record<string, unknown> | null;

  runtimeOptionsModalOpen: boolean;
  runtimeOptions: RuntimeOptionState[];
  runtimeOptionSelectedIndex: number;
  runtimeOptionDraftValues: Map<RuntimeOptionId, RuntimeOptionValue>;

  characterDictionaryModalOpen: boolean;
  characterDictionarySelection: CharacterDictionarySelectionSnapshot | null;
  characterDictionarySelectedIndex: number;
  characterDictionaryStatus: string;

  subsyncModalOpen: boolean;
  subsyncSourceTracks: SubsyncSourceTrack[];
  subsyncSubmitting: boolean;

  controllerSelectModalOpen: boolean;
  controllerDebugModalOpen: boolean;
  subtitleSidebarModalOpen: boolean;
  controllerDeviceSelectedIndex: number;
  controllerConfig: ResolvedControllerConfig | null;
  connectedGamepads: ControllerDeviceInfo[];
  activeGamepadId: string | null;
  controllerRawAxes: number[];
  controllerRawButtons: ControllerButtonSnapshot[];

  sessionHelpModalOpen: boolean;
  sessionHelpSelectedIndex: number;
  playlistBrowserModalOpen: boolean;
  playlistBrowserSnapshot: PlaylistBrowserSnapshot | null;
  playlistBrowserStatus: string;
  playlistBrowserActivePane: 'directory' | 'playlist';
  playlistBrowserSelectedDirectoryIndex: number;
  playlistBrowserSelectedPlaylistIndex: number;
  subtitleSidebarCues: SubtitleCue[];
  subtitleSidebarActiveCueIndex: number;
  subtitleSidebarToggleKey: string;
  subtitleSidebarPauseVideoOnHover: boolean;
  subtitleSidebarAutoScroll: boolean;
  subtitleSidebarConfig: SubtitleSidebarSnapshotConfig | null;
  subtitleSidebarManualScrollUntilMs: number;
  subtitleSidebarPausedByHover: boolean;

  knownWordColor: string;
  nPlusOneColor: string;
  nameMatchEnabled: boolean;
  nameMatchColor: string;
  jlptN1Color: string;
  jlptN2Color: string;
  jlptN3Color: string;
  jlptN4Color: string;
  jlptN5Color: string;
  preserveSubtitleLineBreaks: boolean;
  autoPauseVideoOnSubtitleHover: boolean;
  autoPauseVideoOnYomitanPopup: boolean;
  primaryVisibleOnYomitanPopup: boolean;
  frequencyDictionaryEnabled: boolean;
  frequencyDictionaryTopX: number;
  frequencyDictionaryMode: 'single' | 'banded';
  frequencyDictionarySingleColor: string;
  frequencyDictionaryBand1Color: string;
  frequencyDictionaryBand2Color: string;
  frequencyDictionaryBand3Color: string;
  frequencyDictionaryBand4Color: string;
  frequencyDictionaryBand5Color: string;

  sessionBindings: CompiledSessionBinding[];
  sessionBindingMap: Map<string, CompiledSessionBinding>;
  sessionActionTimeoutMs: number;
  statsToggleKey: string;
  markWatchedKey: string;
  chordPending: boolean;
  chordTimeout: ReturnType<typeof setTimeout> | null;
  keyboardDrivenModeEnabled: boolean;
  keyboardSelectionVisible: boolean;
  keyboardSelectedWordIndex: number | null;
  yomitanPopupVisible: boolean;
  isOverYomitanPopup: boolean;
  primarySubtitleMode: PrimarySubMode;
};

export function createRendererState(): RendererState {
  return {
    isOverSubtitle: false,
    isOverSubtitleSidebar: false,
    isOverOverlayNotification: false,
    isOverNotificationHistory: false,
    notificationHistoryOpen: false,
    isDragging: false,
    dragStartY: 0,
    startYPercent: 0,
    currentYPercent: null,
    persistedSubtitlePosition: { yPercent: 10 },

    jimakuModalOpen: false,
    jimakuEntries: [],
    jimakuFiles: [],
    selectedEntryIndex: 0,
    selectedFileIndex: 0,
    currentEpisodeFilter: null,
    currentEntryId: null,

    animetoshoModalOpen: false,
    animetoshoActiveTab: 'en',
    animetoshoEntries: [],
    animetoshoFiles: [],
    selectedAnimetoshoEntryIndex: 0,
    selectedAnimetoshoFileIndex: 0,
    currentAnimetoshoEntryId: null,

    youtubePickerModalOpen: false,
    youtubePickerPayload: null,
    youtubePickerPrimaryTrackId: null,
    youtubePickerSecondaryTrackId: null,
    youtubePickerStatus: '',

    kikuModalOpen: false,
    kikuSelectedCard: 1,
    kikuOriginalData: null,
    kikuDuplicateData: null,
    kikuModalStep: 'select',
    kikuPreviewMode: 'compact',
    kikuPendingChoice: null,
    kikuPreviewCompactData: null,
    kikuPreviewFullData: null,

    runtimeOptionsModalOpen: false,
    runtimeOptions: [],
    runtimeOptionSelectedIndex: 0,
    runtimeOptionDraftValues: new Map(),

    characterDictionaryModalOpen: false,
    characterDictionarySelection: null,
    characterDictionarySelectedIndex: 0,
    characterDictionaryStatus: '',

    subsyncModalOpen: false,
    subsyncSourceTracks: [],
    subsyncSubmitting: false,

    controllerSelectModalOpen: false,
    controllerDebugModalOpen: false,
    subtitleSidebarModalOpen: false,
    controllerDeviceSelectedIndex: 0,
    controllerConfig: null,
    connectedGamepads: [],
    activeGamepadId: null,
    controllerRawAxes: [],
    controllerRawButtons: [],

    sessionHelpModalOpen: false,
    sessionHelpSelectedIndex: 0,
    playlistBrowserModalOpen: false,
    playlistBrowserSnapshot: null,
    playlistBrowserStatus: '',
    playlistBrowserActivePane: 'playlist',
    playlistBrowserSelectedDirectoryIndex: 0,
    playlistBrowserSelectedPlaylistIndex: 0,
    subtitleSidebarCues: [],
    subtitleSidebarActiveCueIndex: -1,
    subtitleSidebarToggleKey: 'Backslash',
    subtitleSidebarPauseVideoOnHover: false,
    subtitleSidebarAutoScroll: true,
    subtitleSidebarConfig: null,
    subtitleSidebarManualScrollUntilMs: 0,
    subtitleSidebarPausedByHover: false,

    knownWordColor: '#a6da95',
    nPlusOneColor: '#c6a0f6',
    nameMatchEnabled: false,
    nameMatchColor: '#f5bde6',
    jlptN1Color: '#ed8796',
    jlptN2Color: '#f5a97f',
    jlptN3Color: '#f9e2af',
    jlptN4Color: '#a6e3a1',
    jlptN5Color: '#8aadf4',
    preserveSubtitleLineBreaks: false,
    autoPauseVideoOnSubtitleHover: false,
    autoPauseVideoOnYomitanPopup: false,
    primaryVisibleOnYomitanPopup: true,
    frequencyDictionaryEnabled: false,
    frequencyDictionaryTopX: 1000,
    frequencyDictionaryMode: 'single',
    frequencyDictionarySingleColor: '#f5a97f',
    frequencyDictionaryBand1Color: '#ed8796',
    frequencyDictionaryBand2Color: '#f5a97f',
    frequencyDictionaryBand3Color: '#f9e2af',
    frequencyDictionaryBand4Color: '#8bd5ca',
    frequencyDictionaryBand5Color: '#8aadf4',

    sessionBindings: [],
    sessionBindingMap: new Map(),
    sessionActionTimeoutMs: 3000,
    statsToggleKey: 'Backquote',
    markWatchedKey: 'KeyW',
    chordPending: false,
    chordTimeout: null,
    keyboardDrivenModeEnabled: false,
    keyboardSelectionVisible: false,
    keyboardSelectedWordIndex: null,
    yomitanPopupVisible: false,
    isOverYomitanPopup: false,
    primarySubtitleMode: 'visible',
  };
}
