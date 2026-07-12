import { ResolvedConfig } from '../../types/config';

export const CORE_DEFAULT_CONFIG: Pick<
  ResolvedConfig,
  | 'subtitlePosition'
  | 'keybindings'
  | 'websocket'
  | 'annotationWebsocket'
  | 'logging'
  | 'texthooker'
  | 'controller'
  | 'shortcuts'
  | 'secondarySub'
  | 'youtube'
  | 'subsync'
  | 'startupWarmups'
  | 'updates'
  | 'notifications'
  | 'auto_start_overlay'
> = {
  subtitlePosition: { yPercent: 10 },
  keybindings: [],
  websocket: {
    enabled: false,
    port: 6677,
  },
  annotationWebsocket: {
    enabled: false,
    port: 6678,
  },
  logging: {
    level: 'warn',
    rotation: 7,
    files: {
      app: true,
      launcher: true,
      mpv: false,
    },
  },
  texthooker: {
    launchAtStartup: false,
    openBrowser: false,
  },
  controller: {
    enabled: false,
    preferredGamepadId: '',
    preferredGamepadLabel: '',
    smoothScroll: true,
    scrollPixelsPerSecond: 900,
    horizontalJumpPixels: 160,
    stickDeadzone: 0.2,
    triggerInputMode: 'auto',
    triggerDeadzone: 0.5,
    repeatDelayMs: 320,
    repeatIntervalMs: 120,
    buttonIndices: {
      select: 6,
      buttonSouth: 0,
      buttonEast: 1,
      buttonWest: 2,
      buttonNorth: 3,
      leftShoulder: 4,
      rightShoulder: 5,
      leftStickPress: 9,
      rightStickPress: 10,
      leftTrigger: 6,
      rightTrigger: 7,
    },
    bindings: {
      toggleLookup: { kind: 'button', buttonIndex: 0 },
      closeLookup: { kind: 'button', buttonIndex: 1 },
      toggleKeyboardOnlyMode: { kind: 'button', buttonIndex: 3 },
      mineCard: { kind: 'button', buttonIndex: 2 },
      quitMpv: { kind: 'button', buttonIndex: 6 },
      previousAudio: { kind: 'none' },
      nextAudio: { kind: 'button', buttonIndex: 5 },
      playCurrentAudio: { kind: 'button', buttonIndex: 4 },
      toggleMpvPause: { kind: 'button', buttonIndex: 9 },
      leftStickHorizontal: { kind: 'axis', axisIndex: 0, dpadFallback: 'horizontal' },
      leftStickVertical: { kind: 'axis', axisIndex: 1, dpadFallback: 'vertical' },
      rightStickHorizontal: { kind: 'axis', axisIndex: 3, dpadFallback: 'none' },
      rightStickVertical: { kind: 'axis', axisIndex: 4, dpadFallback: 'none' },
    },
    profiles: {},
  },
  shortcuts: {
    toggleVisibleOverlayGlobal: 'Alt+Shift+O',
    copySubtitle: 'CommandOrControl+C',
    copySubtitleMultiple: 'CommandOrControl+Shift+C',
    updateLastCardFromClipboard: 'CommandOrControl+V',
    triggerFieldGrouping: 'CommandOrControl+G',
    triggerSubsync: 'Ctrl+Alt+S',
    mineSentence: 'CommandOrControl+S',
    mineSentenceMultiple: 'CommandOrControl+Shift+S',
    multiCopyTimeoutMs: 3000,
    toggleSecondarySub: 'CommandOrControl+Shift+V',
    markAudioCard: 'CommandOrControl+Shift+A',
    openCharacterDictionaryManager: 'CommandOrControl+D',
    openRuntimeOptions: 'CommandOrControl+Shift+O',
    openJimaku: 'Ctrl+Shift+J',
    openSessionHelp: 'CommandOrControl+Slash',
    openControllerSelect: 'Alt+C',
    openControllerDebug: 'Alt+Shift+C',
    toggleSubtitleSidebar: 'Backslash',
    toggleNotificationHistory: 'CommandOrControl+N',
    appendClipboardVideoToQueue: 'CommandOrControl+A',
  },
  secondarySub: {
    secondarySubLanguages: [],
    autoLoadSecondarySub: false,
    defaultMode: 'hover',
  },
  youtube: {
    primarySubLanguages: ['ja', 'jpn'],
    mediaCache: {
      mode: 'direct',
      maxHeight: 720,
    },
  },
  subsync: {
    alass_path: '',
    ffsubsync_path: '',
    ffmpeg_path: '',
    replace: true,
  },
  startupWarmups: {
    lowPowerMode: false,
    mecab: true,
    yomitanExtension: true,
    subtitleDictionaries: true,
    jellyfinRemoteSession: false,
  },
  updates: {
    enabled: true,
    checkIntervalHours: 24,
    notificationType: 'overlay',
    channel: 'stable',
  },
  notifications: {
    overlayPosition: 'top-right',
  },
  auto_start_overlay: true,
};
