import { ResolvedConfig } from '../../types';

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
  | 'subsync'
  | 'startupWarmups'
  | 'auto_start_overlay'
> = {
  subtitlePosition: { yPercent: 10 },
  keybindings: [],
  websocket: {
    enabled: 'auto',
    port: 6677,
  },
  annotationWebsocket: {
    enabled: true,
    port: 6678,
  },
  logging: {
    level: 'info',
  },
  texthooker: {
    launchAtStartup: true,
    openBrowser: true,
  },
  controller: {
    enabled: true,
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
      toggleLookup: 'buttonSouth',
      closeLookup: 'buttonEast',
      toggleKeyboardOnlyMode: 'buttonNorth',
      mineCard: 'buttonWest',
      quitMpv: 'select',
      previousAudio: 'none',
      nextAudio: 'rightShoulder',
      playCurrentAudio: 'leftShoulder',
      toggleMpvPause: 'leftStickPress',
      leftStickHorizontal: 'leftStickX',
      leftStickVertical: 'leftStickY',
      rightStickHorizontal: 'rightStickX',
      rightStickVertical: 'rightStickY',
    },
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
    openRuntimeOptions: 'CommandOrControl+Shift+O',
    openJimaku: 'Ctrl+Shift+J',
  },
  secondarySub: {
    secondarySubLanguages: [],
    autoLoadSecondarySub: false,
    defaultMode: 'hover',
  },
  subsync: {
    defaultMode: 'auto',
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
    jellyfinRemoteSession: true,
  },
  auto_start_overlay: false,
};
