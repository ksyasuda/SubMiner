import { ResolvedConfig } from '../../types';

export const CORE_DEFAULT_CONFIG: Pick<
  ResolvedConfig,
  | 'subtitlePosition'
  | 'keybindings'
  | 'websocket'
  | 'logging'
  | 'texthooker'
  | 'shortcuts'
  | 'secondarySub'
  | 'subsync'
  | 'auto_start_overlay'
  | 'bind_visible_overlay_to_mpv_sub_visibility'
  | 'invisibleOverlay'
> = {
  subtitlePosition: { yPercent: 10 },
  keybindings: [],
  websocket: {
    enabled: 'auto',
    port: 6677,
  },
  logging: {
    level: 'info',
  },
  texthooker: {
    openBrowser: true,
  },
  shortcuts: {
    toggleVisibleOverlayGlobal: 'Alt+Shift+O',
    toggleInvisibleOverlayGlobal: 'Alt+Shift+I',
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
  },
  auto_start_overlay: false,
  bind_visible_overlay_to_mpv_sub_visibility: true,
  invisibleOverlay: {
    startupVisibility: 'platform-default',
  },
};
