export interface CliArgs {
  background: boolean;
  start: boolean;
  launchMpv: boolean;
  launchMpvTargets: string[];
  youtubePlay?: string;
  youtubeMode?: 'download' | 'generate';
  stop: boolean;
  toggle: boolean;
  toggleVisibleOverlay: boolean;
  settings: boolean;
  setup: boolean;
  show: boolean;
  hide: boolean;
  showVisibleOverlay: boolean;
  hideVisibleOverlay: boolean;
  copySubtitle: boolean;
  copySubtitleMultiple: boolean;
  mineSentence: boolean;
  mineSentenceMultiple: boolean;
  updateLastCardFromClipboard: boolean;
  refreshKnownWords: boolean;
  toggleSecondarySub: boolean;
  triggerFieldGrouping: boolean;
  triggerSubsync: boolean;
  markAudioCard: boolean;
  toggleStatsOverlay: boolean;
  toggleSubtitleSidebar: boolean;
  openRuntimeOptions: boolean;
  openSessionHelp: boolean;
  openControllerSelect: boolean;
  openControllerDebug: boolean;
  openJimaku: boolean;
  openYoutubePicker: boolean;
  openPlaylistBrowser: boolean;
  replayCurrentSubtitle: boolean;
  playNextSubtitle: boolean;
  shiftSubDelayPrevLine: boolean;
  shiftSubDelayNextLine: boolean;
  cycleRuntimeOptionId?: string;
  cycleRuntimeOptionDirection?: 1 | -1;
  copySubtitleCount?: number;
  mineSentenceCount?: number;
  anilistStatus: boolean;
  anilistLogout: boolean;
  anilistSetup: boolean;
  anilistRetryQueue: boolean;
  dictionary: boolean;
  dictionaryTarget?: string;
  stats: boolean;
  statsBackground?: boolean;
  statsStop?: boolean;
  statsCleanup?: boolean;
  statsCleanupVocab?: boolean;
  statsCleanupLifetime?: boolean;
  statsResponsePath?: string;
  jellyfin: boolean;
  jellyfinLogin: boolean;
  jellyfinLogout: boolean;
  jellyfinLibraries: boolean;
  jellyfinItems: boolean;
  jellyfinSubtitles: boolean;
  jellyfinSubtitleUrlsOnly: boolean;
  jellyfinPlay: boolean;
  jellyfinRemoteAnnounce: boolean;
  jellyfinPreviewAuth: boolean;
  texthooker: boolean;
  help: boolean;
  autoStartOverlay: boolean;
  generateConfig: boolean;
  configPath?: string;
  backupOverwrite: boolean;
  socketPath?: string;
  backend?: string;
  texthookerPort?: number;
  jellyfinServer?: string;
  jellyfinUsername?: string;
  jellyfinPassword?: string;
  jellyfinLibraryId?: string;
  jellyfinItemId?: string;
  jellyfinSearch?: string;
  jellyfinLimit?: number;
  jellyfinRecursive?: boolean;
  jellyfinIncludeItemTypes?: string;
  jellyfinAudioStreamIndex?: number;
  jellyfinSubtitleStreamIndex?: number;
  jellyfinResponsePath?: string;
  debug: boolean;
  logLevel?: 'debug' | 'info' | 'warn' | 'error';
}

export type CliCommandSource = 'initial' | 'second-instance';

export function parseArgs(argv: string[]): CliArgs {
  const args: CliArgs = {
    background: false,
    start: false,
    launchMpv: false,
    launchMpvTargets: [],
    youtubePlay: undefined,
    youtubeMode: undefined,
    stop: false,
    toggle: false,
    toggleVisibleOverlay: false,
    settings: false,
    setup: false,
    show: false,
    hide: false,
    showVisibleOverlay: false,
    hideVisibleOverlay: false,
    copySubtitle: false,
    copySubtitleMultiple: false,
    mineSentence: false,
    mineSentenceMultiple: false,
    updateLastCardFromClipboard: false,
    refreshKnownWords: false,
    toggleSecondarySub: false,
    triggerFieldGrouping: false,
    triggerSubsync: false,
    markAudioCard: false,
    toggleStatsOverlay: false,
    toggleSubtitleSidebar: false,
    openRuntimeOptions: false,
    openSessionHelp: false,
    openControllerSelect: false,
    openControllerDebug: false,
    openJimaku: false,
    openYoutubePicker: false,
    openPlaylistBrowser: false,
    replayCurrentSubtitle: false,
    playNextSubtitle: false,
    shiftSubDelayPrevLine: false,
    shiftSubDelayNextLine: false,
    anilistStatus: false,
    anilistLogout: false,
    anilistSetup: false,
    anilistRetryQueue: false,
    dictionary: false,
    stats: false,
    statsBackground: false,
    statsStop: false,
    statsCleanup: false,
    statsCleanupVocab: false,
    statsCleanupLifetime: false,
    jellyfin: false,
    jellyfinLogin: false,
    jellyfinLogout: false,
    jellyfinLibraries: false,
    jellyfinItems: false,
    jellyfinSubtitles: false,
    jellyfinSubtitleUrlsOnly: false,
    jellyfinPlay: false,
    jellyfinRemoteAnnounce: false,
    jellyfinPreviewAuth: false,
    texthooker: false,
    help: false,
    autoStartOverlay: false,
    generateConfig: false,
    backupOverwrite: false,
    debug: false,
  };

  const readValue = (value?: string): string | undefined => {
    if (!value) return undefined;
    if (value.startsWith('--')) return undefined;
    return value;
  };

  const parseCycleRuntimeOption = (
    value: string | undefined,
  ): { id: string; direction: 1 | -1 } | null => {
    if (!value) return null;
    const separatorIndex = value.lastIndexOf(':');
    if (separatorIndex <= 0 || separatorIndex === value.length - 1) return null;
    const id = value.slice(0, separatorIndex).trim();
    const rawDirection = value
      .slice(separatorIndex + 1)
      .trim()
      .toLowerCase();
    if (!id) return null;
    if (rawDirection === 'next' || rawDirection === '1') {
      return { id, direction: 1 };
    }
    if (rawDirection === 'prev' || rawDirection === '-1') {
      return { id, direction: -1 };
    }
    return null;
  };

  for (let i = 0; i < argv.length; i += 1) {
    const arg = argv[i];
    if (!arg || !arg.startsWith('--')) continue;

    if (arg === '--background') args.background = true;
    else if (arg === '--start') args.start = true;
    else if (arg.startsWith('--youtube-play=')) {
      const value = arg.split('=', 2)[1];
      if (value) args.youtubePlay = value;
    } else if (arg === '--youtube-play') {
      const value = readValue(argv[i + 1]);
      if (value) args.youtubePlay = value;
    } else if (arg.startsWith('--youtube-mode=')) {
      const value = arg.split('=', 2)[1];
      if (value === 'download' || value === 'generate') args.youtubeMode = value;
    } else if (arg === '--youtube-mode') {
      const value = readValue(argv[i + 1]);
      if (value === 'download' || value === 'generate') args.youtubeMode = value;
    } else if (arg === '--launch-mpv') {
      args.launchMpv = true;
      args.launchMpvTargets = argv.slice(i + 1).filter((value) => value && !value.startsWith('--'));
      break;
    } else if (arg === '--stop') args.stop = true;
    else if (arg === '--toggle') args.toggle = true;
    else if (arg === '--toggle-visible-overlay') args.toggleVisibleOverlay = true;
    else if (arg === '--settings' || arg === '--yomitan') args.settings = true;
    else if (arg === '--setup') args.setup = true;
    else if (arg === '--show') args.show = true;
    else if (arg === '--hide') args.hide = true;
    else if (arg === '--show-visible-overlay') args.showVisibleOverlay = true;
    else if (arg === '--hide-visible-overlay') args.hideVisibleOverlay = true;
    else if (arg === '--copy-subtitle') args.copySubtitle = true;
    else if (arg === '--copy-subtitle-multiple') args.copySubtitleMultiple = true;
    else if (arg === '--mine-sentence') args.mineSentence = true;
    else if (arg === '--mine-sentence-multiple') args.mineSentenceMultiple = true;
    else if (arg === '--update-last-card-from-clipboard') args.updateLastCardFromClipboard = true;
    else if (arg === '--refresh-known-words') args.refreshKnownWords = true;
    else if (arg === '--toggle-secondary-sub') args.toggleSecondarySub = true;
    else if (arg === '--trigger-field-grouping') args.triggerFieldGrouping = true;
    else if (arg === '--trigger-subsync') args.triggerSubsync = true;
    else if (arg === '--mark-audio-card') args.markAudioCard = true;
    else if (arg === '--toggle-stats-overlay') args.toggleStatsOverlay = true;
    else if (arg === '--toggle-subtitle-sidebar') args.toggleSubtitleSidebar = true;
    else if (arg === '--open-runtime-options') args.openRuntimeOptions = true;
    else if (arg === '--open-session-help') args.openSessionHelp = true;
    else if (arg === '--open-controller-select') args.openControllerSelect = true;
    else if (arg === '--open-controller-debug') args.openControllerDebug = true;
    else if (arg === '--open-jimaku') args.openJimaku = true;
    else if (arg === '--open-youtube-picker') args.openYoutubePicker = true;
    else if (arg === '--open-playlist-browser') args.openPlaylistBrowser = true;
    else if (arg === '--replay-current-subtitle') args.replayCurrentSubtitle = true;
    else if (arg === '--play-next-subtitle') args.playNextSubtitle = true;
    else if (arg === '--shift-sub-delay-prev-line') args.shiftSubDelayPrevLine = true;
    else if (arg === '--shift-sub-delay-next-line') args.shiftSubDelayNextLine = true;
    else if (arg.startsWith('--cycle-runtime-option=')) {
      const parsed = parseCycleRuntimeOption(arg.split('=', 2)[1]);
      if (parsed) {
        args.cycleRuntimeOptionId = parsed.id;
        args.cycleRuntimeOptionDirection = parsed.direction;
      }
    } else if (arg === '--cycle-runtime-option') {
      const parsed = parseCycleRuntimeOption(readValue(argv[i + 1]));
      if (parsed) {
        args.cycleRuntimeOptionId = parsed.id;
        args.cycleRuntimeOptionDirection = parsed.direction;
      }
    } else if (arg.startsWith('--copy-subtitle-count=')) {
      const value = Number(arg.split('=', 2)[1]);
      if (Number.isInteger(value) && value > 0) args.copySubtitleCount = value;
    } else if (arg === '--copy-subtitle-count') {
      const value = Number(readValue(argv[i + 1]));
      if (Number.isInteger(value) && value > 0) args.copySubtitleCount = value;
    } else if (arg.startsWith('--mine-sentence-count=')) {
      const value = Number(arg.split('=', 2)[1]);
      if (Number.isInteger(value) && value > 0) args.mineSentenceCount = value;
    } else if (arg === '--mine-sentence-count') {
      const value = Number(readValue(argv[i + 1]));
      if (Number.isInteger(value) && value > 0) args.mineSentenceCount = value;
    } else if (arg === '--anilist-status') args.anilistStatus = true;
    else if (arg === '--anilist-logout') args.anilistLogout = true;
    else if (arg === '--anilist-setup') args.anilistSetup = true;
    else if (arg === '--anilist-retry-queue') args.anilistRetryQueue = true;
    else if (arg === '--dictionary') args.dictionary = true;
    else if (arg.startsWith('--dictionary-target=')) {
      const value = arg.split('=', 2)[1];
      if (value) args.dictionaryTarget = value;
    } else if (arg === '--dictionary-target') {
      const value = readValue(argv[i + 1]);
      if (value) args.dictionaryTarget = value;
    } else if (arg === '--stats') args.stats = true;
    else if (arg === '--stats-background') {
      args.stats = true;
      args.statsBackground = true;
    } else if (arg === '--stats-stop') {
      args.stats = true;
      args.statsStop = true;
    } else if (arg === '--stats-cleanup') args.statsCleanup = true;
    else if (arg === '--stats-cleanup-vocab') args.statsCleanupVocab = true;
    else if (arg === '--stats-cleanup-lifetime') args.statsCleanupLifetime = true;
    else if (arg.startsWith('--stats-response-path=')) {
      const value = arg.split('=', 2)[1];
      if (value) args.statsResponsePath = value;
    } else if (arg === '--stats-response-path') {
      const value = readValue(argv[i + 1]);
      if (value) args.statsResponsePath = value;
    } else if (arg === '--jellyfin') args.jellyfin = true;
    else if (arg === '--jellyfin-login') args.jellyfinLogin = true;
    else if (arg === '--jellyfin-logout') args.jellyfinLogout = true;
    else if (arg === '--jellyfin-libraries') args.jellyfinLibraries = true;
    else if (arg === '--jellyfin-items') args.jellyfinItems = true;
    else if (arg === '--jellyfin-subtitles') args.jellyfinSubtitles = true;
    else if (arg === '--jellyfin-subtitle-urls') {
      args.jellyfinSubtitles = true;
      args.jellyfinSubtitleUrlsOnly = true;
    } else if (arg === '--jellyfin-play') args.jellyfinPlay = true;
    else if (arg === '--jellyfin-remote-announce') args.jellyfinRemoteAnnounce = true;
    else if (arg === '--jellyfin-preview-auth') args.jellyfinPreviewAuth = true;
    else if (arg === '--texthooker') args.texthooker = true;
    else if (arg === '--auto-start-overlay') args.autoStartOverlay = true;
    else if (arg === '--generate-config') args.generateConfig = true;
    else if (arg === '--backup-overwrite') args.backupOverwrite = true;
    else if (arg === '--help') args.help = true;
    else if (arg === '--debug') args.debug = true;
    else if (arg.startsWith('--log-level=')) {
      const value = arg.split('=', 2)[1]?.toLowerCase();
      if (value === 'debug' || value === 'info' || value === 'warn' || value === 'error') {
        args.logLevel = value;
      }
    } else if (arg === '--log-level') {
      const value = readValue(argv[i + 1])?.toLowerCase();
      if (value === 'debug' || value === 'info' || value === 'warn' || value === 'error') {
        args.logLevel = value;
      }
    } else if (arg.startsWith('--config-path=')) {
      const value = arg.split('=', 2)[1];
      if (value) args.configPath = value;
    } else if (arg === '--config-path') {
      const value = readValue(argv[i + 1]);
      if (value) args.configPath = value;
    } else if (arg.startsWith('--socket=')) {
      const value = arg.split('=', 2)[1];
      if (value) args.socketPath = value;
    } else if (arg === '--socket') {
      const value = readValue(argv[i + 1]);
      if (value) args.socketPath = value;
    } else if (arg.startsWith('--backend=')) {
      const value = arg.split('=', 2)[1];
      if (value) args.backend = value;
    } else if (arg === '--backend') {
      const value = readValue(argv[i + 1]);
      if (value) args.backend = value;
    } else if (arg.startsWith('--port=')) {
      const value = Number(arg.split('=', 2)[1]);
      if (!Number.isNaN(value)) args.texthookerPort = value;
    } else if (arg === '--port') {
      const value = Number(readValue(argv[i + 1]));
      if (!Number.isNaN(value)) args.texthookerPort = value;
    } else if (arg.startsWith('--jellyfin-server=')) {
      const value = arg.split('=', 2)[1];
      if (value) args.jellyfinServer = value;
    } else if (arg === '--jellyfin-server') {
      const value = readValue(argv[i + 1]);
      if (value) args.jellyfinServer = value;
    } else if (arg.startsWith('--jellyfin-username=')) {
      const value = arg.split('=', 2)[1];
      if (value) args.jellyfinUsername = value;
    } else if (arg === '--jellyfin-username') {
      const value = readValue(argv[i + 1]);
      if (value) args.jellyfinUsername = value;
    } else if (arg.startsWith('--jellyfin-password=')) {
      const value = arg.split('=', 2)[1];
      if (value) args.jellyfinPassword = value;
    } else if (arg === '--jellyfin-password') {
      const value = readValue(argv[i + 1]);
      if (value) args.jellyfinPassword = value;
    } else if (arg.startsWith('--jellyfin-library-id=')) {
      const value = arg.split('=', 2)[1];
      if (value) args.jellyfinLibraryId = value;
    } else if (arg === '--jellyfin-library-id') {
      const value = readValue(argv[i + 1]);
      if (value) args.jellyfinLibraryId = value;
    } else if (arg.startsWith('--jellyfin-item-id=')) {
      const value = arg.split('=', 2)[1];
      if (value) args.jellyfinItemId = value;
    } else if (arg === '--jellyfin-item-id') {
      const value = readValue(argv[i + 1]);
      if (value) args.jellyfinItemId = value;
    } else if (arg.startsWith('--jellyfin-search=')) {
      const value = arg.split('=', 2)[1];
      if (value) args.jellyfinSearch = value;
    } else if (arg === '--jellyfin-search') {
      const value = readValue(argv[i + 1]);
      if (value) args.jellyfinSearch = value;
    } else if (arg.startsWith('--jellyfin-limit=')) {
      const value = Number(arg.split('=', 2)[1]);
      if (Number.isFinite(value) && value > 0) args.jellyfinLimit = Math.floor(value);
    } else if (arg === '--jellyfin-limit') {
      const value = Number(readValue(argv[i + 1]));
      if (Number.isFinite(value) && value > 0) args.jellyfinLimit = Math.floor(value);
    } else if (arg.startsWith('--jellyfin-recursive=')) {
      const value = arg.split('=', 2)[1]?.trim().toLowerCase();
      if (value === 'true' || value === '1' || value === 'yes') args.jellyfinRecursive = true;
      if (value === 'false' || value === '0' || value === 'no') args.jellyfinRecursive = false;
    } else if (arg === '--jellyfin-recursive') {
      const value = readValue(argv[i + 1])
        ?.trim()
        .toLowerCase();
      if (value === 'false' || value === '0' || value === 'no') {
        args.jellyfinRecursive = false;
      } else if (value === 'true' || value === '1' || value === 'yes') {
        args.jellyfinRecursive = true;
      }
    } else if (arg === '--jellyfin-non-recursive') {
      args.jellyfinRecursive = false;
    } else if (arg.startsWith('--jellyfin-include-item-types=')) {
      const value = arg.split('=', 2)[1];
      if (value) args.jellyfinIncludeItemTypes = value;
    } else if (arg === '--jellyfin-include-item-types') {
      const value = readValue(argv[i + 1]);
      if (value) args.jellyfinIncludeItemTypes = value;
    } else if (arg.startsWith('--jellyfin-audio-stream-index=')) {
      const value = Number(arg.split('=', 2)[1]);
      if (Number.isInteger(value) && value >= 0) args.jellyfinAudioStreamIndex = value;
    } else if (arg === '--jellyfin-audio-stream-index') {
      const value = Number(readValue(argv[i + 1]));
      if (Number.isInteger(value) && value >= 0) args.jellyfinAudioStreamIndex = value;
    } else if (arg.startsWith('--jellyfin-subtitle-stream-index=')) {
      const value = Number(arg.split('=', 2)[1]);
      if (Number.isInteger(value) && value >= 0) args.jellyfinSubtitleStreamIndex = value;
    } else if (arg === '--jellyfin-subtitle-stream-index') {
      const value = Number(readValue(argv[i + 1]));
      if (Number.isInteger(value) && value >= 0) args.jellyfinSubtitleStreamIndex = value;
    } else if (arg.startsWith('--jellyfin-response-path=')) {
      const value = arg.split('=', 2)[1];
      if (value) args.jellyfinResponsePath = value;
    } else if (arg === '--jellyfin-response-path') {
      const value = readValue(argv[i + 1]);
      if (value) args.jellyfinResponsePath = value;
    }
  }

  return args;
}

export function hasExplicitCommand(args: CliArgs): boolean {
  return (
    args.background ||
    args.start ||
    Boolean(args.youtubePlay) ||
    args.launchMpv ||
    args.stop ||
    args.toggle ||
    args.toggleVisibleOverlay ||
    args.settings ||
    args.setup ||
    args.show ||
    args.hide ||
    args.showVisibleOverlay ||
    args.hideVisibleOverlay ||
    args.copySubtitle ||
    args.copySubtitleMultiple ||
    args.mineSentence ||
    args.mineSentenceMultiple ||
    args.updateLastCardFromClipboard ||
    args.refreshKnownWords ||
    args.toggleSecondarySub ||
    args.triggerFieldGrouping ||
    args.triggerSubsync ||
    args.markAudioCard ||
    args.toggleStatsOverlay ||
    args.toggleSubtitleSidebar ||
    args.openRuntimeOptions ||
    args.openSessionHelp ||
    args.openControllerSelect ||
    args.openControllerDebug ||
    args.openJimaku ||
    args.openYoutubePicker ||
    args.openPlaylistBrowser ||
    args.replayCurrentSubtitle ||
    args.playNextSubtitle ||
    args.shiftSubDelayPrevLine ||
    args.shiftSubDelayNextLine ||
    args.cycleRuntimeOptionId !== undefined ||
    args.copySubtitleCount !== undefined ||
    args.mineSentenceCount !== undefined ||
    args.anilistStatus ||
    args.anilistLogout ||
    args.anilistSetup ||
    args.anilistRetryQueue ||
    args.dictionary ||
    args.stats ||
    args.jellyfin ||
    args.jellyfinLogin ||
    args.jellyfinLogout ||
    args.jellyfinLibraries ||
    args.jellyfinItems ||
    args.jellyfinSubtitles ||
    args.jellyfinPlay ||
    args.jellyfinRemoteAnnounce ||
    args.jellyfinPreviewAuth ||
    args.texthooker ||
    args.generateConfig ||
    args.help
  );
}

export function isHeadlessInitialCommand(args: CliArgs): boolean {
  return args.refreshKnownWords;
}

export function isStandaloneTexthookerCommand(args: CliArgs): boolean {
  return (
    args.texthooker &&
    !args.background &&
    !args.start &&
    !Boolean(args.youtubePlay) &&
    !args.launchMpv &&
    !args.stop &&
    !args.toggle &&
    !args.toggleVisibleOverlay &&
    !args.settings &&
    !args.setup &&
    !args.show &&
    !args.hide &&
    !args.showVisibleOverlay &&
    !args.hideVisibleOverlay &&
    !args.copySubtitle &&
    !args.copySubtitleMultiple &&
    !args.mineSentence &&
    !args.mineSentenceMultiple &&
    !args.updateLastCardFromClipboard &&
    !args.refreshKnownWords &&
    !args.toggleSecondarySub &&
    !args.triggerFieldGrouping &&
    !args.triggerSubsync &&
    !args.markAudioCard &&
    !args.toggleStatsOverlay &&
    !args.toggleSubtitleSidebar &&
    !args.openRuntimeOptions &&
    !args.openSessionHelp &&
    !args.openControllerSelect &&
    !args.openControllerDebug &&
    !args.openJimaku &&
    !args.openYoutubePicker &&
    !args.openPlaylistBrowser &&
    !args.replayCurrentSubtitle &&
    !args.playNextSubtitle &&
    !args.shiftSubDelayPrevLine &&
    !args.shiftSubDelayNextLine &&
    args.cycleRuntimeOptionId === undefined &&
    args.copySubtitleCount === undefined &&
    args.mineSentenceCount === undefined &&
    !args.anilistStatus &&
    !args.anilistLogout &&
    !args.anilistSetup &&
    !args.anilistRetryQueue &&
    !args.dictionary &&
    !args.stats &&
    !args.jellyfin &&
    !args.jellyfinLogin &&
    !args.jellyfinLogout &&
    !args.jellyfinLibraries &&
    !args.jellyfinItems &&
    !args.jellyfinSubtitles &&
    !args.jellyfinPlay &&
    !args.jellyfinRemoteAnnounce &&
    !args.jellyfinPreviewAuth &&
    !args.help &&
    !args.autoStartOverlay &&
    !args.generateConfig
  );
}

export function shouldStartApp(args: CliArgs): boolean {
  if (args.stop && !args.start) return false;
  if (
    args.background ||
    args.start ||
    Boolean(args.youtubePlay) ||
    args.launchMpv ||
    args.toggle ||
    args.toggleVisibleOverlay ||
    args.settings ||
    args.setup ||
    args.copySubtitle ||
    args.copySubtitleMultiple ||
    args.mineSentence ||
    args.mineSentenceMultiple ||
    args.updateLastCardFromClipboard ||
    args.refreshKnownWords ||
    args.toggleSecondarySub ||
    args.triggerFieldGrouping ||
    args.triggerSubsync ||
    args.markAudioCard ||
    args.toggleStatsOverlay ||
    args.toggleSubtitleSidebar ||
    args.openRuntimeOptions ||
    args.openSessionHelp ||
    args.openControllerSelect ||
    args.openControllerDebug ||
    args.openJimaku ||
    args.openYoutubePicker ||
    args.openPlaylistBrowser ||
    args.replayCurrentSubtitle ||
    args.playNextSubtitle ||
    args.shiftSubDelayPrevLine ||
    args.shiftSubDelayNextLine ||
    args.cycleRuntimeOptionId !== undefined ||
    args.copySubtitleCount !== undefined ||
    args.mineSentenceCount !== undefined ||
    args.dictionary ||
    args.stats ||
    args.jellyfin ||
    args.jellyfinPlay ||
    args.texthooker
  ) {
    if (args.launchMpv) {
      return false;
    }
    return true;
  }
  return false;
}

export function shouldRunSettingsOnlyStartup(args: CliArgs): boolean {
  return (
    args.settings &&
    !args.background &&
    !args.start &&
    !args.stop &&
    !args.toggle &&
    !args.toggleVisibleOverlay &&
    !args.show &&
    !args.hide &&
    !args.setup &&
    !args.showVisibleOverlay &&
    !args.hideVisibleOverlay &&
    !args.copySubtitle &&
    !args.copySubtitleMultiple &&
    !args.mineSentence &&
    !args.mineSentenceMultiple &&
    !args.updateLastCardFromClipboard &&
    !args.refreshKnownWords &&
    !args.toggleSecondarySub &&
    !args.triggerFieldGrouping &&
    !args.triggerSubsync &&
    !args.markAudioCard &&
    !args.toggleStatsOverlay &&
    !args.toggleSubtitleSidebar &&
    !args.openRuntimeOptions &&
    !args.openSessionHelp &&
    !args.openControllerSelect &&
    !args.openControllerDebug &&
    !args.openJimaku &&
    !args.openYoutubePicker &&
    !args.openPlaylistBrowser &&
    !args.replayCurrentSubtitle &&
    !args.playNextSubtitle &&
    !args.shiftSubDelayPrevLine &&
    !args.shiftSubDelayNextLine &&
    args.cycleRuntimeOptionId === undefined &&
    args.copySubtitleCount === undefined &&
    args.mineSentenceCount === undefined &&
    !args.anilistStatus &&
    !args.anilistLogout &&
    !args.anilistSetup &&
    !args.anilistRetryQueue &&
    !args.dictionary &&
    !args.stats &&
    !args.jellyfin &&
    !args.jellyfinLogin &&
    !args.jellyfinLogout &&
    !args.jellyfinLibraries &&
    !args.jellyfinItems &&
    !args.jellyfinSubtitles &&
    !args.jellyfinPlay &&
    !args.youtubePlay &&
    !args.jellyfinRemoteAnnounce &&
    !args.jellyfinPreviewAuth &&
    !args.texthooker &&
    !args.help &&
    !args.autoStartOverlay &&
    !args.generateConfig &&
    !args.backupOverwrite &&
    !args.debug
  );
}

export function commandNeedsOverlayRuntime(args: CliArgs): boolean {
  return (
    args.toggle ||
    args.toggleVisibleOverlay ||
    args.show ||
    args.hide ||
    args.showVisibleOverlay ||
    args.hideVisibleOverlay ||
    args.copySubtitle ||
    args.copySubtitleMultiple ||
    args.mineSentence ||
    args.mineSentenceMultiple ||
    args.updateLastCardFromClipboard ||
    args.toggleSecondarySub ||
    args.toggleStatsOverlay ||
    args.toggleSubtitleSidebar ||
    args.triggerFieldGrouping ||
    args.triggerSubsync ||
    args.markAudioCard ||
    args.openRuntimeOptions ||
    args.openSessionHelp ||
    args.openControllerSelect ||
    args.openControllerDebug ||
    args.openJimaku ||
    args.openYoutubePicker ||
    args.openPlaylistBrowser ||
    args.replayCurrentSubtitle ||
    args.playNextSubtitle ||
    args.shiftSubDelayPrevLine ||
    args.shiftSubDelayNextLine ||
    args.cycleRuntimeOptionId !== undefined ||
    args.copySubtitleCount !== undefined ||
    args.mineSentenceCount !== undefined
  );
}

export function commandNeedsOverlayStartupPrereqs(args: CliArgs): boolean {
  return commandNeedsOverlayRuntime(args) || Boolean(args.youtubePlay);
}
