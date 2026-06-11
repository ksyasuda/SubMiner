import { ResolvedConfig } from '../../types/config';
import { getDefaultMpvSocketPath } from '../../shared/mpv-socket-path';

export const INTEGRATIONS_DEFAULT_CONFIG: Pick<
  ResolvedConfig,
  | 'ankiConnect'
  | 'jimaku'
  | 'anilist'
  | 'mpv'
  | 'yomitan'
  | 'jellyfin'
  | 'discordPresence'
  | 'ai'
  | 'youtubeSubgen'
> = {
  ankiConnect: {
    enabled: true,
    url: 'http://127.0.0.1:8765',
    pollingRate: 3000,
    proxy: {
      enabled: true,
      host: '127.0.0.1',
      port: 8766,
      upstreamUrl: 'http://127.0.0.1:8765',
    },
    tags: ['SubMiner'],
    deck: '',
    fields: {
      word: 'Expression',
      audio: 'ExpressionAudio',
      image: 'Picture',
      sentence: 'Sentence',
      miscInfo: 'MiscInfo',
      translation: 'SelectionText',
    },
    ai: {
      enabled: false,
      model: '',
      systemPrompt: '',
    },
    media: {
      generateAudio: true,
      generateImage: true,
      imageType: 'static',
      imageFormat: 'jpg',
      imageQuality: 92,
      imageMaxWidth: 0,
      imageMaxHeight: 0,
      animatedFps: 10,
      animatedMaxWidth: 640,
      animatedMaxHeight: 0,
      animatedCrf: 35,
      syncAnimatedImageToWordAudio: true,
      audioPadding: 0,
      fallbackDuration: 3.0,
      maxMediaDuration: 30,
    },
    knownWords: {
      highlightEnabled: false,
      refreshMinutes: 1440,
      addMinedWordsImmediately: true,
      matchMode: 'headword',
      decks: {},
    },
    behavior: {
      overwriteAudio: true,
      overwriteImage: true,
      mediaInsertMode: 'append',
      highlightWord: true,
      notificationType: 'overlay',
      autoUpdateNewCards: true,
    },
    nPlusOne: {
      enabled: false,
      minSentenceWords: 3,
    },
    metadata: {
      pattern: '[SubMiner] %f (%t)',
    },
    isLapis: {
      enabled: false,
      sentenceCardModel: 'Lapis',
    },
    isKiku: {
      enabled: false,
      fieldGrouping: 'disabled',
      deleteDuplicateInAuto: true,
    },
  },
  jimaku: {
    apiBaseUrl: 'https://jimaku.cc',
    apiKey: '',
    apiKeyCommand: '',
    languagePreference: 'ja',
    maxEntryResults: 10,
  },
  mpv: {
    executablePath: '',
    launchMode: 'normal',
    profile: '',
    socketPath: getDefaultMpvSocketPath(),
    backend: 'auto',
    autoStartSubMiner: true,
    pauseUntilOverlayReady: true,
    subminerBinaryPath: '',
    aniskipEnabled: true,
    aniskipButtonKey: 'TAB',
  },
  anilist: {
    enabled: false,
    accessToken: '',
    characterDictionary: {
      refreshTtlHours: 168,
      maxLoaded: 3,
      evictionPolicy: 'delete',
      profileScope: 'all',
      collapsibleSections: {
        description: false,
        characterInformation: false,
        voicedBy: false,
      },
    },
  },
  yomitan: {
    externalProfilePath: '',
  },
  jellyfin: {
    enabled: false,
    serverUrl: '',
    recentServers: [],
    username: '',
    defaultLibraryId: '',
    remoteControlEnabled: true,
    remoteControlAutoConnect: true,
    autoAnnounce: false,
    pullPictures: false,
    iconCacheDir: '/tmp/subminer-jellyfin-icons',
    directPlayPreferred: true,
    directPlayContainers: ['mkv', 'mp4', 'webm', 'mov', 'flac', 'mp3', 'aac'],
    transcodeVideoCodec: 'h264',
  },
  discordPresence: {
    enabled: true,
    presenceStyle: 'default' as const,
    updateIntervalMs: 3_000,
    debounceMs: 750,
  },
  ai: {
    enabled: false,
    apiKey: '',
    apiKeyCommand: '',
    model: 'openai/gpt-4o-mini',
    baseUrl: 'https://openrouter.ai/api',
    systemPrompt:
      'You are a translation engine. Return only the translated text with no explanations.',
    requestTimeoutMs: 15_000,
  },
  youtubeSubgen: {
    whisperBin: '',
    whisperModel: '',
    whisperVadModel: '',
    whisperThreads: 4,
    fixWithAi: false,
    ai: {
      model: '',
      systemPrompt: '',
    },
  },
};
