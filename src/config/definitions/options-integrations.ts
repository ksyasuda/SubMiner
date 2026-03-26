import { ResolvedConfig } from '../../types';
import { ConfigOptionRegistryEntry, RuntimeOptionRegistryEntry } from './shared';

export function buildIntegrationConfigOptionRegistry(
  defaultConfig: ResolvedConfig,
  runtimeOptionRegistry: RuntimeOptionRegistryEntry[],
): ConfigOptionRegistryEntry[] {
  const runtimeOptionById = new Map(runtimeOptionRegistry.map((entry) => [entry.id, entry]));

  return [
    {
      path: 'ankiConnect.enabled',
      kind: 'boolean',
      defaultValue: defaultConfig.ankiConnect.enabled,
      description: 'Enable AnkiConnect integration.',
    },
    {
      path: 'ankiConnect.pollingRate',
      kind: 'number',
      defaultValue: defaultConfig.ankiConnect.pollingRate,
      description: 'Polling interval in milliseconds.',
    },
    {
      path: 'ankiConnect.proxy.enabled',
      kind: 'boolean',
      defaultValue: defaultConfig.ankiConnect.proxy.enabled,
      description: 'Enable local AnkiConnect-compatible proxy for push-based auto-enrichment.',
    },
    {
      path: 'ankiConnect.proxy.host',
      kind: 'string',
      defaultValue: defaultConfig.ankiConnect.proxy.host,
      description: 'Bind host for local AnkiConnect proxy.',
    },
    {
      path: 'ankiConnect.proxy.port',
      kind: 'number',
      defaultValue: defaultConfig.ankiConnect.proxy.port,
      description: 'Bind port for local AnkiConnect proxy.',
    },
    {
      path: 'ankiConnect.proxy.upstreamUrl',
      kind: 'string',
      defaultValue: defaultConfig.ankiConnect.proxy.upstreamUrl,
      description: 'Upstream AnkiConnect URL proxied by local AnkiConnect proxy.',
    },
    {
      path: 'ankiConnect.tags',
      kind: 'array',
      defaultValue: defaultConfig.ankiConnect.tags,
      description:
        'Tags to add to cards mined or updated by SubMiner. Provide an empty array to disable automatic tagging.',
    },
    {
      path: 'ankiConnect.fields.word',
      kind: 'string',
      defaultValue: defaultConfig.ankiConnect.fields.word,
      description: 'Card field for the mined word or expression text.',
    },
    {
      path: 'ankiConnect.ai.enabled',
      kind: 'boolean',
      defaultValue: defaultConfig.ankiConnect.ai.enabled,
      description: 'Enable AI provider usage for Anki translation/enrichment flows.',
    },
    {
      path: 'ankiConnect.ai.model',
      kind: 'string',
      defaultValue: defaultConfig.ankiConnect.ai.model,
      description: 'Optional model override for Anki AI translation/enrichment flows.',
    },
    {
      path: 'ankiConnect.ai.systemPrompt',
      kind: 'string',
      defaultValue: defaultConfig.ankiConnect.ai.systemPrompt,
      description: 'Optional system prompt override for Anki AI translation/enrichment flows.',
    },
    {
      path: 'ankiConnect.behavior.autoUpdateNewCards',
      kind: 'boolean',
      defaultValue: defaultConfig.ankiConnect.behavior.autoUpdateNewCards,
      description: 'Automatically update newly added cards.',
      runtime: runtimeOptionById.get('anki.autoUpdateNewCards'),
    },
    {
      path: 'ankiConnect.media.syncAnimatedImageToWordAudio',
      kind: 'boolean',
      defaultValue: defaultConfig.ankiConnect.media.syncAnimatedImageToWordAudio,
      description:
        'For animated AVIF images, prepend a frozen first frame matching the existing word-audio duration so motion starts with sentence audio.',
    },
    {
      path: 'ankiConnect.knownWords.matchMode',
      kind: 'enum',
      enumValues: ['headword', 'surface'],
      defaultValue: defaultConfig.ankiConnect.knownWords.matchMode,
      description: 'Known-word matching strategy for subtitle annotations.',
    },
    {
      path: 'ankiConnect.knownWords.highlightEnabled',
      kind: 'boolean',
      defaultValue: defaultConfig.ankiConnect.knownWords.highlightEnabled,
      description: 'Enable fast local highlighting for words already known in Anki.',
    },
    {
      path: 'ankiConnect.knownWords.refreshMinutes',
      kind: 'number',
      defaultValue: defaultConfig.ankiConnect.knownWords.refreshMinutes,
      description: 'Minutes between known-word cache refreshes.',
    },
    {
      path: 'ankiConnect.knownWords.addMinedWordsImmediately',
      kind: 'boolean',
      defaultValue: defaultConfig.ankiConnect.knownWords.addMinedWordsImmediately,
      description: 'Immediately append newly mined card words into the known-word cache.',
    },
    {
      path: 'ankiConnect.nPlusOne.minSentenceWords',
      kind: 'number',
      defaultValue: defaultConfig.ankiConnect.nPlusOne.minSentenceWords,
      description: 'Minimum sentence word count required for N+1 targeting (default: 3).',
    },
    {
      path: 'ankiConnect.knownWords.decks',
      kind: 'object',
      defaultValue: defaultConfig.ankiConnect.knownWords.decks,
      description:
        'Decks and fields for known-word cache. Object mapping deck names to arrays of field names to extract, e.g. { "Kaishi 1.5k": ["Word", "Word Reading"] }.',
    },
    {
      path: 'ankiConnect.nPlusOne.nPlusOne',
      kind: 'string',
      defaultValue: defaultConfig.ankiConnect.nPlusOne.nPlusOne,
      description: 'Color used for the single N+1 target token highlight.',
    },
    {
      path: 'ankiConnect.knownWords.color',
      kind: 'string',
      defaultValue: defaultConfig.ankiConnect.knownWords.color,
      description: 'Color used for known-word highlights.',
    },
    {
      path: 'ankiConnect.isKiku.fieldGrouping',
      kind: 'enum',
      enumValues: ['auto', 'manual', 'disabled'],
      defaultValue: defaultConfig.ankiConnect.isKiku.fieldGrouping,
      description: 'Kiku duplicate-card field grouping mode.',
      runtime: runtimeOptionById.get('anki.kikuFieldGrouping'),
    },
    {
      path: 'jimaku.languagePreference',
      kind: 'enum',
      enumValues: ['ja', 'en', 'none'],
      defaultValue: defaultConfig.jimaku.languagePreference,
      description: 'Preferred language used in Jimaku search.',
    },
    {
      path: 'jimaku.maxEntryResults',
      kind: 'number',
      defaultValue: defaultConfig.jimaku.maxEntryResults,
      description: 'Maximum Jimaku search results returned.',
    },
    {
      path: 'anilist.enabled',
      kind: 'boolean',
      defaultValue: defaultConfig.anilist.enabled,
      description: 'Enable AniList post-watch progress updates.',
    },
    {
      path: 'anilist.accessToken',
      kind: 'string',
      defaultValue: defaultConfig.anilist.accessToken,
      description:
        'Optional explicit AniList access token override; leave empty to use locally stored token from setup.',
    },
    {
      path: 'anilist.characterDictionary.enabled',
      kind: 'boolean',
      defaultValue: defaultConfig.anilist.characterDictionary.enabled,
      description:
        'Enable automatic Yomitan character dictionary sync for currently watched AniList media.',
    },
    {
      path: 'anilist.characterDictionary.refreshTtlHours',
      kind: 'number',
      defaultValue: defaultConfig.anilist.characterDictionary.refreshTtlHours,
      description:
        'Legacy setting; merged character dictionary retention is now usage-based and this value is ignored.',
    },
    {
      path: 'anilist.characterDictionary.maxLoaded',
      kind: 'number',
      defaultValue: defaultConfig.anilist.characterDictionary.maxLoaded,
      description:
        'Maximum number of most-recently-used anime snapshots included in the merged Yomitan character dictionary.',
    },
    {
      path: 'anilist.characterDictionary.evictionPolicy',
      kind: 'enum',
      enumValues: ['disable', 'delete'],
      defaultValue: defaultConfig.anilist.characterDictionary.evictionPolicy,
      description:
        'Legacy setting; merged character dictionary eviction is usage-based and this value is ignored.',
    },
    {
      path: 'anilist.characterDictionary.profileScope',
      kind: 'enum',
      enumValues: ['all', 'active'],
      defaultValue: defaultConfig.anilist.characterDictionary.profileScope,
      description: 'Yomitan profile scope for dictionary enable/disable updates.',
    },
    {
      path: 'anilist.characterDictionary.collapsibleSections.description',
      kind: 'boolean',
      defaultValue: defaultConfig.anilist.characterDictionary.collapsibleSections.description,
      description:
        'Open the Description section by default in character dictionary glossary entries.',
    },
    {
      path: 'anilist.characterDictionary.collapsibleSections.characterInformation',
      kind: 'boolean',
      defaultValue:
        defaultConfig.anilist.characterDictionary.collapsibleSections.characterInformation,
      description:
        'Open the Character Information section by default in character dictionary glossary entries.',
    },
    {
      path: 'anilist.characterDictionary.collapsibleSections.voicedBy',
      kind: 'boolean',
      defaultValue: defaultConfig.anilist.characterDictionary.collapsibleSections.voicedBy,
      description:
        'Open the Voiced by section by default in character dictionary glossary entries.',
    },
    {
      path: 'yomitan.externalProfilePath',
      kind: 'string',
      defaultValue: defaultConfig.yomitan.externalProfilePath,
      description:
        'Optional external Yomitan Electron profile path to use in read-only mode for shared dictionaries/settings. Example: ~/.config/gsm_overlay',
    },
    {
      path: 'jellyfin.enabled',
      kind: 'boolean',
      defaultValue: defaultConfig.jellyfin.enabled,
      description: 'Enable optional Jellyfin integration and CLI control commands.',
    },
    {
      path: 'jellyfin.serverUrl',
      kind: 'string',
      defaultValue: defaultConfig.jellyfin.serverUrl,
      description: 'Base Jellyfin server URL (for example: http://localhost:8096).',
    },
    {
      path: 'jellyfin.username',
      kind: 'string',
      defaultValue: defaultConfig.jellyfin.username,
      description: 'Default Jellyfin username used during CLI login.',
    },
    {
      path: 'jellyfin.defaultLibraryId',
      kind: 'string',
      defaultValue: defaultConfig.jellyfin.defaultLibraryId,
      description: 'Optional default Jellyfin library ID for item listing.',
    },
    {
      path: 'jellyfin.remoteControlEnabled',
      kind: 'boolean',
      defaultValue: defaultConfig.jellyfin.remoteControlEnabled,
      description: 'Enable Jellyfin remote cast control mode.',
    },
    {
      path: 'jellyfin.remoteControlAutoConnect',
      kind: 'boolean',
      defaultValue: defaultConfig.jellyfin.remoteControlAutoConnect,
      description: 'Auto-connect to the configured remote control target.',
    },
    {
      path: 'jellyfin.autoAnnounce',
      kind: 'boolean',
      defaultValue: defaultConfig.jellyfin.autoAnnounce,
      description:
        'When enabled, automatically trigger remote announce/visibility check on websocket connect.',
    },
    {
      path: 'jellyfin.remoteControlDeviceName',
      kind: 'string',
      defaultValue: defaultConfig.jellyfin.remoteControlDeviceName,
      description: 'Device name reported for Jellyfin remote control sessions.',
    },
    {
      path: 'jellyfin.pullPictures',
      kind: 'boolean',
      defaultValue: defaultConfig.jellyfin.pullPictures,
      description: 'Enable Jellyfin poster/icon fetching for launcher menus.',
    },
    {
      path: 'jellyfin.iconCacheDir',
      kind: 'string',
      defaultValue: defaultConfig.jellyfin.iconCacheDir,
      description: 'Directory used by launcher for cached Jellyfin poster icons.',
    },
    {
      path: 'jellyfin.directPlayPreferred',
      kind: 'boolean',
      defaultValue: defaultConfig.jellyfin.directPlayPreferred,
      description: 'Try direct play before server-managed transcoding when possible.',
    },
    {
      path: 'jellyfin.directPlayContainers',
      kind: 'array',
      defaultValue: defaultConfig.jellyfin.directPlayContainers,
      description: 'Container allowlist for direct play decisions.',
    },
    {
      path: 'jellyfin.transcodeVideoCodec',
      kind: 'string',
      defaultValue: defaultConfig.jellyfin.transcodeVideoCodec,
      description: 'Preferred transcode video codec when direct play is unavailable.',
    },
    {
      path: 'discordPresence.enabled',
      kind: 'boolean',
      defaultValue: defaultConfig.discordPresence.enabled,
      description: 'Enable optional Discord Rich Presence updates.',
    },
    {
      path: 'discordPresence.updateIntervalMs',
      kind: 'number',
      defaultValue: defaultConfig.discordPresence.updateIntervalMs,
      description: 'Minimum interval between presence payload updates.',
    },
    {
      path: 'discordPresence.debounceMs',
      kind: 'number',
      defaultValue: defaultConfig.discordPresence.debounceMs,
      description: 'Debounce delay used to collapse bursty presence updates.',
    },
    {
      path: 'ai.enabled',
      kind: 'boolean',
      defaultValue: defaultConfig.ai.enabled,
      description: 'Enable shared OpenAI-compatible AI provider features.',
    },
    {
      path: 'ai.apiKey',
      kind: 'string',
      defaultValue: defaultConfig.ai.apiKey,
      description: 'Static API key for the shared OpenAI-compatible AI provider.',
    },
    {
      path: 'ai.apiKeyCommand',
      kind: 'string',
      defaultValue: defaultConfig.ai.apiKeyCommand,
      description: 'Shell command used to resolve the shared AI provider API key.',
    },
    {
      path: 'ai.baseUrl',
      kind: 'string',
      defaultValue: defaultConfig.ai.baseUrl,
      description: 'Base URL for the shared OpenAI-compatible AI provider.',
    },
    {
      path: 'ai.requestTimeoutMs',
      kind: 'number',
      defaultValue: defaultConfig.ai.requestTimeoutMs,
      description: 'Timeout in milliseconds for shared AI provider requests.',
    },
    {
      path: 'youtubeSubgen.whisperBin',
      kind: 'string',
      defaultValue: defaultConfig.youtubeSubgen.whisperBin,
      description: 'Legacy compatibility path kept for external subtitle fallback tools; not used by default.',
    },
    {
      path: 'youtubeSubgen.whisperModel',
      kind: 'string',
      defaultValue: defaultConfig.youtubeSubgen.whisperModel,
      description: 'Legacy compatibility model path kept for external subtitle fallback tooling; not used by default.',
    },
    {
      path: 'youtubeSubgen.whisperVadModel',
      kind: 'string',
      defaultValue: defaultConfig.youtubeSubgen.whisperVadModel,
      description:
        'Legacy compatibility VAD path kept for external subtitle fallback tooling; not used by default.',
    },
    {
      path: 'youtubeSubgen.whisperThreads',
      kind: 'number',
      defaultValue: defaultConfig.youtubeSubgen.whisperThreads,
      description: 'Legacy thread tuning for subtitle fallback tooling; not used by default.',
    },
    {
      path: 'youtubeSubgen.fixWithAi',
      kind: 'boolean',
      defaultValue: defaultConfig.youtubeSubgen.fixWithAi,
      description:
        'Legacy subtitle fallback post-processing switch kept for compatibility; use is currently disabled by default.',
    },
    {
      path: 'youtubeSubgen.ai.model',
      kind: 'string',
      defaultValue: defaultConfig.youtubeSubgen.ai.model,
      description:
        'Optional model override for legacy subtitle fallback post-processing; not used by default.',
    },
    {
      path: 'youtubeSubgen.ai.systemPrompt',
      kind: 'string',
      defaultValue: defaultConfig.youtubeSubgen.ai.systemPrompt,
      description:
        'Optional system prompt override for legacy subtitle fallback post-processing; not used by default.',
    },
  ];
}
