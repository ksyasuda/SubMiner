import { ResolvedConfig } from '../../types/config';
import { MPV_LAUNCH_MODE_VALUES } from '../../shared/mpv-launch-mode';
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
      path: 'ankiConnect.url',
      kind: 'string',
      defaultValue: defaultConfig.ankiConnect.url,
      description: 'Base URL of the AnkiConnect HTTP server.',
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
      path: 'ankiConnect.fields.audio',
      kind: 'string',
      defaultValue: defaultConfig.ankiConnect.fields.audio,
      description: 'Card field that receives generated sentence audio.',
    },
    {
      path: 'ankiConnect.fields.image',
      kind: 'string',
      defaultValue: defaultConfig.ankiConnect.fields.image,
      description: 'Card field that receives the captured screenshot or animated image.',
    },
    {
      path: 'ankiConnect.fields.sentence',
      kind: 'string',
      defaultValue: defaultConfig.ankiConnect.fields.sentence,
      description: 'Card field that receives the source sentence text.',
    },
    {
      path: 'ankiConnect.fields.miscInfo',
      kind: 'string',
      defaultValue: defaultConfig.ankiConnect.fields.miscInfo,
      description:
        'Card field that receives the miscellaneous info pattern (see ankiConnect.metadata.pattern).',
    },
    {
      path: 'ankiConnect.fields.translation',
      kind: 'string',
      defaultValue: defaultConfig.ankiConnect.fields.translation,
      description: 'Card field that receives the current selection or translated text.',
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
      path: 'ankiConnect.behavior.overwriteAudio',
      kind: 'boolean',
      defaultValue: defaultConfig.ankiConnect.behavior.overwriteAudio,
      description:
        'When updating an existing card, overwrite the audio field instead of skipping it.',
    },
    {
      path: 'ankiConnect.behavior.overwriteImage',
      kind: 'boolean',
      defaultValue: defaultConfig.ankiConnect.behavior.overwriteImage,
      description:
        'When updating an existing card, overwrite the image field instead of skipping it.',
    },
    {
      path: 'ankiConnect.behavior.mediaInsertMode',
      kind: 'enum',
      enumValues: ['append', 'prepend'],
      defaultValue: defaultConfig.ankiConnect.behavior.mediaInsertMode,
      description:
        'Whether new media is appended after or prepended before existing field contents on update.',
    },
    {
      path: 'ankiConnect.behavior.highlightWord',
      kind: 'boolean',
      defaultValue: defaultConfig.ankiConnect.behavior.highlightWord,
      description: 'Bold the mined word inside the sentence field on the saved Anki card.',
    },
    {
      path: 'ankiConnect.behavior.notificationType',
      kind: 'enum',
      enumValues: ['osd', 'system', 'both', 'none'],
      defaultValue: defaultConfig.ankiConnect.behavior.notificationType,
      description: 'Notification surface used to announce mining and update outcomes.',
    },
    {
      path: 'ankiConnect.media.syncAnimatedImageToWordAudio',
      kind: 'boolean',
      defaultValue: defaultConfig.ankiConnect.media.syncAnimatedImageToWordAudio,
      description:
        'For animated AVIF images, prepend a frozen first frame matching the existing word-audio duration so motion starts with sentence audio.',
    },
    {
      path: 'ankiConnect.media.generateAudio',
      kind: 'boolean',
      defaultValue: defaultConfig.ankiConnect.media.generateAudio,
      description: 'Generate sentence audio for mined cards.',
    },
    {
      path: 'ankiConnect.media.generateImage',
      kind: 'boolean',
      defaultValue: defaultConfig.ankiConnect.media.generateImage,
      description: 'Generate screenshot or animated image for mined cards.',
    },
    {
      path: 'ankiConnect.media.imageType',
      kind: 'enum',
      enumValues: ['static', 'avif'],
      defaultValue: defaultConfig.ankiConnect.media.imageType,
      description:
        'Image capture type: "static" for a single still frame, "avif" for an animated AVIF.',
    },
    {
      path: 'ankiConnect.media.imageFormat',
      kind: 'enum',
      enumValues: ['jpg', 'png', 'webp'],
      defaultValue: defaultConfig.ankiConnect.media.imageFormat,
      description: 'Encoding format used when imageType is "static".',
    },
    {
      path: 'ankiConnect.media.imageQuality',
      kind: 'number',
      defaultValue: defaultConfig.ankiConnect.media.imageQuality,
      description: 'Quality (0-100) used for lossy static image encoders.',
    },
    {
      path: 'ankiConnect.media.imageMaxWidth',
      kind: 'number',
      defaultValue: defaultConfig.ankiConnect.media.imageMaxWidth,
      description:
        'Optional maximum width for static images. Leave unset to preserve the source resolution.',
    },
    {
      path: 'ankiConnect.media.imageMaxHeight',
      kind: 'number',
      defaultValue: defaultConfig.ankiConnect.media.imageMaxHeight,
      description:
        'Optional maximum height for static images. Leave unset to preserve the source resolution.',
    },
    {
      path: 'ankiConnect.media.animatedFps',
      kind: 'number',
      defaultValue: defaultConfig.ankiConnect.media.animatedFps,
      description: 'Target frame rate for animated AVIF captures.',
    },
    {
      path: 'ankiConnect.media.animatedMaxWidth',
      kind: 'number',
      defaultValue: defaultConfig.ankiConnect.media.animatedMaxWidth,
      description: 'Maximum width applied to animated AVIF captures.',
    },
    {
      path: 'ankiConnect.media.animatedMaxHeight',
      kind: 'number',
      defaultValue: defaultConfig.ankiConnect.media.animatedMaxHeight,
      description:
        'Optional maximum height for animated AVIF captures. Leave unset to preserve aspect ratio.',
    },
    {
      path: 'ankiConnect.media.animatedCrf',
      kind: 'number',
      defaultValue: defaultConfig.ankiConnect.media.animatedCrf,
      description:
        'Animated AVIF CRF quality target. Lower values produce larger, higher-quality files.',
    },
    {
      path: 'ankiConnect.media.audioPadding',
      kind: 'number',
      defaultValue: defaultConfig.ankiConnect.media.audioPadding,
      description: 'Seconds of padding appended to both ends of generated sentence audio.',
    },
    {
      path: 'ankiConnect.media.fallbackDuration',
      kind: 'number',
      defaultValue: defaultConfig.ankiConnect.media.fallbackDuration,
      description: 'Fallback clip duration in seconds when subtitle timing data is unavailable.',
    },
    {
      path: 'ankiConnect.media.maxMediaDuration',
      kind: 'number',
      defaultValue: defaultConfig.ankiConnect.media.maxMediaDuration,
      description: 'Maximum allowed media clip duration in seconds.',
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
      path: 'ankiConnect.nPlusOne.enabled',
      kind: 'boolean',
      defaultValue: defaultConfig.ankiConnect.nPlusOne.enabled,
      description:
        'Enable N+1 subtitle highlighting (highlights the one unknown word in a sentence). Requires known-word cache data.',
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
      path: 'ankiConnect.isKiku.fieldGrouping',
      kind: 'enum',
      enumValues: ['auto', 'manual', 'disabled'],
      defaultValue: defaultConfig.ankiConnect.isKiku.fieldGrouping,
      description: 'Kiku duplicate-card field grouping mode.',
      runtime: runtimeOptionById.get('anki.kikuFieldGrouping'),
    },
    {
      path: 'ankiConnect.isKiku.enabled',
      kind: 'boolean',
      defaultValue: defaultConfig.ankiConnect.isKiku.enabled,
      description: 'Enable Kiku-specific mining behaviors (duplicate handling, field grouping).',
    },
    {
      path: 'ankiConnect.isKiku.deleteDuplicateInAuto',
      kind: 'boolean',
      defaultValue: defaultConfig.ankiConnect.isKiku.deleteDuplicateInAuto,
      description:
        'When Kiku field grouping is "auto", delete the duplicate source card after grouping completes.',
    },
    {
      path: 'ankiConnect.isLapis.enabled',
      kind: 'boolean',
      defaultValue: defaultConfig.ankiConnect.isLapis.enabled,
      description: 'Enable Lapis-specific mining behaviors and sentence card model targeting.',
    },
    {
      path: 'ankiConnect.isLapis.sentenceCardModel',
      kind: 'string',
      defaultValue: defaultConfig.ankiConnect.isLapis.sentenceCardModel,
      description: 'Note type name used by Lapis sentence cards.',
    },
    {
      path: 'ankiConnect.metadata.pattern',
      kind: 'string',
      defaultValue: defaultConfig.ankiConnect.metadata.pattern,
      description:
        'Template used to render the miscInfo field. Placeholders include %f (filename) and %t (timestamp).',
    },
    {
      path: 'jimaku.apiBaseUrl',
      kind: 'string',
      defaultValue: defaultConfig.jimaku.apiBaseUrl,
      description: 'Base URL of the Jimaku subtitle search API.',
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
      path: 'mpv.executablePath',
      kind: 'string',
      defaultValue: defaultConfig.mpv.executablePath,
      description:
        'Optional absolute path to mpv.exe for Windows launch flows. Leave empty to auto-discover from SUBMINER_MPV_PATH or PATH.',
    },
    {
      path: 'mpv.launchMode',
      kind: 'enum',
      enumValues: MPV_LAUNCH_MODE_VALUES,
      defaultValue: defaultConfig.mpv.launchMode,
      description: 'Default window state for SubMiner-managed mpv launches.',
    },
    {
      path: 'mpv.profile',
      kind: 'string',
      defaultValue: defaultConfig.mpv.profile,
      description:
        'Optional mpv profile name passed to SubMiner-managed mpv launches. Leave empty to pass no profile.',
    },
    {
      path: 'mpv.socketPath',
      kind: 'string',
      defaultValue: defaultConfig.mpv.socketPath,
      description:
        'mpv IPC socket path used by SubMiner-managed playback and the bundled mpv plugin.',
    },
    {
      path: 'mpv.backend',
      kind: 'enum',
      enumValues: ['auto', 'hyprland', 'sway', 'x11', 'macos', 'windows'],
      defaultValue: defaultConfig.mpv.backend,
      description:
        'Window tracking backend passed to the bundled mpv plugin. Auto detects the current platform.',
    },
    {
      path: 'mpv.autoStartSubMiner',
      kind: 'boolean',
      defaultValue: defaultConfig.mpv.autoStartSubMiner,
      description: 'Start SubMiner in the background when SubMiner-managed mpv loads a file.',
    },
    {
      path: 'mpv.pauseUntilOverlayReady',
      kind: 'boolean',
      defaultValue: defaultConfig.mpv.pauseUntilOverlayReady,
      description:
        'Pause mpv on visible-overlay auto-start until SubMiner signals subtitle tokenization readiness.',
    },
    {
      path: 'mpv.subminerBinaryPath',
      kind: 'string',
      defaultValue: defaultConfig.mpv.subminerBinaryPath,
      description:
        'Optional SubMiner app binary path passed to the bundled mpv plugin. Leave empty to use the launcher-detected app path.',
    },
    {
      path: 'mpv.aniskipEnabled',
      kind: 'boolean',
      defaultValue: defaultConfig.mpv.aniskipEnabled,
      description: 'Enable AniSkip intro detection and skip markers in the bundled mpv plugin.',
    },
    {
      path: 'mpv.aniskipButtonKey',
      kind: 'string',
      defaultValue: defaultConfig.mpv.aniskipButtonKey,
      description: 'mpv key used to trigger the AniSkip button while the skip marker is visible.',
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
      path: 'jellyfin.recentServers',
      kind: 'array',
      defaultValue: defaultConfig.jellyfin.recentServers,
      description: 'Recently authenticated Jellyfin server URLs shown in setup.',
    },
    {
      path: 'jellyfin.username',
      kind: 'string',
      defaultValue: defaultConfig.jellyfin.username,
      description: 'Default Jellyfin username used during CLI login.',
    },
    {
      path: 'jellyfin.deviceId',
      kind: 'string',
      defaultValue: defaultConfig.jellyfin.deviceId,
      description:
        'Stable device identifier sent on the Jellyfin authentication handshake; primarily internal.',
    },
    {
      path: 'jellyfin.clientName',
      kind: 'string',
      defaultValue: defaultConfig.jellyfin.clientName,
      description: 'Client name sent on the Jellyfin authentication handshake; primarily internal.',
    },
    {
      path: 'jellyfin.clientVersion',
      kind: 'string',
      defaultValue: defaultConfig.jellyfin.clientVersion,
      description:
        'Client version sent on the Jellyfin authentication handshake; primarily internal.',
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
      path: 'discordPresence.presenceStyle',
      kind: 'enum',
      enumValues: ['default', 'meme', 'japanese', 'minimal'],
      defaultValue: defaultConfig.discordPresence.presenceStyle,
      description:
        'Presence card text preset: "default" (clean bilingual), "meme" (Mining and crafting), "japanese" (fully JP), or "minimal".',
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
      path: 'ai.model',
      kind: 'string',
      defaultValue: defaultConfig.ai.model,
      description: 'Default model identifier requested from the shared AI provider.',
    },
    {
      path: 'ai.systemPrompt',
      kind: 'string',
      defaultValue: defaultConfig.ai.systemPrompt,
      description: 'Default system prompt sent with shared AI provider requests.',
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
      description:
        'Legacy compatibility path kept for external subtitle fallback tools; not used by default.',
    },
    {
      path: 'youtubeSubgen.whisperModel',
      kind: 'string',
      defaultValue: defaultConfig.youtubeSubgen.whisperModel,
      description:
        'Legacy compatibility model path kept for external subtitle fallback tooling; not used by default.',
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
