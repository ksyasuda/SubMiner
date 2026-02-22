import { ResolvedConfig } from '../../types';
import { ConfigOptionRegistryEntry, RuntimeOptionRegistryEntry } from './shared';

export function buildIntegrationConfigOptionRegistry(
  defaultConfig: ResolvedConfig,
  runtimeOptionRegistry: RuntimeOptionRegistryEntry[],
): ConfigOptionRegistryEntry[] {
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
      path: 'ankiConnect.tags',
      kind: 'array',
      defaultValue: defaultConfig.ankiConnect.tags,
      description:
        'Tags to add to cards mined or updated by SubMiner. Provide an empty array to disable automatic tagging.',
    },
    {
      path: 'ankiConnect.behavior.autoUpdateNewCards',
      kind: 'boolean',
      defaultValue: defaultConfig.ankiConnect.behavior.autoUpdateNewCards,
      description: 'Automatically update newly added cards.',
      runtime: runtimeOptionRegistry[0],
    },
    {
      path: 'ankiConnect.nPlusOne.matchMode',
      kind: 'enum',
      enumValues: ['headword', 'surface'],
      defaultValue: defaultConfig.ankiConnect.nPlusOne.matchMode,
      description: 'Known-word matching strategy for N+1 highlighting.',
    },
    {
      path: 'ankiConnect.nPlusOne.highlightEnabled',
      kind: 'boolean',
      defaultValue: defaultConfig.ankiConnect.nPlusOne.highlightEnabled,
      description: 'Enable fast local highlighting for words already known in Anki.',
    },
    {
      path: 'ankiConnect.nPlusOne.refreshMinutes',
      kind: 'number',
      defaultValue: defaultConfig.ankiConnect.nPlusOne.refreshMinutes,
      description: 'Minutes between known-word cache refreshes.',
    },
    {
      path: 'ankiConnect.nPlusOne.minSentenceWords',
      kind: 'number',
      defaultValue: defaultConfig.ankiConnect.nPlusOne.minSentenceWords,
      description: 'Minimum sentence word count required for N+1 targeting (default: 3).',
    },
    {
      path: 'ankiConnect.nPlusOne.decks',
      kind: 'array',
      defaultValue: defaultConfig.ankiConnect.nPlusOne.decks,
      description: 'Decks used for N+1 known-word cache scope. Supports one or more deck names.',
    },
    {
      path: 'ankiConnect.nPlusOne.nPlusOne',
      kind: 'string',
      defaultValue: defaultConfig.ankiConnect.nPlusOne.nPlusOne,
      description: 'Color used for the single N+1 target token highlight.',
    },
    {
      path: 'ankiConnect.nPlusOne.knownWord',
      kind: 'string',
      defaultValue: defaultConfig.ankiConnect.nPlusOne.knownWord,
      description: 'Color used for legacy known-word highlights.',
    },
    {
      path: 'ankiConnect.isKiku.fieldGrouping',
      kind: 'enum',
      enumValues: ['auto', 'manual', 'disabled'],
      defaultValue: defaultConfig.ankiConnect.isKiku.fieldGrouping,
      description: 'Kiku duplicate-card field grouping mode.',
      runtime: runtimeOptionRegistry[1],
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
      path: 'youtubeSubgen.mode',
      kind: 'enum',
      enumValues: ['automatic', 'preprocess', 'off'],
      defaultValue: defaultConfig.youtubeSubgen.mode,
      description: 'YouTube subtitle generation mode for the launcher script.',
    },
    {
      path: 'youtubeSubgen.whisperBin',
      kind: 'string',
      defaultValue: defaultConfig.youtubeSubgen.whisperBin,
      description: 'Path to whisper.cpp CLI used as fallback transcription engine.',
    },
    {
      path: 'youtubeSubgen.whisperModel',
      kind: 'string',
      defaultValue: defaultConfig.youtubeSubgen.whisperModel,
      description: 'Path to whisper model used for fallback transcription.',
    },
    {
      path: 'youtubeSubgen.primarySubLanguages',
      kind: 'string',
      defaultValue: defaultConfig.youtubeSubgen.primarySubLanguages.join(','),
      description: 'Comma-separated primary subtitle language priority used by the launcher.',
    },
  ];
}
