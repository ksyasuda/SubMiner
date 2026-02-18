# Launcher Script Refactor Design

## Problem

The `subminer` launcher script is a 3,682-line single-file Bun TypeScript program. It contains
several independent domains (Jimaku API, Jellyfin browsing, YouTube subtitle generation, mpv IPC,
rofi/fzf picker UI, arg parsing, config loading) all in one file. This makes the script hard to
navigate, modify, and reason about.

## Constraints

- The launcher must remain installable and runnable as a **single script** at `~/.local/bin/subminer`
- It runs under **Bun** (not Node, not Electron) — separate from the app build pipeline
- The launcher does **not import from `src/`** — it spawns the AppImage/binary as a child process
- Jimaku and Jellyfin logic is intentionally duplicated between the launcher and `src/` modules
  (different runtimes, different needs) — this refactor keeps them as separate copies

## Solution

Split the monolithic `subminer` file into a `launcher/` directory with domain-focused modules.
Use `bun build` to bundle them back into a single distributable `subminer` script.

```
launcher/
  main.ts          # entrypoint — orchestration only
  types.ts         # shared types, interfaces, constants
  log.ts           # logging, error output
  util.ts          # pure utilities, child process runner
  config.ts        # config loading from disk, arg parsing
  jimaku.ts        # Jimaku API client, media info parsing
  jellyfin.ts      # Jellyfin API, session management, browsing
  youtube.ts       # YouTube subtitle generation pipeline
  picker.ts        # rofi/fzf menu UI (shared by video + Jellyfin)
  mpv.ts           # mpv process management, IPC, overlay lifecycle
```

Build: `bun build launcher/main.ts --target=bun --outfile=subminer`

## Module Details

### `launcher/types.ts` (~100 lines)

All shared types and constants that multiple modules reference:

- Type aliases: `LogLevel`, `YoutubeSubgenMode`, `Backend`, `JimakuLanguagePreference`
- Interfaces: `Args`, `SubtitleCandidate`, `YoutubeSubgenOutputs`, `MpvTrack`,
  `LauncherYoutubeSubgenConfig`, `LauncherJellyfinConfig`, `PluginRuntimeConfig`,
  `CommandExecOptions`, `CommandExecResult`
- Constants: `VIDEO_EXTENSIONS`, `YOUTUBE_SUB_EXTENSIONS`, `YOUTUBE_AUDIO_EXTENSIONS`,
  `DEFAULT_SOCKET_PATH`, `DEFAULT_YOUTUBE_PRIMARY_SUB_LANGS`,
  `DEFAULT_YOUTUBE_SECONDARY_SUB_LANGS`, `DEFAULT_YOUTUBE_SUBGEN_OUT_DIR`,
  `DEFAULT_MPV_LOG_FILE`, `DEFAULT_YOUTUBE_YTDL_FORMAT`, `DEFAULT_JIMAKU_API_BASE_URL`,
  `DEFAULT_MPV_SUBMINER_ARGS`, `ROFI_THEME_FILE`

### `launcher/log.ts` (~60 lines)

Logging infrastructure:

- `COLORS`, `LOG_PRI` lookup tables
- `shouldLog()`, `log()`, `fail()`, `getMpvLogPath()`, `appendToMpvLog()`

Imports: `LogLevel` from `types.ts`

### `launcher/util.ts` (~120 lines)

Pure utility functions with no domain knowledge:

- Path helpers: `resolvePathMaybe()`, `resolveBinaryPathCandidate()`, `realpathMaybe()`
- Command helpers: `commandExists()`, `isExecutable()`, `runExternalCommand()`
- URL helpers: `isUrlTarget()`, `isYoutubeTarget()`
- String helpers: `sanitizeToken()`, `normalizeBasename()`, `normalizeLangCode()`,
  `uniqueNormalizedLangCodes()`, `escapeRegExp()`, `parseBoolLike()`
- `sleep()`

Imports: types from `types.ts`, logging from `log.ts`

Note: `runExternalCommand()` uses a mutable `state.youtubeSubgenChildren` set for cleanup
tracking. This state reference will be passed in or imported from `mpv.ts` where `state` lives.

### `launcher/config.ts` (~200 lines)

Configuration loading and argument parsing:

- `loadLauncherYoutubeSubgenConfig()` — reads `~/.config/SubMiner/config.jsonc`
- `loadLauncherJellyfinConfig()` — reads jellyfin section from same config
- `readPluginRuntimeConfig()` — reads mpv `subminer.conf` for auto_start/socket_path
- `parseArgs()` — hand-rolled CLI parser
- `usage()` — help text

Imports: types from `types.ts`, logging from `log.ts`, url/path helpers from `util.ts`

### `launcher/jimaku.ts` (~350 lines)

Jimaku subtitle API client and media filename parsing:

Local types (not exported to other modules):

- `JimakuEntry`, `JimakuFileEntry`, `JimakuApiError`, `JimakuApiResponse`,
  `JimakuDownloadResult`, `JimakuConfig`, `JimakuMediaInfo`

Functions:

- API: `resolveJimakuApiKey()`, `jimakuFetchJson()`, `downloadToFile()`, `getRetryAfter()`
- Media parsing: `parseMediaInfo()`, `parseMediaInfoWithGuessit()`, `parseGuessitOutput()`,
  `matchEpisodeFromName()`, `detectSeasonFromDir()`, `cleanupTitle()`
- Scoring/sorting: `sortJimakuFiles()`, `formatLangScore()`, `mapPreferenceToLanguages()`
- Query helpers: `normalizeJimakuSearchInput()`, `sanitizeJimakuQueryInput()`, `buildJimakuConfig()`
- Subtitle validation: `isValidSubtitleCandidateFile()`

Imports: types from `types.ts`, logging from `log.ts`, command helpers from `util.ts`

### `launcher/jellyfin.ts` (~450 lines)

Jellyfin API client and interactive library/item browsing:

Local types:

- `JellyfinSessionConfig`, `JellyfinLibraryEntry`, `JellyfinItemEntry`, `JellyfinGroupEntry`

Functions:

- API: `sanitizeServerUrl()`, `jellyfinApiRequest()`
- Display: `formatJellyfinItemDisplay()`
- Browsing: `resolveJellyfinSelection()`, `promptOptionalJellyfinSearch()`
- Entrypoint: `runJellyfinPlayMenu()`

Imports: types from `types.ts`, logging from `log.ts`, picker functions from `picker.ts`,
utility functions from `util.ts`

### `launcher/youtube.ts` (~350 lines)

YouTube subtitle generation pipeline (yt-dlp + whisper fallback):

Functions:

- Pipeline: `generateYoutubeSubtitles()` (main orchestrator)
- Subtitle scanning: `scanSubtitleCandidates()`, `pickBestCandidate()`,
  `classifyLanguage()`, `filenameHasLanguageTag()`
- Conversion: `convertToSrt()`, `findAudioFile()`
- Whisper: `runWhisper()`, `convertAudioForWhisper()`, `resolveWhisperBinary()`
- Language helpers: `toYtdlpLangPattern()`, `preferredLangLabel()`,
  `inferWhisperLanguage()`, `sourceTag()`

Imports: types from `types.ts`, logging from `log.ts`, `runExternalCommand()` from `util.ts`

### `launcher/picker.ts` (~250 lines)

Rofi and fzf menu UI used by both video browsing and Jellyfin browsing:

Functions:

- Generic menus: `showRofiFlatMenu()`, `showFzfFlatMenu()`
- Video menus: `showRofiMenu()`, `showFzfMenu()`, `buildRofiMenu()`, `buildFzfMenu()`
- Video discovery: `collectVideos()`
- Theme: `findRofiTheme()`
- Jellyfin pickers: `pickLibrary()`, `pickItem()`, `pickGroup()`
- Helpers: `formatPickerLaunchError()`, `escapeShellSingle()`,
  `parseSelectionId()`, `parseSelectionLabel()`

Imports: types from `types.ts`, logging from `log.ts`, `commandExists()` from `util.ts`

### `launcher/mpv.ts` (~250 lines)

mpv process lifecycle, IPC socket communication, and overlay management:

Exports:

- `state` object (`mpvProc`, `overlayProc`, `youtubeSubgenChildren`, `appPath`, etc.)
- IPC: `sendMpvCommand()`, `sendMpvCommandWithResponse()`
- Track detection: `getMpvTracks()`, `waitForSubtitleTrackList()`, `findPreferredSubtitleTrack()`,
  `isPreferredStreamLang()`
- Subtitle loading: `loadSubtitleIntoMpv()`
- Socket: `waitForSocket()`
- Process management: `startMpv()`, `startOverlay()`, `stopOverlay()`, `launchTexthookerOnly()`
- Binary resolution: `findAppBinary()`, `resolveMacAppBinaryCandidate()`
- Backend detection: `detectBackend()`

Imports: types from `types.ts`, logging from `log.ts`, helpers from `util.ts`

### `launcher/main.ts` (~100 lines)

Entrypoint — pure orchestration with no domain logic:

Functions:

- `main()` — top-level flow
- `chooseTarget()` — resolve file/URL target
- `registerCleanup()` — SIGINT/SIGTERM handlers
- `checkDependencies()`, `checkPickerDependencies()`
- `runAppCommandWithInherit()`

Imports: everything it needs from the other modules

## Dependency Graph

```
main.ts
  ├── config.ts     (parseArgs, loadConfig, readPluginRuntimeConfig)
  ├── mpv.ts        (startMpv, startOverlay, stopOverlay, state, waitForSocket, ...)
  ├── picker.ts     (chooseTarget delegates to showRofiMenu/showFzfMenu)
  ├── youtube.ts    (generateYoutubeSubtitles)
  ├── jellyfin.ts   (runJellyfinPlayMenu)
  ├── util.ts       (isYoutubeTarget, commandExists, ...)
  ├── log.ts        (log, fail)
  └── types.ts      (Args, etc.)

jellyfin.ts
  ├── picker.ts     (pickLibrary, pickItem, pickGroup, promptOptionalJellyfinSearch)
  ├── util.ts
  ├── log.ts
  └── types.ts

youtube.ts
  ├── util.ts       (runExternalCommand, uniqueNormalizedLangCodes, ...)
  ├── log.ts
  └── types.ts

picker.ts
  ├── util.ts       (commandExists)
  ├── log.ts
  └── types.ts

config.ts
  ├── util.ts       (resolvePathMaybe, uniqueNormalizedLangCodes, ...)
  ├── log.ts
  └── types.ts

mpv.ts
  ├── util.ts
  ├── log.ts
  └── types.ts
```

No circular dependencies. `types.ts` and `log.ts` are leaf nodes.

## Build Integration

Add to Makefile:

```makefile
build-launcher:
	bun build launcher/main.ts --target=bun --outfile=subminer
	chmod +x subminer
```

The existing `install-linux` target will depend on `build-launcher` to ensure the bundled
script is up to date before copying to `~/.local/bin/`.

If `bun build` does not preserve the `#!/usr/bin/env bun` shebang, the build step will
prepend it to the output.

## Migration Strategy

1. Create the `launcher/` directory and all module files
2. Move functions from `subminer` into the appropriate modules (pure moves, no logic changes)
3. Add imports/exports to wire modules together
4. Add `build-launcher` Makefile target
5. Build and verify the bundled output matches current behavior
6. Remove the old monolithic `subminer` file (it becomes a build artifact)

## What This Refactor Does NOT Do

- No logic changes, no bug fixes, no feature additions
- No sharing of code between launcher and `src/` modules
- No changes to the Electron app build pipeline
- No changes to CLI flags, behavior, or output
