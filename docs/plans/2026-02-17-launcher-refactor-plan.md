# Launcher Script Refactor Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Split the 3,682-line monolithic `subminer` Bun script into 10 domain-focused modules under `launcher/`, bundled back into a single distributable script via `bun build`.

**Architecture:** Each module owns one domain (types, logging, utilities, config, Jimaku, Jellyfin, YouTube, picker UI, mpv, orchestration). The `launcher/main.ts` entrypoint imports from all others. `bun build` bundles everything into the `subminer` output script. No logic changes — pure code moves with import/export wiring.

**Tech Stack:** Bun (runtime + bundler), TypeScript, Make

**Design doc:** `docs/plans/2026-02-17-launcher-refactor-design.md`

---

## Verified Build Recipe

Tested and confirmed working:

```bash
bun build ./launcher/main.ts --target=bun --packages=bundle --outfile=subminer
sed -i '1s|^// @bun|#!/usr/bin/env bun\n// @bun|' subminer
chmod +x subminer
```

`bun build` emits `// @bun` as line 1 which breaks the shebang. The `sed` command inserts the
shebang before it. The bundled output inlines the `jsonc-parser` dependency and all launcher
modules into a single file.

---

### Task 1: Create launcher/ directory and types.ts

**Files:**

- Create: `launcher/types.ts`

**Step 1: Create the launcher directory**

```bash
mkdir -p launcher
```

**Step 2: Create `launcher/types.ts`**

Extract from `subminer` (lines 17-66, 67-73, 519-582, 1462-1489, 607-612):

- All `const` sets: `VIDEO_EXTENSIONS`, `YOUTUBE_SUB_EXTENSIONS`, `YOUTUBE_AUDIO_EXTENSIONS`
- All `DEFAULT_*` constants: `ROFI_THEME_FILE`, `DEFAULT_SOCKET_PATH`, `DEFAULT_YOUTUBE_PRIMARY_SUB_LANGS`, `DEFAULT_YOUTUBE_SECONDARY_SUB_LANGS`, `DEFAULT_YOUTUBE_SUBGEN_OUT_DIR`, `DEFAULT_MPV_LOG_FILE`, `DEFAULT_YOUTUBE_YTDL_FORMAT`, `DEFAULT_JIMAKU_API_BASE_URL`, `DEFAULT_MPV_SUBMINER_ARGS`
- Type aliases: `LogLevel`, `YoutubeSubgenMode`, `Backend`, `JimakuLanguagePreference`
- Interfaces: `Args`, `LauncherYoutubeSubgenConfig`, `LauncherJellyfinConfig`, `PluginRuntimeConfig`, `CommandExecOptions`, `CommandExecResult`, `SubtitleCandidate`, `YoutubeSubgenOutputs`, `MpvTrack`

All items must be `export`ed. Add necessary imports at the top: `import path from "node:path"` and `import os from "node:os"` (for `DEFAULT_YOUTUBE_SUBGEN_OUT_DIR` and `DEFAULT_MPV_LOG_FILE`).

**Step 3: Verify it compiles**

```bash
bun build ./launcher/types.ts --target=bun --no-bundle --outdir=/tmp/launcher-check 2>&1
```

Expected: no errors.

**Step 4: Commit**

```bash
git add launcher/types.ts
git commit -m "refactor(launcher): extract types and constants to launcher/types.ts"
```

---

### Task 2: Create launcher/log.ts

**Files:**

- Create: `launcher/log.ts`

**Step 1: Create `launcher/log.ts`**

Extract from `subminer` (lines 583-724):

- `COLORS` object
- `LOG_PRI` record
- `shouldLog()`, `log()`, `getMpvLogPath()`, `appendToMpvLog()`, `fail()`

Add at top:

```typescript
import type { LogLevel } from './types.js';
```

All functions must be `export`ed. Keep `import fs from "node:fs"` and `import path from "node:path"` for the logging file operations.

**Step 2: Verify it compiles**

```bash
bun build ./launcher/log.ts --target=bun --no-bundle --outdir=/tmp/launcher-check 2>&1
```

Expected: no errors.

**Step 3: Commit**

```bash
git add launcher/log.ts
git commit -m "refactor(launcher): extract logging to launcher/log.ts"
```

---

### Task 3: Create launcher/util.ts

**Files:**

- Create: `launcher/util.ts`

**Step 1: Create `launcher/util.ts`**

Extract from `subminer`:

- `sleep()` (line 1491)
- `isExecutable()` (line 726)
- `commandExists()` (line 770)
- `resolvePathMaybe()` (line 780)
- `resolveBinaryPathCandidate()` (line 787)
- `realpathMaybe()` (line 1443)
- `isUrlTarget()` (line 1451)
- `isYoutubeTarget()` (line 1455)
- `sanitizeToken()` (line 1495)
- `normalizeBasename()` (line 1503)
- `normalizeLangCode()` (line 1511)
- `uniqueNormalizedLangCodes()` (line 1515)
- `escapeRegExp()` (line 1531)
- `parseBoolLike()` (line 3193)
- `runExternalCommand()` (line 1576)

Add at top:

```typescript
import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import { spawn } from 'node:child_process';
import type { CommandExecOptions, CommandExecResult } from './types.js';
import { log } from './log.js';
```

`runExternalCommand()` currently references `state.youtubeSubgenChildren`. Change this to accept the set as a parameter:

```typescript
export function runExternalCommand(
  executable: string,
  args: string[],
  opts: CommandExecOptions = {},
  childTracker?: Set<ReturnType<typeof spawn>>,
): Promise<CommandExecResult> {
```

Inside: replace `state.youtubeSubgenChildren.add(child)` with `childTracker?.add(child)` and `state.youtubeSubgenChildren.delete(child)` with `childTracker?.delete(child)`.

All functions must be `export`ed.

**Step 2: Verify it compiles**

```bash
bun build ./launcher/util.ts --target=bun --no-bundle --outdir=/tmp/launcher-check 2>&1
```

Expected: no errors.

**Step 3: Commit**

```bash
git add launcher/util.ts
git commit -m "refactor(launcher): extract utilities to launcher/util.ts"
```

---

### Task 4: Create launcher/config.ts

**Files:**

- Create: `launcher/config.ts`

**Step 1: Create `launcher/config.ts`**

Extract from `subminer`:

- `loadLauncherYoutubeSubgenConfig()` (line 794)
- `loadLauncherJellyfinConfig()` (line 903)
- `readPluginRuntimeConfig()` (line 3225) — also needs `getPluginConfigCandidates()` (line 3214)
- `parseArgs()` (line 2725)
- `usage()` (line 614)

Add imports:

```typescript
import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import { parse as parseJsonc } from 'jsonc-parser';
import type {
  LogLevel,
  YoutubeSubgenMode,
  Backend,
  Args,
  LauncherYoutubeSubgenConfig,
  LauncherJellyfinConfig,
  PluginRuntimeConfig,
  JimakuLanguagePreference,
} from './types.js';
import {
  DEFAULT_SOCKET_PATH,
  DEFAULT_YOUTUBE_PRIMARY_SUB_LANGS,
  DEFAULT_YOUTUBE_SECONDARY_SUB_LANGS,
  DEFAULT_YOUTUBE_SUBGEN_OUT_DIR,
  DEFAULT_JIMAKU_API_BASE_URL,
} from './types.js';
import { log, fail } from './log.js';
import { resolvePathMaybe, isUrlTarget, uniqueNormalizedLangCodes, parseBoolLike } from './util.js';
```

`parseArgs()` calls `inferWhisperLanguage()` — that function will live in `youtube.ts`. For now, inline it or import it. Since `youtube.ts` hasn't been created yet, the simplest approach is to move `inferWhisperLanguage()` into `util.ts` (it's a pure string function with no domain coupling) and import it from there.

All functions must be `export`ed.

**Step 2: Verify it compiles**

```bash
bun build ./launcher/config.ts --target=bun --no-bundle --outdir=/tmp/launcher-check 2>&1
```

Expected: no errors (may warn about unresolved imports to sibling modules; that's fine).

**Step 3: Commit**

```bash
git add launcher/config.ts
git commit -m "refactor(launcher): extract config loading and arg parsing to launcher/config.ts"
```

---

### Task 5: Create launcher/jimaku.ts

**Files:**

- Create: `launcher/jimaku.ts`

**Step 1: Create `launcher/jimaku.ts`**

Extract from `subminer`:

Local types (keep unexported or export as needed):

- `JimakuEntry` (line 74), `JimakuFileEntry` (line 88), `JimakuApiError` (line 95),
  `JimakuApiResponse` (line 101), `JimakuDownloadResult` (line 105),
  `JimakuConfig` (line 109), `JimakuMediaInfo` (line 117)

Functions:

- `getRetryAfter()` (line 126)
- `matchEpisodeFromName()` (line 135)
- `detectSeasonFromDir()` (line 184)
- `parseGuessitOutput()` (line 192)
- `parseMediaInfoWithGuessit()` (line 254)
- `cleanupTitle()` (line 272)
- `formatLangScore()` (line 280)
- `resolveJimakuApiKey()` (line 301)
- `jimakuFetchJson()` (line 322)
- `parseMediaInfo()` (line 399)
- `sortJimakuFiles()` (line 441)
- `downloadToFile()` (line 453)
- `isValidSubtitleCandidateFile()` (line 2020)
- `mapPreferenceToLanguages()` (line 2031)
- `normalizeJimakuSearchInput()` (line 2315)
- `sanitizeJimakuQueryInput()` (line 2345)
- `buildJimakuConfig()` (line 2353)

Add imports:

```typescript
import fs from 'node:fs';
import path from 'node:path';
import http from 'node:http';
import https from 'node:https';
import { spawnSync } from 'node:child_process';
import type { Args, JimakuLanguagePreference } from './types.js';
import { DEFAULT_JIMAKU_API_BASE_URL } from './types.js';
import { commandExists } from './util.js';
```

Export all functions that are used by other modules (at minimum `buildJimakuConfig`, `normalizeJimakuSearchInput`, `isValidSubtitleCandidateFile`, `mapPreferenceToLanguages`, `parseMediaInfo`, `resolveJimakuApiKey`, `jimakuFetchJson`, `sortJimakuFiles`, `downloadToFile`).

**Step 2: Verify it compiles**

```bash
bun build ./launcher/jimaku.ts --target=bun --no-bundle --outdir=/tmp/launcher-check 2>&1
```

**Step 3: Commit**

```bash
git add launcher/jimaku.ts
git commit -m "refactor(launcher): extract Jimaku API to launcher/jimaku.ts"
```

---

### Task 6: Create launcher/picker.ts

**Files:**

- Create: `launcher/picker.ts`

**Step 1: Create `launcher/picker.ts`**

Extract from `subminer`:

- `escapeShellSingle()` (line 995)
- `showRofiFlatMenu()` (line 999)
- `showFzfFlatMenu()` (line 1039)
- `parseSelectionId()` (line 1082)
- `parseSelectionLabel()` (line 1089)
- `promptOptionalJellyfinSearch()` (line 1095)
- `pickLibrary()` (line 1147) — also needs `libraryPreviewUrl()` (line 1074)
- `pickItem()` (line 1189) — also needs `itemPreviewUrl()` (line 1078)
- `pickGroup()` (line 1224)
- `formatPickerLaunchError()` (line 2252)
- `collectVideos()` (line 2095)
- `buildRofiMenu()` (line 2125)
- `showRofiMenu()` (line 2141) — also needs `findRofiTheme()` (line 2061)
- `buildFzfMenu()` (line 2185)
- `showFzfMenu()` (line 2189)

Add imports:

```typescript
import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import { spawnSync } from 'node:child_process';
import type { LogLevel } from './types.js';
import { VIDEO_EXTENSIONS, ROFI_THEME_FILE } from './types.js';
import { log, fail } from './log.js';
import { commandExists, realpathMaybe } from './util.js';
```

The Jellyfin picker functions (`pickLibrary`, `pickItem`, `pickGroup`) take a `JellyfinSessionConfig` parameter. Since that type is local to `jellyfin.ts`, define a minimal interface locally in picker.ts or accept it as a generic object with `serverUrl` and `accessToken` fields. The cleanest approach: export `JellyfinSessionConfig` from `jellyfin.ts` and import it in `picker.ts`. But since jellyfin.ts doesn't exist yet, define the shape inline for now:

```typescript
interface PickerSessionConfig {
  serverUrl: string;
  accessToken: string;
}
```

Then `jellyfin.ts` will use this same type or its own `JellyfinSessionConfig` that extends it.

Actually, simpler: move `JellyfinSessionConfig` to `types.ts` since it's shared between picker and jellyfin modules. Add it in Task 1 or patch `types.ts` here.

Export all functions.

**Step 2: Verify it compiles**

```bash
bun build ./launcher/picker.ts --target=bun --no-bundle --outdir=/tmp/launcher-check 2>&1
```

**Step 3: Commit**

```bash
git add launcher/picker.ts
git commit -m "refactor(launcher): extract picker UI to launcher/picker.ts"
```

---

### Task 7: Create launcher/mpv.ts

**Files:**

- Create: `launcher/mpv.ts`

**Step 1: Create `launcher/mpv.ts`**

Extract from `subminer`:

State object (line 598):

- `state` — export it so `main.ts` and `util.ts` (via `runExternalCommand`) can access it

IPC functions:

- `sendMpvCommand()` (line 1802)
- `sendMpvCommandWithResponse()` (line 1872) — also needs `MpvResponseEnvelope` (line 1866)
- `getMpvTracks()` (line 1938)
- `waitForSubtitleTrackList()` (line 2000)
- `isPreferredStreamLang()` (line 1971)
- `findPreferredSubtitleTrack()` (line 1982)
- `loadSubtitleIntoMpv()` (line 1816)
- `waitForSocket()` (line 3280)

Process management:

- `detectBackend()` (line 2041)
- `resolveMacAppBinaryCandidate()` (line 735)
- `findAppBinary()` (line 2264)
- `startMpv()` (line 3300)
- `startOverlay()` (line 3112)
- `stopOverlay()` (line 3150)
- `launchTexthookerOnly()` (line 3140)
- `makeTempDir()` (line 2037) — if only used by mpv.ts, otherwise put in util.ts

Add imports:

```typescript
import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import net from 'node:net';
import { spawn, spawnSync } from 'node:child_process';
import type { LogLevel, Backend, Args } from './types.js';
import { DEFAULT_SOCKET_PATH, DEFAULT_MPV_SUBMINER_ARGS } from './types.js';
import { log, fail, getMpvLogPath } from './log.js';
import {
  commandExists,
  isExecutable,
  resolvePathMaybe,
  realpathMaybe,
  isYoutubeTarget,
  uniqueNormalizedLangCodes,
  sleep,
} from './util.js';
```

Export `state` and all functions.

**Step 2: Verify it compiles**

```bash
bun build ./launcher/mpv.ts --target=bun --no-bundle --outdir=/tmp/launcher-check 2>&1
```

**Step 3: Commit**

```bash
git add launcher/mpv.ts
git commit -m "refactor(launcher): extract mpv management to launcher/mpv.ts"
```

---

### Task 8: Create launcher/youtube.ts

**Files:**

- Create: `launcher/youtube.ts`

**Step 1: Create `launcher/youtube.ts`**

Extract from `subminer`:

Language/classification helpers:

- `toYtdlpLangPattern()` (line 1527)
- `filenameHasLanguageTag()` (line 1535)
- `classifyLanguage()` (line 1541)
- `preferredLangLabel()` (line 1558)
- `sourceTag()` (line 1570)

Subtitle operations:

- `pickBestCandidate()` (line 1673)
- `scanSubtitleCandidates()` (line 1687)
- `convertToSrt()` (line 1715)
- `findAudioFile()` (line 1726)

Whisper:

- `runWhisper()` (line 1749)
- `convertAudioForWhisper()` (line 1780)
- `resolveWhisperBinary()` (line 2394)

Main pipeline:

- `generateYoutubeSubtitles()` (line 2401)

Add imports:

```typescript
import fs from 'node:fs';
import path from 'node:path';
import os from 'node:os';
import type { Args, SubtitleCandidate, YoutubeSubgenOutputs } from './types.js';
import { YOUTUBE_SUB_EXTENSIONS, YOUTUBE_AUDIO_EXTENSIONS } from './types.js';
import { log } from './log.js';
import {
  resolvePathMaybe,
  uniqueNormalizedLangCodes,
  normalizeLangCode,
  escapeRegExp,
  normalizeBasename,
  sanitizeToken,
  runExternalCommand,
} from './util.js';
import { state } from './mpv.js';
```

`generateYoutubeSubtitles()` calls `runExternalCommand()` — pass `state.youtubeSubgenChildren` as the `childTracker` parameter.

Export `generateYoutubeSubtitles`, `resolveWhisperBinary`, and any helpers needed externally.

**Step 2: Verify it compiles**

```bash
bun build ./launcher/youtube.ts --target=bun --no-bundle --outdir=/tmp/launcher-check 2>&1
```

**Step 3: Commit**

```bash
git add launcher/youtube.ts
git commit -m "refactor(launcher): extract YouTube subtitle pipeline to launcher/youtube.ts"
```

---

### Task 9: Create launcher/jellyfin.ts

**Files:**

- Create: `launcher/jellyfin.ts`

**Step 1: Create `launcher/jellyfin.ts`**

Extract from `subminer`:

Local types (or imported from types.ts if moved in Task 6):

- `JellyfinSessionConfig` (line 944)
- `JellyfinLibraryEntry` (line 951)
- `JellyfinItemEntry` (line 957)
- `JellyfinGroupEntry` (line 964)

Functions:

- `sanitizeServerUrl()` (line 971)
- `jellyfinApiRequest()` (line 975)
- `formatJellyfinItemDisplay()` (line 1264)
- `resolveJellyfinSelection()` (line 1282)
- `runJellyfinPlayMenu()` (line 3436)

Add imports:

```typescript
import type { Args } from './types.js';
import { log, fail } from './log.js';
import { commandExists } from './util.js';
import {
  pickLibrary,
  pickItem,
  pickGroup,
  promptOptionalJellyfinSearch,
  findRofiTheme,
  showRofiFlatMenu,
} from './picker.js';
import { loadLauncherJellyfinConfig } from './config.js';
import { runAppCommandWithInherit } from './main.js';
```

Note: `runJellyfinPlayMenu()` calls `runAppCommandWithInherit()` which lives in `main.ts`. This creates a potential circular dependency (main -> jellyfin -> main). Solution: move `runAppCommandWithInherit()` to `mpv.ts` (it spawns a process and exits, which is process management). Or keep it in `util.ts`. The cleanest placement is `mpv.ts` since it calls the app binary.

Export `runJellyfinPlayMenu`, `sanitizeServerUrl`, `jellyfinApiRequest`, `resolveJellyfinSelection`.

**Step 2: Verify it compiles**

```bash
bun build ./launcher/jellyfin.ts --target=bun --no-bundle --outdir=/tmp/launcher-check 2>&1
```

**Step 3: Commit**

```bash
git add launcher/jellyfin.ts
git commit -m "refactor(launcher): extract Jellyfin integration to launcher/jellyfin.ts"
```

---

### Task 10: Create launcher/main.ts

**Files:**

- Create: `launcher/main.ts`

**Step 1: Create `launcher/main.ts`**

Extract from `subminer`:

- `checkDependencies()` (line 2369)
- `checkPickerDependencies()` (line 2716)
- `chooseTarget()` (line 3379)
- `registerCleanup()` (line 3411)
- `main()` (line 3470)

Add imports:

```typescript
import fs from 'node:fs';
import path from 'node:path';
import type { Args } from './types.js';
import { log, fail } from './log.js';
import {
  commandExists,
  isUrlTarget,
  isYoutubeTarget,
  resolvePathMaybe,
  realpathMaybe,
} from './util.js';
import {
  parseArgs,
  loadLauncherYoutubeSubgenConfig,
  loadLauncherJellyfinConfig,
  readPluginRuntimeConfig,
  usage,
} from './config.js';
import { showRofiMenu, showFzfMenu, collectVideos } from './picker.js';
import {
  state,
  startMpv,
  startOverlay,
  stopOverlay,
  launchTexthookerOnly,
  findAppBinary,
  waitForSocket,
  loadSubtitleIntoMpv,
  runAppCommandWithInherit,
} from './mpv.js';
import { generateYoutubeSubtitles } from './youtube.js';
import { runJellyfinPlayMenu } from './jellyfin.js';
```

End with:

```typescript
main().catch((error: unknown) => {
  const message = error instanceof Error ? error.message : String(error);
  fail(message);
});
```

**Step 2: Verify it compiles**

```bash
bun build ./launcher/main.ts --target=bun --no-bundle --outdir=/tmp/launcher-check 2>&1
```

**Step 3: Commit**

```bash
git add launcher/main.ts
git commit -m "refactor(launcher): extract main entrypoint to launcher/main.ts"
```

---

### Task 11: Add build-launcher Makefile target and verify

**Files:**

- Modify: `Makefile`

**Step 1: Add build-launcher target to Makefile**

Add after the existing `build-macos-unsigned` target (around line 133):

```makefile
build-launcher:
	@printf '%s\n' "[INFO] Bundling launcher script"
	@bun build ./launcher/main.ts --target=bun --packages=bundle --outfile=subminer
	@sed -i '1s|^// @bun|#!/usr/bin/env bun\n// @bun|' subminer
	@chmod +x subminer
```

Update the `install-linux` target to depend on `build-launcher`:

```makefile
install-linux: build-launcher
```

Update the `install-macos` target similarly:

```makefile
install-macos: build-launcher
```

Add `build-launcher` to the `.PHONY` list.

**Step 2: Run the build**

```bash
make build-launcher
```

Expected: produces `./subminer` with shebang on line 1.

**Step 3: Verify the bundled script works**

```bash
./subminer --help
```

Expected: prints the full help text identical to the original.

**Step 4: Commit**

```bash
git add Makefile
git commit -m "build: add build-launcher Makefile target for bundled subminer script"
```

---

### Task 12: Delete old monolithic subminer and final verification

**Files:**

- Delete: the old `subminer` file content (it's now a build artifact generated by `make build-launcher`)

**Step 1: Verify the launcher/ source is complete**

Check that every function from the original `subminer` exists in exactly one `launcher/*.ts` file:

```bash
grep -rn 'export function\|export async function' launcher/ | wc -l
```

Cross-reference with the original function count.

**Step 2: Build and run end-to-end**

```bash
make build-launcher
./subminer --help
```

Expected: identical help output.

**Step 3: Add subminer to .gitignore**

Since `subminer` is now a build artifact, add it to `.gitignore`:

```
subminer
```

Note: if the project needs the built `subminer` checked in for users who don't have bun, skip this step and keep it tracked. Discuss with the user.

**Step 4: Commit**

```bash
git add launcher/ Makefile .gitignore
git commit -m "refactor(launcher): complete split into modular launcher/ directory

- Split 3,682-line monolithic subminer script into 10 focused modules
- launcher/types.ts: shared types and constants
- launcher/log.ts: logging infrastructure
- launcher/util.ts: pure utilities and child process runner
- launcher/config.ts: config loading and arg parsing
- launcher/jimaku.ts: Jimaku API client and media parsing
- launcher/picker.ts: rofi/fzf menu UI
- launcher/mpv.ts: mpv process management and IPC
- launcher/youtube.ts: YouTube subtitle generation pipeline
- launcher/jellyfin.ts: Jellyfin API and browsing
- launcher/main.ts: orchestration entrypoint
- Add build-launcher Makefile target using bun build
- subminer is now a build artifact produced by make build-launcher"
```

---

## Cross-Cutting Concerns

### Circular Dependency Prevention

The main risk is `main.ts` <-> `jellyfin.ts`. `runJellyfinPlayMenu()` calls `runAppCommandWithInherit()`. Solution: place `runAppCommandWithInherit()` in `mpv.ts` (it's process management). Both `main.ts` and `jellyfin.ts` import from `mpv.ts` — no cycle.

### The `state` Object

`state` is defined in `mpv.ts` and exported. It's imported by:

- `main.ts` (reads `mpvProc`, calls `stopOverlay`)
- `youtube.ts` (passes `state.youtubeSubgenChildren` to `runExternalCommand`)
- `mpv.ts` itself (mutates all fields)

### `inferWhisperLanguage()` Placement

Used by `parseArgs()` in `config.ts`. It's a pure string function — place it in `util.ts` to avoid config depending on youtube.

### `JellyfinSessionConfig` Placement

Used by both `jellyfin.ts` and `picker.ts` (the pick functions accept it). Place it in `types.ts` to avoid picker depending on jellyfin.

### Import Path Convention

All imports between launcher modules use `./module.js` suffix (Bun ESM resolution).
