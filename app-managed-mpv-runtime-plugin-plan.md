# App-Managed mpv Runtime Plugin Plan

## Summary

Shift SubMiner’s default mpv integration from “install and maintain a plugin in the user’s mpv scripts directory” to “inject the bundled plugin at launch time when SubMiner controls mpv.” Existing user-installed plugins remain supported and take precedence. If SubMiner detects an installed plugin, it will not inject the bundled runtime plugin and will show/log a helpful message with the detected path.

This keeps existing users working, avoids double-loading, and removes the need to update user plugin files for new installs.

## Goals

- New users should not need to install the mpv plugin manually for SubMiner-managed mpv launches.
- Existing users with an installed mpv plugin should keep working without behavior changes.
- SubMiner should avoid loading both the installed plugin and bundled runtime plugin in the same mpv session.
- SubMiner should tell users exactly where an installed plugin was detected so they can remove it if they want app-managed runtime loading.
- Windows app-managed mpv launch should support this.
- macOS/Linux launcher-managed mpv launch should support this.

## Non-Goals

- Do not remove the current first-run plugin installer yet.
- Do not auto-delete, rename, or overwrite user-installed mpv plugin files.
- Do not force `--load-scripts=no` by default.
- Do not make the runtime plugin override an installed user plugin.
- Do not solve arbitrary already-running mpv sessions without IPC or prior SubMiner launch control.

## Runtime Policy

When SubMiner is about to launch mpv:

1. Detect whether a SubMiner plugin is already installed in mpv’s user scripts directory.
2. If an installed plugin is found:
   - Do not pass `--script=<bundled plugin path>`.
   - Continue passing SubMiner script options needed by the installed plugin.
   - Show/log a warning with the detected installed plugin path and detected version if available.
   - Defer to the installed plugin for that mpv session.
3. If no installed plugin is found:
   - Pass `--script=<bundled runtime plugin path>`.
   - Pass runtime script options for binary path, socket path, log level, and any existing launch metadata.
   - Use the plugin bundled with the currently running SubMiner app.

## Plugin Detection

Add a shared detector module, likely under `src/shared/` or `src/core/services/`, usable by both Electron main runtime and launcher code.

Suggested type:

```ts
export type InstalledMpvPluginSource =
  | 'default-config'
  | 'xdg-config'
  | 'portable-config'
  | 'legacy-file';

export interface InstalledMpvPluginDetection {
  installed: boolean;
  path: string | null;
  version: string | null;
  source: InstalledMpvPluginSource | null;
  message: string | null;
}
```

Detection candidates:

Windows:

```text
%APPDATA%\mpv\scripts\subminer\main.lua
%APPDATA%\mpv\scripts\subminer.lua
%APPDATA%\mpv\scripts\subminer-loader.lua
```

If an mpv executable path is known, also check:

```text
<mpv.exe directory>\portable_config\scripts\subminer\main.lua
<mpv.exe directory>\portable_config\scripts\subminer.lua
<mpv.exe directory>\portable_config\scripts\subminer-loader.lua
```

macOS/Linux:

```text
$XDG_CONFIG_HOME/mpv/scripts/subminer/main.lua
$XDG_CONFIG_HOME/mpv/scripts/subminer.lua
$XDG_CONFIG_HOME/mpv/scripts/subminer-loader.lua
~/.config/mpv/scripts/subminer/main.lua
~/.config/mpv/scripts/subminer.lua
~/.config/mpv/scripts/subminer-loader.lua
```

Rules:

- Return the first existing candidate using the same precedence mpv would most likely use.
- Prefer `subminer/main.lua` over legacy single-file loaders within the same config root.
- Include the absolute detected path in the warning message.
- If multiple plugin candidates are found, log all candidates at debug level, but use the first one as the active detection result.

## Plugin Versioning

Add version metadata to the bundled Lua plugin.

New file:

```text
plugin/subminer/version.lua
```

Contents shape:

```lua
return {
  name = "SubMiner mpv plugin",
  version = "0.12.0",
}
```

Update `plugin/subminer/init.lua` or `bootstrap.lua` to load this version and expose it in logs.

Version detection should try:

1. Read sibling `version.lua` next to detected `main.lua`.
2. Parse a simple `version = "..."` value.
3. If unavailable, return `version: null`.

For legacy single-file installs, return `version: null` unless a parseable marker is present.

Important: old installed plugins will not have version metadata, so `null` is expected and should be messaged as “unknown or legacy version,” not treated as an error.

## Bundled Runtime Plugin Path Resolution

Add a helper for resolving the packaged plugin entrypoint/directory.

For Electron main process:

- Reuse `resolvePackagedFirstRunPluginAssets()` or extract a narrower helper from `src/main/runtime/first-run-setup-plugin.ts`.
- Preferred runtime script path should be the plugin directory if mpv supports directory script loading:
  ```text
  <resourcesPath>/plugin/subminer
  ```
- Fallback:
  ```text
  <resourcesPath>/plugin/subminer/main.lua
  ```

For launcher:

- Accept a resolved runtime plugin path where possible.
- If only `appPath` is available, resolve relative packaged resources using existing app path conventions.
- If runtime plugin path cannot be resolved and no installed plugin exists, fail with a clear message telling the user the packaged mpv plugin assets were not found.

## Windows Implementation

Update `src/main/runtime/windows-mpv-launch.ts`.

Current behavior already supports `pluginEntrypointPath`.

Change launch preparation so the caller passes either:

- `pluginEntrypointPath` when no installed plugin is detected.
- `undefined` when an installed plugin is detected.

Add or update a dependency boundary so detection is testable:

```ts
detectInstalledMpvPlugin?: () => InstalledMpvPluginDetection;
notifyInstalledPluginDetected?: (detection: InstalledMpvPluginDetection) => void;
```

Behavior:

- Resolve `mpv.exe` first.
- Run installed plugin detection, including portable config checks when `mpv.exe` path is known.
- If installed plugin exists:
  - show a non-fatal notification or warning dialog/log entry:
    ```text
    SubMiner detected an installed mpv plugin at:
    <path>

    This mpv session will use the installed plugin. Remove it to use SubMiner's bundled runtime plugin automatically.
    ```
  - launch without `--script=<runtime plugin>`.
- If no installed plugin exists:
  - launch with `--script=<bundled runtime plugin path>`.

Keep current script options:

```text
subminer-binary_path=<process.execPath>
subminer-socket_path=<socket>
```

## macOS/Linux Launcher Implementation

Update `launcher/mpv.ts`.

Affected functions:

- `startMpv()`
- `launchMpvIdleDetached()`

Current launcher passes `--script-opts=...` but does not explicitly pass `--script=...`.

New behavior:

- Detect installed plugin before building mpv args.
- If installed plugin exists:
  - do not add `--script=<bundled runtime plugin path>`.
  - keep passing SubMiner script opts.
  - log warning with detected path/version.
- If no installed plugin exists:
  - add `--script=<bundled runtime plugin path>`.
  - keep passing SubMiner script opts.

Prefer repeated `--script-opt=key=value` in new code where practical, but do not require a full conversion if it risks broad parser churn. If keeping the existing `--script-opts=...` helper, preserve current escaping behavior.

## User Messaging

Add user-facing copy in two places:

1. Runtime warning/log when launching mpv and installed plugin is detected.
2. First-run/setup UI copy update, if applicable, to explain that plugin installation is now optional for normal SubMiner-managed launch.

Suggested warning:

```text
SubMiner detected an installed mpv plugin at:
<path>

This session will use the installed plugin. Remove that plugin to use the bundled runtime plugin that ships with this SubMiner version.
```

If version is known:

```text
Detected plugin version: <version>
Bundled plugin version: <version>
```

If version is unknown:

```text
Detected plugin version: unknown or legacy
```

## Compatibility Behavior

Existing installed plugin:

- Used as-is.
- Receives script options from the launch command.
- Not modified.

No installed plugin:

- Bundled runtime plugin is injected.
- No files are written to mpv config.

Old installed plugin with no version:

- Used as-is.
- Warning says version is unknown or legacy.
- User can remove it to switch to runtime injection.

Portable Windows mpv:

- If `mpv.exe` path is known and `portable_config` exists beside it, detect installed plugin there before checking `%APPDATA%`.
- If `mpv.exe` path is unknown, only `%APPDATA%` detection is possible.

## Public Interface / Type Changes

Add shared detection exports:

```ts
export interface InstalledMpvPluginDetection {
  installed: boolean;
  path: string | null;
  version: string | null;
  source: InstalledMpvPluginSource | null;
  message: string | null;
}

export function detectInstalledMpvPlugin(options: {
  platform: NodeJS.Platform;
  homeDir: string;
  xdgConfigHome?: string;
  appDataDir?: string;
  mpvExecutablePath?: string;
  existsSync?: (path: string) => boolean;
  readFileSync?: (path: string, encoding: BufferEncoding) => string;
}): InstalledMpvPluginDetection;
```

Add runtime plugin path helper:

```ts
export function resolvePackagedRuntimePluginPath(options: {
  dirname: string;
  appPath: string;
  resourcesPath: string;
  existsSync?: (path: string) => boolean;
  joinPath?: (...parts: string[]) => string;
}): string | null;
```

Lua plugin version module:

```text
plugin/subminer/version.lua
```

## Tests

Add detector unit tests covering:

- Windows `%APPDATA%\mpv\scripts\subminer\main.lua`.
- Windows legacy `subminer.lua`.
- Windows `portable_config` next to known `mpv.exe`.
- Windows portable config precedence over `%APPDATA%`.
- Linux `$XDG_CONFIG_HOME/mpv/scripts/subminer/main.lua`.
- Linux fallback to `~/.config/mpv/scripts/subminer/main.lua`.
- macOS `~/.config/mpv/scripts/subminer/main.lua`.
- No installed plugin.
- Version parsed from `version.lua`.
- Missing version returns `null`.

Update Windows launch tests:

- Installed plugin detected means no `--script=<runtime plugin>`.
- Installed plugin detected still includes `--script-opts` or equivalent script options.
- No installed plugin means `--script=<runtime plugin>`.
- Warning callback receives detected path.

Update launcher tests:

- `startMpv()` includes `--script=<runtime plugin>` when no installed plugin exists.
- `startMpv()` omits `--script=<runtime plugin>` when installed plugin exists.
- `launchMpvIdleDetached()` follows the same policy.
- Warning/log path includes detected installed plugin path.

Lua/plugin tests:

- Version module loads successfully.
- Bootstrap logs or exposes version without breaking existing plugin startup.

## Verification

Minimum targeted verification:

```text
bun test src/shared/<new-detector>.test.ts
bun test src/main/runtime/windows-mpv-launch.test.ts
bun test launcher/mpv.test.ts launcher/main.test.ts
lua scripts/test-plugin-lua-compat.lua
bun run typecheck
```

Broader handoff verification if implementation touches setup or launch wiring substantially:

```text
bun run test:launcher
bun run test:fast
```

## Rollout Notes

- Keep first-run plugin install available for one release as an optional compatibility path.
- Update docs/release notes to explain:
  - SubMiner-managed mpv launch no longer requires plugin installation.
  - Existing installed plugin takes precedence.
  - Remove the installed plugin to use the bundled runtime plugin.
- Add a `changes/*.md` fragment because this is user-visible behavior.

## Assumptions and Defaults

- Default behavior preserves user-installed plugins and does not override them.
- Runtime injection is only guaranteed when SubMiner launches mpv or has IPC access to an existing mpv.
- No automatic deletion or migration of installed plugin files.
- Windows portable mpv detection is implemented only when SubMiner knows the `mpv.exe` path.
- macOS/Linux use default mpv config paths unless future work adds custom config-dir discovery.
- Use the bundled plugin from `process.resourcesPath/plugin/subminer` as the runtime source of truth.
