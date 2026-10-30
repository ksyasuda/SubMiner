# Plugins

Plugins are JavaScript files that SubMiner loads at startup. A plugin can react to playback, run code on a timer, add actions you bind to keys, run programs, and show results in an overlay panel, all without rebuilding SubMiner.

::: warning Only run plugins you trust
Plugins run inside SubMiner with the same access to your files and system as SubMiner itself. There is no sandbox.
:::

These are SubMiner app plugins. The Lua script that runs inside mpv is the separate [mpv plugin](/mpv-plugin).

## Install a plugin

1. Put the plugin in its own folder in `plugins/` next to your [config file](/configuration#configuration-file), for example `~/.config/SubMiner/plugins/ai-explain/`. The folder name is the plugin id.
2. Enable it in `config.jsonc`. `settings` is passed to the plugin as-is:

   ```jsonc
   {
     "plugins": {
       "ai-explain": { "enabled": true, "settings": { "reasoningEffort": "medium" } },
     },
   }
   ```

3. Bind its actions to keys with the `__plugin-action:<plugin id>/<action>` command:

   ```jsonc
   {
     "keybindings": [
       { "key": "Ctrl+KeyE", "command": ["__plugin-action:ai-explain/explain"] },
       { "key": "Ctrl+KeyT", "command": ["__plugin-action:ai-explain/translate"] },
     ],
   }
   ```

   To run an action from a script or a window-manager shortcut, use `subminer app --plugin-action ai-explain/explain`. It goes to the running SubMiner, or starts one if none is running.

4. Restart SubMiner. Plugin changes apply on the next start; keybinding changes apply immediately.

Load errors and plugin log output go to the app log with a `plugin:<id>` scope.

## Example: explain the current line with Codex

The repository ships [`assets/plugins/ai-explain`](https://github.com/ksyasuda/SubMiner/tree/main/assets/plugins/ai-explain). It sends the current line, plus the lines around it, to [`codex exec`](https://github.com/openai/codex) and shows the answer in a panel. It needs a logged-in Codex CLI.

| Action      | What it does                                                    |
| ----------- | --------------------------------------------------------------- |
| `explain`   | Translation plus a breakdown of the grammar and vocabulary      |
| `translate` | Natural English translation, with a note only when it's unclear |

| Setting           | Default   | What it does                                                      |
| ----------------- | --------- | ----------------------------------------------------------------- |
| `codexPath`       | `"codex"` | Codex binary. Use a full path if SubMiner can't find it on `PATH` |
| `model`           | `""`      | Model passed to `codex exec -m`; empty uses your Codex config     |
| `reasoningEffort` | `"low"`   | Codex reasoning effort; empty uses your Codex config              |
| `before`, `after` | `2`, `1`  | Lines of context sent before and after the current line           |
| `pause`           | `true`    | Pause playback when the panel opens                               |
| `timeoutMs`       | `120000`  | Give up after this long                                           |
| `prompts`         | `{}`      | Replace an action's instructions, e.g. `{ "translate": "..." }`   |

Press `Esc` to close the panel. Closing it, or starting another request, stops the running Codex process.

## Example: sync Anki after mining

[`assets/plugins/anki-sync`](https://github.com/ksyasuda/SubMiner/tree/main/assets/plugins/anki-sync) runs AnkiConnect's `sync` after you mine cards, so they reach AnkiWeb without opening Anki. It waits until you stop mining for a while, so a burst of cards becomes one sync. It also adds a `sync` action you can bind to a key.

| Setting     | Default | What it does                                            |
| ----------- | ------- | ------------------------------------------------------- |
| `delayMs`   | `30000` | Sync after no card has been mined for this long         |
| `syncOnEnd` | `true`  | Sync right away when playback ends or mpv quits         |
| `timeoutMs` | `60000` | Give up on a sync after this long                       |
| `notify`    | `true`  | Show a notification after each sync. Failures always do |

Anki must be running and logged in to AnkiWeb. Cards mined from the stats dashboard don't trigger a sync.

## Example: tag cards with the episode

[`assets/plugins/episode-tags`](https://github.com/ksyasuda/SubMiner/tree/main/assets/plugins/episode-tags) tags each mined card with the show and episode it came from, such as `show::Sousou_no_Frieren::E07`. It reads them from the file name (`[Group] Show - 07 (1080p).mkv`, `Show.S01E07.mkv`, `Show 第7話.mkv`), or from the title mpv shows for streams. It uses the source captured when mining started, so delayed media and field grouping into an older note keep the correct episode. Notes with unknown provenance, including cards discovered by polling, are left alone. Spaces become `_`.

| Setting  | Default                  | What it does                                    |
| -------- | ------------------------ | ----------------------------------------------- |
| `tags`   | `["show::{show}::{ep}"]` | Tags to add, one template each                  |
| `notify` | `false`                  | Show a notification with the tags for each card |

Templates can use `{show}`, `{season}`, `{episode}`, and `{ep}` (`S01E07`, or `E07` when there is no season). When a placeholder has no value, the tag stops before that `::` level, so a movie gets `show::<title>`. The parsing is a guess from the name: a title ending in a number, like `Mob Psycho 100`, reads that number as the episode.

## Example: open Jimaku when a file has no Japanese subtitles

[`assets/plugins/auto-jimaku`](https://github.com/ksyasuda/SubMiner/tree/main/assets/plugins/auto-jimaku) opens the Jimaku search a few seconds after a file starts if none of its subtitle tracks is Japanese. A track counts as Japanese when its language code is in `languages`, or its title contains "Japanese" or 日本語.

| Setting     | Default               | What it does                                                    |
| ----------- | --------------------- | --------------------------------------------------------------- |
| `action`    | `"openJimaku"`        | Built-in action to run, e.g. `openTsukihime`                    |
| `languages` | `["ja", "jp", "jpn"]` | Track language codes that count as Japanese                     |
| `delayMs`   | `3000`                | Wait this long after the file starts, so external subs load     |
| `streams`   | `false`               | Also check `http(s)` streams, which Jimaku usually can't search |

## Write a plugin

A plugin folder needs a `plugin.json`:

```json
{ "name": "My Plugin", "version": "1.0.0", "apiVersion": 1, "main": "main.js" }
```

`apiVersion` declares the required SubMiner plugin API version. The current version is `1`; omitted values also mean v1 for existing plugins. An unsupported version is rejected before the plugin's entry module runs. `version` is the plugin's own version and does not control compatibility.

`main` points to a CommonJS module (`.js`) or an ES module (`.mjs`) that exports `activate(ctx)`. Activation may be asynchronous and has a 10-second timeout. A plugin's actions wait for its own activation, so a slow plugin does not hold up other plugins. If activation fails or times out, its registrations are removed and its signal aborts. Synchronous code still shares the app's main thread.

Activation may return a cleanup function, or the module can export `deactivate()`. On quit, SubMiner aborts `ctx.signal` first and waits up to 2 seconds in total for loading and deactivation. A plugin still reading its files is skipped. Everything a plugin registers through `ctx` is cleaned up automatically. A late activation cannot leave new listeners, timers, or actions behind, and its returned cleanup still runs once. Deactivation can call Anki, mpv, built-in actions, and notifications until disposal. After disposal, those calls throw or reject, including work that resumes after an activation or deactivation timeout.

```js
function activate(ctx) {
  ctx.media.on('change', (media) => {
    if (media) ctx.log.info(`Now playing ${media.path}`);
  });
  ctx.every(60_000, () => ctx.log.info('Still watching'));
  ctx.actions.register('hello', () => {
    const line = ctx.subtitles.current();
    ctx.ui.panel({ title: 'Hello', markdown: line ? `> ${line.text}` : '_No subtitle_' });
  });
}

module.exports = { activate };
```

An error thrown by a plugin is logged and does not crash SubMiner. A plugin that blocks with a long synchronous loop still freezes the app, so keep slow work asynchronous.

### Types and cleanup

The public types live in [`src/plugin-api.ts`](https://github.com/ksyasuda/SubMiner/blob/main/src/plugin-api.ts). A build emits `dist/plugin-api.d.ts`; copy that file into your plugin as `subminer.d.ts`, or copy the source file with that name. It has no imports from SubMiner internals. Electron's types must be available to your editor because `ui.openWindow` returns a `BrowserWindow`.

JavaScript plugins can use the types through JSDoc:

```js
/** @param {import('./subminer').PluginContext} ctx */
function activate(ctx) {
  const { ipcMain } = require('electron');
  ipcMain.handle(`${ctx.id}:info`, () => ({ media: ctx.media.current() }));
  ctx.onDispose(() => ipcMain.removeHandler(`${ctx.id}:info`));
}

module.exports = { activate };
```

Use `ctx.signal` with `fetch` or other cancellable work. It aborts when unloading starts, before `deactivate` runs. `ctx.onDispose(fn)` registers synchronous cleanup for resources outside `ctx`, such as file watchers and Electron IPC handlers. The returned function runs the cleanup early, at most once. Registering cleanup after unloading starts runs it immediately. Timers and listeners registered after that point are removed immediately too. Use the returned activation cleanup or `deactivate` for asynchronous teardown, within the shared quit timeout.

### `ctx` reference

| Member                                 | What it does                                                                                                              |
| -------------------------------------- | ------------------------------------------------------------------------------------------------------------------------- |
| `id`, `dir`, `settings`                | Plugin id, plugin folder path, and `plugins.<id>.settings`                                                                |
| `apiVersion`                           | The host's plugin API version, currently `1`                                                                              |
| `signal`, `onDispose(fn)`              | Lifetime cancellation and cleanup for resources created outside `ctx`                                                     |
| `dataDir`                              | `plugin-data/<id>/` next to your config file, for caches and state. Exists before `activate` runs                         |
| `log.debug/info/warn/error(message)`   | Write to the app log                                                                                                      |
| `every(ms, fn)`, `after(ms, fn)`       | Repeating and one-shot timers. Return a function that cancels them                                                        |
| `mpv.on(event, fn)`                    | Listen to mpv events (below). Returns an unsubscribe function                                                             |
| `mpv.command(args)`                    | Send an mpv command, e.g. `['set', 'pause', 'yes']`                                                                       |
| `mpv.request(args)`                    | Await an mpv command result; rejects on command error, disconnect, or timeout                                             |
| `mpv.getProperty(name)`                | Read an mpv property (Promise)                                                                                            |
| `mpv.connected`                        | Whether mpv is connected                                                                                                  |
| `mpv.ensureRunning()`                  | Connect to mpv, starting an idle mpv if none is running. Resolves `false` if that fails                                   |
| `media.current()`                      | Cached `{ playbackId, path, title }`, or `null` when there is no connected media                                          |
| `media.on('change', fn, options)`      | Subscribe to playback snapshots; immediately delivers the current snapshot unless `emitCurrent: false`                    |
| `anki.on(event, fn)`                   | Listen to Anki events (below). Returns an unsubscribe function                                                            |
| `anki.invoke(action, params, options)` | Send an AnkiConnect action to `ankiConnect.url`, e.g. `('sync')`. Resolves its result (below)                             |
| `subtitles.current()`                  | The displayed line as `{ text, startTime, endTime }`, or `null`                                                           |
| `subtitles.context({ before, after })` | `{ current, before, after }`. Neighbors come from the loaded subtitle file, or recent lines for streams                   |
| `actions.register(name, fn)`           | Add an action, bindable as `__plugin-action:<id>/<name>`                                                                  |
| `actions.run(actionId, payload)`       | Run a built-in action, the same ids keybindings use, e.g. `openJimaku` or `mineSentence`                                  |
| `notify(message, { variant })`         | Overlay notification. `variant` is `info`, `warning`, or `error`                                                          |
| `ui.panel({ title, markdown })`        | Open the plugin panel. Returns `{ update(content), close(), closed, signal }`                                             |
| `ui.openWindow({ html, preload })`     | Open a window from the plugin folder (below). Returns the Electron `BrowserWindow`                                        |
| `exec(command, args, options)`         | Run a program without a shell (Windows `.cmd`/`.bat` shims go through `cmd.exe`). Resolves `{ exitCode, stdout, stderr }` |
| `unstable`                             | SubMiner internals. These change between releases without notice                                                          |

A panel's `signal` aborts when the user closes it or another panel replaces it. Pass it to `exec` to stop work nobody will see. Only one panel is open at a time. Panel markdown supports headings, lists, quotes, code, bold, and italics; HTML is shown as plain text.

`exec` options are `stdin`, `cwd`, `env`, `timeoutMs`, and `signal`. It resolves for any exit code and rejects when the program can't start, times out, or is aborted.

`ui.openWindow` takes `html` and an optional `preload`, both relative to the plugin folder, plus `title`, `width`, and `height`. Opening the same `html` again focuses the open window. The page runs with context isolation and no Node access. The preload is sandboxed too: it can `require('electron')` but nothing else, so bundle anything more into it. Expose what the page needs from the preload with `contextBridge`, and handle its messages with `ipcMain` from `require('electron')` in the plugin. Prefix channel names with the plugin id. Plugin windows close when SubMiner quits.

`media` snapshots are immutable. `playbackId` changes on a new load, including reopening the same path. A title may arrive later with the same id. Disconnecting or ending playback delivers `null`. Use the id to discard stale asynchronous results. Snapshots describe cached playback state; they do not promise that external subtitle discovery has finished.

`mpv.request` accepts string, number, and boolean arguments, e.g. `await ctx.mpv.request(['set_property', 'pause', true])`. It returns mpv's result data and rejects after 4 seconds. A timeout does not cancel the command in mpv, so a slow command may still complete. `mpv.command` remains available for commands whose result you do not need.

mpv events: `connection-change`, `subtitle-change`, `subtitle-timing`, `time-pos-change`, `duration-change`, `pause-change`, `fullscreen-change`, `media-path-change`, `media-title-change`, `subtitle-track-change`, `subtitle-track-list-change`, `secondary-subtitle-change`, `secondary-subtitle-track-change`, `secondary-subtitle-delay-change`, `secondary-subtitle-visibility`, `client-message`, `file-loaded`, and `end-file`. `file-loaded` has no payload. `end-file` gives `{ reason, playlistEntryId, error }`, where `reason` is `eof`, `stop`, `quit`, `error`, `redirect`, or `unknown`.

Anki events: `card-mined` gives `{ noteIds, notes }` once SubMiner has finished filling in a mined card (sentence, audio, image, and field grouping). For YouTube cards whose media waits on the download, it fires once that media is added or the download fails. With `ankiConnect.behavior.autoUpdateNewCards` off, it fires as soon as SubMiner sees the card. If field grouping merged the card into an older one, `noteIds` contains the older card's id.

Each `notes` entry has `{ originalNoteId, noteId, provenance }`. `originalNoteId` identifies the note before merges; `noteId` is the note that kept its content. `provenance` is `{ minedAt, media, subtitle }`, captured before asynchronous creation or the AnkiConnect proxy's upstream request. `minedAt` is Unix time in milliseconds. `media` uses the snapshot format above, and `subtitle` uses `{ text, startTime, endTime }`. Either may be `null`. Polling cannot establish the original source, so its provenance is `null` rather than the currently playing file.

Several original notes can merge into one kept note. `notes` preserves each source, while `noteIds` contains unique kept ids. Existing listeners using only `noteIds` continue to work. Cards mined from the stats dashboard do not emit this event.

`anki.invoke` takes `{ timeoutMs }` (default 10 seconds), rejects on an Anki error, and never retries, so an action like `sync` or `addNote` runs at most once.
