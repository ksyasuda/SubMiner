# Keyboard Shortcuts

This page is the complete reference for every keystroke SubMiner responds to. If you are just getting started, focus on the **Mining Shortcuts** and **Overlay Controls** sections - those cover the day-to-day mining loop. The rest can wait until you need them.

A few terms used throughout:

- **Overlay** - the transparent SubMiner window that sits on top of mpv and shows the interactive subtitles. Most shortcuts only work while this window has focus (click the video once if a shortcut seems to do nothing).
- **`Ctrl/Cmd`** - use `Ctrl` on Windows/Linux and `Cmd` (⌘) on macOS. In the config file this is written as `CommandOrControl`.
- **Accelerator** - Electron's name for a shortcut string like `Alt+Shift+O`.

All shortcuts are configurable in `config.jsonc` under `shortcuts` and `keybindings`. Set any shortcut to `null` to disable it.

## App-Wide Shortcuts

| Shortcut      | Action                 | Scope                                       | Configurable                           |
| ------------- | ---------------------- | -------------------------------------------- | -------------------------------------- |
| `Alt+Shift+O` | Toggle visible overlay | Works while the overlay or mpv has focus     | `shortcuts.toggleVisibleOverlayGlobal` |
| `Alt+Shift+Y` | Open Yomitan settings  | OS-global (registered with the OS)           | Fixed (not configurable)               |

::: tip
`Alt+Shift+O` is dispatched by the overlay window and the mpv plugin, so it works from either surface without OS registration. Only `Alt+Shift+Y` is registered with the OS; if it conflicts with another application, that binding cannot be changed. All `shortcuts.*` keys hot-reload - no restart needed.
:::

## Mining Shortcuts

These work when the overlay window has focus.

| Shortcut           | Action                                          | Config key                              |
| ------------------ | ----------------------------------------------- | --------------------------------------- |
| `Ctrl/Cmd+S`       | Mine current subtitle as sentence card          | `shortcuts.mineSentence`                |
| `Ctrl/Cmd+Shift+S` | Mine multiple lines (press 1–9 to select count) | `shortcuts.mineSentenceMultiple`        |
| `Ctrl/Cmd+C`       | Copy current subtitle text                      | `shortcuts.copySubtitle`                |
| `Ctrl/Cmd+Shift+C` | Copy multiple lines (press 1–9 to select count) | `shortcuts.copySubtitleMultiple`        |
| `Ctrl/Cmd+V`       | Update last Anki card from clipboard text       | `shortcuts.updateLastCardFromClipboard` |
| `Ctrl/Cmd+G`       | Trigger field grouping (Kiku merge check)       | `shortcuts.triggerFieldGrouping`        |
| `Ctrl/Cmd+Shift+A` | Mark last card as audio card                    | `shortcuts.markAudioCard`               |

The multi-line shortcuts open a digit selector with a 3-second timeout (`shortcuts.multiCopyTimeoutMs`). Press `1`–`9` to select how many recent subtitle lines to combine. When the shortcut starts from mpv, SubMiner focuses the visible overlay for that selector instead of reserving the number keys in the mpv plugin.

## Overlay Controls

These control playback and subtitle display. They require overlay window focus.

| Shortcut             | Action                                                     |
| -------------------- | ---------------------------------------------------------- |
| `Space`              | Toggle mpv pause                                           |
| `F`                  | Toggle fullscreen                                          |
| `V`                  | Cycle primary subtitle bar mode (hidden → visible → hover) |
| `J`                  | Cycle primary subtitle track                               |
| `Shift+J`            | Cycle secondary subtitle track                             |
| `Ctrl+Alt+P`         | Open playlist browser for current directory + queue        |
| `ArrowRight`         | Seek forward 5 seconds                                     |
| `ArrowLeft`          | Seek backward 5 seconds                                    |
| `ArrowUp`            | Seek forward 60 seconds                                    |
| `ArrowDown`          | Seek backward 60 seconds                                   |
| `Shift+H`            | Jump to previous subtitle                                  |
| `Shift+L`            | Jump to next subtitle                                      |
| `Ctrl+Shift+Left`    | Shift subtitle delay to previous subtitle cue              |
| `Ctrl+Shift+Right`   | Shift subtitle delay to next subtitle cue                  |
| `z`                  | Shift subtitles 100 ms earlier                             |
| `Shift+Z`            | Delay subtitles by 100 ms                                  |
| `x`                  | Delay subtitles by 100 ms                                  |
| `Ctrl+Shift+H`       | Replay current subtitle (play to end, then pause)          |
| `Ctrl+Shift+L`       | Play next subtitle (jump, play to end, then pause)         |
| `Q`                  | Quit mpv                                                   |
| `Ctrl+W`             | Quit mpv                                                   |
| `Right-click`        | Toggle pause (outside subtitle area)                       |
| `Right-click + drag` | Reposition subtitles (on subtitle area)                    |

The mpv-command rows above (`Space`, `F`, `J`, `Shift+J`, the seek/sub-seek/sub-step/sub-delay keys, replay/play-next, and quit) are merged from the `keybindings` config array and can be remapped or disabled there. `V` and the mouse actions are built-in overlay behaviors and are not part of the `keybindings` array. The playlist browser opens a split overlay modal with sibling video files on the left and the live mpv playlist on the right.

On macOS managed playback, SubMiner disables mpv's menu-bar shortcuts so configured SubMiner shortcuts like `Cmd+Shift+O` reach the mpv plugin instead of opening native mpv menu actions.

Mouse-hover playback behavior is configured separately from shortcuts: `subtitleStyle.autoPauseVideoOnHover` defaults to `true` (pause on subtitle hover, resume on leave).

## Subtitle & Feature Shortcuts

| Shortcut           | Action                                                   | Config key                                 |
| ------------------ | -------------------------------------------------------- | ------------------------------------------ |
| `Ctrl/Cmd+Shift+V` | Cycle secondary subtitle mode (hidden → visible → hover) | `shortcuts.toggleSecondarySub`             |
| `Ctrl/Cmd+D`       | Open loaded character dictionary manager                 | `shortcuts.openCharacterDictionaryManager` |
| `Ctrl/Cmd+Shift+O` | Open runtime options palette                             | `shortcuts.openRuntimeOptions`             |
| `Ctrl/Cmd+/`       | Open session help modal                                  | `shortcuts.openSessionHelp`                |
| `Ctrl+Shift+J`     | Open Jimaku subtitle search modal                        | `shortcuts.openJimaku`                     |
| `Ctrl+Shift+T`     | Open Animetosho subtitle search modal (EN/JA tabs)       | `shortcuts.openAnimetosho`                 |
| `Ctrl/Cmd+N`       | Toggle overlay notification history panel                | `shortcuts.toggleNotificationHistory`      |
| `Ctrl+Alt+C`       | Open the manual YouTube subtitle picker                  | `keybindings`                              |
| `Ctrl+Alt+S`       | Open subtitle sync (subsync) modal                       | `shortcuts.triggerSubsync`                 |
| `Ctrl/Cmd+A`       | Append clipboard video path to mpv playlist              | `shortcuts.appendClipboardVideoToQueue`    |
| `\`                | Toggle subtitle sidebar                                  | `subtitleSidebar.toggleKey` (overlay) / `shortcuts.toggleSubtitleSidebar` (mpv session binding) |
| `` ` ``            | Toggle stats overlay                                     | `stats.toggleKey`                          |
| `W`                | Mark current video watched and advance to next in queue  | `stats.markWatchedKey`                     |

The stats toggle is handled inside the focused visible overlay window. It is configurable through the top-level `stats.toggleKey` setting and defaults to `Backquote`.

The subtitle sidebar toggle is overlay-local and only opens when SubMiner has a parsed cue list for the active subtitle source.

## Controller Shortcuts

These overlay-local shortcuts open controller utilities for the Chrome Gamepad API integration.

| Shortcut      | Action                               | Configurable                     |
| ------------- | ------------------------------------ | -------------------------------- |
| `Alt+C`       | Open controller config + remap modal | `shortcuts.openControllerSelect` |
| `Alt+Shift+C` | Open controller debug modal          | `shortcuts.openControllerDebug`  |

Controller input only drives the overlay while keyboard-only mode is enabled. The controller mapping and tuning live under the top-level `controller` config block; keyboard-only mode still works normally without a controller.

## MPV Plugin Chords

When the mpv plugin is installed, all commands use a `y` chord prefix - press `y`, then the second key (the overlay-side chord times out after 1 second; the mpv plugin uses native mpv key sequences).

| Chord | Action                                                     |
| ----- | ---------------------------------------------------------- |
| `y-y` | Open SubMiner menu (OSD)                                   |
| `y-s` | Start overlay                                              |
| `y-S` | Stop overlay                                               |
| `y-t` | Toggle visible overlay                                     |
| `v`   | Cycle primary subtitle bar mode (hidden → visible → hover) |
| `y-o` | Open Yomitan settings                                      |
| `y-r` | Restart overlay                                            |
| `y-c` | Check overlay status                                       |
| `y-h` | Open session help                                          |

The bare `v` plugin binding intentionally overrides mpv's native primary subtitle visibility toggle so it cycles the SubMiner primary subtitle bar (hidden → visible → hover) instead.

When the overlay has focus, press `y` then `d` to toggle DevTools (debugging helper).

## Drag-and-Drop

| Gesture                   | Action                                           |
| ------------------------- | ------------------------------------------------ |
| Drop file(s) onto overlay | Replace current mpv playlist with dropped files  |
| `Shift` + drop file(s)    | Append all dropped files to current mpv playlist |

## Customizing Shortcuts

All `shortcuts.*` keys accept [Electron accelerator strings](https://www.electronjs.org/docs/latest/tutorial/keyboard-shortcuts), for example `"CommandOrControl+D"`. Use `null` to disable a shortcut.

```jsonc
{
  "shortcuts": {
    "mineSentence": "CommandOrControl+S",
    "copySubtitle": "CommandOrControl+C",
    "toggleVisibleOverlayGlobal": "Alt+Shift+O",
    "openJimaku": null, // disabled
  },
}
```

The `keybindings` array overrides or extends the overlay's built-in key handling for mpv commands:

```jsonc
{
  "keybindings": [
    { "key": "f", "command": ["cycle", "fullscreen"] },
    { "key": "m", "command": ["cycle", "mute"] },
    { "key": "MBTN_BACK", "command": ["sub-seek", -1] },
    { "key": "MBTN_FORWARD", "command": ["sub-seek", 1] },
    { "key": "Space", "command": null }, // disable default Space → pause
  ],
}
```

Mouse keybinding names are `MBTN_LEFT`, `MBTN_MID`, `MBTN_RIGHT`, `MBTN_BACK`, and `MBTN_FORWARD`.

Both `shortcuts`, `keybindings`, and `subtitleSidebar` are [hot-reloadable](/configuration#hot-reload-behavior) - changes take effect without restarting SubMiner.
