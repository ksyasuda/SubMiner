# Keyboard Shortcuts

All shortcuts are configurable in `config.jsonc` under `shortcuts` and `keybindings`. Set any shortcut to `null` to disable it.

## Global Shortcuts

These work system-wide regardless of which window has focus.

| Shortcut      | Action                 | Configurable                           |
| ------------- | ---------------------- | -------------------------------------- |
| `Alt+Shift+O` | Toggle visible overlay | `shortcuts.toggleVisibleOverlayGlobal` |
| `Alt+Shift+Y` | Open Yomitan settings  | Fixed (not configurable)               |

::: tip
Global shortcuts are registered with the OS. If they conflict with another application, update them in `shortcuts` config and restart SubMiner.
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

The multi-line shortcuts open a digit selector with a 3-second timeout (`shortcuts.multiCopyTimeoutMs`). Press `1`–`9` to select how many recent subtitle lines to combine.

## Overlay Controls

These control playback and subtitle display. They require overlay window focus.

| Shortcut             | Action                                              |
| -------------------- | --------------------------------------------------- |
| `Space`              | Toggle mpv pause                                    |
| `J`                  | Cycle primary subtitle track                        |
| `Shift+J`            | Cycle secondary subtitle track                      |
| `Ctrl+Alt+P`         | Open playlist browser for current directory + queue |
| `ArrowRight`         | Seek forward 5 seconds                              |
| `ArrowLeft`          | Seek backward 5 seconds                             |
| `ArrowUp`            | Seek forward 60 seconds                             |
| `ArrowDown`          | Seek backward 60 seconds                            |
| `Shift+H`            | Jump to previous subtitle                           |
| `Shift+L`            | Jump to next subtitle                               |
| `Shift+[`            | Shift subtitle delay to previous subtitle cue       |
| `Shift+]`            | Shift subtitle delay to next subtitle cue           |
| `Ctrl+Shift+H`       | Replay current subtitle (play to end, then pause)   |
| `Ctrl+Shift+L`       | Play next subtitle (jump, play to end, then pause)  |
| `Q`                  | Quit mpv                                            |
| `Ctrl+W`             | Quit mpv                                            |
| `Right-click`        | Toggle pause (outside subtitle area)                |
| `Right-click + drag` | Reposition subtitles (on subtitle area)             |
| `Ctrl/Cmd+A`         | Append clipboard video path to mpv playlist         |

These keybindings can be overridden or disabled via the `keybindings` config array. The playlist browser opens a split overlay modal with sibling video files on the left and the live mpv playlist on the right.

Mouse-hover playback behavior is configured separately from shortcuts: `subtitleStyle.autoPauseVideoOnHover` defaults to `true` (pause on subtitle hover, resume on leave).

## Subtitle & Feature Shortcuts

| Shortcut           | Action                                                   | Config key                          |
| ------------------ | -------------------------------------------------------- | ----------------------------------- |
| `Ctrl/Cmd+Shift+V` | Cycle secondary subtitle mode (hidden → visible → hover) | `shortcuts.toggleSecondarySub`      |
| `Ctrl/Cmd+Alt+A` | Open character dictionary AniList selector                 | `shortcuts.openCharacterDictionary` |
| `Ctrl/Cmd+Shift+O` | Open runtime options palette                             | `shortcuts.openRuntimeOptions`      |
| `Ctrl/Cmd+Shift+H` | Open session help modal                                  | `shortcuts.openSessionHelp`         |
| `Ctrl+Shift+J`     | Open Jimaku subtitle search modal                        | `shortcuts.openJimaku`              |
| `Ctrl+Alt+C`       | Open the manual YouTube subtitle picker                  | `keybindings`                       |
| `Ctrl+Alt+S`       | Open subtitle sync (subsync) modal                       | `shortcuts.triggerSubsync`          |
| `\`                | Toggle subtitle sidebar                                  | `subtitleSidebar.toggleKey`         |
| `` ` ``            | Toggle stats overlay                                     | `stats.toggleKey`                   |

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

When the mpv plugin is installed, all commands use a `y` chord prefix — press `y`, then the second key within 1 second.

| Chord | Action                   |
| ----- | ------------------------ |
| `y-y` | Open SubMiner menu (OSD) |
| `y-s` | Start overlay            |
| `y-S` | Stop overlay             |
| `y-t` | Toggle visible overlay   |
| `y-o` | Open Yomitan settings    |
| `y-r` | Restart overlay          |
| `y-c` | Check overlay status     |
| `y-h` | Open session help        |

When the overlay has focus, press `y` then `d` to toggle DevTools (debugging helper).

## Drag-and-Drop

| Gesture                   | Action                                           |
| ------------------------- | ------------------------------------------------ |
| Drop file(s) onto overlay | Replace current mpv playlist with dropped files  |
| `Shift` + drop file(s)    | Append all dropped files to current mpv playlist |

## Customizing Shortcuts

All `shortcuts.*` keys accept [Electron accelerator strings](https://www.electronjs.org/docs/latest/tutorial/keyboard-shortcuts), for example `"CommandOrControl+Alt+A"`. Use `null` to disable a shortcut.

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
    { "key": "Space", "command": null }, // disable default Space → pause
  ],
}
```

Both `shortcuts`, `keybindings`, and `subtitleSidebar` are [hot-reloadable](/configuration#hot-reload-behavior) — changes take effect without restarting SubMiner.
