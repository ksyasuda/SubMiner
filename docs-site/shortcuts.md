# Keyboard shortcuts

Every key SubMiner responds to, with its default binding. `Ctrl/Cmd` means `Ctrl` on Windows and Linux and `Cmd` on macOS (`CommandOrControl` in config).

Shortcuts work when the overlay has focus. With the [mpv plugin](/mpv-plugin), `shortcuts.*` and `keybindings` entries also work while mpv has focus. If a key does nothing, click the video once. Set any shortcut to `null` to disable it. Changes to `shortcuts`, `keybindings`, and `subtitleSidebar` apply without a restart.

## Global

| Shortcut      | Action                          | Config key                             |
| ------------- | ------------------------------- | -------------------------------------- |
| `Alt+Shift+O` | Toggle visible overlay          | `shortcuts.toggleVisibleOverlayGlobal` |
| `Alt+Shift+Y` | Open active dictionary settings | Fixed                                  |

`Alt+Shift+Y` opens Yomitan or Hachidori settings, whichever backend is running. It is registered with the OS and works from any app. If another app already uses it, SubMiner cannot take it and you cannot rebind it.

## Mining

| Shortcut           | Action                                              | Config key                              |
| ------------------ | --------------------------------------------------- | --------------------------------------- |
| `Ctrl/Cmd+S`       | Mine current line as a sentence card                | `shortcuts.mineSentence`                |
| `Ctrl/Cmd+Shift+S` | Mine several lines as one sentence card             | `shortcuts.mineSentenceMultiple`        |
| `Ctrl/Cmd+C`       | Copy current line                                   | `shortcuts.copySubtitle`                |
| `Ctrl/Cmd+Shift+C` | Copy several lines                                  | `shortcuts.copySubtitleMultiple`        |
| `Ctrl/Cmd+V`       | Update last-added card from the clipboard           | `shortcuts.updateLastCardFromClipboard` |
| `Ctrl/Cmd+G`       | Run the field grouping check on the last-added card | `shortcuts.triggerFieldGrouping`        |
| `Ctrl/Cmd+Shift+A` | Mark last-added card as an audio card               | `shortcuts.markAudioCard`               |

After a multi-line shortcut, press `1` to `9` for how many lines to combine, counting back from and including the current line. The prompt closes after `shortcuts.multiCopyTimeoutMs`.

When text is selected in the [subtitle sidebar](/subtitle-sidebar), `Ctrl/Cmd+C` copies that selection instead.

## Playback

These are the default `keybindings` entries. Remap or disable them in the `keybindings` array.

| Shortcut           | Action                                   |
| ------------------ | ---------------------------------------- |
| `Space`            | Pause or resume                          |
| `F`                | Toggle fullscreen                        |
| `J`                | Cycle primary subtitle track             |
| `Shift+J`          | Cycle secondary subtitle track           |
| `ArrowRight`       | Seek forward 5 seconds                   |
| `ArrowLeft`        | Seek back 5 seconds                      |
| `ArrowUp`          | Seek forward 60 seconds                  |
| `ArrowDown`        | Seek back 60 seconds                     |
| `Shift+H`          | Jump to previous subtitle                |
| `Shift+L`          | Jump to next subtitle                    |
| `Ctrl+Shift+Left`  | Shift subtitle delay to the previous cue |
| `Ctrl+Shift+Right` | Shift subtitle delay to the next cue     |
| `Z`                | Subtitle delay -100 ms                   |
| `Shift+Z`          | Subtitle delay +100 ms                   |
| `X`                | Subtitle delay +100 ms                   |
| `Ctrl+Shift+H`     | Replay current subtitle, then pause      |
| `Ctrl+Shift+L`     | Play next subtitle, then pause           |
| `Ctrl+Alt+P`       | Open playlist browser                    |
| `Ctrl+Alt+C`       | Open YouTube subtitle picker             |
| `Q`                | Quit mpv                                 |
| `Ctrl+W`           | Quit mpv                                 |

Built into the overlay, not configurable:

| Input                     | Action                                             |
| ------------------------- | -------------------------------------------------- |
| `V`                       | Cycle primary subtitle bar: hidden, visible, hover |
| Right-click               | Pause or resume (outside the subtitle area)        |
| Right-click + drag        | Move the subtitles                                 |
| Drop files on the overlay | Replace the mpv playlist                           |
| `Shift` + drop files      | Append to the mpv playlist                         |

## Overlay features

| Shortcut           | Action                                                 | Config key                                 |
| ------------------ | ------------------------------------------------------ | ------------------------------------------ |
| `Ctrl/Cmd+Shift+V` | Cycle secondary subtitle bar: hidden, visible, hover   | `shortcuts.toggleSecondarySub`             |
| `Ctrl/Cmd+Shift+O` | Open runtime options                                   | `shortcuts.openRuntimeOptions`             |
| `Ctrl/Cmd+/`       | Open session help                                      | `shortcuts.openSessionHelp`                |
| `Ctrl/Cmd+D`       | Open character dictionary manager                      | `shortcuts.openCharacterDictionaryManager` |
| `Ctrl/Cmd+N`       | Toggle notification history                            | `shortcuts.toggleNotificationHistory`      |
| `Ctrl/Cmd+A`       | Append the video path on the clipboard to the playlist | `shortcuts.appendClipboardVideoToQueue`    |
| `Ctrl+Shift+J`     | Open Jimaku subtitle search                            | `shortcuts.openJimaku`                     |
| `Ctrl+Shift+T`     | Open TsukiHime subtitle search                         | `shortcuts.openTsukihime`                  |
| `Ctrl+Shift+G`     | Open Japanese subtitle generation                      | `shortcuts.openSubtitleGeneration`         |
| `Ctrl+Alt+S`       | Open subtitle sync (subsync)                           | `shortcuts.triggerSubsync`                 |
| `g` then `s`       | Pick primary and secondary subtitles (when enabled)    | `shortcuts.openSubtitleSelection`          |
| `\`                | Toggle subtitle sidebar                                | `subtitleSidebar.toggleKey`                |
| `` ` ``            | Toggle stats overlay                                   | `stats.toggleKey`                          |
| `W`                | Mark video watched and play the next one in the queue  | `stats.markWatchedKey`                     |
| `Alt+C`            | Open controller setup and remapping                    | `shortcuts.openControllerSelect`           |
| `Alt+Shift+C`      | Open controller debug view                             | `shortcuts.openControllerDebug`            |

The sidebar key has a separate mpv-side binding, `shortcuts.toggleSubtitleSidebar`. The sidebar only opens when SubMiner has parsed the active subtitle file. In the sidebar, `Enter` seeks to the focused line.

The subtitle picker (`g` then `s`) is off until you turn it on in **Settings, Behavior, Subtitle Selection**. Press the second key within one second. If `g` already has an action in SubMiner or mpv, the sequence is disabled and a warning is shown. See [subtitle selection](/configuration#subtitle-selection).

## mpv plugin keys

Press `y`, then the second key.

| Keys  | Action                          |
| ----- | ------------------------------- |
| `y-y` | Open the SubMiner menu          |
| `y-s` | Start the overlay               |
| `y-S` | Stop the overlay                |
| `y-t` | Toggle the visible overlay      |
| `y-o` | Open active dictionary settings |
| `y-r` | Restart the overlay             |
| `y-c` | Show overlay status             |
| `y-h` | Open session help               |
| `v`   | Cycle primary subtitle bar      |

The plugin's `v` replaces mpv's own subtitle visibility toggle. When the overlay has focus, `y` then `d` toggles DevTools.

## Customizing

`shortcuts.*` values are [Electron accelerator strings](https://www.electronjs.org/docs/latest/tutorial/keyboard-shortcuts).

```jsonc
{
  "shortcuts": {
    "mineSentence": "CommandOrControl+S",
    "openJimaku": null, // disabled
  },
}
```

`keybindings` entries map a key to an mpv command. They are merged with the defaults above. Set `command` to `null` to disable a default.

```jsonc
{
  "keybindings": [
    { "key": "m", "command": ["cycle", "mute"] },
    { "key": "MBTN_BACK", "command": ["sub-seek", -1] },
    { "key": "Space", "command": null },
  ],
}
```

Mouse button names are `MBTN_LEFT`, `MBTN_MID`, `MBTN_RIGHT`, `MBTN_BACK`, and `MBTN_FORWARD`. See [keybindings](/configuration#keybindings) and [shortcuts configuration](/configuration#shortcuts-configuration) in the config reference.

## Automatic mpv bindings

The overlay reads single-key bindings from the running mpv (`input.conf`, mpv defaults, and scripts). If SubMiner does not handle a key, it passes it to mpv. SubMiner shortcuts and `keybindings` entries win, including ones set to `null`. Keys are not forwarded while you type in a text field, use an overlay menu, or have a Yomitan popup open.

Mouse buttons, keypad and media keys, and key sequences are not imported. Bindings imported this way do not appear in session help. If you add an mpv binding while SubMiner runs, refocus the overlay to pick it up.
