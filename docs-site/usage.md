# Usage

This page covers everyday use: starting playback, working with the overlay, and the commands you will reach for most. For every `subminer` subcommand and flag, see [Launcher script](/launcher-script).

## Play a video

```bash
subminer video.mkv
```

On Windows, double-click the **SubMiner mpv** shortcut or drag a video onto it.

SubMiner starts mpv, connects to it, and opens the overlay. Subtitle lines appear as hoverable words. Hover a word to look it up, then mine it into Anki. [Mining workflow](/mining-workflow) covers lookup and card creation in detail.

Run `subminer` with no file to pick one from the current directory instead. See [Picking files](#picking-files).

### Yomitan setup

Lookups need at least one dictionary in SubMiner's bundled Yomitan. First-run setup asks you to import one. To add more later, open Yomitan settings with `Alt+Shift+Y` or `subminer app --yomitan`.

The bundled Yomitan is separate from any Yomitan in your browser. It has its own dictionaries and settings.

## Picking files

```bash
subminer                    # fzf picker for the current directory
subminer -d ~/Anime -r      # pick from a directory, searching subfolders
subminer -R                 # rofi picker instead of fzf (Linux)
subminer -H                 # watch history: replay, next, or previous episode
```

See [Launcher script](/launcher-script#video-picker) for picker and history details.

## Overlay basics

| Key           | Action                                                                |
| ------------- | --------------------------------------------------------------------- |
| `Alt+Shift+O` | Show or hide the overlay (works while the overlay or mpv has focus)   |
| `Alt+Shift+Y` | Open Yomitan settings (works from any window, not configurable)       |
| `V`           | Cycle the subtitle bar through hidden, visible, and hover-only        |
| `Ctrl+Alt+P`  | Open the playlist browser to queue, reorder, or jump between episodes |
| `Ctrl/Cmd+/`  | Show every overlay and mpv keybinding for this session                |

Hovering subtitle text pauses mpv, and moving away resumes it. An open Yomitan popup also keeps playback paused. Turn these off with `subtitleStyle.autoPauseVideoOnHover` and `subtitleStyle.autoPauseVideoOnYomitanPopup`.

You can drop files onto the overlay:

- A video replaces what is playing. Hold `Shift` to add it to the playlist instead.
- A subtitle file loads as a new subtitle track.

The full list is in [Keyboard shortcuts](/shortcuts). The in-player `y` key chords are in [mpv plugin](/mpv-plugin).

## YouTube playback

Pass a URL or a search. Install `yt-dlp` first.

```bash
subminer https://youtu.be/...
subminer ytsearch:"jp news"     # play the first search result
```

SubMiner picks subtitles during startup while mpv is paused. It selects a Japanese primary track and an English secondary track, downloads whatever is missing, and resumes once the primary subtitles are ready. If the choice is wrong, press `Ctrl+Alt+C` to open the YouTube subtitle picker and choose again.

Language preferences live under `youtube` and `secondarySub` in the config. See [YouTube integration](/youtube-integration).

## Common commands

```bash
subminer stats                     # start the immersion stats dashboard
subminer settings                  # open the settings window
subminer doctor                    # check dependencies, config, and the mpv socket
subminer generate-subs video.mkv   # make Japanese subtitles from the audio
subminer logs -e                   # export a log ZIP for bug reports
subminer app --setup               # reopen first-run setup
subminer -u                        # update SubMiner
```

Two flags help early on:

- `-a/--args` passes options to mpv, for example `subminer --args "--volume=80" video.mkv`.
- `--log-level debug` turns on verbose logs when something is wrong.

[Launcher script](/launcher-script) lists every command. Jellyfin, sync, and character dictionary commands are covered in [Jellyfin](/jellyfin-integration), [Sync between machines](/launcher-script#sync-between-machines), and [Character dictionary](/character-dictionary).

### Generate Japanese subtitles locally

`subminer generate-subs` transcribes audio with whisper.cpp and writes a Japanese SRT file. If that file is playing in mpv, it loads the new subtitles right away. Leave out the path to use the file mpv is playing.

```bash
subminer generate-subs video.mkv --download-model   # download a model on first use
subminer generate-subs video.mkv --model-path ~/models/ggml-medium.bin
```

You need `whisper-cli`, `ffmpeg`, and `ffprobe`. Check the output before mining, since speech recognition makes mistakes over music and overlapping voices. See [Subtitle generation](/subtitle-generation) for models, timing references, and settings.

## Windows mpv shortcut

First-run setup can create a **SubMiner mpv** shortcut in the Start menu and on the desktop. It is the easiest way to play local files on Windows:

- Double-click it to open mpv with SubMiner attached.
- Drag a video onto it to play that file.
- Run it from a terminal:

```powershell
& "C:\Program Files\SubMiner\SubMiner.exe" --launch-mpv "C:\Videos\episode 01.mkv"
```

mpv must be on `PATH`, or `mpv.executablePath` must point to `mpv.exe`. The `subminer` terminal command also works on Windows if you installed it during setup.

## Tray menu

The tray icon gives you:

- **Export Logs**: saves a log ZIP and shows its path. Usernames, IP addresses, emails, tokens, passwords, and cookies are masked in the exported copy. Your log files on disk stay unchanged.
- **View Changelog**: release notes, including versions newer than yours. Use `J`/`K` to move between versions, `Enter` to expand one, and `Esc` to close.
- **Sync Stats & History**: opens the [sync window](/launcher-script#sync-between-machines).
- **Jellyfin Discovery**: turns cast discovery on or off for this session, once [Jellyfin](/jellyfin-integration) is set up.

On Wayland, the tray icon only appears if your panel provides a StatusNotifier (AppIndicator) tray.

## Controller support {#controller-support}

You can drive the overlay with a gamepad.

1. Set `controller.enabled` to `true` in your config.
2. Connect a controller. SubMiner uses the first one it sees.
3. Press `Y` on the controller to turn on keyboard-only mode. The controller only works in this mode.
4. Move between words with the left stick, press `A` to look one up, and `X` to mine it.

Press `Alt+C` to choose a controller and remap buttons. Click an action's **Learn** button, then press the button you want. `Alt+Shift+C` shows raw input values for unusual pads.

| Button                | Action                               |
| --------------------- | ------------------------------------ |
| `A` (South)           | Look up the selected word            |
| `B` (East)            | Close the lookup                     |
| `X` (West)            | Mine a card                          |
| `Y` (North)           | Toggle keyboard-only mode            |
| `L1`                  | Play the current Yomitan audio       |
| `R1`                  | Next Yomitan audio source            |
| `L3`                  | Pause or resume mpv                  |
| `Select` / `Minus`    | Quit mpv                             |
| Left stick            | Move between words, scroll the popup |
| Right stick (up/down) | Jump through the popup               |

On controllers that report the W3C standard layout, the default quit button lands on `L2` instead of `Select`. Remap it with `Alt+C`. All options are in [Configuration](/configuration#controller-support).

## Changing settings while you watch

SubMiner watches your config file and applies most changes without a restart, including subtitle style, keybindings, and most Anki settings. If a change needs a restart, SubMiner tells you. If the file has an error, it keeps the last working config and shows a notification. See [Configuration](/configuration).

Next: [Mining workflow](/mining-workflow).
