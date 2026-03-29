# Subtitle Sidebar

The subtitle sidebar displays the full parsed cue list for the active subtitle file as a scrollable panel alongside mpv. It lets you review past and upcoming lines, click any cue to seek directly to that moment, and follow along without depending on the transient overlay subtitles.

The sidebar is opt-in and disabled by default. Enable it under `subtitleSidebar.enabled` in your config.

## How It Works

When SubMiner parses the active subtitle source into a cue list, the sidebar becomes available. Toggle it with the `\` key (configurable via `subtitleSidebar.toggleKey`). While open:

- The active cue is highlighted and kept in view as playback advances (when `autoScroll` is `true`).
- Clicking any cue seeks mpv to that timestamp.
- The sidebar stays synchronized with the overlay — media transitions and subtitle source changes update both simultaneously.

The sidebar only appears when a parsed cue list is available. External subtitle sources that SubMiner cannot parse (for example, embedded ASS tracks rendered directly by mpv) will not populate the sidebar.

## Layout Modes

Two layout modes are available via `subtitleSidebar.layout`:

**`overlay`** (default) — The sidebar floats over mpv as a panel. It does not affect the player window size or position.

**`embedded`** — Reserves space on the right side of the player and shifts the video area to mimic a split-pane layout. Useful if you want the cue list visible without it covering the video. If you see unexpected positioning in your environment, switch back to `overlay` to isolate the issue.

## Configuration

Enable and configure the sidebar under `subtitleSidebar` in your config file:

```json
{
  "subtitleSidebar": {
    "enabled": false,
    "autoOpen": false,
    "layout": "overlay",
    "toggleKey": "Backslash",
    "pauseVideoOnHover": false,
    "autoScroll": true,
    "fontFamily": "\"M PLUS 1\", \"Noto Sans CJK JP\", sans-serif",
    "fontSize": 16
  }
}
```

| Option                      | Type    | Default      | Description                                                                                        |
| --------------------------- | ------- | ------------ | -------------------------------------------------------------------------------------------------- |
| `enabled`                   | boolean | `false`      | Enable subtitle sidebar support                                                                    |
| `autoOpen`                  | boolean | `false`      | Open the sidebar automatically on overlay startup                                                  |
| `layout`                    | string  | `"overlay"`  | `"overlay"` floats over mpv; `"embedded"` reserves right-side player space                        |
| `toggleKey`                 | string  | `"Backslash"` | `KeyboardEvent.code` for the toggle shortcut                                                      |
| `pauseVideoOnHover`         | boolean | `false`      | Pause playback while hovering the cue list                                                         |
| `autoScroll`                | boolean | `true`       | Keep the active cue in view during playback                                                        |
| `maxWidth`                  | number  | `420`        | Maximum sidebar width in CSS pixels                                                                |
| `opacity`                   | number  | `0.95`       | Sidebar opacity between `0` and `1`                                                                |
| `backgroundColor`           | string  | —            | Sidebar shell background color                                                                     |
| `textColor`                 | string  | —            | Default cue text color                                                                             |
| `fontFamily`                | string  | —            | CSS `font-family` applied to cue text                                                              |
| `fontSize`                  | number  | `16`         | Base cue font size in CSS pixels                                                                   |
| `timestampColor`            | string  | —            | Cue timestamp color                                                                                |
| `activeLineColor`           | string  | —            | Active cue text color                                                                              |
| `activeLineBackgroundColor` | string  | —            | Active cue background color                                                                        |
| `hoverLineBackgroundColor`  | string  | —            | Hovered cue background color                                                                       |

Default colors use Catppuccin Macchiato with a semi-transparent shell so the panel stays readable without feeling like a solid overlay.

## Keyboard Shortcut

| Key | Action                  | Config key                     |
| --- | ----------------------- | ------------------------------ |
| `\` | Toggle subtitle sidebar | `subtitleSidebar.toggleKey`    |

The toggle is overlay-local and only opens when SubMiner has a parsed cue list for the active subtitle source. See [Keyboard Shortcuts](/shortcuts) for the full shortcut reference.
