# Subtitle Sidebar

The subtitle sidebar displays the full parsed cue list for the active subtitle file as a scrollable panel alongside mpv. It lets you review past and upcoming lines, click any cue to seek directly to that moment, and follow along without depending on the transient overlay subtitles.

The sidebar is enabled by default. Set `subtitleSidebar.enabled` to `false` if you want to turn it off.

## How It Works

When SubMiner parses the active subtitle source into a cue list, the sidebar becomes available. Toggle it with the `\` key (configurable via `subtitleSidebar.toggleKey`). While open:

- The active cue is highlighted and kept in view as playback advances (when `autoScroll` is `true`).
- Clicking any cue seeks mpv to that timestamp.
- The sidebar stays synchronized with the overlay - media transitions and subtitle source changes update both simultaneously.

For typeset ASS karaoke and animated signs, SubMiner collapses generated animation frames and repeated full-line color phases before they reach the sidebar. It recovers a clean complete line from a matching timed authoring comment or from full-line events surrounding generated fragments. Ordinary ASS comments, editor notes, alternate lines, repeated dialogue, and separately positioned signs remain distinct.

The sidebar only appears when a parsed cue list is available. External subtitle sources that SubMiner cannot parse (for example, embedded ASS tracks rendered directly by mpv) will not populate the sidebar.

## Layout Modes

Two layout modes are available via `subtitleSidebar.layout`:

**`overlay`** (default) - The sidebar floats over mpv as a panel. It does not affect the player window size or position.

**`embedded`** - Reserves space on the right side of the player and shifts the video area to mimic a split-pane layout. Useful if you want the cue list visible without it covering the video. If you see unexpected positioning in your environment, switch back to `overlay` to isolate the issue.

## Configuration

Enable and configure the sidebar under `subtitleSidebar` in your config file:

```json
{
  "subtitleSidebar": {
    "enabled": true,
    "autoOpen": false,
    "layout": "overlay",
    "toggleKey": "Backslash",
    "pauseVideoOnHover": true,
    "autoScroll": true,
    "css": {
      "font-family": "Hiragino Sans, M PLUS 1, Source Han Sans JP, Noto Sans CJK JP",
      "color": "#cad3f5",
      "background-color": "rgba(73, 77, 100, 0.9)",
      "font-size": "16px",
      "opacity": "0.95",
      "--subtitle-sidebar-max-width": "420px",
      "--subtitle-sidebar-timestamp-color": "#a5adcb",
      "--subtitle-sidebar-active-line-color": "#f5bde6",
      "--subtitle-sidebar-active-background-color": "rgba(138, 173, 244, 0.22)",
      "--subtitle-sidebar-hover-background-color": "rgba(54, 58, 79, 0.84)"
    }
  }
}
```

Styling lives under the `css` object, using CSS property names and CSS custom properties (the same pattern as `subtitleStyle.css`).

| Option              | Type    | Default       | Description                                                                |
| ------------------- | ------- | ------------- | -------------------------------------------------------------------------- |
| `enabled`           | boolean | `true`        | Enable subtitle sidebar support                                            |
| `autoOpen`          | boolean | `false`       | Open the sidebar automatically on overlay startup                          |
| `layout`            | string  | `"overlay"`   | `"overlay"` floats over mpv; `"embedded"` reserves right-side player space |
| `toggleKey`         | string  | `"Backslash"` | `KeyboardEvent.code` for the toggle shortcut                               |
| `pauseVideoOnHover` | boolean | `true`        | Pause playback while hovering the cue list                                 |
| `autoScroll`        | boolean | `true`        | Keep the active cue in view during playback                                |

| `css` property                              | Default                     | Description                  |
| ------------------------------------------- | --------------------------- | ---------------------------- |
| `font-family`                               | `Hiragino Sans, M PLUS 1, Source Han Sans JP, Noto Sans CJK JP` | Cue text font family |
| `color`                                     | `#cad3f5`                   | Default cue text color       |
| `background-color`                          | `rgba(73, 77, 100, 0.9)`    | Sidebar shell background color |
| `font-size`                                 | `16px`                      | Base cue font size           |
| `opacity`                                   | `0.95`                      | Sidebar opacity between `0` and `1` |
| `--subtitle-sidebar-max-width`              | `420px`                     | Maximum sidebar width        |
| `--subtitle-sidebar-timestamp-color`        | `#a5adcb`                   | Cue timestamp color          |
| `--subtitle-sidebar-active-line-color`      | `#f5bde6`                   | Active cue text color        |
| `--subtitle-sidebar-active-background-color`| `rgba(138, 173, 244, 0.22)` | Active cue background color  |
| `--subtitle-sidebar-hover-background-color` | `rgba(54, 58, 79, 0.84)`    | Hovered cue background color |

## Keyboard Shortcut

| Key | Action                  | Config key                     |
| --- | ----------------------- | ------------------------------ |
| `\` | Toggle subtitle sidebar | `subtitleSidebar.toggleKey`    |

The toggle is overlay-local and only opens when SubMiner has a parsed cue list for the active subtitle source. See [Keyboard Shortcuts](/shortcuts) for the full shortcut reference.
