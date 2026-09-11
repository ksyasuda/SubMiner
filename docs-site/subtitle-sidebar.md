# Subtitle sidebar

The subtitle sidebar puts the whole parsed cue list for the active subtitle file in a scrollable panel next to mpv. Scroll back through lines you already passed, look ahead at what is coming, and click any cue to seek straight to it. The overlay only ever shows the current line; the sidebar shows the rest.

The sidebar is enabled by default. Set `subtitleSidebar.enabled` to `false` if you want to turn it off.

## How it works

When the sidebar has no subtitle lines loaded, the **Generate Japanese subtitles** button opens [local subtitle generation](/subtitle-generation). The button hides once subtitle lines are loaded and stays hidden between lines. Press **Ctrl+Shift+G** to open generation at any time.

When SubMiner parses the active subtitle source into a cue list, the sidebar becomes available. Toggle it with the `\` key (configurable via `subtitleSidebar.toggleKey`). While open:

- The active cue is highlighted and kept in view as playback advances (when `autoScroll` is `true`).
- Between subtitle lines, the sidebar follows playback to the next cue without jumping back to a cue at the start of the file.
- Clicking any cue seeks mpv into that line. For overlapping ASS karaoke, SubMiner moves past the previous line's exit animation when the selected cue has enough time remaining.
- The sidebar and the overlay share one cue list, so a media change or subtitle source switch updates both at once.

For typeset ASS karaoke and animated signs, SubMiner collapses generated animation frames and repeated full-line color phases before they reach the sidebar. It recovers a clean complete line from a matching timed authoring comment or from full-line events surrounding generated fragments. Ordinary ASS comments, editor notes, alternate lines, repeated dialogue, and separately positioned signs remain distinct.

The sidebar only opens when a parsed cue list exists. Subtitle sources SubMiner cannot parse, such as embedded ASS tracks that mpv renders itself, leave it empty.

## Selecting and copying dialogue

Drag across subtitle text to select an excerpt, including across multiple rows. Scroll to extend a selection through a longer conversation. `Ctrl/Cmd+C` or the **Copy** button copies the highlighted text in subtitle order, without timestamps. Partial first and last lines are preserved, with a blank line between subtitle cues.

Dragging to select does not seek playback. Playback-following auto-scroll stops while you drag or have a selection, so the excerpt stays in view. Press `Escape` to clear the selection. An ordinary click with no selection still seeks to that cue.

Selection survives playback updates and Yomitan popup dismissal. Changing media or subtitle sources, refreshing the cue list, or closing the sidebar clears it. Copying an excerpt does not require creating an Anki card.

## Layout modes

Two layout modes are available via `subtitleSidebar.layout`:

**`overlay`** (default) - The sidebar floats over mpv as a panel. It does not affect the player window size or position.

**`embedded`** - Reserves space on the right side of the player and shifts the video area over, giving you a split pane. Use this when you want the cue list up without it covering the video. Positioning depends on the compositor, so switch back to `overlay` if the geometry comes out wrong.

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

## Keyboard shortcut

| Key | Action                  | Config key                     |
| --- | ----------------------- | ------------------------------ |
| `\` | Toggle subtitle sidebar | `subtitleSidebar.toggleKey`    |

The toggle is overlay-local and only opens when SubMiner has a parsed cue list for the active subtitle source. See [Keyboard Shortcuts](/shortcuts) for the full shortcut reference.
