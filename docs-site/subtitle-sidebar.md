# Subtitle sidebar

The subtitle sidebar lists every line of the current subtitle file in a scrollable panel next to mpv. Use it to reread lines you missed, look ahead, jump to any line, or copy a stretch of dialogue.

## Using the sidebar

Press `\` to open or close it. The sidebar is on by default. Set `subtitleSidebar.enabled` to `false` to turn it off, or `subtitleSidebar.autoOpen` to `true` to open it at startup.

- Click a line to seek to it. With a line focused from the keyboard, `Enter` seeks to it.
- The current line is highlighted and kept in view as playback moves (`autoScroll`).
- Hovering the list pauses playback (`pauseVideoOnHover`).
- Switching media or subtitle track updates the list.

The sidebar needs a subtitle file SubMiner can parse. Tracks that mpv renders itself, such as embedded ASS tracks, leave it empty. With no lines loaded, the sidebar shows a **Generate Japanese subtitles** button that opens [subtitle generation](/subtitle-generation). You can also open generation any time with `Ctrl+Shift+G`.

For karaoke and animated ASS subtitles, SubMiner merges the per-frame effect lines into one clean line per cue.

## Copying dialogue {#selecting-and-copying-dialogue}

1. Drag across the text to select it. The selection can span several lines, and you can scroll to extend it.
2. Press `Ctrl/Cmd+C` or click **Copy**.

SubMiner copies the text in subtitle order, without timestamps, with a blank line between cues. Dragging does not seek, and auto-scroll pauses while you have a selection. Press `Escape` to clear it. Changing media or subtitle track, or closing the sidebar, also clears it.

## Layout

`subtitleSidebar.layout` has two modes:

- `overlay`: the sidebar floats over mpv and does not change the player window.
- `embedded`: reserves space on the right of the player and moves the video over, so the list doesn't cover it. Placement depends on your compositor. If the geometry comes out wrong, switch back to `overlay`.

## Configuration

All keys live under `subtitleSidebar`. Defaults are in the [configuration reference](/configuration).

| Key                 | What it does                                                    |
| ------------------- | --------------------------------------------------------------- |
| `enabled`           | Turn the sidebar on or off                                      |
| `autoOpen`          | Open the sidebar when the overlay starts                        |
| `layout`            | `overlay` or `embedded`                                         |
| `toggleKey`         | Toggle key, as a `KeyboardEvent.code` value such as `Backslash` |
| `pauseVideoOnHover` | Pause playback while the pointer is over the list               |
| `autoScroll`        | Keep the current line in view                                   |
| `css`               | Styling, see below                                              |

`css` takes CSS properties (`font-family`, `font-size`, `color`, `background-color`, `opacity`) and these custom properties:

| Property                                     | Styles                         |
| -------------------------------------------- | ------------------------------ |
| `--subtitle-sidebar-max-width`               | Maximum sidebar width          |
| `--subtitle-sidebar-timestamp-color`         | Timestamp text                 |
| `--subtitle-sidebar-active-line-color`       | Current line text              |
| `--subtitle-sidebar-active-background-color` | Current line background        |
| `--subtitle-sidebar-hover-background-color`  | Background of the hovered line |

```jsonc
{
  "subtitleSidebar": {
    "layout": "embedded",
    "css": {
      "font-size": "18px",
      "--subtitle-sidebar-max-width": "480px",
    },
  },
}
```

Your `css` object replaces the default one as a whole. To keep a default value, copy it from the configuration reference into your object.

See [keyboard shortcuts](/shortcuts) for all overlay keys.
