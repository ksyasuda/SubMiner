type: fixed
area: overlay

- Fixed macOS subtitle bars being uninteractive (no hover, Yomitan lookups, or drag) after autoplay starts with "wait for overlay to be ready" enabled, until the user clicked a subtitle. The overlay window is now activated when pointer recovery fires.
- Fixed the macOS overlay and subtitles (and the subtitle sidebar) staying hidden after a modal closes until the user clicked the mpv window, and mpv keyboard shortcuts (e.g. pause/unpause) not working because focus was left in limbo. Focus is now restored to mpv when the last modal closes, so the overlay reappears above mpv and playback keys reach mpv again — including when mpv is in native fullscreen (the overlay no longer requires a manual click to come back).
