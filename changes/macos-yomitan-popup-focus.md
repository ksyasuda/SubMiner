type: fixed
area: overlay

- Fixed macOS Yomitan popup focus after card mining or popup reload while still allowing click-away to close the popup without a hide/reappear cycle.
- Fixed macOS Yomitan popups staying open when clicking transparent overlay space; click-away is captured for popup close, then passthrough returns to mpv.
