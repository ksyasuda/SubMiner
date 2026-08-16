type: fixed
area: overlay

- Dedicated overlay modals are prewarmed on macOS and Windows so shortcuts open them promptly on the first press. Windows now refreshes the hidden modal renderer between sessions to keep later modals interactive. On macOS, reused modals and the in-app stats window also open above fullscreen mpv on its current Space instead of appearing on another desktop or forcing a Space change.
- Updated subtitle ASS observation to mpv's current `sub-text/ass` property, removing its deprecation warning.
