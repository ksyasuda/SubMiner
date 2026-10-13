type: internal
area: overlay

- Removed 16 overlay preload/IPC methods that nothing called, plus their main-process handlers. SubMiner no longer observes mpv's `sub-text/ass` or sends every ASS line to overlay windows that never listened.
