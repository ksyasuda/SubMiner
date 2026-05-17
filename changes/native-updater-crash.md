type: fixed
area: updates

- Restored the standard macOS `electron-updater`/Squirrel update path and routed supplemental GitHub updater requests through Electron networking instead of Node fetch.
- macOS update checks now skip local build-output apps outside Applications before touching Squirrel, and macOS tray checks no longer perform the supplemental GitHub asset lookup.
- macOS `electron-updater` metadata and full ZIP downloads now use `/usr/bin/curl` under the hood to avoid the Electron network crash seen during tray update checks while preserving Squirrel installation.
