type: added
area: config

- Added a dedicated Configuration window with launcher entry points via `subminer --config` and `subminer config`.
- Fixed the Configuration window preload so launcher-opened windows can initialize even when Electron sandboxing is active.
- Kept config-window startup lightweight by skipping AniList token refresh and automatic update polling.
