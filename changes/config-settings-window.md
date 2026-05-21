type: added
area: config

- Added a dedicated Settings window with launcher entry points via `subminer --settings` and `subminer settings`.
- Fixed the Settings window preload so launcher-opened windows can initialize even when Electron sandboxing is active.
- Kept settings-window startup lightweight by skipping AniList token refresh and automatic update polling.
- Marked safe live config options in the Settings window and expanded hot reload for stats keys, logging level, Jimaku, Subsync, YouTube language defaults, Anki media/sentence/misc field mappings, sentence card model, and selected Anki annotation/runtime options.
- Hid AI and translation fields from the Settings window while keeping them supported in config files.
