type: added
area: overlay

- Added complete multi-language i18n framework with English and Simplified Chinese (zh-CN) support.
- UI language auto-detects from the operating system, with manual override via the new `uiLanguage` config option (system/en/zh-CN).
- Translated all user-visible surfaces: overlay modals (jimaku, youtube picker, kiku, runtime options, character dictionary, subsync, controller config, subtitle sidebar, session help, playlist browser), settings window (categories, search, save, validation, Anki controls, keybinding controls), stats dashboard (overview, library, trends, vocabulary, search, sessions, episode detail), tray menu (all 12 items), notification toasts, native dialog boxes (update, fatal error, log export, config validation, mpv plugin detection, legacy plugin removal), and mpv on-screen display messages.
- Translated first-run setup wizard: 28 section titles, 17 status labels, 60+ action buttons and inline form labels, and all service-layer error messages.
- Translated 50+ shortcut/action descriptions in the session help modal across MPV commands, key bindings, and section taxonomy (Mining, Stats, Overlay, Modals, Y chords, Global, etc.).
- Translated 25+ IPC error payloads (character dictionary, AniList, playlist browser, runtime options, subsync) and 25+ AniList/dialog strings in the character dictionary modal.
- Translated controller standard button/axis names (17 buttons + 6 axes), debug text, and picker tags.
- Translated stats component section headers, empty states, and error messages across 14 component files.
- Added ~1700 translation keys total across the en.json and zh-CN.json locale files with perfect parity.
- Fixed a bug where the first-run setup wizard injected hardcoded Chinese characters (是/否) regardless of the user's selected language.
- Fixed hardcoded `title: 'Jellyfin'/'AniList'` strings in main-process notifications.
- Documented the `uiLanguage` configuration option in `docs-site/configuration.md`.
