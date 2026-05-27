type: added
area: config

- Added a dedicated Settings window via `subminer --settings` or `subminer settings`, organized into Appearance, Behavior, Anki, Input, and Integration sections with click-to-learn keybinding controls (including the AniSkip button key) and AnkiConnect-backed deck, field, and note-type pickers.
- Expanded live reload so Settings saves apply immediately for stats keys, logging level, Jimaku, Subsync, YouTube language defaults, Anki field mappings, sentence card model, and selected annotation/runtime options.
- Settings search works across all categories, narrows on multi-word terms, and hides settings owned by richer editors.
- The note-fields note type picker defaults to the configured Anki deck's note type, then exact `Kiku`, then exact `Lapis`, leaving it blank for manual selection otherwise.
- AI and translation settings remain config-file only.
