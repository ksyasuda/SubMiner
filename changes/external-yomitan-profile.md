type: changed
area: config

- Added `yomitan.externalProfilePath` to reuse another Electron app's Yomitan profile in read-only mode.
- SubMiner now reuses external Yomitan dictionaries/settings without writing back to that profile.
- SubMiner now seeds `config.jsonc` even when the default config directory already exists.
