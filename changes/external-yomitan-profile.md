type: added
area: config

- Added `yomitan.externalProfilePath` to reuse another Electron app's Yomitan profile in read-only mode.
- SubMiner now reuses external Yomitan dictionaries/settings without writing back to that profile.
- Fixed default config bootstrap so `config.jsonc` is seeded even when the config directory already exists.
