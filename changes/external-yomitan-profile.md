type: changed
area: config

- Added `yomitan.externalProfilePath` to reuse another Electron app's Yomitan profile in read-only mode.
- SubMiner now reuses external Yomitan dictionaries/settings without writing back to that profile.
- SubMiner now seeds `config.jsonc` even when the default config directory already exists.
- First-run setup now allows zero internal dictionaries when `yomitan.externalProfilePath` is configured, and falls back to requiring at least one internal dictionary if that external profile is later removed.
