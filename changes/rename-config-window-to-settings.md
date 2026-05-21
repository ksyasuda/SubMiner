type: changed
area: launcher
breaking: true

- Renamed the SubMiner Configuration window to the Settings window across the UI, tray menu, docs, and CLI verbiage.
- Replaced the `--config` flag and `subminer config` (no action) entry points with `--settings` and `subminer settings`. The `subminer config` subcommand now only accepts `path` or `show`.
- Removed the `--settings` alias that previously opened the bundled Yomitan settings popup. Use `--yomitan` to open Yomitan settings.
