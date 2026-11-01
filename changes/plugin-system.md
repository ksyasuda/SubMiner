type: added
area: plugins

- Added JavaScript plugins loaded at startup from `<config dir>/plugins/<id>/` and enabled under `plugins` in the config. Plugins can listen to mpv events, run timers, register actions bound with `__plugin-action:<id>/<action>` keybindings or run with `--plugin-action`, run built-in actions, react to mined Anki cards and call AnkiConnect, start mpv, run programs, keep state in a per-plugin data folder, and show results in an overlay panel or their own window.
- Added an example `ai-explain` plugin that sends the current line and its neighbors to `codex exec` and shows a translation or grammar explanation.
- Added an example `anki-sync` plugin that syncs Anki to AnkiWeb after you mine cards, once mining goes quiet or playback ends.
- Added an example `episode-tags` plugin that tags mined cards with the show and episode, such as `show::Frieren::S01E07`.
- Added an example `auto-jimaku` plugin that opens the Jimaku search when a file starts without a Japanese subtitle track.
- Added a versioned plugin API with standalone authoring types, cached playback snapshots, lifecycle cancellation and cleanup, and awaitable mpv commands. Async activation is bounded and actions wait only for their own plugin. Shutdown skips pending loads, runs each deactivation once, and blocks calls from disposed plugins.
- Mining events preserve original and kept note ids with the media and subtitle captured when mining starts, so delayed enrichment and merges into older notes keep the source episode. Stats-dashboard cards do not trigger plugin events. Plugin subtitle lookups and mining provenance account for subtitle delay and return timing on the subtitle file's timeline.
