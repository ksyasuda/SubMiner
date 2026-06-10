type: changed
area: playback

- AniSkip intro detection now runs in the SubMiner app instead of the mpv plugin: lookups cover every local file loaded during an mpv session (including playlist advances), and the plugin no longer performs any network calls.
- `mpv.aniskipEnabled` and `mpv.aniskipButtonKey` now hot-reload without restarting playback.
- AniSkip now requires the SubMiner app to be connected to mpv; plugin-only mpv sessions without the app no longer fetch skip windows.
