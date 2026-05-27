type: fixed
area: overlay

- Primed the first startup subtitle before autoplay resumes so the overlay renders text before video playback begins.
- Kept the visible overlay and subtitle stream alive after restarting SubMiner from the mpv `y-r` shortcut, with correct Linux bounds reapplication and user-paused playback preserved through readiness gates.
- Fixed managed mpv startup so launcher-owned videos quit SubMiner when playback ends while background/tray sessions stay alive, with pause-until-ready waiting for overlay and tokenization readiness.
- Fixed the subtitle sync modal on macOS so it no longer flashes and hides on the first attempt or leaves stale state after syncing.
- Fixed Windows managed mpv launches from a background instance so the warm app receives the start command, retargets the new mpv socket, binds to the player window, and receives startup overlay options.
