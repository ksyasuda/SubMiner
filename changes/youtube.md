type: fixed
area: youtube

- Downloaded selected YouTube primary subtitles to temporary local files so the primary bar and sidebar read the same source, with cleanup on reload and quit, and suppressed false load-failure notifications by re-checking live mpv subtitle state.
- Launcher-managed playback commands create the tray icon even when attaching to an already-running process, and app-owned YouTube playback no longer lets the mpv plugin start a second SubMiner instance.
- Logged Linux tray registration failures with a StatusNotifier/AppIndicator hint and documented the Hyprland tray-host requirement.
