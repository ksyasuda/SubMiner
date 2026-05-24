type: fixed
area: linux

- Suppressed false YouTube primary subtitle failure notifications after SubMiner confirms the selected primary track loaded successfully.
- Ensured launcher-managed playback commands create the tray icon even when they attach to an already-running SubMiner process.
- Prevented app-owned YouTube playback from letting the mpv plugin start a second SubMiner process after the launcher already started one.
- Logged Linux tray registration failures with a StatusNotifier/AppIndicator hint and documented the Hyprland tray-host requirement.
