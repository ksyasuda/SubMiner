type: fixed
area: launcher

- Launcher-opened videos reuse an already-running background SubMiner instance, reapply preferred subtitles on warm launches, and close launcher-owned tray apps after playback ends.
- Videos stay paused when attaching to a running background app until subtitle priming and tokenization readiness complete, with mpv plugin subtitle auto-selection moved to pre-load so launch-time choices are not reset.
- `subminer settings` on macOS no longer emits Electron menu diagnostics and exits cleanly when the window is closed.
- `subminer app` on Linux returns control to the terminal immediately, and Linux first-run launcher installs build with a valid Bun shebang.
- `subminer app --setup` opens the setup flow when SubMiner is already running in the background.
