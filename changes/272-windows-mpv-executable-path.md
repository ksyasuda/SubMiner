type: fixed
area: launcher

- Added a blank-by-default `mpv.executablePath` override for Windows playback so users can point SubMiner at `mpv.exe` when it is not on `PATH`.
- Kept the Windows shortcut and `--launch-mpv` flow simple by preserving PATH auto-discovery as the default and exposing the override in first-run setup.
