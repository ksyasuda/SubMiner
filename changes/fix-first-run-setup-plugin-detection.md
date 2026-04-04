type: fixed
area: launcher

Fixed first-run setup blocking playback on macOS when the SubMiner mpv plugin was already installed at the canonical `~/.config/mpv` path.
Fixed launcher setup gating so stale cancelled setup state no longer prevents playback when the canonical mpv plugin entrypoint already exists.
