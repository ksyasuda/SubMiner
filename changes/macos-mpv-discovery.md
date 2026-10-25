type: fixed
area: jellyfin

- Fixed casting from Jellyfin on macOS when SubMiner is opened from Finder, the Dock, or Spotlight. SubMiner now honors `mpv.executablePath` and `SUBMINER_MPV_PATH` on macOS and Linux, and on macOS also finds mpv in `/opt/homebrew/bin`, `/usr/local/bin`, and `/Applications/mpv.app` instead of failing with `connect ENOENT /tmp/subminer-socket`.
- The subtitle timing preview uses the same mpv lookup.
