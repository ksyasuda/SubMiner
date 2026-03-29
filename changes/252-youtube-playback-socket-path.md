type: fixed
area: main

- Resolve the YouTube playback socket path lazily so startup honors CLI and config overrides.
- Add regression coverage for the lazy socket-path lookup during Windows mpv startup.
