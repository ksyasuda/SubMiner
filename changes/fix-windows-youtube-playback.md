type: fixed
area: launcher

- Fixed Windows direct `--youtube-play` startup so MPV boots reliably, stays paused until the app-owned subtitle flow is ready, and reuses an already-running SubMiner instance when available.
- Fixed standalone Windows `--youtube-play` sessions so closing MPV fully exits SubMiner instead of leaving hidden overlay windows or a background process behind.
