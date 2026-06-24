type: fixed
area: youtube

- Improved YouTube card media generation by sending safer ffmpeg request options for resolved streams and skipping stale stream maps, including cached YouTube files.
- Added `youtube.mediaCache.mode` with `direct` and `background` modes so YouTube card audio/image extraction can optionally use a background yt-dlp media cache when direct stream extraction is unreliable; background mode now announces cache download start/readiness through queued overlay/OSD notifications, creates text-only cards while the cache downloads, queues media updates for the mined note IDs, fills audio/image fields once the cached file is ready, and caps background downloads at `youtube.mediaCache.maxHeight` 720p by default.
