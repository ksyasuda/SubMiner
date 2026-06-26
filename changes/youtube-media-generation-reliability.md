type: fixed
area: youtube

- Improved YouTube card media generation by sending safer ffmpeg request options for resolved streams and skipping stale stream maps, including cached YouTube files.
- Added `youtube.mediaCache.mode` with `direct` and `background` modes so YouTube card audio/image extraction can optionally use a background yt-dlp media cache when direct stream extraction is unreliable; background mode now announces cache download start/readiness through queued overlay/OSD notifications, creates text-only cards while the cache downloads, queues media updates for the mined note IDs, fills audio/image fields once the cached file is ready, and caps background downloads at `youtube.mediaCache.maxHeight` 720p by default while honoring height changes between downloads.
- Cleaned stale YouTube background media cache files on startup and before new background downloads so only the active cached media file remains.
- Hardened YouTube media cache downloads with IPv4/extractor retry flags and made failed background downloads notify users and clear queued media updates instead of leaving them pending silently.
