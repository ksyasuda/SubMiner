type: fixed
area: youtube

- Improved YouTube card media generation by sending safer ffmpeg request options for resolved streams and skipping stale stream maps.
- Added `youtube.mediaCache.mode` with `direct` and `background` modes so YouTube card audio/image extraction can optionally use a background yt-dlp media cache when direct stream extraction is unreliable.
