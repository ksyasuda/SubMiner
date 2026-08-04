type: fixed
area: streaming

- Jellyfin playback now seeds the subtitle tokenization prefetch straight from the subtitle file it downloads, instead of waiting on an mpv track-selection event that could be missed or coalesced and leave a whole episode tokenizing line by line.
- Streamed media no longer drops its parsed subtitle cues when the active subtitle track briefly cannot be resolved, such as when cycling onto a subtitle track embedded in the stream.
- Raised the subtitle tokenization cache from 256 to 2500 lines so a full-length episode or film stays warm end to end; prefetching previously stopped once the cache filled, leaving the tail of longer media uncached.
