type: fixed
area: streaming

- Jellyfin playback now seeds the subtitle tokenization prefetch straight from the subtitle file it downloads, instead of waiting on an mpv track-selection event that could be missed or coalesced and leave a whole episode tokenizing line by line.
- Streamed media no longer drops its parsed subtitle cues when the active subtitle track briefly cannot be resolved, such as when cycling onto a subtitle track embedded in the stream.
- Subtitle prefetching now runs to the end of a file instead of stopping as soon as the tokenization cache fills, so the back half of an episode no longer gets tokenized line by line during playback. Previously the cache was also never cleared between episodes, so the stall carried over to every later title in a session.
- Raised the tokenization cache from 256 to 2500 lines. It is now purely a memory bound rather than a limit on how much gets prefetched, and it leaves room for lines that repeat across episodes so openings and endings stay warm between titles.
