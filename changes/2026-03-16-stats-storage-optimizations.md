type: changed
area: immersion

- Kept immersion tracking history by default while preserving daily/monthly rollup maintenance.
- Added exact lifetime summary reads for overview/anime/media stats so dashboard totals no longer depend on rescanning raw telemetry.
- Reduced tracker storage overhead by removing duplicated subtitle text from subtitle-line event payloads.
- Deduplicated episode cover-art blobs through a shared blob store and updated cover-art reads/writes to resolve shared images correctly.
- Added indexes for large-history session, telemetry, vocabulary, kanji, and cover-art queries to keep dashboard reads fast as the SQLite database grows.
