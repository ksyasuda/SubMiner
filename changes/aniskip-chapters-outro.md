type: added
area: aniskip

- Added outro skipping: the AniSkip key now also skips the ending, with its own prompt and chapter markers.
- Intro and outro detection now reads the file's own chapters first (named OP/ED chapters, or song-length chapters near the start and end) and only asks AniSkip for what is missing, so it also works for Jellyfin and other streams.
- AniSkip lookups now send the episode length so the times match your TV or Blu-ray cut.
- Fixed AniSkip rate limits, server errors, and network failures being remembered as "no intro" until restart; they are now retried on the next load.
