type: fixed
area: anilist

- Used fresh mpv time-position, duration, and subtitle timing events for AniList post-watch threshold checks so progress updates still fire when playback reaches or skips past the watched threshold.
- Prefer season-specific AniList search results for multi-season files before falling back to the base title.
- Show a clear AniList message when the matched season is not in Planning or Watching instead of silently queueing an impossible progress update.
