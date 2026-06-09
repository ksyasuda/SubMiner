type: changed
area: stats

- Split local and Jellyfin library entries by detected season, using season folders first and filename parsing as fallback.
- Repaired older combined-series stats rows by moving parsed episodes into season-specific library entries, rebuilding summaries, and deleting now-empty legacy rows.
- Refresh anime detail and library cover art immediately after manually changing an AniList entry.
