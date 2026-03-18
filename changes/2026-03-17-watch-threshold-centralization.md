type: changed
area: anilist

- Standardized episode completion threshold by introducing `DEFAULT_MIN_WATCH_RATIO` and using it for both local watched state transitions and AniList post-watch progress updates.
- Episode auto-marking now uses the same threshold as AniList (`85%`), removing divergent completion behavior.
