type: added
area: stats

- Live-action dramas and movies in the stats Library now get posters, synopses, and titles from TMDB. Release builds include a project key, so it works out of the box; `tmdb.apiKey` (or `tmdb.apiKeyCommand`) overrides it, and is required when running from source.
- Unlinked titles that AniList cannot match are looked up on TMDB automatically when the parsed filename matches a Japanese live-action title exactly; otherwise use the new **Link to TMDB** action on a title to pick it by hand.
- Entries linked to the same TMDB title are merged into one card even when they came from different season folders, and the Library gained an Anime / Live Action filter.
- Provider reassignment preserves the previous link and artwork if the replacement download fails. Merges and sync keep conflicting AniList and TMDB identities separate.
