type: fixed
area: jellyfin

- Set the mpv title before loading Jellyfin streams and reject URL-derived titles from metadata lookups, Anki source fields, Discord presence, and stats.
- Keep authenticated stream URLs out of stats identities even when playback metadata has not arrived.
- Remove previously cached credential-bearing anime parser metadata without changing watch history or library assignments.
- Use safe media identities for persisted AniList retry keys and discard URL-derived queued searches.
