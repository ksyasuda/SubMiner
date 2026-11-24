type: added
area: jellyfin

- Jellyfin playback now adds the episodes around the playing one to mpv's playlist, so mpv's playlist keys and end-of-file advance move through the series (across seasons) in the same session, with resume, progress reporting, stats, and subtitles handled per episode.
- New `jellyfin.episodePlaylist.enabled` (on by default) and `jellyfin.episodePlaylist.maxEpisodes` settings, under Settings > Integrations > Jellyfin > Episode Playlist.
- mpv's playlist lists the queued episodes by title instead of stream URLs that carry the API key. Before mpv 0.42 this needs the SubMiner mpv plugin.
