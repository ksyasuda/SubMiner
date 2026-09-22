type: fixed
area: jellyfin

- Authenticate Jellyfin playback, subtitle, artwork, and remote-control socket URLs with the `ApiKey` query parameter and stop sending the legacy `X-Emby-Token` and `X-Emby-Authorization` headers, so the integration keeps working on Jellyfin 12 where legacy authorization is disabled by default.
- Keep the cast-target websocket alive by answering Jellyfin keep-alive requests and reconnect when the server stops replying, so "Play on SubMiner" keeps working on Jellyfin 12 instead of silently dying about a minute after connecting. Failed playback progress and stop reports are now logged as warnings.

- Send the playback stop report only after any in-flight progress report has finished and stop reporting progress the moment playback ends, so the Jellyfin "now playing" bar clears when you close or finish a cast video instead of running on to the end of the episode.
- Anki cards mined from Jellyfin playback get the episode title in the misc info field again instead of "Unknown media"; the title mpv was given before loading was being discarded when the stream path changed.
