type: fixed
area: jellyfin

- Authenticate Jellyfin playback, subtitle, artwork, and remote-control socket URLs with the `ApiKey` query parameter and stop sending the legacy `X-Emby-Token` and `X-Emby-Authorization` headers, so the integration keeps working on Jellyfin 12 where legacy authorization is disabled by default.
- Keep the cast-target websocket alive by answering Jellyfin keep-alive requests and reconnect when the server stops replying, so "Play on SubMiner" keeps working on Jellyfin 12 instead of silently dying about a minute after connecting. Failed playback progress and stop reports are now logged as warnings.
