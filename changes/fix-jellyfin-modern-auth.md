type: fixed
area: jellyfin

- Authenticate Jellyfin playback, subtitle, artwork, and remote-control socket URLs with the `ApiKey` query parameter and stop sending the legacy `X-Emby-Token` and `X-Emby-Authorization` headers, so the integration keeps working on Jellyfin 12 where legacy authorization is disabled by default.
