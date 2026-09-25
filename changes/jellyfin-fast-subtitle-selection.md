type: fixed
area: jellyfin

- Jellyfin playback now selects the Japanese subtitle track, and starts subtitle annotations, as soon as that track downloads instead of waiting for every other subtitle track. Tracks now download in parallel, so a slow embedded track Jellyfin has to extract no longer delays the primary subtitles.
