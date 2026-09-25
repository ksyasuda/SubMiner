type: fixed
area: jellyfin

- Jellyfin playback now selects the Japanese subtitle track, and starts subtitle annotations, as soon as that track downloads instead of waiting for every other subtitle track. Tracks now download in parallel, so a slow embedded track Jellyfin has to extract no longer delays the primary subtitles.
- Jellyfin episodes with image-based embedded subtitles (PGS, DVD, DVB) now auto-select the Japanese and English tracks. SubMiner no longer requests those tracks as text, and a subtitle track that fails to download no longer cancels selection of the others.
