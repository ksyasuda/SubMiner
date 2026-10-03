type: fixed
area: annotations

- Subtitle lines shown while two cues overlap (common in Jimaku and Jellyfin SRTs, and usually the long lines) are now prefetched, so their annotations appear immediately instead of after a delay.
- Subtitle prefetching no longer restarts each time the same subtitle file is re-selected during startup, and a line tokenized before a restart or seek is kept instead of discarded.
