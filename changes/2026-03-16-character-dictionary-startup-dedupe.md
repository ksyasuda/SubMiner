type: fixed
area: anki

- Fixed repeated character-dictionary startup work by scheduling auto-sync only from mpv media-path changes instead of also re-triggering it from connection and media-title events for the same title.
