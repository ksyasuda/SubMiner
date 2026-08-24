type: fixed
area: subtitles

- Embedded subtitle tracks on network-mounted (SMB/NFS) media are extracted and parsed again, restoring full karaoke reconstruction, sidebar cues, and mining for releases that ship subtitles only inside the container. Extraction reads the whole file once per episode (roughly 10 seconds per GB on gigabit), its timeout now accommodates large Bluray remuxes, and duplicate extraction requests share one ffmpeg process. Only true remote URLs keep the live-text-only path.
- Live subtitle text from per-glyph typeset karaoke no longer shows a wall of scattered letters in the overlays while extraction is still running or when no parsed cues exist (remote URLs, unreadable sources); the glyph wall and its typed-syllable fragments are suppressed while concurrent dialogue lines remain.
