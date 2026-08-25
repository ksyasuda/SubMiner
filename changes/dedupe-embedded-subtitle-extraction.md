type: fixed
area: subtitles

- Embedded subtitle tracks on network-mounted (SMB/NFS) media are extracted and parsed again, restoring karaoke reconstruction, sidebar cues, and mining for releases that ship subtitles only inside the container. Extraction reads the file once per episode (roughly 10 seconds per GB on gigabit), its timeout accommodates large Bluray remuxes, and duplicate requests share one ffmpeg process. Only true remote URLs keep the live-text-only path.
- While extraction is still running, or when no parsed cues exist (remote URLs, unreadable sources), per-glyph typeset karaoke no longer shows a wall of scattered letters in the overlays; concurrent dialogue lines still display.
