type: changed
area: launcher

- Changed app-owned YouTube playback to auto-load the default primary subtitle and a best-effort secondary subtitle at startup instead of forcing the picker modal first.
- Kept startup playback gating tied only to primary subtitle load and tokenization readiness; secondary failures no longer block resume.
- Added a default `Ctrl+Alt+C` keybinding that opens the YouTube subtitle picker manually during active YouTube playback.
- Routed primary YouTube subtitle auto-load failures through the configured notification output path so failed startup still resumes cleanly.
- Kept `youtubeSubgen.whisper*` settings as legacy compatibility configuration only; normal startup now resolves tracks through YouTube source download/selection.
- Fixed startup-launched YouTube playback so primary subtitle overlay updates continue after auto-load completes.
- Fixed auto-loaded YouTube primary subtitles so parsed cues appear in the subtitle sidebar without needing a manual picker retry.
- Prefer existing authoritative YouTube subtitle tracks when available, fall back to downloaded tracks only for missing sides, and keep native mpv secondary subtitle rendering hidden so the overlay remains the sole secondary display.
