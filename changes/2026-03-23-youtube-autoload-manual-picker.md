type: changed
area: launcher

- Changed app-owned YouTube playback to auto-load the default primary subtitle and a best-effort secondary subtitle at startup instead of forcing the picker modal first.
- Kept startup playback gating tied only to primary subtitle load and tokenization readiness; secondary failures no longer block resume.
- Added a default `Ctrl+Shift+J` keybinding that opens the YouTube subtitle picker manually during active YouTube playback.
- Routed primary YouTube subtitle auto-load failures through the configured notification output path so failed startup still resumes cleanly.
- Fixed startup-launched YouTube playback so primary subtitle overlay updates continue after auto-load completes.
- Fixed auto-loaded YouTube primary subtitles so parsed cues appear in the subtitle sidebar without needing a manual picker retry.
