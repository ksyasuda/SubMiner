type: fixed
area: overlay

- Fixed visible overlay startup/resume so subtitle bars can be hovered and clicked as soon as the first subtitle line appears, without waiting for the next subtitle update.
- Released playback after the first overlay measurement instead of waiting for cold subtitle annotation warmup, so overlay notifications and subtitle controls do not freeze during visible-overlay startup.
- Primed Linux overlay input from the first measured subtitle/notification surface before playback resumes, so first-line subtitles and startup notifications are clickable immediately.
- Restored visible-overlay loading feedback as an mpv OSD spinner that stops once the overlay is content-ready and visible.
- Starts that OSD spinner when mpv connects, opens media, or the visible overlay is requested, so cold startup shows feedback before the overlay is almost ready.
- Shows an immediate plugin-side mpv OSD on `start-file` for visible overlay startup, even when normal plugin status OSD messages are disabled or the launcher owns the overlay start, and keeps it spinning until Electron reports the visible overlay is content-ready.
