type: fixed
area: launcher

- Fixed the Windows `SubMiner mpv` shortcut idle launch so loading a video after opening the shortcut keeps mpv in the expected SubMiner-managed session, auto-starts the overlay, and re-arms subtitle auto-selection for the newly opened file.
- Removed the redundant `.` subtitle search path from the Windows shortcut launch args and deduped repeated subtitle source tracks in the manual sync picker so duplicate external subtitle entries no longer appear from the shortcut path.
