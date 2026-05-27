type: fixed
area: logging

- Forward SubMiner `logging.level` into launcher-started and Windows shortcut-started mpv sessions, covering mpv log verbosity, plugin script logging, and plugin-launched app logging.
- Added numeric `logging.rotation` (default 7 days of retained daily app, launcher, and mpv logs) and `logging.files` toggles per component, with mpv logs disabled by default unless explicitly enabled for debugging.
- Added Windows mpv launch, IPC socket, subtitle track, and Yomitan/dictionary diagnostics, including expected/active IPC socket values when plugin auto-start skips on a socket mismatch.
- Stop repeated mpv IPC socket warning spam while the app waits in the background for mpv to recreate the socket.
