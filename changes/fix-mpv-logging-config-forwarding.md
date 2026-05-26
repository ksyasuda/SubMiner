type: fixed
area: logging

- Forward SubMiner `logging.level` into launcher-started and Windows shortcut-started mpv sessions, including mpv log verbosity, plugin script logging, and plugin-launched app logging.
- Add numeric `logging.rotation`, defaulting to 7 days of retained daily app, launcher, and mpv logs.
- Log Windows mpv launch diagnostics, IPC socket connection state, subtitle track summaries, Yomitan extension load state, dictionary counts, and expected/active IPC socket values when plugin auto-start skips due to a socket mismatch.
- Add `logging.files` toggles for app, launcher, and mpv logs, with mpv logs disabled by default unless explicitly enabled for debugging.
