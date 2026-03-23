type: changed
area: launcher

- Stopped forcing `--ytdl-raw-options=` before user-provided MPV options during YouTube playback so existing YouTube cookie integrations in user configs are no longer clobbered.
- Reordered launcher argument application so user `--args` are appended after SubMiner’s internal YouTube defaults, preserving explicit runtime overrides for `--ytdl-raw-options-*`.
