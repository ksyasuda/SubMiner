type: fixed
area: overlay

- Overlay: Fixed the default replay/next subtitle keybindings by moving the session-help shortcut to `Ctrl/Cmd+/`, leaving `Ctrl+Shift+H` and `Ctrl+Shift+L` free for subtitle playback controls. The mpv plugin now registers shifted letter chords with mpv's uppercase key form so `Ctrl+Shift+L` reaches the play-next-subtitle action instead of falling through as `Ctrl+L`, and play-next now starts playback from a paused state before pausing again at the subtitle end.
