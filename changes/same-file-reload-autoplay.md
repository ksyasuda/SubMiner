type: fixed
area: playback

- Reloading the same file in mpv (`loadfile <current path> replace`) could intermittently leave playback paused forever. SubMiner now treats that reload like any other newly loaded file, so it resets the subtitle state and resumes once the file is ready.
