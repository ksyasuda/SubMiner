type: fixed
area: overlay

- The open subtitle sidebar no longer re-fetches the whole subtitle list and re-queries mpv about 12 times a second. It polls only the current time and line, and reloads the list when the subtitle track, file, or parsed cues change (plus a slow ~2 s fallback), cutting idle sidebar IPC and renderer work during playback.
