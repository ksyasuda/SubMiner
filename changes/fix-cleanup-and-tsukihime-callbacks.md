type: fixed
area: runtime

- Shutdown finishes independent media and Discord cleanup and waits for sync shutdown even when Jellyfin cleanup fails, then reports the first cleanup error.
- Tsukihime ignores delayed media-info results and errors from a closed modal session.
