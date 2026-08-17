type: fixed
area: overlay

- Fixed the overlay getting stuck on "Overlay loading" forever when startup stalls: mpv IPC connection attempts now time out and retry, switching sockets aborts obsolete attempts, and the plugin replaces its spinner with an actionable error if overlay content is still not ready after 30 seconds.
