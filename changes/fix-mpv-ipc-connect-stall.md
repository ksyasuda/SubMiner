type: fixed
area: overlay

- Fixed the overlay getting stuck on "Overlay loading" forever when the app's mpv IPC connection stalled silently: connection attempts now time out after 5 seconds and retry instead of latching, and switching to a new mpv socket aborts any stalled attempt to the old one.
