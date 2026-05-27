type: fixed
area: jellyfin

- Keep the discovery tray checkbox in sync on Linux after tray, CLI, or startup remote-session changes, and restart stale discovery sessions when the server no longer lists the SubMiner cast target.
- Fixed remote controller visibility and progress sync for mpv/SubMiner seek jumps, stopped sessions, startup path changes, and Linux websocket reconnect windows.
- Kept Play and Resume distinct: Play starts from the beginning while Resume starts at the saved position, and final progress reports reuse SubMiner's last known position when mpv resets during stop.
- Derived cast device identity from the OS hostname and always report the client as SubMiner, ignoring legacy configurable identity fields so multiple installs no longer share a remote-session identity.
- Fixed the Windows setup login flow with an IPC bridge, immediate progress feedback, and a timeout with an inline error for unreachable servers.
