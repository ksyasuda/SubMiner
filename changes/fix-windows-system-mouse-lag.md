type: fixed
area: overlay

- Fixed system-wide mouse lag on Windows while SubMiner is running: the overlay no longer installs Electron's global mouse hook for click-through forwarding, and the mpv window tracker no longer blocks the app on repeated PowerShell command-line lookups.
