type: fixed
area: app

- Fixed "Service Crash" desktop notifications (KDE DrKonqi) after closing a video when running the Linux AppImage: on quit, the AppImage runtime unmounted the FUSE squashfs while Chromium utility children (notably the network service) were still shutting down, killing them with SIGBUS. Background launches (`--background`, used by the mpv plugin and the launcher) now run through a small supervisor that mounts the AppImage via `--appimage-mount`, executes `AppRun` from that mount, and releases the mount only after no process is still executing from it. Set `SUBMINER_NO_APPIMAGE_MOUNT_KEEPALIVE=1` to restore the old direct launch.
