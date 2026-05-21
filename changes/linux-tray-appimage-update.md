type: changed
area: updater

- Linux tray "Check for Updates" now installs the new AppImage automatically via `electron-updater`, matching the macOS and Windows tray flow, instead of stopping at a "manual update required" dialog. AppImages managed by a system package (AUR `/opt/SubMiner/SubMiner.AppImage`) and non-AppImage launches (no `APPIMAGE` env) still fall back to the GitHub-asset flow.
- Routed `electron-updater` HTTP through `/usr/bin/curl` on Linux and disabled differential downloads, matching the macOS path, so background update checks stay off Electron's network service.
