type: fixed
area: updater

- Linux: `subminer -u` performs release updates independently of any running tray app (reporting `up to date` without downloading when not newer), and update checks use GitHub release metadata/assets instead of the native Electron updater to avoid network-service crashes during startup.
- macOS: update dialogs from `subminer -u` reliably appear in the foreground; builds that cannot apply native updates show a manual-install message instead of a restart prompt; `electron-updater` metadata and ZIP downloads route through `/usr/bin/curl` to avoid Electron network crashes while preserving the Squirrel install path; and metadata mismatches from conflicting ZIP filenames are resolved.
- Windows: automatic updates keep the native `electron-updater`/NSIS install path while routing updater HTTP through main-process fetch, avoiding the delayed app exit after launch.
