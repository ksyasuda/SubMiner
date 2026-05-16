type: fixed
area: updater

- Stopped Linux tray update checks from invoking the native Electron updater, using GitHub release metadata/assets instead so checks do not crash the tray app.
