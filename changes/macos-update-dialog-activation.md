type: fixed
area: updater

- macOS update dialogs triggered by `subminer -u` now reliably appear in the foreground. SubMiner now shows the dock icon and activates itself via `osascript` (LaunchServices) before opening the modal alert; `app.focus({ steal: true })` alone was unreliable when SubMiner was reached through single-instance forwarding from the CLI-spawned child, leaving the dialog stranded behind other apps with a bouncing dock icon.
