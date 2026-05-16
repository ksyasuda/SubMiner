type: fixed
area: tray

- Kept the tray app running when closing tray-launched Yomitan settings.
- Kept tray-launched Yomitan settings loading from blocking other tray actions.
- Replaced the default native Yomitan settings menu with a close-only menu so closing settings does not quit the tray app.
- Added an in-page close button for Yomitan settings on Hyprland, where native window controls are not available.
- Disabled Yomitan's embedded popup preview in the tray-launched settings window to avoid renderer hangs during normal sidebar navigation.
- Serialized copied Yomitan extension refreshes so startup cannot race itself and leave extension loading in an error state.
- Fixed tray-launched session help focus handling so the modal can close without mpv running.
