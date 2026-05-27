type: fixed
area: tray

- Kept the tray app running when closing tray-launched Yomitan settings, with a close-only menu so closing settings does not quit the tray, and an in-page close button on Hyprland where native window controls are unavailable.
- Kept settings loading from blocking other tray actions, serialized copied Yomitan extension refreshes at startup, and disabled the embedded popup preview to avoid renderer hangs during sidebar navigation.
- Fixed session help focus handling so the modal can close without mpv running.
- Fixed the Windows tray "Open SubMiner Setup" action so it opens the setup window after first-run setup is complete.
