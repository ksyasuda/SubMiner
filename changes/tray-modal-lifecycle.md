type: fixed
area: tray

- Kept the tray app running when closing tray-launched Yomitan settings.
- Kept tray-launched Yomitan settings loading from blocking other tray actions.
- Removed the default native app menu from Yomitan settings so File > Quit cannot put the tray app into a stuck quit state.
- Disabled Yomitan's embedded popup preview in the tray-launched settings window to avoid renderer hangs during normal sidebar navigation.
- Skipped heavy Yomitan settings startup preview, storage, dictionary, and Anki controllers when launched from SubMiner to avoid renderer hangs with large dictionary databases.
- Cached Yomitan settings dictionary metadata after explicit loads to avoid repeated large IndexedDB reads.
- Serialized copied Yomitan extension refreshes so startup cannot race itself and leave extension loading in an error state.
- Fixed tray-launched session help focus handling so the modal can close without mpv running.
