type: fixed
area: jellyfin

- Fixed `subminer jellyfin -p` reporting "No Jellyfin libraries found." when SubMiner was already running (for example right after `subminer jellyfin` login). The app now hands library, item, and preview-auth results to the launcher directly instead of through log files, so picker image previews work again and real errors (expired session, unreachable server) are shown instead of an empty list.
- The rofi "Jellyfin Search (optional)" step after picking a library is now a compact input box with a "Type to search, or press Enter to browse all" hint instead of a blank list.
- Playing from the terminal (`subminer jellyfin -p`) now reports watch progress to Jellyfin even when cast discovery is off, so Jellyfin remembers where you stopped and marks finished episodes as played. Previously progress was only sent while the cast connection was active.
- When `subminer jellyfin -p` has to start SubMiner itself, the app now shows its tray icon and quits when you close mpv, the same as playing a local file. Before, it had no tray icon and kept running hidden after playback. A SubMiner that was already running is still reused and left running.
