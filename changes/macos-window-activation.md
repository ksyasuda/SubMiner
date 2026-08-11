type: fixed
area: macos

- `subminer anime` (and `--settings` / `--sync`) now bring their window to the front on macOS. `show()`/`focus()` only reorder windows inside the app that is already active, so the window opened behind the terminal that launched it; SubMiner now activates itself when opening one. The anime browser also restores its Dock icon before showing rather than after, because the overlay's fullscreen transform leaves the app as an accessory process that macOS refuses to bring forward at all.
