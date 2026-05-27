type: fixed
area: overlay

- Hid the macOS visible overlay when mpv loses focus, is minimized, or is no longer the foreground target so other apps and Spaces are not covered, treating the frontmost mpv app as the focus signal.
- Kept the overlay stable through transient window-tracking misses, when mpv remains frontmost but window geometry temporarily disappears from macOS APIs, and when clicking from the overlay back into mpv.
- Kept the overlay correctly layered during stats mouse passthrough, opened the stats overlay inactive so it appears over fullscreen mpv without switching Spaces, and fixed passthrough so mpv controls stay clickable before hovering a subtitle bar.
- Reduced window-tracker background work by preferring the compiled helper and slowing polls while mpv is stably focused.
