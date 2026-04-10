type: fixed
area: overlay

- Fixed Windows overlay z-order so the visible subtitle overlay stops staying above unrelated apps after mpv loses focus.
- Fixed Windows overlay tracking to use native window polling and owner/z-order binding, which keeps the subtitle overlay aligned to the active mpv window more reliably.
- Fixed Windows overlay hide/restore behavior so minimizing mpv hides the overlay and restoring mpv brings it back aligned to the tracked window.
- Fixed stats overlay layering so the in-player stats page now stays above mpv and the subtitle overlay while it is open.
- Fixed Windows subtitle overlay stability so transient tracker misses and restore events keep the current subtitle visible instead of waiting for the next subtitle line.
- Fixed Windows focus handoff from the interactive subtitle overlay back to mpv so the overlay no longer drops behind mpv and briefly disappears.
- Fixed Windows visible-overlay startup so it no longer briefly opens as an interactive or opaque surface before the tracked transparent overlay state settles.
