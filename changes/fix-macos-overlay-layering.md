type: fixed
area: overlay

- Hid the macOS visible overlay when mpv is no longer the foreground target so other apps and Spaces are not covered by SubMiner subtitles.
- Kept the macOS overlay layered above active mpv while stats mouse passthrough is enabled, and treated the frontmost mpv app as the focus signal.
- Opened the stats overlay inactive on macOS so it appears over fullscreen mpv instead of switching back to SubMiner's original desktop.
- Preserved the active mpv focus state through transient macOS helper misses so subtitles do not flicker while mpv remains foreground.
- Kept fullscreen macOS overlays stable when mpv remains frontmost but window geometry temporarily disappears from the macOS window APIs.
