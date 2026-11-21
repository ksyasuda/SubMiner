type: fixed
area: overlay

- mpv keys from `input.conf` (for example arrow keys bound to `sub-seek`) are no longer sent to a Yomitan popup that has already closed. The overlay decided whether the popup was open from a cached flag that could lag behind a close.
- mpv keys no longer get dropped while the overlay reloads mpv's bindings after it regains focus.
- With `--log-level debug`, the log names why the overlay kept an mpv-bound key instead of forwarding it.
