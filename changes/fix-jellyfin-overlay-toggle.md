type: fixed
area: jellyfin

- Fixed Jellyfin `y-t` overlay hide so the plugin sends an explicit hide command when it knows the overlay is visible, avoiding overlay reloads and paused playback resumes.
- Kept that manual hide sticky across Jellyfin stream redirects that change mpv's path, even when the redirected URL drops mpv's media title.
- Re-armed managed subtitle defaults during those path-changing redirects so Japanese primary subtitles can load on the redirected stream.
- Routed visible-overlay shortcuts and app-side visibility changes back through the mpv plugin so SubMiner overlay toggling stays independent of Jellyfin playback controls.
- Collapsed duplicate visible-overlay toggle events so Hyprland does not process one physical shortcut as hide-then-show.
- Kept passive Linux/Hyprland visible-overlay shows from taking keyboard focus away from mpv/Jellyfin.
- Made Jellyfin external subtitle selection tolerate transient mpv `track-list` read failures and numeric string track IDs so Japanese subtitles are selected after preload on Linux.
- Fixed AppImage-launched Jellyfin playback controls so mpv sends overlay commands to the running SubMiner app-control socket instead of the mounted Electron binary.
