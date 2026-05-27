type: fixed
area: jellyfin

- Show the visible subtitle overlay automatically during Jellyfin playback so `subtitleStyle` appearance applies, and inject the bundled mpv plugin when SubMiner auto-launches mpv so mpv-side keybindings work without overlay focus.
- Made the `y-t` overlay toggle reliable and sticky across stream redirects that change mpv's path, re-arming managed subtitle defaults on redirect so Japanese primary subtitles load, collapsing duplicate toggle events on Hyprland, and keeping passive Linux/Hyprland overlay shows from stealing keyboard focus from mpv.
- Improved subtitle timing by preferring default embedded streams over external sidecars, stripping Jellyfin's server-selected stream from playback URLs, suppressing mpv auto-selection while SubMiner stages managed tracks, correcting clear Japanese-vs-English cue offsets, and restoring per-stream subtitle delay shifts. Track selection tolerates transient `track-list` read failures and numeric string track IDs on Linux.
