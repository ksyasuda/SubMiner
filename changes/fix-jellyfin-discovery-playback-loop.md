type: fixed
area: jellyfin

- Prevented Jellyfin discovery playback from reloading the active item, misreporting paused mpv playback as still playing, retrying startup unpause after playback is paused again, unpausing after a manual `y-t` overlay toggle during startup, repeatedly restoring the overlay from duplicate ready signals, missing delayed Japanese subtitle selection on startup, letting later German/Russian subtitle loads steal the selected Japanese track, and spawning long-lived sidebar ffmpeg extractors against Jellyfin stream URLs.
