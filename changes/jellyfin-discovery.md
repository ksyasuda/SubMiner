type: fixed
area: jellyfin

- Fixed Jellyfin discovery playback: the active item is no longer reloaded on startup, paused mpv is no longer misreported as playing, startup unpause no longer repeats after a manual pause or `y-t` toggle, duplicate ready signals no longer re-show the overlay, delayed Japanese subtitle selection is handled correctly, later-loading foreign tracks no longer steal the active Japanese track, and long-lived sidebar ffmpeg extractors no longer run against stream URLs.
- Fixed discovery resume when a remote play command sends `StartPositionTicks: 0` despite saved progress on the item.
- Kept Jellyfin picker library discovery working when the app log level is above info.
