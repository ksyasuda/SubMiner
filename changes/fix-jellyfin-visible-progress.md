type: fixed
area: jellyfin

- Preserved Jellyfin-visible resume progress when mpv resets its position during playback stop by reusing SubMiner's last known playback position for final progress and stopped reports.
- Kept Jellyfin remote Play and Resume distinct so normal Play starts from the beginning, while Resume starts at Jellyfin's requested position without an early mpv seek race.
