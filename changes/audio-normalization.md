type: added
area: mining

- Generated sentence audio is now normalized to `-23 LUFS` by default. Clips mined from playback also mirror mpv's cubic software-volume curve, with a `-1 dBFS` limiter for amplification. The live `ankiConnect.media.normalizeAudio` and `ankiConnect.media.mirrorMpvVolume` settings control each behavior independently.
