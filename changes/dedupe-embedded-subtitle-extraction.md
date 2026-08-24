type: fixed
area: subtitles

- Prevented embedded subtitle parsing from starving network playback: mounted SMB/NFS media now uses deduplicated mpv live text, while duplicate extraction requests for local media share one ffmpeg process.
- Live subtitle text from per-glyph typeset karaoke (network-mounted media without parsed cues) no longer shows a wall of scattered letters in the overlays; the glyph wall and its typed-syllable fragments are suppressed while concurrent dialogue lines remain.
