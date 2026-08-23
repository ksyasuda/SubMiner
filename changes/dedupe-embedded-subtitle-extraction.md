type: fixed
area: subtitles

- Prevented embedded subtitle parsing from starving network playback: mounted SMB/NFS media now uses deduplicated mpv live text, while duplicate extraction requests for local media share one ffmpeg process.
