type: fixed
area: overlay

- Kept the visible overlay and subtitle stream alive after restarting SubMiner from the mpv `y-r` shortcut by transporting Linux AppImage control args safely, restoring mpv subtitle visibility during shutdown, snapshotting subtitles before overlay suppression resumes, reapplying Linux overlay bounds after the restarted window maps, allowing Hyprland to resize the visible overlay window, and preserving user-paused playback while readiness gates clear.
