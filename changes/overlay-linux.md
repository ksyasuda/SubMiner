type: fixed
area: overlay

- Refreshed overlay placement after leaving mpv fullscreen so the visible overlay stays aligned to the player on Hyprland.
- Kept the visible overlay stacked above mpv after mpv regains focus from clicks or overlay movement, and suspended it while the in-player stats window is open, restoring it mouse-passive afterward.
- Promoted SubMiner and Yomitan settings windows above the subtitle overlay on Hyprland instead of opening behind it, without hiding subtitles.
- Hid the visible overlay as soon as the character dictionary modal opens, including while AniList lookup is loading or returns no results.
