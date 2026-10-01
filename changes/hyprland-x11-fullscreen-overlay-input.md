type: fixed
area: overlay

- Fixed the overlay (subtitles and sidebar) becoming unclickable after mpv goes fullscreen on Hyprland with `mpv.backend: x11`. The launcher no longer strips `HYPRLAND_INSTANCE_SIGNATURE` from the X11 environment, so the overlay can lift Hyprland's fullscreen input block again.
