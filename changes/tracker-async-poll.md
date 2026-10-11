type: fixed
area: overlay

- Hyprland and Sway window tracking no longer freezes the app when `hyprctl` or `swaymsg` is slow or hangs: polls now run in the background with a timeout and no longer pile up.
- The Hyprland/Sway trackers read each mpv process's command line once instead of spawning `ps` on every poll.
