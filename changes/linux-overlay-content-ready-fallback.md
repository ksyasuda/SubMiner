type: fixed
area: overlay

- Fixed the overlay never loading (stuck on the "Overlay loading" OSD spinner) when the Yomitan content-script reload raced overlay window creation at startup, most visible when playing from the anime browser on Linux/Wayland: a hidden window stops painting after that reload, so the ready-to-show signal that gates showing the overlay never fired. Content-ready now falls back to did-finish-load after a short grace period.
