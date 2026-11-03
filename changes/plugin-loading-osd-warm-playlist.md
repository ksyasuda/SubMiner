type: fixed
area: mpv-plugin

- Fixed the "Overlay loading" message staying on screen when mpv moves to the next file in a playlist, then turning into "Overlay did not become ready; check SubMiner logs" after 30 seconds even though the overlay was working. The overlay window is reused between files, so the plugin no longer restarts the loading message once the overlay has reported ready.
