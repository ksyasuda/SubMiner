type: fixed
area: updates

- Fixed Linux updates so the managed support-asset install now creates and refreshes both the launcher runtime plugin copy and the rofi theme alongside AppImage and launcher updates.
- Fixed Linux support-asset refreshes so unrelated SubMiner data directories are left alone and plugin copies are staged before replacing the live runtime plugin.
- Fixed first Linux launcher playback on fresh installs by auto-installing the managed runtime plugin copy from the bundled app before mpv starts when that copy is missing.
