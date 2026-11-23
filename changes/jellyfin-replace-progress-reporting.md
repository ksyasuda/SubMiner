type: fixed
area: jellyfin

- Starting Jellyfin playback while another file was playing could drop the new item's playback tracking, so Jellyfin never received its progress or stop reports. mpv's unload of the previous file no longer cancels the new playback.
