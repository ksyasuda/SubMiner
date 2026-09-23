type: fixed
area: subsync

- Subsync no longer fails with `Protocol "file:" not supported` on a subtitle that was dropped onto mpv. mpv reports such a track as a percent-encoded `file://` URL, which subsync read as a stream and tried to fetch over HTTP; the URL is now decoded back to its path, so both the retimed target and an alass reference work. A `file://` video path is treated as local too, which restores the video reference and ffsubsync for a dropped file.
