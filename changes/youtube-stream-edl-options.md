type: fixed
area: youtube

- Fixed direct YouTube stream media extraction by parsing mpv EDL stream URLs with their byte-length guards, preventing trailing EDL segment options from corrupting signed googlevideo URLs and causing ffmpeg 403 errors.
