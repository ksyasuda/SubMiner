type: fixed
area: overlay

- Fixed `mpv.pauseUntilOverlayReady` releasing playback seconds before tokenization warmup finished: startup subtitle priming emits the current cue untokenized so the overlay can paint early, and that emission was treated as the autoplay-readiness signal as soon as the overlay window loaded. The autoplay gate now ignores untokenized subtitle payloads while tokenization warmup is pending, so playback resumes only after the first tokenized delivery (or the post-warmup release). Most visible when resuming mid-episode or when a subtitle cue starts within the first two seconds.
