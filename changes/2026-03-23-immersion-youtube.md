type: changed
area: launcher

- Added an app-owned YouTube subtitle flow that pauses mpv, lets the overlay picker choose tracks, and injects downloaded subtitle files before playback resumes.
- Added absPlayer-style YouTube timedtext parsing/conversion so downloaded subtitle tracks load as parsed cues for the sidebar, tokenization, and mining flows.
- Added yt-dlp metadata probing so YouTube playback and immersion tracking keep canonical video and channel metadata.
- Hardened the YouTube picker against duplicate submissions and tightened YouTube URL detection so follow-up runtime flows only treat real YouTube hosts as YouTube playback.
