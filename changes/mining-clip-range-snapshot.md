type: fixed
area: anki

- Mined audio and animated AVIF clips now capture the subtitle line that was actually mined. The clip range is snapshotted once at Yomitan lookup time (and reused for both audio and image), instead of each generator reading the live mpv subtitle when it starts — which clipped whatever line was on screen after slow audio extraction finished, producing too-short or misaligned AVIF clips.
