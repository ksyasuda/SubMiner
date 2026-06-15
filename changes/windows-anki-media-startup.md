type: fixed
area: anki

- Fixed known-word cache refreshes without a configured deck by using AnkiConnect's valid all-notes query instead of `is:note`.
- Fixed Windows media generation after background launches by recreating missing FFmpeg temp output directories before clipping audio or images.
