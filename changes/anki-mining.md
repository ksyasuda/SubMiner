type: fixed
area: anki

- Made sentence-audio padding opt-in by default, and kept animated AVIF freeze-frame duration aligned to the word audio length without double-counting sentence audio padding.
- Kept multi-line sentence mining aligned when repeated subtitle text appears in the selected history range.
- Fixed Kiku duplicate-card detection so local duplicate sentence cards trigger the manual modal or auto merge, modal-open acknowledgement races no longer cancel the flow, and merged fields follow Kiku's group ordering, sentence-audio, furigana, and tag semantics.
- Fixed manual clipboard card updates from YouTube playback so generated audio and images use mpv's resolved stream URLs instead of the YouTube page URL.
- Sentence cards now refresh the current secondary subtitle before saving, so the translation field uses the loaded subtitle instead of repeating the primary text.
