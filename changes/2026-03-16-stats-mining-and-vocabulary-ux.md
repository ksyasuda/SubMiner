type: added
area: immersion

- Added Mine Word, Mine Sentence, and Mine Audio buttons to word detail example lines in the stats dashboard.
- Mine Word creates a full Yomitan card (definition, reading, pitch accent) via the hidden search page bridge, then enriches with sentence audio, screenshot, and metadata extracted from the source video.
- Mine Sentence and Mine Audio create cards directly with appropriate Lapis/Kiku flags, sentence highlighting, and media from the source file.
- Media generation (audio + image/AVIF) runs in parallel and respects all AnkiConnect config options.
- Added word exclusion list to the Vocabulary tab with localStorage persistence and a management modal.
- Fixed truncated readings in the frequency rank table (e.g. お前 now shows おまえ instead of まえ).
- Clicking a bar in the Top Repeated Words chart now opens the word detail panel.
- Secondary subtitle text is now stored alongside primary subtitle lines for use as translation when mining cards from the stats page.
