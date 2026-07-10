type: changed
area: overlay

- Known-word highlighting now compares subtitle and Anki-card readings, preventing false matches between homographs such as 床/とこ and 床/ゆか, unrelated kanji words with the same reading, and single-kana grammar tokens that only share a card's reading.
- Cards without readings retain word-only matching, and matching still works across kana and kanji spellings. The cache refreshes from v2 to v3 automatically, and Stats reads the new format without migration.
