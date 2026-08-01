type: added
area: anki

- Added `ankiConnect.lapisKiku.wordCardKind` (Settings > Mining/Anki > Kiku/Lapis Features > "Word Card Type") to choose which card-type flag SubMiner marks on Kiku/Lapis word cards: `word-and-sentence` (default, unchanged behavior), `click`, `sentence`, `audio`, or `none` to leave the flags alone. Requested by Kiku users who only want `IsClickCard` set.
- Word cards can now be flagged as Kiku click cards (`IsClickCard`), which SubMiner previously never set.
- Setting a card-type flag now clears `IsClickCard` alongside the other card-type flags, so a note can no longer claim two card types at once.
