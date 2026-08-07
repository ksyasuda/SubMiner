type: fixed
area: character dictionary

- A large character dictionary no longer fails to install. The Yomitan import (and the delete that precedes it) used a flat 7 second budget, so a series like One Piece failed with `importYomitanDictionary(merged.zip) timed out after 7000ms` while the import was still healthy. The budget is now sized from the merged ZIP (2 minutes plus 6 seconds per MB, capped at 30 minutes); the quick queries keep the 7 second budget.
- The "Generating character dictionary" notification now reports what it is doing instead of one static message for the whole run: the AniList page and character count while characters download (`page 12, 587 characters`), then `image 240/1220, ~4m left` per downloaded character/voice actor image, then `name 800/1220` for MeCab name splits and `saving snapshot` at the end. Updates are throttled to one per second, and every stage change and final item always reports.
- Every long phase also carries an elapsed clock (`· 1m 35s`), refreshed every 5 seconds even when nothing else moves, so a stalled step is visibly different from a frozen app. Long imports tick the same way.
- Image downloads are also logged every 100 files, and the import logs the timeout it computed.
