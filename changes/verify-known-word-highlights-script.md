type: internal
area: tokenizer

- Added `verify-known-word-highlights:electron` script: tokenizes a real subtitle file through the app's Yomitan/MeCab pipeline with the live known-word cache, prints each line in the configured tier colors, and summarizes the tier counts so highlighting can be checked outside of playback.
- Added `--audit`, which re-derives every highlighted tier from live Anki card data (`notesInfo` + `cardsInfo` intervals) and reports each token whose rendered tier disagrees, catching both stale cache entries and tier-classification bugs.
- Added `--profile-copy` so the check can run while SubMiner is open (Electron locks the Yomitan userData dir), plus `--refresh`, `--limit`, `--json`, and `--quiet`.
- Added `KnownWordCacheManager.getKnownWordMatchNoteIds`, exposing the note ids behind a known-word match so an audit can trace a rendered tier back to the exact Anki notes.
