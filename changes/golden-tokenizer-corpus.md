type: internal
area: tokenizer

- Added a golden-file regression corpus for the tokenizer/annotation pipeline: recorded Yomitan backend responses and MeCab tokens replay through the real tokenizeSubtitle pipeline in bun tests without Electron or dictionaries.
- Added `record-tokenizer-fixture:electron` script to capture new fixtures from a live Yomitan/MeCab session, with flags for known words, JLPT levels, and annotation toggles.
- Seeded eleven fixtures covering the #147–#156 regression classes (grammar-helper suppression, lexical くれる, kanji non-independent nouns, N+1 targeting, reading collisions, unparsed runs, ordinal/honorific prefixes).
- Added `compare-yomitan-api:electron` script that diffs SubMiner tokenization against a stock Yomitan instance via the yomitan-api bridge (segmentation, readings, headword forms).
