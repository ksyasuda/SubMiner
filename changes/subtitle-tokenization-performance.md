type: changed
area: subtitles

- Subtitle tokenization no longer runs a duplicate full `parseText` pass per line: the termsFind scanner walk is now the only tokenizer and emits its own hoverable filler runs for unmatched text (parseText is kept only as an error fallback). This roughly halves the dictionary work per line.
- The Yomitan scanning helpers are now installed once per parser window (`__subminerYomitanScan`) instead of re-shipping and re-parsing a ~500-line script for every subtitle line; each line only evaluates a tiny call.
- termsFind lookups are cached across subtitle lines in a window-persistent LRU keyed by substring, so repeated particles and verb forms stop costing backend round trips. The cache invalidates on dictionary/settings changes and window reloads.
- The scanner walk now skips lookups at punctuation and whitespace positions (latin letters and digits still look up, e.g. Tシャツ) and caps the shrinking-window retry ladder at four extra lookups per position.
- Tokenizer runtime dependencies are built once instead of per line, fixing a JLPT lookup cache that never hit (it was keyed on a per-call closure identity and leaked a Map per line) and a `which mecab` availability check that re-ran synchronously on every line when MeCab is absent.
- Subtitle changes no longer restart the prefetch run per line (which discarded in-flight tokenization work); prefetch now only pauses for the live line and restarts on real seeks, cache invalidation, or option changes. Prefetch also stays paused across a provisional raw-subtitle emit and resumes only after the tokenized payload lands, so it never competes with the on-screen line for the parser window.
- Added per-stage debug timings (`scanMs`, `mecabMs`, `frequencyMs`, `annotateMs`) to the subtitle tokenization pipeline log.
