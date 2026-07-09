type: fixed
area: overlay

- Fixed character names being swallowed by longer generic dictionary matches during subtitle tokenization (e.g. ヨータ in 美姫とヨータ was lost because とヨー normalizes to とよう and matched 渡洋), so the name never tokenized and got no highlight, portrait, or hover lookup. Character-dictionary matches now claim their spans in a greedy pre-pass before the left-to-right scan, and re-segmented name spans override the parse tokenization on merge. A name only claims its span when no strictly longer generic word starts at the same position, so a character named 空 no longer splits 空気 (ties still go to the name). The pre-pass runs only when a SubMiner character dictionary is enabled in the active Yomitan profile.
