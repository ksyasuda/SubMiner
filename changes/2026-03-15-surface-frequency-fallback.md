type: fixed
area: overlay

- Fixed subtitle frequency tagging for merged lookup-backed tokens like `陰に` by falling back to exact surface-form Yomitan frequencies when the normalized headword lookup misses.
- Fixed MeCab merged-token position mapping across line breaks so merged content-plus-particle tokens like `陰に` keep their matched Yomitan frequency instead of inheriting shifted POS tags.
- Fixed grouped frequency parsing in both Yomitan and fallback frequency-dictionary lookups so display values like `118,121` use the leading rank instead of collapsing the rank and occurrence count into `118121`.
- Fixed frequency-rank ingestion to ignore Yomitan dictionaries explicitly marked `occurrence-based`, so raw occurrence counts are no longer treated as subtitle rank values.
- Fixed inflected headword frequency tagging to prefer ranks from the selected Yomitan `termsFind` popup entry itself, ordered by configured dictionary priority, so forms like `潜み` use primary-dictionary ranks like `4073` before falling back to lower-priority raw lemma metadata such as `CC100`.
- Fixed annotation-stage frequency filtering so exact kanji noun tokens like `者` keep their matched rank even when MeCab labels them `名詞/非自立`, instead of dropping the highlight after scan-time frequency lookup succeeds.
