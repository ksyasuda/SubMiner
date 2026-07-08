type: fixed
area: overlay

- Fixed single-kana grammar tokens counting as known words by borrowing the reading of an unrelated Anki note (よ in 全然いいよ matched a card read よ such as 夜, standalone え matched 絵), which painted them with the known-word highlight. Reading-only known-word matching now requires at least two kana; single-kana cards still match by their word field, and tokens genuinely present in the known-words cache (e.g. です) keep their highlight.
- Standalone suffix tokens (MeCab pos2 接尾, e.g. さん, れる) are now excluded from JLPT/frequency/N+1 annotations by default, matching how particles and interjections are treated. Cache-backed known-word highlighting still applies; override via the pos2 exclusion config if you want them annotated.
