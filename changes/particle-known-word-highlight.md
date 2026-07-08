type: fixed
area: overlay

- Fixed grammar particles, interjections, and suffixes rendering the known-word highlight even though they are excluded from annotations (e.g. よ in 全然いいよ, standalone え, non-independent forms like ですよ/れる). The annotation noise filter now clears the known-word flag along with JLPT/frequency/N+1 metadata, so excluded tokens render as plain text; the N+1 detector still counts their real known status when sizing sentences.
- Fixed single-kana tokens counting as known words by borrowing the reading of an unrelated Anki note (よ matched a card read よ such as 夜, making the particle "known"). Reading-only known-word matching now requires at least two kana; single-kana cards still match by their word field.
- Standalone suffix tokens (MeCab pos2 接尾, e.g. さん, れる) are now excluded from annotations by default, matching how particles and interjections are treated. Override via the pos2 exclusion config if you want them annotated.
