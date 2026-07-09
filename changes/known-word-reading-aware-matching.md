type: changed
area: overlay

- New reading-aware subtitle parsing and known-word matching: the known-word cache now stores each Anki card's word together with its reading (cache format v3), and a token only gets the known-word highlight when its parsed reading agrees with the card. The cache and the stats server upgrade automatically.
- Fixes words being highlighted green as known when a same-spelled Anki card taught a different reading (e.g. とこ parsed as 床 "bed" no longer matches a known 床/ゆか "floor" card). Cards without a reading field keep matching in any reading as before.
- Single-kana grammar tokens (よ in 全然いいよ, standalone え) no longer borrow the reading of an unrelated card (such as 夜 or 絵) and get painted as known: reading-only matching requires at least two kana, while single-kana cards still match by their word field.
