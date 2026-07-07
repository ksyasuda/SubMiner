type: fixed
area: overlay

- Fixed words being highlighted green as known when a same-spelled Anki card taught a different reading (e.g. とこ parsed as 床 "bed" matching a known 床/ゆか "floor" card). The known-word cache now stores each card's word together with its reading and only matches when the token's reading agrees; cards without a reading field keep matching in any reading as before.
