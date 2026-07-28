type: fixed
area: anilist

- Season 2 and later files now resolve to the right AniList entry. AniList has no numbered seasons (sequels are separate entries titled `Zoku`, `Kan`, `2nd Season`), so the old `<title> Season N` search returned nothing and silently fell back to the season 1 entry: an Oregairu S03 episode loaded the season 1 character dictionary and pushed watch progress onto season 1. SubMiner now searches the base title and walks AniList `SEQUEL` relations to reach the requested season, ordering the franchise's TV entries by air date when the relation chain is incomplete.
- When the season still cannot be located, SubMiner no longer writes progress to the season 1 entry; it skips the update and points at the manual AniList override instead. The character dictionary logs the same warning and states which entry it fell back to.
- Character dictionary overrides are now scoped by detected season as well as directory, so seasons kept in one flat folder no longer share (or overwrite) each other's override, and an override now also pins the entry used for AniList watch progress: correcting a wrong match once fixes both the dictionary and progress tracking.
- Cover art lookup uses the same season-aware resolution, so per-season art no longer falls back to the season 1 cover.
- An unresolvable season is no longer cached as if it were a normal match, so the fallback cannot silently stick for every later episode of that series.
