type: fixed
area: anilist

- AniList progress sync now adds a show to your Watching list (or Completed on the final episode) when it is not on your list yet, instead of failing with "AniList update not possible". This only happens when the file title exactly matches an AniList title or the entry is pinned with the character dictionary override; loose matches still require adding the show yourself.
- Entries marked Completed, Paused, Dropped, or Rewatching are still never changed.
