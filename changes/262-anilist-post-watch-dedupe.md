type: fixed
area: anilist

- Stop AniList post-watch from sending a second progress update when the current episode was already satisfied by a ready retry item in the same watch-completion pass.
- Add regression coverage for the retry-queue plus live-update duplicate path.
