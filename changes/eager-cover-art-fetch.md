type: fixed
area: stats

- Cover art is now fetched eagerly when a new series starts playing, instead of waiting for the first visit to its series detail page, so the stats timeline shows the best-guess AniList image right away. The stats covers endpoint also backfills missing series art in the background, so existing series without an image pick one up on the next stats page load.
