type: fixed
area: stats

- Relinking a title to a different AniList entry now updates its cover in the stats Library grid too. The detail view picked up the new art immediately while the grid kept the previous entry's image, so several unrelated titles could end up sharing one wrong cover. The library list is refetched after a relink, and cover responses now carry an ETag and revalidate instead of being cached for a day, so a stale image can no longer be served from the browser cache.
