type: fixed
area: stats

- The stats dashboard loads much faster on large databases. Overview, media and library detail, and trends no longer re-count every recorded word in the whole database on each request, so these views only cost as much as the sessions they show. On a ~1 GB database the overview went from ~1.4 s to ~0.14 s, media detail from ~1.5 s to ~0.08 s, the 30-day trends from ~1.5 s to ~0.06 s, and all-time trends from ~1.6 s to ~0.2 s.
- The Vocabulary tab's summary and word list load several times faster (summary ~2 s to ~0.2 s, word list ~0.6 s to ~0.2 s).
- The Trends tab downloads far less data: the per-title library charts now arrive as daily values and are totalled in the dashboard, so the all-time response shrank from ~3.9 MB to ~0.2 MB.
- Sentence search with headword matching returns several times faster (~0.3 s to ~0.08 s per query).
- Vocabulary totals and new-word charts stay correct after running vocabulary cleanup. Words it repaired could be counted or hidden by mistake; the first launch after updating corrects them in the background.
