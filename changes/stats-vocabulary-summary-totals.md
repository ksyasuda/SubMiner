type: fixed
area: stats

- Fixed Vocabulary totals and charts counting only the first browsing page instead of all tracked vocabulary, without delaying the rest of the page.
- New-word history now uses permanent daily lexical rollups, backfilled in the background and repaired when tracked material is removed or reprocessed; playback writes queue safely during the one-time rebuild and resume afterward.
- Calendar-day chart labels now preserve the recorded local date in time zones west of UTC.
- Vocabulary summary cards and charts refresh automatically after the word exclusion list changes, and failed loads retry with backoff before showing an inline error with a Retry control.
