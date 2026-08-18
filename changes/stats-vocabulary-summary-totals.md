type: fixed
area: stats

- Fixed Vocabulary totals and charts counting only the first browsing page instead of all tracked vocabulary, without delaying the rest of the page.
- New-word history now uses permanent daily lexical rollups that apply the same vocabulary filters as the totals and normalize legacy second/millisecond timestamps; versioned background rebuilds repair existing history without dropping playback writes.
- Calendar-day chart labels now preserve the recorded local date in time zones west of UTC.
- Vocabulary summary cards and charts refresh automatically after the word exclusion list changes, and failed or unfinished loads use bounded retries before showing an inline error with a Retry control.
- Rapid exclusion edits no longer race each other; writes are sent in order so a slower earlier save cannot overwrite a newer list.
