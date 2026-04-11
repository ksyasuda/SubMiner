type: fixed
area: stats

- Fixed immersion-tracker timestamp handling under Bun/libsql so library rows, session timelines, and lifetime summaries keep real wall-clock millisecond values instead of truncating to invalid negative timestamps.
