type: fixed
area: stats

- Session known-word counts no longer show 0 everywhere. The stats server's known-word cache parser only understood the v1/v2 cache formats, so after the reading-aware v3 cache upgrade it silently treated the cache as missing; it now flattens v3 note entries into the headword set.
