type: fixed
area: stats

- Stats reported 0 known words for every session after the known-word cache gained maturity tiers. The stats server carried its own copy of the cache parser that only recognized versions up to 3, so the new v4 file was read as "no cache" rather than as a format it should understand.
- The cache format, its parser, and the derived known-word set now live in one module that both the cache manager and the stats server read, and the version dispatch ends in an exhaustive check so a future format bump fails the build instead of silently reporting zero. A cache that exists but does not parse now logs a warning rather than passing for an empty one.
