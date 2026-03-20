type: fixed
area: anki

- Known-word cache refreshes now reconcile Anki changes incrementally instead of wiping and rebuilding on startup, mined cards can append their word into the cache immediately through a new default-enabled config flag, and explicit refreshes now run through `subminer doctor --refresh-known-words`.
