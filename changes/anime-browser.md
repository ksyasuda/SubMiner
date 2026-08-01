type: added
area: anime

- Added an anime browser window that searches Aniyomi extension sources, shows cover art and episode lists, and plays an episode in mpv so the overlay and mining tools attach as usual.
- Added `subminer anime` and the `--anime` flag to open the browser, plus a "Browse Anime" tray entry.
- Anime extensions are read from `<userData>/anime-extensions`; drop Aniyomi `.apk` files there to add sources.
- Added a source settings tab so extensions that need configuration (server address, credentials, quality) can be set up from the browser; values persist per source.
- Added an Extensions tab for adding repository URLs and installing, updating, or removing extensions in place; extensions that fail to load are listed with the reason.
- Browse, Extensions, and Source settings are tabs, so each one gets the full window instead of sharing it with the search results.
- Repository URLs only need to be an https URL to a `.json` index; the file name is not restricted to `index.min.json`.
- The source picker offers "All sources", which searches every installed source at once. Results stream in as each source answers — a fast source is on screen while a slow one is still resolving, with per-source progress in the status bar. Results are tagged with the source they came from, and a source that fails is named in the status bar instead of blanking the grid.
- The Extensions tab opens with an Installed section listing every extension on disk with Remove — including ones added by hand or whose repository has since been removed — and Update where a configured repository still carries it.
- Added `anime.repos`, `anime.extensionsDir`, and `anime.preferredQuality` config keys. SubMiner ships no extension repositories and performs no discovery.
- Anime playback targets Japanese audio: dub-labelled entries are skipped when the source offers an alternative, `alang` prefers Japanese, and the source's own audio and subtitle tracks are loaded into mpv (Japanese selected) instead of being discarded, so all of them can be switched from mpv's track menu.
- The primary subtitle slot stays reserved for Japanese: a source that only has, say, English subtitles gets them added with a normalized language tag (`English` → `en`) but not selected, so the regular `secondarySub` auto-load can route them to the secondary slot instead.
- The Linux x64 bridge bundle is verified and pinned, so the anime browser starts on Linux instead of refusing with "No pinned checksum for linux-x64-bundle.zip".
- The bridge bundle is fetched from the pinned release tag rather than whatever release is newest, so an upstream publish no longer breaks every install with a checksum mismatch.
