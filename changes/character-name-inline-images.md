type: added
area: subtitles

- Added optional inline AniList portraits for character-name subtitle matches, including automatic refresh of cached character dictionary snapshots that do not contain portrait data.
- Scoped manual AniList overrides by parent media directory, so separate season folders can keep separate character dictionary selections.
- Fixed large character dictionary imports by serving the merged ZIP through a local URL when supported, with a base64 fallback for older bundled Yomitan builds.
- Allowed subtitle overlay data image sources so inline character portraits render instead of showing a broken image icon.
