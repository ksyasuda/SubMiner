type: fixed
area: character-dictionary

- Reused cached character-dictionary media matches so loading a title with an existing snapshot no longer sends another AniList search request.
- Block the character dictionary manager when annotations are disabled, with a notice through the configured OSD/system notification surfaces.
- Added surname honorific matches for Japanese localized character aliases embedded in AniList alternative names (e.g. Korean-source characters with Japanese names in parentheses), and refresh cached snapshots so those aliases are regenerated.
- Use `subtitleStyle.nameMatchEnabled` as the only switch for character-dictionary sync/builds, hiding the legacy `anilist.characterDictionary.enabled` option.
- Forward character dictionary manager session-action keybindings to the mpv plugin.
