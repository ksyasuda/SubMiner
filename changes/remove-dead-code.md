type: internal
area: cleanup

- Removed dead code: the unused launcher Jimaku client (the `SUBMINER_JIMAKU_*` env vars it read were documented but had no effect), the mpv plugin's hover-highlight module (no sender since Overlay 2.0; custom mpv scripts that still send `subminer-hover-token` or `yomipv-hover-token` are now ignored), the never-read `anilist.characterDictionary.refreshTtlHours` / `evictionPolicy` config keys, unused root dependencies (eslint, geist fonts), and a committed draft plan doc.
- Typecheck now runs with `noUnusedLocals`; removed the unused locals, imports, and the Windows overlay focus-handoff values they fed, which were computed but never used.
