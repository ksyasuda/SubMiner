# Subagents Collaboration

Shared notes. Append-only.

- [YYYY-MM-DDTHH:MM:SSZ] [agent_id|alias] note, question, dependency, conflict, decision.
- [2026-02-19T08:21:11Z] [codex-main|planner-exec] conflict note: `docs/subagents/INDEX.md` and `docs/subagents/agents/codex-main.md` were externally updated to TASK-69 while TASK-38 work was in-flight; reconciled own row/file back to TASK-38 handoff state.
- [2026-02-20T00:01:40Z] [codex-anilist-deeplink|anilist-deeplink] preparing commit; scoping staged set to repo changes, excluding external reference dirs (vendor/yomitan-jlpt-vocab, mpv-anilist-updater).
- [2026-02-20T05:42:54Z] [codex-subtitle-bg-20260220T054247Z-h9cu|codex-subtitle-bg] short config tweak requested: update default subtitle background color; scoping to config defaults/tests only.
- [2026-02-20T05:44:45Z] [codex-subtitle-bg-20260220T054247Z-h9cu|codex-subtitle-bg] completed TASK-89; updated default subtitle background in config defaults/docs/examples/renderer CSS; config tests green.
- [2026-02-20T06:17:31Z] [codex-narrow-space-tokenizer-20260220T061716Z-p97s|codex-narrow-space-tokenizer] potential overlap notice: investigating tokenizer whitespace normalization and tests (likely `src/core/services/tokenizer-service.ts` + tests); coordinating to avoid clobber with ongoing TASK-85 refactor touches.
- [2026-02-20T06:20:07Z] [codex-narrow-space-tokenizer-20260220T061716Z-p97s|codex-narrow-space-tokenizer] completed TASK-90 fix in `src/subtitle/stages/normalize.ts`; normalize `U+200B/U+2060/U+FEFF` to spaces for tokenizer input; added regression test `src/subtitle/stages/normalize.test.ts`; targeted tokenizer suite green.
- [2026-02-20T06:35:38Z] [codex-preserve-linebreaks-20260220T063538Z-s4nd|codex-preserve-linebreaks] overlap note: touching subtitle config + renderer render path (`src/types.ts`, `src/config/*`, `src/renderer/subtitle-render.ts`, docs/config examples) to add optional preserve-line-breaks behavior while keeping default normalization unchanged.
- [2026-02-20T06:42:51Z] [codex-preserve-linebreaks-20260220T063538Z-s4nd|codex-preserve-linebreaks] completed TASK-91; added `subtitleStyle.preserveLineBreaks` config (default false), renderer token/source alignment helper to preserve visible overlay line breaks when enabled, config+renderer tests green.
- [2026-02-20T09:24:08Z] [codex-bundle-config-example-20260220T092408Z-a1b2|codex-bundle-config-example] conflict note: target file .github/workflows/release.yml already modified by codex-release-mpv-plugin session; applying minimal additive delta for config example bundling only.
