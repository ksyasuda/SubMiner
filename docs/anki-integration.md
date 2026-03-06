# Anki Integration

read_when:
- changing `src/anki-integration.ts`
- changing Anki transport/config hot-reload behavior
- tracing note update, field grouping, or proxy ownership

## Ownership

- `src/anki-integration.ts`: thin facade; wires dependencies; exposes public Anki API used by runtime/services.
- `src/anki-integration/runtime.ts`: normalized config state, polling-vs-proxy transport lifecycle, runtime config patch handling.
- `src/anki-integration/card-creation.ts`: sentence/audio card creation and clipboard update flow.
- `src/anki-integration/note-update-workflow.ts`: enrich newly added notes.
- `src/anki-integration/field-grouping.ts`: preview/build helpers for Kiku field grouping.
- `src/anki-integration/field-grouping-workflow.ts`: auto/manual merge execution.
- `src/anki-integration/anki-connect-proxy.ts`: local proxy transport for post-add enrichment.
- `src/anki-integration/known-word-cache.ts`: known-word cache lifecycle and persistence.

## Refactor seam

`AnkiIntegrationRuntime` owns the cluster that previously mixed:

- config normalization/defaulting
- polling vs proxy startup/shutdown
- transport restart decisions during runtime patches
- known-word cache lifecycle toggles tied to config changes

Keep new orchestration work in `runtime.ts` when it changes process-level Anki state. Keep note/card behavior in the workflow/service modules.
