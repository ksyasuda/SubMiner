# Backlog

Purpose: lightweight repo-local task board. Seeded with current testing / coverage work.

Status keys:

- `todo`: not started
- `doing`: in progress
- `blocked`: waiting
- `done`: shipped

Priority keys:

- `P0`: urgent / release-risk
- `P1`: high value
- `P2`: useful cleanup
- `P3`: nice-to-have

## Active

| ID     | Pri | Status | Area           | Title                                               |
| ------ | --- | ------ | -------------- | --------------------------------------------------- |
| SM-013 | P1  | doing  | review-followup | Address PR #36 CodeRabbit action items              |

## Ready

| ID     | Pri | Status | Area              | Title                                                            |
| ------ | --- | ------ | ----------------- | ---------------------------------------------------------------- |
| SM-001 | P1  | todo   | launcher          | Add tests for CLI parser and args normalizer                     |
| SM-002 | P1  | todo   | immersion-tracker | Backfill tests for uncovered query exports                       |
| SM-003 | P1  | todo   | anki              | Add focused field-grouping service + merge edge-case tests       |
| SM-004 | P2  | todo   | tests             | Extract shared test utils for deps factories and polling helpers |
| SM-005 | P2  | todo   | tests             | Strengthen weak assertions in app-ready and IPC tests            |
| SM-006 | P2  | todo   | tests             | Break up monolithic youtube-flow and subtitle-sidebar tests      |
| SM-007 | P2  | todo   | anilist           | Add tests for AniList rate limiter                               |
| SM-008 | P3  | todo   | subtitles         | Add core subtitle-position persistence/path tests                |
| SM-009 | P3  | todo   | tokenizer         | Add tests for JLPT token filter                                  |
| SM-010 | P1  | todo   | immersion-tracker | Refactor storage + immersion-tracker service into focused modules    |
| SM-011 | P1  | done   | tests             | Add coverage reporting for maintained test lanes                 |
| SM-012 | P2  | done   | config/runtime    | Replace JSON serialize-clone helpers with structured cloning     |

## Icebox

None.

## Ticket Details

### SM-001

Title: Add tests for CLI parser and args normalizer
Priority: P1
Status: done
Scope:

- `launcher/config/cli-parser-builder.ts`
- `launcher/config/args-normalizer.ts`
  Acceptance:
- root options parsing covered
- subcommand routing covered
- invalid action / invalid log level / invalid backend cases covered
- target classification covered: file, directory, URL, invalid

### SM-002

Title: Backfill tests for uncovered query exports
Priority: P1
Status: todo
Scope:

- `src/core/services/immersion-tracker/query-*.ts`
  Targets:
- headword helpers
- anime/media detail helpers not covered by existing wrapper tests
- lexical detail / appearance helpers
- maintenance helpers beyond `deleteSession` and `upsertCoverArt`
  Acceptance:
- every exported query helper either directly tested or explicitly justified as covered elsewhere
- at least one focused regression per complex SQL branch / aggregation branch

### SM-003

Title: Add focused field-grouping service + merge edge-case tests
Priority: P1
Status: todo
Scope:

- `src/anki-integration/field-grouping.ts`
- `src/anki-integration/field-grouping-merge.ts`
  Acceptance:
- auto/manual/disabled flow branches covered
- duplicate-card preview failure path covered
- merge edge cases covered: empty fields, generated media fallback, strict grouped spans, audio synchronization

### SM-004

Title: Extract shared test utils for deps factories and polling helpers
Priority: P2
Status: todo
Scope:

- common `makeDeps` / `createDeps` helpers
- common `waitForCondition`
  Acceptance:
- shared helper module added
- at least 3 duplicated polling helpers removed
- at least 5 duplicated deps factories consolidated or clearly prepared for follow-up migration

### SM-005

Title: Strengthen weak assertions in app-ready and IPC tests
Priority: P2
Status: todo
Scope:

- `src/core/services/app-ready.test.ts`
- `src/core/services/ipc.test.ts`
  Acceptance:
- replace broad `assert.ok(...)` presence checks with exact value / order assertions where expected value known
- handler registration tests assert channel-specific behavior, not only existence

### SM-006

Title: Break up monolithic youtube-flow and subtitle-sidebar tests
Priority: P2
Status: todo
Scope:

- `src/main/runtime/youtube-flow.test.ts`
- `src/renderer/modals/subtitle-sidebar.test.ts`
  Acceptance:
- reduce single-test breadth
- split largest tests into focused cases by behavior
- keep semantics unchanged

### SM-007

Title: Add tests for AniList rate limiter
Priority: P2
Status: todo
Scope:

- `src/core/services/anilist/rate-limiter.ts`
  Acceptance:
- capacity-window wait behavior covered
- `x-ratelimit-remaining` + reset handling covered
- `retry-after` handling covered

### SM-008

Title: Add core subtitle-position persistence/path tests
Priority: P3
Status: todo
Scope:

- `src/core/services/subtitle-position.ts`
  Acceptance:
- save/load persistence covered
- fallback behavior covered
- path normalization behavior covered for URL vs local target

### SM-009

Title: Add tests for JLPT token filter
Priority: P3
Status: todo
Scope:

- `src/core/services/jlpt-token-filter.ts`
  Acceptance:
- excluded term membership covered
- ignored POS1 membership covered
- exported list / entry consistency covered

### SM-010

Title: Refactor storage + immersion-tracker service into focused layers without API changes
Priority: P1
Status: todo
Scope:

- `src/core/database/storage/storage.ts`
- `src/core/database/storage/schema.ts`
- `src/core/database/storage/cover-blob.ts`
- `src/core/database/storage/records.ts`
- `src/core/database/storage/write-path.ts`
- `src/core/services/immersion-tracker/youtube.ts`
- `src/core/services/immersion-tracker/youtube-manager.ts`
- `src/core/services/immersion-tracker/write-queue.ts`
- `src/core/services/immersion-tracker/immersion-tracker-service.ts`

Acceptance:

- behavior and public API remain unchanged for all callers
- `storage.ts` responsibilities split into DDL/migrations, cover blob helpers, record CRUD, and write-path execution
- `immersion-tracker-service.ts` reduces to session state, media change orchestration, query proxies, and lifecycle
- YouTube code split into pure utilities, a stateful manager (`YouTubeManager`), and a dedicated write queue (`WriteQueue`)
- removed `storage.ts` is replaced with focused modules and updated imports
- no API or migration regressions; existing tests for trackers/storage coverage remain green or receive focused updates

### SM-011

Title: Add coverage reporting for maintained test lanes
Priority: P1
Status: done
Scope:

- `package.json`
- CI workflow files under `.github/`
- `docs/workflow/verification.md`
  Acceptance:
- at least one maintained test lane emits machine-readable coverage output
- CI surfaces coverage as an artifact, summary, or check output
- local contributor path for coverage is documented
- chosen coverage path works with Bun/TypeScript lanes already maintained by the repo
Implementation note:
- Added `bun run test:coverage:src` for the maintained source lane via a sharded coverage runner, with merged LCOV output at `coverage/test-src/lcov.info` and CI/release artifact upload as `coverage-test-src`.

### SM-012

Title: Replace JSON serialize-clone helpers with structured cloning
Priority: P2
Status: todo
Scope:

- `src/runtime-options.ts`
- `src/config/definitions.ts`
- `src/config/service.ts`
- `src/main/controller-config-update.ts`
  Acceptance:
- runtime/config clone helpers stop using `JSON.parse(JSON.stringify(...))`
- replacement preserves current behavior for plain config/runtime objects
- focused tests cover clone/merge behavior that could regress during the swap
- no new clone helper is introduced in these paths without a documented reason

Done:

- replaced JSON serialize-clone call sites in runtime/config/controller update paths with `structuredClone`
- updated focused tests and fixtures to cover detached clone behavior and guard against regressions

### SM-013

Title: Address PR #36 CodeRabbit action items
Priority: P1
Status: doing
Scope:

- `plugins/subminer-workflow/skills/subminer-change-verification/scripts/verify_subminer_change.sh`
- `scripts/subminer-change-verification.test.ts`
- `src/core/services/immersion-tracker/query-sessions.ts`
- `src/core/services/immersion-tracker/query-trends.ts`
- `src/core/services/immersion-tracker/maintenance.ts`
- `src/main/boot/services.ts`
- `src/main/character-dictionary-runtime/zip.test.ts`
  Acceptance:
- fix valid open CodeRabbit findings on PR #36
- add focused regression coverage for behavior changes where practical
- verify touched tests plus typecheck stay green
