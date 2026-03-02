---
id: TASK-81
title: 'Tokenization performance: disable Yomitan MeCab parser, gate local MeCab init, and add persistent MeCab process'
status: Done
assignee: []
created_date: '2026-03-02 07:44'
updated_date: '2026-03-02 20:37'
labels: []
dependencies: []
priority: high
ordinal: 9001
---

## Description

<!-- SECTION:DESCRIPTION:BEGIN -->

Reduce subtitle annotation latency by:
- disabling Yomitan-side MeCab parser requests (`useMecabParser=false`);
- initializing local MeCab only when POS-dependent annotations are enabled (N+1 / JLPT / frequency);
- replacing per-line local MeCab process spawning with a persistent parser process that auto-shuts down after idle time and restarts on demand.

<!-- SECTION:DESCRIPTION:END -->

## Acceptance Criteria

<!-- AC:BEGIN -->

- [x] #1 Yomitan parse requests disable MeCab parser path.
- [x] #2 MeCab warmup/init is skipped when all POS-dependent annotation toggles are off.
- [x] #3 Local MeCab tokenizer uses persistent process across subtitle lines.
- [x] #4 Persistent MeCab process auto-shuts down after idle timeout and restarts on next tokenize activity.
- [x] #5 Tests cover parser flag, warmup gating, and persistent MeCab lifecycle behavior.

<!-- AC:END -->

## Final Summary

<!-- SECTION:FINAL_SUMMARY:BEGIN -->

Implemented tokenizer latency optimizations:
- switched Yomitan parse requests to `useMecabParser: false`;
- added annotation-aware MeCab initialization gating in runtime warmup flow;
- added persistent local MeCab process (default idle shutdown: 30s) with queued requests, retry-on-process-end, idle auto-shutdown, and automatic restart on new work;
- added regression tests for Yomitan parse flag, MeCab warmup gating, and persistent/idle lifecycle behavior;
- fixed tokenization warmup gate so first-use warmup completion is sticky (`tokenizationWarmupCompleted`) and sequential `tokenizeSubtitle` calls no longer re-run Yomitan/dictionary warmup path;
- added regression coverage in `src/main/runtime/composers/mpv-runtime-composer.test.ts` for sequential tokenize calls (`warmup` side effects run once);
- post-review critical fix: treat Yomitan default-profile Anki server sync `no-change` as successful check, so `lastSyncedYomitanAnkiServer` is cached and expensive sync checks do not repeat on every subtitle line;
- added regression assertion in `src/core/services/tokenizer/yomitan-parser-runtime.test.ts` for `updated: false` path returning sync success;
- post-review performance fix: refactored POS enrichment to pre-index MeCab tokens by surface/start position and use sliding overlap window + binary-search cursor fallback, removing repeated full MeCab scans per token (`O(n*m)` hotspot);
- added regression test in `src/core/services/tokenizer/parser-enrichment-stage.test.ts` that fails on repeated distant-token scan access and passes with indexed lookup;
- validated with targeted tests and `tsc --noEmit`.

<!-- SECTION:FINAL_SUMMARY:END -->
