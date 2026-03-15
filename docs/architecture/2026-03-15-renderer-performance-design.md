# Renderer Performance Optimizations

**Date:** 2026-03-15
**Status:** Draft

## Goal

Minimize the time between a subtitle line appearing and annotations being displayed. Three optimizations target different pipeline stages to achieve this.

## Current Pipeline (Warm State)

```
MPV subtitle change (0ms)
  -> IPC to main (5ms)
  -> Cache check (2ms)
  -> [CACHE MISS] Yomitan parser (35-180ms)
  -> Parallel: MeCab enrichment (20-80ms) + Frequency lookup (15-50ms)
  -> Annotation stage: 4 sequential passes (25-70ms)
  -> IPC to renderer (10ms)
  -> DOM render: createElement per token (15-50ms)
  ─────────────────────────────────
  Total: ~200-320ms (cache miss)
  Total: ~72ms (cache hit)
```

## Target Pipeline

```
MPV subtitle change (0ms)
  -> IPC to main (5ms)
  -> Cache check (2ms)
  -> [CACHE HIT via prefetch] (0ms)
  -> IPC to renderer (10ms)
  -> DOM render: cloneNode from template (10-30ms)
  ─────────────────────────────────
  Total: ~30-50ms (prefetch-warmed, normal playback)

  [CACHE MISS, e.g. immediate seek]
  -> Yomitan parser (35-180ms)
  -> Parallel: MeCab enrichment + Frequency lookup
  -> Annotation stage: 1 batched pass (10-25ms)
  -> IPC to renderer (10ms)
  -> DOM render: cloneNode from template (10-30ms)
  ─────────────────────────────────
  Total: ~150-260ms (cache miss, still improved)
```

---

## Optimization 1: Subtitle Prefetching

### Summary

A new `SubtitlePrefetchService` parses external subtitle files and tokenizes upcoming lines in the background before they appear on screen. This converts most cache misses into cache hits during normal playback.

### Scope

External subtitle files only (SRT, VTT, ASS). Embedded subtitle tracks are out of scope since Japanese subtitles are virtually always external files.

### Architecture

#### Subtitle File Parsing

A new cue parser that extracts both timing and text content from subtitle files. The existing `parseSrtOrVttStartTimes` in `subtitle-delay-shift.ts` only extracts timing; this needs a companion that also extracts the dialogue text.

**Parsed cue structure:**
```typescript
interface SubtitleCue {
  startTime: number;  // seconds
  endTime: number;    // seconds
  text: string;       // raw subtitle text
}
```

**Supported formats:**
- SRT/VTT: Regex-based parsing of timing lines + text content between timing blocks.
- ASS: Parse `[Events]` section, extract `Dialogue:` lines, split on commas to get timing and text fields.

#### Prefetch Service Lifecycle

1. **Activation trigger:** When a subtitle track is activated (or changes), check if it's external via MPV's `track-list` property. If `external === true`, read the file via `external-filename` using the existing `loadSubtitleSourceText` infrastructure.
2. **Parse phase:** Parse all cues from the file content. Sort by start time. Store as an ordered array.
3. **Priority window:** Determine the current playback position. Identify the next 10 cues as the priority window.
4. **Priority tokenization:** Tokenize the priority window cues sequentially, storing results into the `SubtitleProcessingController`'s tokenization cache.
5. **Background tokenization:** After the priority window is done, tokenize remaining cues working forward from the current position, then wrapping around to cover earlier cues.
6. **Seek handling:** On seek (detected via playback position jump), re-compute the priority window from the new position. The current in-flight tokenization finishes naturally, then the new priority window takes over.
7. **Teardown:** When the subtitle track changes or playback ends, stop all prefetch work and discard state.

#### Live Priority

The prefetcher and live subtitle handler share the Yomitan parser (single-threaded IPC). Live subtitle requests must always take priority. The prefetcher:

- Pauses when a live subtitle change arrives.
- Resumes after the live subtitle has been processed and emitted.
- Yields between each background cue tokenization (e.g., via `setTimeout(0)` or checking a pause flag) so live processing is never blocked.

#### Cache Integration

The prefetcher writes into the existing `SubtitleProcessingController` tokenization cache. This requires exposing a method to insert pre-computed results:

```typescript
// New method on SubtitleProcessingController
preCacheTokenization: (text: string, data: SubtitleData) => void;
```

This uses the same `setCachedTokenization` logic internally (LRU eviction, Map-based storage).

#### Integration Points

- **MPV property subscriptions:** Needs `track-list` (to detect external subtitle file path) and `time-pos` or `sub-start`/`sub-end` (to track playback position for window calculation).
- **File loading:** Uses existing `loadSubtitleSourceText` dependency.
- **Tokenization:** Calls the same `tokenizeSubtitle` function used by live processing.
- **Cache:** Writes into `SubtitleProcessingController`'s cache.

### Files Affected

- **New:** `src/core/services/subtitle-prefetch.ts` -- the prefetch service
- **New:** `src/core/services/subtitle-cue-parser.ts` -- SRT/VTT/ASS cue parser (text + timing)
- **Modified:** `src/core/services/subtitle-processing-controller.ts` -- expose `preCacheTokenization` method
- **Modified:** `src/main.ts` -- wire up the prefetch service, listen to track changes

---

## Optimization 2: Batched Annotation Pass

### Summary

Collapse the 4 sequential annotation passes (`applyKnownWordMarking` -> `applyFrequencyMarking` -> `applyJlptMarking` -> `markNPlusOneTargets`) into a single iteration over the token array, followed by N+1 marking.

### Current Flow (4 passes, 4 array copies)

```
tokens
  -> applyKnownWordMarking()   // .map() -> new array
  -> applyFrequencyMarking()   // .map() -> new array
  -> applyJlptMarking()        // .map() -> new array
  -> markNPlusOneTargets()     // .map() -> new array
```

### Dependency Analysis

All annotations either depend on MeCab POS data or benefit from running after it:
- **Known word marking:** Needs base tokens (surface/headword). No POS dependency, but no reason to run separately.
- **Frequency marking:** Uses `pos1Exclusions` and `pos2Exclusions` to filter out particles and noise tokens. Depends on MeCab POS data.
- **JLPT marking:** Uses `shouldIgnoreJlptForMecabPos1` to filter. Depends on MeCab POS data.
- **N+1 marking:** Uses POS exclusion sets to filter candidates. Depends on known word status + MeCab POS.

Since frequency and JLPT filtering both depend on POS data from MeCab enrichment, and MeCab enrichment already happens before the annotation stage, all four can run in a single pass after MeCab completes.

### New Flow (1 pass + N+1)

```typescript
function annotateTokens(tokens, deps, options): MergedToken[] {
  const pos1Exclusions = resolvePos1Exclusions(options);
  const pos2Exclusions = resolvePos2Exclusions(options);

  // Single pass: known word + frequency + JLPT computed together
  const annotated = tokens.map((token) => {
    const isKnown = nPlusOneEnabled
      ? token.isKnown || computeIsKnown(token, deps)
      : false;

    const frequencyRank = frequencyEnabled
      ? computeFrequencyRank(token, pos1Exclusions, pos2Exclusions)
      : undefined;

    const jlptLevel = jlptEnabled
      ? computeJlptLevel(token, deps.getJlptLevel)
      : undefined;

    return { ...token, isKnown, frequencyRank, jlptLevel };
  });

  // N+1 must run after known word status is set for all tokens
  if (nPlusOneEnabled) {
    return markNPlusOneTargets(annotated, minSentenceWords, pos1Exclusions, pos2Exclusions);
  }

  return annotated;
}
```

### What Changes

- The individual `applyKnownWordMarking`, `applyFrequencyMarking`, `applyJlptMarking` functions are refactored into per-token computation helpers (pure functions that compute a single field).
- The `annotateTokens` orchestrator runs one `.map()` call that invokes all three helpers per token.
- `markNPlusOneTargets` remains a separate pass because it needs the full array with `isKnown` set (it examines sentence-level context).
- Net: 4 array copies + 4 iterations become 1 array copy + 1 iteration + N+1 pass.

### Expected Savings

~15-45ms saved (3 fewer array allocations + 3 fewer full iterations). Annotation drops from ~25-70ms to ~10-25ms.

### Files Affected

- **Modified:** `src/core/services/tokenizer/annotation-stage.ts` -- refactor into batched single-pass

---

## Optimization 3: DOM Template Pooling

### Summary

Replace `document.createElement('span')` calls in the renderer with `templateSpan.cloneNode(false)` from a pre-created template element.

### Current Behavior

In `renderWithTokens` (`subtitle-render.ts`), each render cycle:
1. Clears DOM with `innerHTML = ''`
2. Creates a `DocumentFragment`
3. Calls `document.createElement('span')` for each token (~10-15 per subtitle)
4. Sets `className`, `textContent`, `dataset.*` individually
5. Appends fragment to root

### New Behavior

1. At renderer initialization (`createSubtitleRenderer`), create a single template:
   ```typescript
   const templateSpan = document.createElement('span');
   ```
2. In `renderWithTokens`, replace every `document.createElement('span')` with:
   ```typescript
   const span = templateSpan.cloneNode(false) as HTMLSpanElement;
   ```
3. Everything else stays the same (setting className, textContent, dataset, appending to fragment).

### Why cloneNode Over Full Node Recycling

Full recycling (collecting old nodes, clearing attributes, reusing them) requires carefully resetting every `dataset.*` property that might have been set on a previous render. This is error-prone -- a stale `data-frequency-rank` from a previous subtitle appearing on a new token would cause incorrect styling. `cloneNode(false)` on a bare template is nearly as fast and produces a clean node every time.

### Expected Savings

`cloneNode(false)` is ~2-3x faster than `createElement` in most browser engines. For 10-15 tokens per subtitle: ~3-8ms saved per render cycle.

### Files Affected

- **Modified:** `src/renderer/subtitle-render.ts` -- template creation + cloneNode usage

---

## Combined Impact Summary

| Scenario | Before | After | Improvement |
|----------|--------|-------|-------------|
| Normal playback (prefetch-warmed) | ~200-320ms | ~30-50ms | ~80-85% |
| Cache hit (repeated subtitle) | ~72ms | ~55-65ms | ~10-20% |
| Cache miss (immediate seek) | ~200-320ms | ~150-260ms | ~20-25% |

---

## Files Summary

### New Files
- `src/core/services/subtitle-prefetch.ts`
- `src/core/services/subtitle-cue-parser.ts`

### Modified Files
- `src/core/services/subtitle-processing-controller.ts` (expose `preCacheTokenization`)
- `src/core/services/tokenizer/annotation-stage.ts` (batched single-pass)
- `src/renderer/subtitle-render.ts` (template cloneNode)
- `src/main.ts` (wire up prefetch service)

### Test Files
- New tests for subtitle cue parser (SRT, VTT, ASS formats)
- New tests for subtitle prefetch service (priority window, seek, pause/resume)
- Updated tests for annotation stage (same behavior, new implementation)
- Updated tests for subtitle render (template cloning)
