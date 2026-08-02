# Renderer Performance Optimizations

**Date:** 2026-03-15
**Status:** Draft

## Goal

Minimize the time between a subtitle line appearing and annotations being displayed. Three optimizations target different pipeline stages to achieve this.

## Current Pipeline (Warm State)

```text
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

```text
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

A cue parser extracts both timing and text content from subtitle files for prefetching.

**Parsed cue structure:**
```typescript
interface SubtitleCue {
  startTime: number; // seconds
  endTime: number; // seconds
  text: string; // plain text, decoded from the source format
}
```

**Supported formats:**
- SRT/VTT: Regex-based parsing of timing lines + text content between timing blocks.
- ASS: Parse `[Events]` section, extract `Dialogue:` lines, read the field order from the `Format:` row, and take everything after the Text field index as the text (Text can itself contain commas).

**ASS decoding.** The parser is where ASS text is decoded, once, via `assToPlainText()` in `src/core/services/ass-text.ts`. That decoder mirrors mpv's `ass_to_plaintext` so a cue read from a file reads identically to the same line arriving live on `sub-text`: `{...}` override blocks are markup, `\pN … \p0` vector drawing runs are dropped rather than shown as text, `\N`/`\n`/`\h` are the only escapes (`\{`, `\}` and `\\` are not), and an unclosed `{` is rendered verbatim. Every layer downstream — renderer, timing tracker, tokenizer, tokenization cache keys — receives plain text and uses `normalizePlainSubtitleText()` for whitespace only, so nothing decodes the same string twice and one authored line always maps to one cache key.

**Duplicate collapsing.** Typeset scripts emit one `Dialogue:` event per animation frame, plus layered copies of the same line. The parser collapses identical text over an identical span unconditionally, and collapses contiguous same-text runs of at least three events when the run looks like an animation. For ASS that means shared style and actor plus authoring evidence: a temporal tag (`\t`, `\move`, `\k`/`\kf`/`\ko`/`\K`, or anything wrapped in `\t(...)`), an animated `Effect` column (`Karaoke`, `Banner`, `Scroll`), or override values that change across the run. Static tags shared by every event (`\pos`, an identical `\clip`) are not evidence. SRT/VTT carry no such metadata, so there collapsing needs at least five contiguous events all under 0.1s — the frame timing left behind by ASS-to-SRT conversion. The parser keeps this authoring metadata (style, actor, layer, `Effect`, parsed override commands, source order) private; `parseSubtitleCues()` returns only `SubtitleCue`.

#### Prefetch Service Lifecycle

1. **Activation trigger:** When a subtitle track is activated (or changes), check if it's external via MPV's `track-list` property. If `external === true`, read the file via `external-filename` using the existing `loadSubtitleSourceText` infrastructure.
2. **Parse phase:** Parse all cues from the file content. Sort by start time. Store as an ordered array.
3. **Priority window:** Determine the current playback position. Identify the next 10 cues as the priority window.
4. **Priority tokenization:** Tokenize the priority window cues sequentially, storing results into the `SubtitleProcessingController`'s tokenization cache.
5. **Background tokenization:** After the priority window is done, tokenize remaining cues working forward from the current position, then wrapping around to cover earlier cues. The prefetcher stops once it has tokenized all cues or the cache is full (whichever comes first) to avoid wasteful eviction churn. For files with more cues than the cache limit, background tokenization focuses on cues ahead of the current position.
6. **Seek handling:** On seek, re-compute the priority window from the new position. A seek is detected by observing MPV's `time-pos` property and checking if the delta from the last observed position exceeds a threshold (e.g., > 3 seconds forward or any backward jump). The current in-flight tokenization finishes naturally, then the new priority window takes over.
7. **Teardown:** When the subtitle track changes or playback ends, stop all prefetch work and discard state.

#### Live Priority

The prefetcher and live subtitle handler share the Yomitan parser (single-threaded IPC). Live subtitle requests must always take priority. The prefetcher:

- Checks a `paused` flag before each cue tokenization. The live handler sets `paused = true` on subtitle change and clears it after emission.
- Yields between each background cue tokenization (via `setTimeout(0)` or equivalent) so the live handler can set the pause flag between cues.
- When paused, the prefetcher waits (polling the flag on a short interval or awaiting a resume signal) before continuing with the next cue.

#### Cache Integration

The prefetcher calls the same `tokenizeSubtitle` function used by live processing to produce `SubtitleData` results, then stores them into the existing `SubtitleProcessingController` tokenization cache via a new method:

```typescript
// New methods on SubtitleProcessingController
preCacheTokenization: (text: string, data: SubtitleData) => void;
isCacheFull: () => boolean;
```

`preCacheTokenization` uses the same `setCachedTokenization` logic internally (LRU eviction, Map-based storage). `isCacheFull` returns `true` when the cache has reached its limit, allowing the prefetcher to stop background tokenization and avoid wasteful eviction churn.

#### Cache Invalidation

When the user marks a word as known (or any event triggers `invalidateTokenizationCache()`), all cached results are cleared -- including prefetched ones, since they share the same cache. After invalidation, the prefetcher re-computes the priority window from the current playback position and re-tokenizes those cues to restore warm cache state.

#### Error Handling

If the subtitle file is malformed or partially parseable, the cue parser uses what it can extract. A file that yields zero cues disables prefetching silently (falls back to live-only processing). Encoding errors from `loadSubtitleSourceText` are caught and logged; prefetching is skipped for that track.

#### Integration Points

- **MPV property subscriptions:** Needs `track-list` (to detect external subtitle file path) and `time-pos` (to track playback position for window calculation and seek detection).
- **File loading:** Uses existing `loadSubtitleSourceText` dependency.
- **Tokenization:** Calls the same `tokenizeSubtitle` function used by live processing.
- **Cache:** Writes into `SubtitleProcessingController`'s cache.
- **Cache invalidation:** Listens for cache invalidation events to re-prefetch the priority window.

### Files Affected

- **New:** `src/core/services/subtitle-prefetch.ts` -- the prefetch service
- **New:** `src/core/services/subtitle-cue-parser.ts` -- SRT/VTT/ASS cue parser (text + timing)
- **Modified:** `src/core/services/subtitle-processing-controller.ts` -- expose `preCacheTokenization` method
- **Modified:** `src/main.ts` -- wire up the prefetch service, listen to track changes

---

## Optimization 2: Batched Annotation Pass

### Summary

Collapse the 4 sequential annotation passes (`applyKnownWordMarking` -> `applyFrequencyMarking` -> `applyJlptMarking` -> `markNPlusOneTargets`) into a single iteration over the token array, followed by N+1 marking.

**Important context:** Frequency rank _values_ (`token.frequencyRank`) are already assigned at the parser level by `applyFrequencyRanks()` in `tokenizer.ts`, before the annotation stage is called. The annotation stage's `applyFrequencyMarking` only performs POS-based _filtering_ -- clearing `frequencyRank` to `undefined` for tokens that should be excluded (particles, noise tokens, etc.) and normalizing valid ranks. This optimization does not change the parser-level frequency rank assignment; it only batches the annotation-level filtering.

### Current Flow (4 passes, 4 array copies)

```text
tokens (already have frequencyRank values from parser-level applyFrequencyRanks)
  -> applyKnownWordMarking()   // .map() -> new array
  -> applyFrequencyMarking()   // .map() -> new array (POS-based filtering only)
  -> applyJlptMarking()        // .map() -> new array
  -> markNPlusOneTargets()     // .map() -> new array
```

### Dependency Analysis

All annotations either depend on MeCab POS data or benefit from running after it:
- **Known word marking:** Needs base tokens (surface/headword). No POS dependency, but no reason to run separately.
- **Frequency filtering:** Uses `pos1Exclusions` and `pos2Exclusions` to clear frequency ranks on excluded tokens (particles, noise). Depends on MeCab POS data.
- **JLPT marking:** Uses `shouldIgnoreJlptForMecabPos1` to filter. Depends on MeCab POS data.
- **N+1 marking:** Uses POS exclusion sets to filter candidates. Depends on known word status + MeCab POS.

Since frequency filtering and JLPT marking both depend on POS data from MeCab enrichment, and MeCab enrichment already happens before the annotation stage, all four can run in a single pass after MeCab completes.

### New Flow (1 pass + N+1)

```typescript
function annotateTokens(tokens, deps, options): MergedToken[] {
  const pos1Exclusions = resolvePos1Exclusions(options);
  const pos2Exclusions = resolvePos2Exclusions(options);

  // Single pass: known word + frequency filtering + JLPT computed together
  const annotated = tokens.map((token) => {
    const isKnown = nPlusOneEnabled
      ? token.isKnown || computeIsKnown(token, deps)
      : false;

    // Filter frequency rank using POS exclusions (rank values already set at parser level)
    const frequencyRank = frequencyEnabled
      ? filterFrequencyRank(token, pos1Exclusions, pos2Exclusions)
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

- The individual `applyKnownWordMarking`, `applyFrequencyMarking`, `applyJlptMarking` functions are refactored into per-token computation helpers (pure functions that compute a single field). The frequency helper is named `filterFrequencyRank` to clarify it performs POS-based exclusion, not rank computation.
- The `annotateTokens` orchestrator runs one `.map()` call that invokes all three helpers per token.
- `markNPlusOneTargets` remains a separate pass because it needs the full array with `isKnown` set (it examines sentence-level context).
- The parser-level `applyFrequencyRanks()` call in `tokenizer.ts` is unchanged -- it remains a separate step outside the annotation stage.
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
3. Replace all `innerHTML = ''` calls with `root.replaceChildren()` to avoid the HTML parser invocation on clear. This applies to `renderSubtitle` (primary subtitle root), `renderSecondarySub` (secondary subtitle root), and `renderCharacterLevel` if applicable.
4. Everything else stays the same (setting className, textContent, dataset, appending to fragment).

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
