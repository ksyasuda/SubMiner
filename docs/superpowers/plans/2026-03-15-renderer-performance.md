# Renderer Performance Optimizations Implementation Plan

> **For agentic workers:** REQUIRED: Use superpowers:subagent-driven-development (if subagents available) or superpowers:executing-plans to implement this plan. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Minimize subtitle-to-annotation latency via prefetching, batched annotation, and DOM template pooling.

**Architecture:** Three independent optimizations targeting different pipeline stages: (1) a subtitle prefetch service that parses external subtitle files and tokenizes upcoming lines in the background, (2) collapsing 4 sequential annotation passes into a single loop, and (3) using `cloneNode(false)` from a template span instead of `createElement`.

**Tech Stack:** TypeScript, Electron, Bun test runner (`node:test` + `node:assert/strict`)

**Spec:** `docs/architecture/2026-03-15-renderer-performance-design.md`

**Test command:** `bun test <file>`

---

## File Structure

### New Files
| File | Responsibility |
|------|---------------|
| `src/core/services/subtitle-cue-parser.ts` | Parse SRT/VTT/ASS files into `SubtitleCue[]` (timing + text) |
| `src/core/services/subtitle-cue-parser.test.ts` | Tests for cue parser |
| `src/core/services/subtitle-prefetch.ts` | Background tokenization service with priority window + seek handling |
| `src/core/services/subtitle-prefetch.test.ts` | Tests for prefetch service |

### Modified Files
| File | Change |
|------|--------|
| `src/core/services/subtitle-processing-controller.ts` | Add `preCacheTokenization` + `isCacheFull` to public interface |
| `src/core/services/subtitle-processing-controller.test.ts` | Add tests for new methods |
| `src/core/services/tokenizer/annotation-stage.ts` | Refactor 4 passes into 1 batched pass + N+1 |
| `src/core/services/tokenizer/annotation-stage.test.ts` | Existing tests must still pass (behavioral equivalence) |
| `src/renderer/subtitle-render.ts` | `cloneNode` template + `replaceChildren()` |
| `src/main.ts` | Wire up prefetch service |

---

## Chunk 1: Batched Annotation Pass + DOM Template Pooling

### Task 1: Extend SubtitleProcessingController with `preCacheTokenization` and `isCacheFull`

**Files:**
- Modify: `src/core/services/subtitle-processing-controller.ts:10-14` (interface), `src/core/services/subtitle-processing-controller.ts:112-133` (return object)
- Test: `src/core/services/subtitle-processing-controller.test.ts`

- [ ] **Step 1: Write failing tests for `preCacheTokenization` and `isCacheFull`**

Add to end of `src/core/services/subtitle-processing-controller.test.ts`:

```typescript
test('preCacheTokenization stores entry that is returned on next subtitle change', async () => {
  const emitted: SubtitleData[] = [];
  let tokenizeCalls = 0;
  const controller = createSubtitleProcessingController({
    tokenizeSubtitle: async (text) => {
      tokenizeCalls += 1;
      return { text, tokens: [] };
    },
    emitSubtitle: (payload) => emitted.push(payload),
  });

  controller.preCacheTokenization('予め', { text: '予め', tokens: [] });
  controller.onSubtitleChange('予め');
  await flushMicrotasks();

  assert.equal(tokenizeCalls, 0, 'should not call tokenize when pre-cached');
  assert.deepEqual(emitted, [{ text: '予め', tokens: [] }]);
});

test('isCacheFull returns false when cache is below limit', () => {
  const controller = createSubtitleProcessingController({
    tokenizeSubtitle: async (text) => ({ text, tokens: null }),
    emitSubtitle: () => {},
  });

  assert.equal(controller.isCacheFull(), false);
});

test('isCacheFull returns true when cache reaches limit', async () => {
  const controller = createSubtitleProcessingController({
    tokenizeSubtitle: async (text) => ({ text, tokens: [] }),
    emitSubtitle: () => {},
  });

  // Fill cache to the 256 limit
  for (let i = 0; i < 256; i += 1) {
    controller.preCacheTokenization(`line-${i}`, { text: `line-${i}`, tokens: [] });
  }

  assert.equal(controller.isCacheFull(), true);
});
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `bun test src/core/services/subtitle-processing-controller.test.ts`
Expected: FAIL — `preCacheTokenization` and `isCacheFull` are not defined on the controller interface.

- [ ] **Step 3: Add `preCacheTokenization` and `isCacheFull` to the interface and implementation**

In `src/core/services/subtitle-processing-controller.ts`, update the interface (line 10-14):

```typescript
export interface SubtitleProcessingController {
  onSubtitleChange: (text: string) => void;
  refreshCurrentSubtitle: (textOverride?: string) => void;
  invalidateTokenizationCache: () => void;
  preCacheTokenization: (text: string, data: SubtitleData) => void;
  isCacheFull: () => boolean;
}
```

Add to the return object (after `invalidateTokenizationCache` at line 130-132):

```typescript
    invalidateTokenizationCache: () => {
      tokenizationCache.clear();
    },
    preCacheTokenization: (text: string, data: SubtitleData) => {
      setCachedTokenization(text, data);
    },
    isCacheFull: () => {
      return tokenizationCache.size >= SUBTITLE_TOKENIZATION_CACHE_LIMIT;
    },
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `bun test src/core/services/subtitle-processing-controller.test.ts`
Expected: All tests PASS.

- [ ] **Step 5: Commit**

```bash
git add src/core/services/subtitle-processing-controller.ts src/core/services/subtitle-processing-controller.test.ts
git commit -m "feat: add preCacheTokenization and isCacheFull to SubtitleProcessingController"
```

---

### Task 2: Batch annotation passes into a single loop

This refactors `annotateTokens` in `annotation-stage.ts` to perform known-word marking, frequency filtering, and JLPT marking in a single `.map()` call instead of 3 separate passes. `markNPlusOneTargets` remains a separate pass (needs full array with `isKnown` set).

**Files:**
- Modify: `src/core/services/tokenizer/annotation-stage.ts:448-502`
- Test: `src/core/services/tokenizer/annotation-stage.test.ts` (existing tests must pass unchanged)

- [ ] **Step 1: Run existing annotation tests to capture baseline**

Run: `bun test src/core/services/tokenizer/annotation-stage.test.ts`
Expected: All tests PASS. Note the count for regression check.

- [ ] **Step 2: Extract per-token helper functions**

These are pure functions extracted from the existing `applyKnownWordMarking`, `applyFrequencyMarking`, and `applyJlptMarking`. They compute a single field per token. Add them above the `annotateTokens` function in `annotation-stage.ts`.

Add `computeTokenKnownStatus` (extracted from `applyKnownWordMarking` lines 46-59):

```typescript
function computeTokenKnownStatus(
  token: MergedToken,
  isKnownWord: (text: string) => boolean,
  knownWordMatchMode: NPlusOneMatchMode,
): boolean {
  const matchText = resolveKnownWordText(token.surface, token.headword, knownWordMatchMode);
  return token.isKnown || (matchText ? isKnownWord(matchText) : false);
}
```

Add `filterTokenFrequencyRank` (extracted from `applyFrequencyMarking` lines 147-167):

```typescript
function filterTokenFrequencyRank(
  token: MergedToken,
  pos1Exclusions: ReadonlySet<string>,
  pos2Exclusions: ReadonlySet<string>,
): number | undefined {
  if (isFrequencyExcludedByPos(token, pos1Exclusions, pos2Exclusions)) {
    return undefined;
  }

  if (typeof token.frequencyRank === 'number' && Number.isFinite(token.frequencyRank)) {
    return Math.max(1, Math.floor(token.frequencyRank));
  }

  return undefined;
}
```

Add `computeTokenJlptLevel` (extracted from `applyJlptMarking` lines 428-446):

```typescript
function computeTokenJlptLevel(
  token: MergedToken,
  getJlptLevel: (text: string) => JlptLevel | null,
): JlptLevel | undefined {
  if (!isJlptEligibleToken(token)) {
    return undefined;
  }

  const primaryLevel = getCachedJlptLevel(resolveJlptLookupText(token), getJlptLevel);
  const fallbackLevel =
    primaryLevel === null ? getCachedJlptLevel(token.surface, getJlptLevel) : null;

  const level = primaryLevel ?? fallbackLevel ?? token.jlptLevel;
  return level ?? undefined;
}
```

- [ ] **Step 3: Rewrite `annotateTokens` to use single-pass batching**

Replace the `annotateTokens` function body (lines 448-502) with:

```typescript
export function annotateTokens(
  tokens: MergedToken[],
  deps: AnnotationStageDeps,
  options: AnnotationStageOptions = {},
): MergedToken[] {
  const pos1Exclusions = resolvePos1Exclusions(options);
  const pos2Exclusions = resolvePos2Exclusions(options);
  const nPlusOneEnabled = options.nPlusOneEnabled !== false;
  const frequencyEnabled = options.frequencyEnabled !== false;
  const jlptEnabled = options.jlptEnabled !== false;

  // Single pass: compute known word status, frequency filtering, and JLPT level together
  const annotated = tokens.map((token) => {
    const isKnown = nPlusOneEnabled
      ? computeTokenKnownStatus(token, deps.isKnownWord, deps.knownWordMatchMode)
      : false;

    const frequencyRank = frequencyEnabled
      ? filterTokenFrequencyRank(token, pos1Exclusions, pos2Exclusions)
      : undefined;

    const jlptLevel = jlptEnabled
      ? computeTokenJlptLevel(token, deps.getJlptLevel)
      : undefined;

    return {
      ...token,
      isKnown,
      isNPlusOneTarget: nPlusOneEnabled ? token.isNPlusOneTarget : false,
      frequencyRank,
      jlptLevel,
    };
  });

  if (!nPlusOneEnabled) {
    return annotated;
  }

  const minSentenceWordsForNPlusOne = options.minSentenceWordsForNPlusOne;
  const sanitizedMinSentenceWordsForNPlusOne =
    minSentenceWordsForNPlusOne !== undefined &&
    Number.isInteger(minSentenceWordsForNPlusOne) &&
    minSentenceWordsForNPlusOne > 0
      ? minSentenceWordsForNPlusOne
      : 3;

  return markNPlusOneTargets(
    annotated,
    sanitizedMinSentenceWordsForNPlusOne,
    pos1Exclusions,
    pos2Exclusions,
  );
}
```

- [ ] **Step 4: Run existing tests to verify behavioral equivalence**

Run: `bun test src/core/services/tokenizer/annotation-stage.test.ts`
Expected: All tests PASS with same count as baseline.

- [ ] **Step 5: Remove dead code**

Delete the now-unused `applyKnownWordMarking` (lines 46-59), `applyFrequencyMarking` (lines 147-167), and `applyJlptMarking` (lines 428-446) functions. These are replaced by the per-token helpers.

- [ ] **Step 6: Run tests again after cleanup**

Run: `bun test src/core/services/tokenizer/annotation-stage.test.ts`
Expected: All tests PASS.

- [ ] **Step 7: Run full test suite**

Run: `bun run test`
Expected: Same results as baseline (500 pass, 1 pre-existing fail).

- [ ] **Step 8: Commit**

```bash
git add src/core/services/tokenizer/annotation-stage.ts
git commit -m "perf: batch annotation passes into single loop

Collapse applyKnownWordMarking, applyFrequencyMarking, and
applyJlptMarking into a single .map() call. markNPlusOneTargets
remains a separate pass (needs full array with isKnown set).

Eliminates 3 intermediate array allocations and 3 redundant
iterations over the token array."
```

---

### Task 3: DOM template pooling and replaceChildren

**Files:**
- Modify: `src/renderer/subtitle-render.ts:289,325` (`createElement` calls), `src/renderer/subtitle-render.ts:473,481` (`renderCharacterLevel` createElement calls), `src/renderer/subtitle-render.ts:506,555` (`innerHTML` calls)

- [ ] **Step 1: Replace `innerHTML = ''` with `replaceChildren()` in all render functions**

In `src/renderer/subtitle-render.ts`:

At line 506 in `renderSubtitle`:
```typescript
// Before:
ctx.dom.subtitleRoot.innerHTML = '';
// After:
ctx.dom.subtitleRoot.replaceChildren();
```

At line 555 in `renderSecondarySub`:
```typescript
// Before:
ctx.dom.secondarySubRoot.innerHTML = '';
// After:
ctx.dom.secondarySubRoot.replaceChildren();
```

- [ ] **Step 2: Add template span and replace `createElement('span')` with `cloneNode`**

In `renderWithTokens` (line 250), add the template as a module-level constant near the top of the file (after the type declarations around line 20):

```typescript
const SPAN_TEMPLATE = document.createElement('span');
```

Then replace the two `document.createElement('span')` calls in `renderWithTokens`:

At line 289 (sourceText branch):
```typescript
// Before:
const span = document.createElement('span');
// After:
const span = SPAN_TEMPLATE.cloneNode(false) as HTMLSpanElement;
```

At line 325 (no-sourceText branch):
```typescript
// Before:
const span = document.createElement('span');
// After:
const span = SPAN_TEMPLATE.cloneNode(false) as HTMLSpanElement;
```

Also in `renderCharacterLevel` at line 481:
```typescript
// Before:
const span = document.createElement('span');
// After:
const span = SPAN_TEMPLATE.cloneNode(false) as HTMLSpanElement;
```

- [ ] **Step 3: Run full test suite**

Run: `bun run test`
Expected: Same results as baseline. The renderer changes don't have direct unit tests (they run in Electron's renderer process), but we verify no compilation or type errors break existing tests.

- [ ] **Step 4: Commit**

```bash
git add src/renderer/subtitle-render.ts
git commit -m "perf: use cloneNode template and replaceChildren for DOM rendering

Replace createElement('span') with cloneNode(false) from a pre-created
template span. Replace innerHTML='' with replaceChildren() to avoid
HTML parser invocation on clear."
```

---

## Chunk 2: Subtitle Cue Parser

### Task 4: Create SRT/VTT cue parser

**Files:**
- Create: `src/core/services/subtitle-cue-parser.ts`
- Create: `src/core/services/subtitle-cue-parser.test.ts`

- [ ] **Step 1: Write failing tests for SRT parsing**

Create `src/core/services/subtitle-cue-parser.test.ts`:

```typescript
import assert from 'node:assert/strict';
import test from 'node:test';
import { parseSrtCues, parseAssCues, parseSubtitleCues } from './subtitle-cue-parser';
import type { SubtitleCue } from './subtitle-cue-parser';

test('parseSrtCues parses basic SRT content', () => {
  const content = [
    '1',
    '00:00:01,000 --> 00:00:04,000',
    'こんにちは',
    '',
    '2',
    '00:00:05,000 --> 00:00:08,500',
    '元気ですか',
    '',
  ].join('\n');

  const cues = parseSrtCues(content);

  assert.equal(cues.length, 2);
  assert.equal(cues[0]!.startTime, 1.0);
  assert.equal(cues[0]!.endTime, 4.0);
  assert.equal(cues[0]!.text, 'こんにちは');
  assert.equal(cues[1]!.startTime, 5.0);
  assert.equal(cues[1]!.endTime, 8.5);
  assert.equal(cues[1]!.text, '元気ですか');
});

test('parseSrtCues handles multi-line subtitle text', () => {
  const content = [
    '1',
    '00:01:00,000 --> 00:01:05,000',
    'これは',
    'テストです',
    '',
  ].join('\n');

  const cues = parseSrtCues(content);

  assert.equal(cues.length, 1);
  assert.equal(cues[0]!.text, 'これは\nテストです');
});

test('parseSrtCues handles hours in timestamps', () => {
  const content = [
    '1',
    '01:30:00,000 --> 01:30:05,000',
    'テスト',
    '',
  ].join('\n');

  const cues = parseSrtCues(content);

  assert.equal(cues[0]!.startTime, 5400.0);
  assert.equal(cues[0]!.endTime, 5405.0);
});

test('parseSrtCues handles VTT-style dot separator', () => {
  const content = [
    '1',
    '00:00:01.000 --> 00:00:04.000',
    'VTTスタイル',
    '',
  ].join('\n');

  const cues = parseSrtCues(content);

  assert.equal(cues.length, 1);
  assert.equal(cues[0]!.startTime, 1.0);
});

test('parseSrtCues returns empty array for empty content', () => {
  assert.deepEqual(parseSrtCues(''), []);
  assert.deepEqual(parseSrtCues('  \n\n  '), []);
});

test('parseSrtCues skips malformed timing lines gracefully', () => {
  const content = [
    '1',
    'NOT A TIMING LINE',
    'テスト',
    '',
    '2',
    '00:00:01,000 --> 00:00:02,000',
    '有効',
    '',
  ].join('\n');

  const cues = parseSrtCues(content);

  assert.equal(cues.length, 1);
  assert.equal(cues[0]!.text, '有効');
});
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `bun test src/core/services/subtitle-cue-parser.test.ts`
Expected: FAIL — module not found.

- [ ] **Step 3: Implement `parseSrtCues`**

Create `src/core/services/subtitle-cue-parser.ts`:

```typescript
export interface SubtitleCue {
  startTime: number;
  endTime: number;
  text: string;
}

const SRT_TIMING_PATTERN =
  /^\s*(?:(\d{1,2}):)?(\d{2}):(\d{2})[,.](\d{1,3})\s*-->\s*(?:(\d{1,2}):)?(\d{2}):(\d{2})[,.](\d{1,3})/;

function parseTimestamp(
  hours: string | undefined,
  minutes: string,
  seconds: string,
  millis: string,
): number {
  return (
    Number(hours || 0) * 3600 +
    Number(minutes) * 60 +
    Number(seconds) +
    Number(millis.padEnd(3, '0')) / 1000
  );
}

export function parseSrtCues(content: string): SubtitleCue[] {
  const cues: SubtitleCue[] = [];
  const lines = content.split(/\r?\n/);
  let i = 0;

  while (i < lines.length) {
    // Skip blank lines and cue index numbers
    const line = lines[i]!;
    const timingMatch = SRT_TIMING_PATTERN.exec(line);
    if (!timingMatch) {
      i += 1;
      continue;
    }

    const startTime = parseTimestamp(timingMatch[1], timingMatch[2]!, timingMatch[3]!, timingMatch[4]!);
    const endTime = parseTimestamp(timingMatch[5], timingMatch[6]!, timingMatch[7]!, timingMatch[8]!);

    // Collect text lines until blank line or end of file
    i += 1;
    const textLines: string[] = [];
    while (i < lines.length && lines[i]!.trim() !== '') {
      textLines.push(lines[i]!);
      i += 1;
    }

    const text = textLines.join('\n').trim();
    if (text) {
      cues.push({ startTime, endTime, text });
    }
  }

  return cues;
}
```

- [ ] **Step 4: Run SRT tests to verify they pass**

Run: `bun test src/core/services/subtitle-cue-parser.test.ts`
Expected: All SRT tests PASS.

- [ ] **Step 5: Commit**

```bash
git add src/core/services/subtitle-cue-parser.ts src/core/services/subtitle-cue-parser.test.ts
git commit -m "feat: add SRT/VTT subtitle cue parser"
```

---

### Task 5: Add ASS cue parser

**Files:**
- Modify: `src/core/services/subtitle-cue-parser.ts`
- Modify: `src/core/services/subtitle-cue-parser.test.ts`

- [ ] **Step 1: Write failing tests for ASS parsing**

Add to `src/core/services/subtitle-cue-parser.test.ts`:

```typescript
test('parseAssCues parses basic ASS dialogue lines', () => {
  const content = [
    '[Script Info]',
    'Title: Test',
    '',
    '[Events]',
    'Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text',
    'Dialogue: 0,0:00:01.00,0:00:04.00,Default,,0,0,0,,こんにちは',
    'Dialogue: 0,0:00:05.00,0:00:08.50,Default,,0,0,0,,元気ですか',
  ].join('\n');

  const cues = parseAssCues(content);

  assert.equal(cues.length, 2);
  assert.equal(cues[0]!.startTime, 1.0);
  assert.equal(cues[0]!.endTime, 4.0);
  assert.equal(cues[0]!.text, 'こんにちは');
  assert.equal(cues[1]!.startTime, 5.0);
  assert.equal(cues[1]!.endTime, 8.5);
  assert.equal(cues[1]!.text, '元気ですか');
});

test('parseAssCues strips override tags from text', () => {
  const content = [
    '[Events]',
    'Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text',
    'Dialogue: 0,0:00:01.00,0:00:04.00,Default,,0,0,0,,{\\b1}太字{\\b0}テスト',
  ].join('\n');

  const cues = parseAssCues(content);

  assert.equal(cues[0]!.text, '太字テスト');
});

test('parseAssCues handles text containing commas', () => {
  const content = [
    '[Events]',
    'Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text',
    'Dialogue: 0,0:00:01.00,0:00:04.00,Default,,0,0,0,,はい、そうです、ね',
  ].join('\n');

  const cues = parseAssCues(content);

  assert.equal(cues[0]!.text, 'はい、そうです、ね');
});

test('parseAssCues handles \\N line breaks', () => {
  const content = [
    '[Events]',
    'Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text',
    'Dialogue: 0,0:00:01.00,0:00:04.00,Default,,0,0,0,,一行目\\N二行目',
  ].join('\n');

  const cues = parseAssCues(content);

  assert.equal(cues[0]!.text, '一行目\\N二行目');
});

test('parseAssCues returns empty for content without Events section', () => {
  const content = [
    '[Script Info]',
    'Title: Test',
  ].join('\n');

  assert.deepEqual(parseAssCues(content), []);
});

test('parseAssCues skips Comment lines', () => {
  const content = [
    '[Events]',
    'Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text',
    'Comment: 0,0:00:01.00,0:00:04.00,Default,,0,0,0,,これはコメント',
    'Dialogue: 0,0:00:05.00,0:00:08.00,Default,,0,0,0,,これは字幕',
  ].join('\n');

  const cues = parseAssCues(content);

  assert.equal(cues.length, 1);
  assert.equal(cues[0]!.text, 'これは字幕');
});

test('parseAssCues handles hour timestamps', () => {
  const content = [
    '[Events]',
    'Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text',
    'Dialogue: 0,1:30:00.00,1:30:05.00,Default,,0,0,0,,テスト',
  ].join('\n');

  const cues = parseAssCues(content);

  assert.equal(cues[0]!.startTime, 5400.0);
  assert.equal(cues[0]!.endTime, 5405.0);
});
```

- [ ] **Step 2: Run tests to verify ASS tests fail**

Run: `bun test src/core/services/subtitle-cue-parser.test.ts`
Expected: SRT tests PASS, ASS tests FAIL — `parseAssCues` not defined.

- [ ] **Step 3: Implement `parseAssCues`**

Add to `src/core/services/subtitle-cue-parser.ts`:

```typescript
const ASS_OVERRIDE_TAG_PATTERN = /\{[^}]*\}/g;

const ASS_TIMING_PATTERN = /^(\d+):(\d{2}):(\d{2})\.(\d{1,2})$/;

function parseAssTimestamp(raw: string): number | null {
  const match = ASS_TIMING_PATTERN.exec(raw.trim());
  if (!match) {
    return null;
  }
  const hours = Number(match[1]);
  const minutes = Number(match[2]);
  const seconds = Number(match[3]);
  const centiseconds = Number(match[4]!.padEnd(2, '0'));
  return hours * 3600 + minutes * 60 + seconds + centiseconds / 100;
}

export function parseAssCues(content: string): SubtitleCue[] {
  const cues: SubtitleCue[] = [];
  const lines = content.split(/\r?\n/);
  let inEventsSection = false;

  for (const line of lines) {
    const trimmed = line.trim();

    if (trimmed.startsWith('[') && trimmed.endsWith(']')) {
      inEventsSection = trimmed.toLowerCase() === '[events]';
      continue;
    }

    if (!inEventsSection) {
      continue;
    }

    if (!trimmed.startsWith('Dialogue:')) {
      continue;
    }

    // Split on first 9 commas (ASS v4+ has 10 fields; last is Text which can contain commas)
    const afterPrefix = trimmed.slice('Dialogue:'.length);
    const fields: string[] = [];
    let remaining = afterPrefix;
    for (let fieldIndex = 0; fieldIndex < 9; fieldIndex += 1) {
      const commaIndex = remaining.indexOf(',');
      if (commaIndex < 0) {
        break;
      }
      fields.push(remaining.slice(0, commaIndex));
      remaining = remaining.slice(commaIndex + 1);
    }

    if (fields.length < 9) {
      continue;
    }

    // fields[1] = Start, fields[2] = End (0-indexed: Layer, Start, End, ...)
    const startTime = parseAssTimestamp(fields[1]!);
    const endTime = parseAssTimestamp(fields[2]!);
    if (startTime === null || endTime === null) {
      continue;
    }

    // remaining = Text field (everything after the 9th comma)
    const rawText = remaining.replace(ASS_OVERRIDE_TAG_PATTERN, '').trim();
    if (rawText) {
      cues.push({ startTime, endTime, text: rawText });
    }
  }

  return cues;
}
```

- [ ] **Step 4: Run tests to verify all pass**

Run: `bun test src/core/services/subtitle-cue-parser.test.ts`
Expected: All SRT + ASS tests PASS.

- [ ] **Step 5: Commit**

```bash
git add src/core/services/subtitle-cue-parser.ts src/core/services/subtitle-cue-parser.test.ts
git commit -m "feat: add ASS subtitle cue parser"
```

---

### Task 6: Add unified `parseSubtitleCues` with format detection

**Files:**
- Modify: `src/core/services/subtitle-cue-parser.ts`
- Modify: `src/core/services/subtitle-cue-parser.test.ts`

- [ ] **Step 1: Write failing tests for `parseSubtitleCues`**

Add to `src/core/services/subtitle-cue-parser.test.ts`:

```typescript
test('parseSubtitleCues auto-detects SRT format', () => {
  const content = [
    '1',
    '00:00:01,000 --> 00:00:04,000',
    'SRTテスト',
    '',
  ].join('\n');

  const cues = parseSubtitleCues(content, 'test.srt');
  assert.equal(cues.length, 1);
  assert.equal(cues[0]!.text, 'SRTテスト');
});

test('parseSubtitleCues auto-detects ASS format', () => {
  const content = [
    '[Events]',
    'Format: Layer, Start, End, Style, Name, MarginL, MarginR, MarginV, Effect, Text',
    'Dialogue: 0,0:00:01.00,0:00:04.00,Default,,0,0,0,,ASSテスト',
  ].join('\n');

  const cues = parseSubtitleCues(content, 'test.ass');
  assert.equal(cues.length, 1);
  assert.equal(cues[0]!.text, 'ASSテスト');
});

test('parseSubtitleCues auto-detects VTT format', () => {
  const content = [
    '1',
    '00:00:01.000 --> 00:00:04.000',
    'VTTテスト',
    '',
  ].join('\n');

  const cues = parseSubtitleCues(content, 'test.vtt');
  assert.equal(cues.length, 1);
  assert.equal(cues[0]!.text, 'VTTテスト');
});

test('parseSubtitleCues returns empty for unknown format', () => {
  assert.deepEqual(parseSubtitleCues('random content', 'test.xyz'), []);
});

test('parseSubtitleCues returns cues sorted by start time', () => {
  const content = [
    '1',
    '00:00:10,000 --> 00:00:14,000',
    '二番目',
    '',
    '2',
    '00:00:01,000 --> 00:00:04,000',
    '一番目',
    '',
  ].join('\n');

  const cues = parseSubtitleCues(content, 'test.srt');
  assert.equal(cues[0]!.text, '一番目');
  assert.equal(cues[1]!.text, '二番目');
});
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `bun test src/core/services/subtitle-cue-parser.test.ts`
Expected: New tests FAIL — `parseSubtitleCues` not defined.

- [ ] **Step 3: Implement `parseSubtitleCues`**

Add to `src/core/services/subtitle-cue-parser.ts`:

```typescript
function detectSubtitleFormat(filename: string): 'srt' | 'vtt' | 'ass' | 'ssa' | null {
  const ext = filename.split('.').pop()?.toLowerCase() ?? '';
  if (ext === 'srt') return 'srt';
  if (ext === 'vtt') return 'vtt';
  if (ext === 'ass' || ext === 'ssa') return 'ass';
  return null;
}

export function parseSubtitleCues(content: string, filename: string): SubtitleCue[] {
  const format = detectSubtitleFormat(filename);
  let cues: SubtitleCue[];

  switch (format) {
    case 'srt':
    case 'vtt':
      cues = parseSrtCues(content);
      break;
    case 'ass':
    case 'ssa':
      cues = parseAssCues(content);
      break;
    default:
      return [];
  }

  cues.sort((a, b) => a.startTime - b.startTime);
  return cues;
}
```

- [ ] **Step 4: Run tests to verify all pass**

Run: `bun test src/core/services/subtitle-cue-parser.test.ts`
Expected: All tests PASS.

- [ ] **Step 5: Commit**

```bash
git add src/core/services/subtitle-cue-parser.ts src/core/services/subtitle-cue-parser.test.ts
git commit -m "feat: add unified parseSubtitleCues with format auto-detection"
```

---

## Chunk 3: Subtitle Prefetch Service

### Task 7: Create prefetch service core (priority window + background tokenization)

**Files:**
- Create: `src/core/services/subtitle-prefetch.ts`
- Create: `src/core/services/subtitle-prefetch.test.ts`

- [ ] **Step 1: Write failing tests for priority window computation and basic prefetching**

Create `src/core/services/subtitle-prefetch.test.ts`:

```typescript
import assert from 'node:assert/strict';
import test from 'node:test';
import {
  computePriorityWindow,
  createSubtitlePrefetchService,
} from './subtitle-prefetch';
import type { SubtitleCue } from './subtitle-cue-parser';
import type { SubtitleData } from '../../types';

function makeCues(count: number, startOffset = 0): SubtitleCue[] {
  return Array.from({ length: count }, (_, i) => ({
    startTime: startOffset + i * 5,
    endTime: startOffset + i * 5 + 4,
    text: `line-${i}`,
  }));
}

test('computePriorityWindow returns next N cues from current position', () => {
  const cues = makeCues(20);
  const window = computePriorityWindow(cues, 12.0, 5);

  assert.equal(window.length, 5);
  // Position 12.0 is during cue index 2 (start=10, end=14). Priority window starts from index 3.
  assert.equal(window[0]!.text, 'line-3');
  assert.equal(window[4]!.text, 'line-7');
});

test('computePriorityWindow clamps to remaining cues at end of file', () => {
  const cues = makeCues(5);
  const window = computePriorityWindow(cues, 18.0, 10);

  // Position 18.0 is during cue 3 (start=15). Only cue 4 is ahead.
  assert.equal(window.length, 1);
  assert.equal(window[0]!.text, 'line-4');
});

test('computePriorityWindow returns empty when past all cues', () => {
  const cues = makeCues(3);
  const window = computePriorityWindow(cues, 999.0, 10);
  assert.equal(window.length, 0);
});

test('computePriorityWindow at position 0 returns first N cues', () => {
  const cues = makeCues(20);
  const window = computePriorityWindow(cues, 0, 5);

  assert.equal(window.length, 5);
  assert.equal(window[0]!.text, 'line-0');
});

function flushMicrotasks(): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, 0));
}

test('prefetch service tokenizes priority window cues and caches them', async () => {
  const cues = makeCues(20);
  const cached: Map<string, SubtitleData> = new Map();
  let tokenizeCalls = 0;

  const service = createSubtitlePrefetchService({
    cues,
    tokenizeSubtitle: async (text) => {
      tokenizeCalls += 1;
      return { text, tokens: [] };
    },
    preCacheTokenization: (text, data) => {
      cached.set(text, data);
    },
    isCacheFull: () => false,
    priorityWindowSize: 3,
  });

  service.start(0);
  // Allow all async tokenization to complete
  for (let i = 0; i < 25; i += 1) {
    await flushMicrotasks();
  }
  service.stop();

  // Priority window (first 3) should be cached
  assert.ok(cached.has('line-0'));
  assert.ok(cached.has('line-1'));
  assert.ok(cached.has('line-2'));
});

test('prefetch service stops when cache is full', async () => {
  const cues = makeCues(20);
  let tokenizeCalls = 0;
  let cacheSize = 0;

  const service = createSubtitlePrefetchService({
    cues,
    tokenizeSubtitle: async (text) => {
      tokenizeCalls += 1;
      return { text, tokens: [] };
    },
    preCacheTokenization: () => {
      cacheSize += 1;
    },
    isCacheFull: () => cacheSize >= 5,
    priorityWindowSize: 3,
  });

  service.start(0);
  for (let i = 0; i < 30; i += 1) {
    await flushMicrotasks();
  }
  service.stop();

  // Should have stopped at 5 (cache full), not tokenized all 20
  assert.ok(tokenizeCalls <= 6, `Expected <= 6 tokenize calls, got ${tokenizeCalls}`);
});

test('prefetch service can be stopped mid-flight', async () => {
  const cues = makeCues(100);
  let tokenizeCalls = 0;

  const service = createSubtitlePrefetchService({
    cues,
    tokenizeSubtitle: async (text) => {
      tokenizeCalls += 1;
      return { text, tokens: [] };
    },
    preCacheTokenization: () => {},
    isCacheFull: () => false,
    priorityWindowSize: 3,
  });

  service.start(0);
  await flushMicrotasks();
  await flushMicrotasks();
  service.stop();
  const callsAtStop = tokenizeCalls;

  // Wait more to confirm no further calls
  for (let i = 0; i < 10; i += 1) {
    await flushMicrotasks();
  }

  assert.equal(tokenizeCalls, callsAtStop, 'No further tokenize calls after stop');
  assert.ok(tokenizeCalls < 100, 'Should not have tokenized all cues');
});
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `bun test src/core/services/subtitle-prefetch.test.ts`
Expected: FAIL — module not found.

- [ ] **Step 3: Implement the prefetch service**

Create `src/core/services/subtitle-prefetch.ts`:

```typescript
import type { SubtitleCue } from './subtitle-cue-parser';
import type { SubtitleData } from '../../types';

export interface SubtitlePrefetchServiceDeps {
  cues: SubtitleCue[];
  tokenizeSubtitle: (text: string) => Promise<SubtitleData | null>;
  preCacheTokenization: (text: string, data: SubtitleData) => void;
  isCacheFull: () => boolean;
  priorityWindowSize?: number;
}

export interface SubtitlePrefetchService {
  start: (currentTimeSeconds: number) => void;
  stop: () => void;
  onSeek: (newTimeSeconds: number) => void;
  pause: () => void;
  resume: () => void;
}

const DEFAULT_PRIORITY_WINDOW_SIZE = 10;

export function computePriorityWindow(
  cues: SubtitleCue[],
  currentTimeSeconds: number,
  windowSize: number,
): SubtitleCue[] {
  if (cues.length === 0) {
    return [];
  }

  // Find the first cue whose start time is >= current position.
  // This includes cues that start exactly at the current time (they haven't
  // been displayed yet and should be prefetched).
  let startIndex = -1;
  for (let i = 0; i < cues.length; i += 1) {
    if (cues[i]!.startTime >= currentTimeSeconds) {
      startIndex = i;
      break;
    }
  }

  if (startIndex < 0) {
    // All cues are before current time
    return [];
  }

  return cues.slice(startIndex, startIndex + windowSize);
}

export function createSubtitlePrefetchService(
  deps: SubtitlePrefetchServiceDeps,
): SubtitlePrefetchService {
  const windowSize = deps.priorityWindowSize ?? DEFAULT_PRIORITY_WINDOW_SIZE;
  let stopped = true;
  let paused = false;
  let currentRunId = 0;

  async function tokenizeCueList(
    cuesToProcess: SubtitleCue[],
    runId: number,
  ): Promise<void> {
    for (const cue of cuesToProcess) {
      if (stopped || runId !== currentRunId) {
        return;
      }

      // Wait while paused
      while (paused && !stopped && runId === currentRunId) {
        await new Promise((resolve) => setTimeout(resolve, 10));
      }

      if (stopped || runId !== currentRunId) {
        return;
      }

      if (deps.isCacheFull()) {
        return;
      }

      try {
        const result = await deps.tokenizeSubtitle(cue.text);
        if (result && !stopped && runId === currentRunId) {
          deps.preCacheTokenization(cue.text, result);
        }
      } catch {
        // Skip failed cues, continue prefetching
      }

      // Yield to allow live processing to take priority
      await new Promise((resolve) => setTimeout(resolve, 0));
    }
  }

  async function startPrefetching(currentTimeSeconds: number, runId: number): Promise<void> {
    const cues = deps.cues;

    // Phase 1: Priority window
    const priorityCues = computePriorityWindow(cues, currentTimeSeconds, windowSize);
    await tokenizeCueList(priorityCues, runId);

    if (stopped || runId !== currentRunId) {
      return;
    }

    // Phase 2: Background - remaining cues forward from current position
    const priorityTexts = new Set(priorityCues.map((c) => c.text));
    const remainingCues = cues.filter(
      (cue) => cue.startTime > currentTimeSeconds && !priorityTexts.has(cue.text),
    );
    await tokenizeCueList(remainingCues, runId);

    if (stopped || runId !== currentRunId) {
      return;
    }

    // Phase 3: Background - earlier cues (for rewind support)
    const earlierCues = cues.filter(
      (cue) => cue.startTime <= currentTimeSeconds && !priorityTexts.has(cue.text),
    );
    await tokenizeCueList(earlierCues, runId);
  }

  return {
    start(currentTimeSeconds: number) {
      stopped = false;
      paused = false;
      currentRunId += 1;
      const runId = currentRunId;
      void startPrefetching(currentTimeSeconds, runId);
    },

    stop() {
      stopped = true;
      currentRunId += 1;
    },

    onSeek(newTimeSeconds: number) {
      // Cancel current run and restart from new position
      currentRunId += 1;
      const runId = currentRunId;
      void startPrefetching(newTimeSeconds, runId);
    },

    pause() {
      paused = true;
    },

    resume() {
      paused = false;
    },
  };
}
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `bun test src/core/services/subtitle-prefetch.test.ts`
Expected: All tests PASS.

- [ ] **Step 5: Commit**

```bash
git add src/core/services/subtitle-prefetch.ts src/core/services/subtitle-prefetch.test.ts
git commit -m "feat: add subtitle prefetch service with priority window

Implements background tokenization of upcoming subtitle cues with a
configurable priority window. Supports stop, pause/resume, seek
re-prioritization, and cache-full stopping condition."
```

---

### Task 8: Add seek detection and pause/resume tests

**Files:**
- Modify: `src/core/services/subtitle-prefetch.test.ts`

- [ ] **Step 1: Write tests for seek re-prioritization and pause/resume**

Add to `src/core/services/subtitle-prefetch.test.ts`:

```typescript
test('prefetch service onSeek re-prioritizes from new position', async () => {
  const cues = makeCues(20);
  const cachedTexts: string[] = [];

  const service = createSubtitlePrefetchService({
    cues,
    tokenizeSubtitle: async (text) => ({ text, tokens: [] }),
    preCacheTokenization: (text) => {
      cachedTexts.push(text);
    },
    isCacheFull: () => false,
    priorityWindowSize: 3,
  });

  service.start(0);
  // Let a few cues process
  for (let i = 0; i < 5; i += 1) {
    await flushMicrotasks();
  }

  // Seek to near the end
  service.onSeek(80.0);
  for (let i = 0; i < 30; i += 1) {
    await flushMicrotasks();
  }
  service.stop();

  // After seek to 80.0, cues starting after 80.0 (line-17, line-18, line-19) should appear in cached
  const hasPostSeekCue = cachedTexts.some((t) => t === 'line-17' || t === 'line-18' || t === 'line-19');
  assert.ok(hasPostSeekCue, 'Should have cached cues after seek position');
});

test('prefetch service pause/resume halts and continues tokenization', async () => {
  const cues = makeCues(20);
  let tokenizeCalls = 0;

  const service = createSubtitlePrefetchService({
    cues,
    tokenizeSubtitle: async (text) => {
      tokenizeCalls += 1;
      return { text, tokens: [] };
    },
    preCacheTokenization: () => {},
    isCacheFull: () => false,
    priorityWindowSize: 3,
  });

  service.start(0);
  await flushMicrotasks();
  await flushMicrotasks();
  service.pause();

  const callsWhenPaused = tokenizeCalls;
  // Wait while paused
  for (let i = 0; i < 5; i += 1) {
    await flushMicrotasks();
  }
  // Should not have advanced much (may have 1 in-flight)
  assert.ok(tokenizeCalls <= callsWhenPaused + 1, 'Should not tokenize much while paused');

  service.resume();
  for (let i = 0; i < 30; i += 1) {
    await flushMicrotasks();
  }
  service.stop();

  assert.ok(tokenizeCalls > callsWhenPaused + 1, 'Should resume tokenizing after unpause');
});
```

- [ ] **Step 2: Run tests to verify they pass**

Run: `bun test src/core/services/subtitle-prefetch.test.ts`
Expected: All tests PASS (the implementation from Task 7 already handles these cases).

- [ ] **Step 3: Commit**

```bash
git add src/core/services/subtitle-prefetch.test.ts
git commit -m "test: add seek and pause/resume tests for prefetch service"
```

---

### Task 9: Wire up prefetch service in main.ts

This is the integration task that connects the prefetch service to the actual MPV events and subtitle processing controller.

**Files:**
- Modify: `src/main.ts`
- Modify: `src/main/runtime/mpv-main-event-actions.ts` (extend time-pos handler)
- Modify: `src/main/runtime/mpv-main-event-bindings.ts` (pass new dep through)

**Architecture context:** MPV events flow through a layered system:
1. `MpvIpcClient` (`src/core/services/mpv.ts`) emits events like `'time-pos-change'`
2. `mpv-client-event-bindings.ts` binds these to handler functions (e.g., `deps.onTimePosChange`)
3. `mpv-main-event-bindings.ts` wires up the bindings with concrete handler implementations created by `mpv-main-event-actions.ts`
4. The handlers are constructed via factory functions like `createHandleMpvTimePosChangeHandler`

Key existing locations:
- `createHandleMpvTimePosChangeHandler` is in `src/main/runtime/mpv-main-event-actions.ts:89-99`
- The `onSubtitleChange` callback is at `src/main.ts:2841-2843`
- The `emitSubtitle` callback is at `src/main.ts:1051-1064`
- `invalidateTokenizationCache` is called at `src/main.ts:1433` (onSyncComplete) and `src/main.ts:2600` (onOptionsChanged)
- `loadSubtitleSourceText` is an inline closure at `src/main.ts:3496-3513` — it needs to be extracted or duplicated

- [ ] **Step 1: Add imports**

At the top of `src/main.ts`, add imports for the new modules:

```typescript
import { parseSubtitleCues } from './core/services/subtitle-cue-parser';
import { createSubtitlePrefetchService } from './core/services/subtitle-prefetch';
import type { SubtitlePrefetchService } from './core/services/subtitle-prefetch';
```

- [ ] **Step 2: Extract `loadSubtitleSourceText` into a reusable function**

The subtitle file loading logic at `src/main.ts:3496-3513` is currently an inline closure passed to `createShiftSubtitleDelayToAdjacentCueHandler`. Extract it into a standalone function above that usage site so it can be shared with the prefetcher:

```typescript
async function loadSubtitleSourceText(source: string): Promise<string> {
  if (/^https?:\/\//i.test(source)) {
    const controller = new AbortController();
    const timeoutId = setTimeout(() => controller.abort(), 4000);
    try {
      const response = await fetch(source, { signal: controller.signal });
      if (!response.ok) {
        throw new Error(`Failed to download subtitle source (${response.status})`);
      }
      return await response.text();
    } finally {
      clearTimeout(timeoutId);
    }
  }

  const filePath = source.startsWith('file://') ? decodeURI(new URL(source).pathname) : source;
  return fs.promises.readFile(filePath, 'utf8');
}
```

Then update the `createShiftSubtitleDelayToAdjacentCueHandler` call to use `loadSubtitleSourceText: loadSubtitleSourceText,` instead of the inline closure.

- [ ] **Step 3: Add prefetch service state and init helper**

Near the other service state declarations (near `tokenizeSubtitleDeferred` at line 1046), add:

```typescript
let subtitlePrefetchService: SubtitlePrefetchService | null = null;
let lastObservedTimePos = 0;
const SEEK_THRESHOLD_SECONDS = 3;

async function initSubtitlePrefetch(
  externalFilename: string,
  currentTimePos: number,
): Promise<void> {
  subtitlePrefetchService?.stop();
  subtitlePrefetchService = null;

  try {
    const content = await loadSubtitleSourceText(externalFilename);
    const cues = parseSubtitleCues(content, externalFilename);
    if (cues.length === 0) {
      return;
    }

    subtitlePrefetchService = createSubtitlePrefetchService({
      cues,
      tokenizeSubtitle: async (text) =>
        tokenizeSubtitleDeferred ? await tokenizeSubtitleDeferred(text) : null,
      preCacheTokenization: (text, data) => {
        subtitleProcessingController.preCacheTokenization(text, data);
      },
      isCacheFull: () => subtitleProcessingController.isCacheFull(),
    });

    subtitlePrefetchService.start(currentTimePos);
    logger.info(`[subtitle-prefetch] started prefetching ${cues.length} cues from ${externalFilename}`);
  } catch (error) {
    logger.warn('[subtitle-prefetch] failed to initialize:', (error as Error).message);
  }
}
```

- [ ] **Step 4: Hook seek detection into the time-pos handler**

The existing `createHandleMpvTimePosChangeHandler` in `src/main/runtime/mpv-main-event-actions.ts:89-99` fires on every `time-pos` update. Extend its deps interface to accept an optional seek callback:

In `src/main/runtime/mpv-main-event-actions.ts`, add to the deps type:

```typescript
export function createHandleMpvTimePosChangeHandler(deps: {
  recordPlaybackPosition: (time: number) => void;
  reportJellyfinRemoteProgress: (forceImmediate: boolean) => void;
  refreshDiscordPresence: () => void;
  onTimePosUpdate?: (time: number) => void;  // NEW
}) {
  return ({ time }: { time: number }): void => {
    deps.recordPlaybackPosition(time);
    deps.reportJellyfinRemoteProgress(false);
    deps.refreshDiscordPresence();
    deps.onTimePosUpdate?.(time);  // NEW
  };
}
```

Then in `src/main/runtime/mpv-main-event-bindings.ts` (around line 122), pass the new dep when constructing the handler:

```typescript
const handleMpvTimePosChange = createHandleMpvTimePosChangeHandler({
  recordPlaybackPosition: (time) => deps.recordPlaybackPosition(time),
  reportJellyfinRemoteProgress: (forceImmediate) =>
    deps.reportJellyfinRemoteProgress(forceImmediate),
  refreshDiscordPresence: () => deps.refreshDiscordPresence(),
  onTimePosUpdate: (time) => deps.onTimePosUpdate?.(time),  // NEW
});
```

And add `onTimePosUpdate` to the deps interface of the main event bindings function, passing it through from `main.ts`.

Finally, in `src/main.ts`, where the MPV main event bindings are wired up, provide the `onTimePosUpdate` callback:

```typescript
onTimePosUpdate: (time) => {
  const delta = time - lastObservedTimePos;
  if (subtitlePrefetchService && (delta > SEEK_THRESHOLD_SECONDS || delta < 0)) {
    subtitlePrefetchService.onSeek(time);
  }
  lastObservedTimePos = time;
},
```

- [ ] **Step 5: Hook prefetch pause/resume into live subtitle processing**

At `src/main.ts:2841-2843`, the `onSubtitleChange` callback:

```typescript
onSubtitleChange: (text) => {
  subtitlePrefetchService?.pause();  // NEW: pause prefetch during live processing
  subtitleProcessingController.onSubtitleChange(text);
},
```

At `src/main.ts:1051-1064`, inside the `emitSubtitle` callback, add resume at the end:

```typescript
emitSubtitle: (payload) => {
  appState.currentSubtitleData = payload;
  broadcastToOverlayWindows('subtitle:set', payload);
  subtitleWsService.broadcast(payload, { /* ... existing ... */ });
  annotationSubtitleWsService.broadcast(payload, { /* ... existing ... */ });
  subtitlePrefetchService?.resume();  // NEW: resume prefetch after emission
},
```

- [ ] **Step 6: Hook into cache invalidation**

At `src/main.ts:1433` (onSyncComplete) after `invalidateTokenizationCache()`:

```typescript
subtitleProcessingController.invalidateTokenizationCache();
subtitlePrefetchService?.onSeek(lastObservedTimePos);  // NEW: re-prefetch after invalidation
```

At `src/main.ts:2600` (onOptionsChanged) after `invalidateTokenizationCache()`:

```typescript
subtitleProcessingController.invalidateTokenizationCache();
subtitlePrefetchService?.onSeek(lastObservedTimePos);  // NEW: re-prefetch after invalidation
```

- [ ] **Step 7: Trigger prefetch on subtitle track activation**

Find where the subtitle track is activated in `main.ts` (where media path changes or subtitle track changes are handled). When a new external subtitle track is detected, call `initSubtitlePrefetch`. The exact location depends on how track changes are wired — search for where `track-list` is processed or where the subtitle track ID changes. Use MPV's `requestProperty('track-list')` to get the external filename, then call:

```typescript
// When external subtitle track is detected:
const trackList = await appState.mpvClient?.requestProperty('track-list');
// Find the active sub track, get its external-filename
// Then:
void initSubtitlePrefetch(externalFilename, lastObservedTimePos);
```

**Note:** The exact wiring location for track activation needs to be determined during implementation. Search for `media-path-change` handler or `updateCurrentMediaPath` (line 2854) as the likely trigger point — when a new media file is loaded, the subtitle track becomes available.

- [ ] **Step 8: Run full test suite**

Run: `bun run test`
Expected: Same results as baseline (500 pass, 1 pre-existing fail).

- [ ] **Step 9: Commit**

```bash
git add src/main.ts src/main/runtime/mpv-main-event-actions.ts src/main/runtime/mpv-main-event-bindings.ts
git commit -m "feat: wire up subtitle prefetch service to MPV events

Initializes prefetch on external subtitle track activation, detects
seeks via time-pos delta threshold, pauses prefetch during live
subtitle processing, and restarts on cache invalidation."
```

---

## Final Verification

### Task 10: Full test suite and type check

- [ ] **Step 1: Run full test suite**

Run: `bun run test`
Expected: 500+ pass, 1 pre-existing fail in immersion-tracker-service.

- [ ] **Step 2: Run TypeScript type check**

Run: `bun run tsc`
Expected: No type errors.

- [ ] **Step 3: Review all commits on the branch**

Run: `git log --oneline main..HEAD`
Expected: ~8 focused commits, one per logical change.
