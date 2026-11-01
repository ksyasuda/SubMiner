// Interim source-text guards: each test regexes src/main.ts (or a runtime file) to
// pin a startup/overlay invariant that main.ts closures make hard to exercise.
// Replace each with a behavioral test as the logic moves out of main.ts closures
// (see docs/workflow/testing.md, Known Debt). Do not add new tests here.
import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';

function readMainSource(): string {
  return fs.readFileSync(path.join(process.cwd(), 'src/main.ts'), 'utf8');
}

function readSource(relPath: string): string {
  return fs.readFileSync(path.join(process.cwd(), relPath), 'utf8');
}

test('subtitle sidebar media path tag is assigned after prefetch succeeds', () => {
  const source = readSource('src/main/runtime/autoplay-subtitle-priming-runtime.ts');
  const actionBlock = source.match(
    /async function refreshSubtitleSidebarFromSource\([\s\S]*?\): Promise<void> \{(?<body>[\s\S]*?)\n  \}/,
  )?.groups?.body;

  assert.ok(actionBlock);
  assert.match(
    actionBlock,
    /const nextMediaPath = mediaPath\?\.trim\(\) \|\| getCurrentAutoplayMediaPath\(\);/,
  );
  assert.ok(
    actionBlock.indexOf('deps.initSubtitlePrefetch(') <
      actionBlock.indexOf('deps.setActiveParsedSubtitleMediaPath(nextMediaPath);'),
  );
});

test('autoplay subtitle prime emits cached annotations and avoids raw fallback overlay flashes', () => {
  const source = readSource('src/main/runtime/autoplay-subtitle-priming-runtime.ts');
  const actionBlock = source.match(
    /function emitAutoplayPrimedSubtitle\([\s\S]*?\): boolean \{(?<body>[\s\S]*?)\n  \}/,
  )?.groups?.body;

  assert.ok(actionBlock);
  assert.match(
    actionBlock,
    /const cachedPayload = subtitleProcessingController\.consumeCachedSubtitle\(text\);/,
  );
  assert.match(actionBlock, /if \(cachedPayload\) \{/);
  assert.match(actionBlock, /emitSubtitlePayload\(cachedPayload\);/);
  assert.match(actionBlock, /subtitleProcessingController\.onSubtitleChange\(text\);/);
  assert.ok(
    actionBlock.indexOf('consumeCachedSubtitle(text)') <
      actionBlock.indexOf('subtitleProcessingController.onSubtitleChange(text);'),
  );
});

test('startup autoplay release is tied to visible overlay measurement readiness', () => {
  const source = readMainSource();
  const gateBlock = source.match(
    /const autoplayReadyGate = createAutoplayReadyGate\(\{(?<body>[\s\S]*?)\n\}\);/,
  )?.groups?.body;
  assert.ok(gateBlock);
  assert.match(gateBlock, /isSignalTargetReady:\s*\(signal\) =>/);
  // Untokenized signals are filtered by the gate's isTokenizationReady dep, not
  // by the target-readiness predicate (which must stay warmup-free so warm and
  // tokenized releases are never deferred behind the global warmup flag).
  assert.match(gateBlock, /isTokenizationReady:\s*\(\) => isTokenizationWarmupReady\(\)/);
  const signalTargetReadyBlock = gateBlock.match(
    /isSignalTargetReady:\s*\(signal\) =>(?<body>[\s\S]*?)\n  schedule:/,
  )?.groups?.body;
  assert.ok(signalTargetReadyBlock);
  assert.doesNotMatch(signalTargetReadyBlock, /isTokenizationWarmupReady\(\)/);
  assert.match(gateBlock, /isVisibleOverlayAutoplayTargetReady\(/);
  assert.match(gateBlock, /getLatestVisibleMeasurement:/);
});

test('stats server Yomitan note creation honors configured Anki server override policy', () => {
  const source = readSource('src/main/runtime/stats-server-runtime.ts');
  const statsServerConfigBlock = source.match(
    /const buildStatsServerConfig[\s\S]*?return \{(?<body>[\s\S]*?)\n    \};\n  \};/,
  )?.groups?.body;
  const addYomitanNoteBlock = statsServerConfigBlock?.match(
    /addYomitanNote:\s*async\s*\(word: string\)\s*=>\s*\{(?<body>[\s\S]*?)\n      \},/,
  )?.groups?.body;

  assert.ok(addYomitanNoteBlock);
  assert.match(
    addYomitanNoteBlock,
    /const ankiConnectConfig = deps\.getResolvedConfig\(\)\.ankiConnect;/,
  );
  assert.match(addYomitanNoteBlock, /shouldForceOverrideYomitanAnkiServer\(ankiConnectConfig\)/);
  assert.doesNotMatch(addYomitanNoteBlock, /forceOverride:\s*true/);
});

test('visible overlay input reset clears the visible content measurement', () => {
  const runtimeSource = readSource('src/main/runtime/visible-overlay-interaction-runtime.ts');
  const resetBlock = runtimeSource.match(
    /function resetVisibleOverlayInputState\(\): void \{(?<body>[\s\S]*?)\n  \}/,
  )?.groups?.body;

  assert.ok(resetBlock);
  assert.match(resetBlock, /overlayContentMeasurementStore\.clear\('visible'\);/);
});

test('Linux visible overlay startup reapplies passive passthrough after input reset', () => {
  const pointerSource = readSource('src/main/runtime/linux-overlay-pointer-interaction.ts');
  const runtimeSource = readSource('src/main/runtime/visible-overlay-interaction-runtime.ts');
  const resetPrimerBlock = runtimeSource.match(
    /function resetLinuxVisibleOverlayStartupInputPrimer\(\): void \{(?<body>[\s\S]*?)\n  \}/,
  )?.groups?.body;
  const startGraceBlock = runtimeSource.match(
    /function startLinuxVisibleOverlayStartupInputGrace\(\): void \{(?<body>[\s\S]*?)\n  \}/,
  )?.groups?.body;
  const depsBlock = runtimeSource.match(
    /const linuxOverlayPointerInteractionDeps = \{(?<body>[\s\S]*?)\n  \};/,
  )?.groups?.body;

  assert.ok(resetPrimerBlock);
  assert.ok(startGraceBlock);
  assert.ok(depsBlock);
  assert.match(resetPrimerBlock, /visibleOverlayInteractionActive = false;/);
  assert.match(resetPrimerBlock, /linuxOverlayPointerInteractionStateApplied = false;/);
  assert.match(
    startGraceBlock,
    /linuxVisibleOverlayStartupInputGraceUntilMs =\s+Date\.now\(\) \+ LINUX_VISIBLE_OVERLAY_STARTUP_INPUT_GRACE_MS;/,
  );
  assert.match(startGraceBlock, /linuxOverlayPointerInteractionStateApplied = false;/);
  assert.match(
    depsBlock,
    /isInteractionStateApplied:\s*\(\) => linuxOverlayPointerInteractionStateApplied/,
  );
  assert.match(
    pointerSource,
    /deps\.getInteractionActive\(\) === desired && deps\.isInteractionStateApplied\?\.\(\) !== false/,
  );
});

test('Linux visible overlay bounds refresh restores X11 shape after applying mpv geometry', () => {
  const source = readSource('src/main/runtime/overlay-geometry-runtime.ts');
  const afterBoundsBlock = source.match(
    /afterSetOverlayWindowBounds:\s*\(\) => \{(?<body>[\s\S]*?)\n      \},/,
  )?.groups?.body;

  assert.ok(afterBoundsBlock);
  assert.match(afterBoundsBlock, /restoreLinuxOverlayWindowShape\(mainWindow\);/);
  assert.ok(
    afterBoundsBlock.indexOf('restoreLinuxOverlayWindowShape(mainWindow);') <
      afterBoundsBlock.indexOf('ensureOverlayWindowLevel(mainWindow);'),
  );
});
