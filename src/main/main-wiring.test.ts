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

test('manual watched session action starts immersion tracker before marking watched', () => {
  const source = readMainSource();
  const actionBlock = source.match(
    /markActiveVideoWatched:\s*async\s*\(\)\s*=>\s*\{(?<body>[\s\S]*?)\}\s*,/,
  )?.groups?.body;

  assert.ok(actionBlock);
  assert.match(actionBlock, /ensureImmersionTrackerStarted\(\);/);
  assert.ok(
    actionBlock.indexOf('ensureImmersionTrackerStarted();') <
      actionBlock.indexOf('markActiveVideoWatched()'),
  );
});

test('visible overlay show paths start the overlay loading OSD', () => {
  const source = readMainSource();
  const setVisibleBlock = source.match(
    /function setVisibleOverlayVisible\(visible: boolean\): void \{(?<body>[\s\S]*?)\n\}/,
  )?.groups?.body;

  assert.ok(setVisibleBlock);
  assert.match(setVisibleBlock, /if \(visible\) \{\s+maybeStartOverlayLoadingOsd\(\);/);
  assert.match(
    source,
    /function toggleVisibleOverlay\(\): void \{[\s\S]*?else \{\s+maybeStartOverlayLoadingOsd\(\);/,
  );
  assert.match(
    source,
    /function setOverlayVisible\(visible: boolean\): void \{[\s\S]*?if \(visible\) \{\s+maybeStartOverlayLoadingOsd\(\);/,
  );
});

test('manual visible overlay toggles only release current-media autoplay when hiding', () => {
  const source = readMainSource();
  const actionBlock = source.match(
    /function toggleVisibleOverlay\(\): void \{(?<body>[\s\S]*?)\n\}/,
  )?.groups?.body;

  assert.ok(actionBlock);
  assert.match(
    actionBlock,
    /if \(!nextVisible\) \{[\s\S]*?autoplayReadyGate\.markCurrentMediaAutoplayReady\(\);[\s\S]*?cancelVisibleOverlaySubtitleRefreshAfterFirstPaint\(\);[\s\S]*?cancelPendingLinuxMpvFullscreenOverlayRefreshBurst\(\);/,
  );
});

test('all visible overlay hide paths clear stale overlay input state', () => {
  const source = readMainSource();
  const setVisibleBlock = source.match(
    /function setVisibleOverlayVisible\(visible: boolean\): void \{(?<body>[\s\S]*?)\n\}/,
  )?.groups?.body;
  const toggleBlock = source.match(
    /function toggleVisibleOverlay\(\): void \{(?<body>[\s\S]*?)\n\}/,
  )?.groups?.body;
  const setOverlayBlock = source.match(
    /function setOverlayVisible\(visible: boolean\): void \{(?<body>[\s\S]*?)\n\}/,
  )?.groups?.body;

  assert.ok(setVisibleBlock);
  assert.ok(toggleBlock);
  assert.ok(setOverlayBlock);
  assert.match(
    setVisibleBlock,
    /if \(!visible\) \{[\s\S]*?autoplayReadyGate\.markCurrentMediaAutoplayReady\(\);[\s\S]*?cancelVisibleOverlaySubtitleRefreshAfterFirstPaint\(\);[\s\S]*?cancelPendingLinuxMpvFullscreenOverlayRefreshBurst\(\);[\s\S]*?resetVisibleOverlayInputState\(\);/,
  );
  assert.match(
    toggleBlock,
    /if \(!nextVisible\) \{[\s\S]*?autoplayReadyGate\.markCurrentMediaAutoplayReady\(\);[\s\S]*?cancelVisibleOverlaySubtitleRefreshAfterFirstPaint\(\);[\s\S]*?cancelPendingLinuxMpvFullscreenOverlayRefreshBurst\(\);[\s\S]*?resetVisibleOverlayInputState\(\);/,
  );
  assert.match(
    setOverlayBlock,
    /if \(!visible\) \{[\s\S]*?cancelVisibleOverlaySubtitleRefreshAfterFirstPaint\(\);[\s\S]*?resetVisibleOverlayInputState\(\);[\s\S]*?autoplayReadyGate\.markCurrentMediaAutoplayReady\(\);[\s\S]*?cancelPendingLinuxMpvFullscreenOverlayRefreshBurst\(\);/,
  );
  // Hiding must also dismiss the loading OSD, or it lingers over mpv.
  assert.match(
    setVisibleBlock,
    /if \(!visible\) \{[\s\S]*?dismissOverlayLoadingStatusNotification\(\);/,
  );
  assert.match(
    toggleBlock,
    /if \(!nextVisible\) \{[\s\S]*?dismissOverlayLoadingStatusNotification\(\);/,
  );
  assert.match(
    setOverlayBlock,
    /if \(!visible\) \{[\s\S]*?dismissOverlayLoadingStatusNotification\(\);/,
  );
});

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
  const measurementBlock = source.match(
    /reportOverlayContentBounds:\s*\(payload: unknown\)\s*=>\s*\{(?<body>[\s\S]*?)\n      \},/,
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

  assert.ok(measurementBlock);
  assert.match(measurementBlock, /overlayContentMeasurementStore\.report\(payload\)/);
  assert.match(measurementBlock, /autoplayReadyGate\.flushPendingAutoplayReadySignal\(\)/);
});

test('visible overlay content-ready does not tokenize before first measurement', () => {
  const source = readMainSource();
  const contentReadyBlock = source.match(
    /onWindowContentReady:\s*\(\)\s*=>\s*\{(?<body>[\s\S]*?)\n      \},/,
  )?.groups?.body;
  const measurementBlock = source.match(
    /reportOverlayContentBounds:\s*\(payload: unknown\)\s*=>\s*\{(?<body>[\s\S]*?)\n      \},/,
  )?.groups?.body;

  assert.ok(contentReadyBlock);
  assert.doesNotMatch(contentReadyBlock, /subtitleProcessingController\.refreshCurrentSubtitle/);
  assert.match(contentReadyBlock, /autoplayReadyGate\.flushPendingAutoplayReadySignal\(\)/);
  assert.match(contentReadyBlock, /primeLinuxOverlayPointerInteractionAfterFirstMeasurement\(\)/);
  assert.ok(
    contentReadyBlock.indexOf('overlayVisibilityRuntime.updateVisibleOverlayVisibility();') <
      contentReadyBlock.indexOf('primeLinuxOverlayPointerInteractionAfterFirstMeasurement();'),
  );
  assert.ok(
    contentReadyBlock.indexOf('primeLinuxOverlayPointerInteractionAfterFirstMeasurement();') <
      contentReadyBlock.indexOf('autoplayReadyGate.flushPendingAutoplayReadySignal();'),
  );

  assert.ok(measurementBlock);
  assert.match(measurementBlock, /autoplayReadyGate\.flushPendingAutoplayReadySignal\(\)/);
  assert.match(measurementBlock, /scheduleVisibleOverlaySubtitleRefreshAfterFirstPaint\(\)/);
  assert.ok(
    measurementBlock.indexOf('autoplayReadyGate.flushPendingAutoplayReadySignal();') <
      measurementBlock.indexOf('scheduleVisibleOverlaySubtitleRefreshAfterFirstPaint();'),
  );
});

test('accepted visible overlay measurement immediately refreshes pointer interaction', () => {
  const source = readMainSource();
  const measurementBlock = source.match(
    /reportOverlayContentBounds:\s*\(payload: unknown\)\s*=>\s*\{(?<body>[\s\S]*?)\n      \},/,
  )?.groups?.body;

  assert.ok(measurementBlock);
  assert.match(measurementBlock, /overlayContentMeasurementStore\.report\(payload\)/);
  assert.match(measurementBlock, /tickLinuxOverlayPointerInteractionNow\(\)/);
  assert.match(measurementBlock, /tickWindowsOverlayPointerInteractionNow\(\)/);
  assert.match(measurementBlock, /primeLinuxOverlayPointerInteractionAfterFirstMeasurement\(\)/);
  assert.ok(
    measurementBlock.indexOf('overlayContentMeasurementStore.report(payload)') <
      measurementBlock.indexOf('tickLinuxOverlayPointerInteractionNow();'),
  );
  assert.ok(
    measurementBlock.indexOf('tickLinuxOverlayPointerInteractionNow();') <
      measurementBlock.indexOf('tickWindowsOverlayPointerInteractionNow();'),
  );
  assert.ok(
    measurementBlock.indexOf('tickWindowsOverlayPointerInteractionNow();') <
      measurementBlock.indexOf('primeLinuxOverlayPointerInteractionAfterFirstMeasurement();'),
  );
});

test('subtitle sidebar open state is restored for replacement visible overlay windows', () => {
  const source = readMainSource();
  const openedBlock = source.match(
    /onOverlayModalOpened:\s*\(modal,\s*senderWindow\)\s*=>\s*\{(?<body>[\s\S]*?)\n      \},/,
  )?.groups?.body;
  const closedBlock = source.match(
    /onOverlayModalClosed:\s*\(modal,\s*senderWindow(?:,\s*instanceId)?\)\s*=>\s*\{(?<body>[\s\S]*?)\n      \},/,
  )?.groups?.body;
  const depsBlock = source.match(/getSubtitleSidebarOpen:\s*\(\)\s*=>\s*(?<body>[^\n,]+)/)?.groups
    ?.body;

  assert.ok(openedBlock);
  assert.ok(closedBlock);
  assert.ok(depsBlock);
  assert.match(openedBlock, /if \(modal === 'subtitle-sidebar'/);
  assert.match(openedBlock, /subtitleSidebarRequestedOpen = true;/);
  assert.match(closedBlock, /if \(modal === 'subtitle-sidebar'/);
  assert.match(closedBlock, /subtitleSidebarRequestedOpen = false;/);
  assert.match(depsBlock, /subtitleSidebarRequestedOpen/);
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

test('Linux visible overlay recreation clears stale input state before creating replacement window', () => {
  const source = readMainSource();
  const runtimeSource = readSource('src/main/runtime/visible-overlay-interaction-runtime.ts');
  const actionBlock = source.match(
    /const linuxOverlayModeRuntime = createLinuxOverlayModeRuntime\(\{[\s\S]*?createWindow: \(\) => \{(?<body>[\s\S]*?)\n  \},/,
  )?.groups?.body;
  const resetBlock = runtimeSource.match(
    /function resetVisibleOverlayInputState\(\): void \{(?<body>[\s\S]*?)\n  \}/,
  )?.groups?.body;

  assert.ok(actionBlock);
  assert.ok(resetBlock);
  assert.match(actionBlock, /resetVisibleOverlayInputState\(\);/);
  assert.match(resetBlock, /overlayContentMeasurementStore\.clear\('visible'\);/);
  assert.ok(
    actionBlock.indexOf('resetVisibleOverlayInputState();') <
      actionBlock.indexOf('createMainWindow();'),
  );
});

test('Linux visible overlay recreation avoids display fallback before tracked geometry exists', () => {
  const source = readMainSource();
  const actionBlock = source.match(
    /const linuxOverlayModeRuntime = createLinuxOverlayModeRuntime\(\{[\s\S]*?refreshWindow: \(\) => \{(?<body>[\s\S]*?)\n  \},/,
  )?.groups?.body;

  assert.ok(actionBlock);
  assert.match(
    actionBlock,
    /const trackedGeometry = overlayGeometryRuntime\.getCurrentTrackedOverlayGeometry\(\);/,
  );
  assert.match(
    actionBlock,
    /if \(trackedGeometry\) overlayManager\.setOverlayWindowBounds\(trackedGeometry\);/,
  );
  assert.match(actionBlock, /overlayManager\.setOverlayWindowBounds\(trackedGeometry\);/);
  assert.doesNotMatch(actionBlock, /setOverlayWindowBounds\(getCurrentOverlayGeometry\(\)\)/);
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

test('Linux visible overlay show starts input grace before first measurement', () => {
  const source = readMainSource();
  const setVisibleBlock = source.match(
    /function setVisibleOverlayVisible\(visible: boolean\): void \{(?<body>[\s\S]*?)\n\}/,
  )?.groups?.body;
  const toggleBlock = source.match(
    /function toggleVisibleOverlay\(\): void \{(?<body>[\s\S]*?)\n\}/,
  )?.groups?.body;
  const setOverlayBlock = source.match(
    /function setOverlayVisible\(visible: boolean\): void \{(?<body>[\s\S]*?)\n\}/,
  )?.groups?.body;

  for (const block of [setVisibleBlock, toggleBlock, setOverlayBlock]) {
    assert.ok(block);
    const resetIndex = block.indexOf(
      'visibleOverlayInteractionRuntime.resetLinuxVisibleOverlayStartupInputPrimer();',
    );
    const graceIndex = block.indexOf(
      'visibleOverlayInteractionRuntime.startLinuxVisibleOverlayStartupInputGrace();',
    );
    const primeIndex = block.indexOf(
      'void autoplaySubtitlePrimingRuntime.primeCurrentSubtitleForVisibleOverlay();',
    );
    assert.ok(resetIndex >= 0);
    assert.ok(graceIndex >= 0);
    assert.ok(primeIndex >= 0);
    assert.ok(resetIndex < graceIndex);
    assert.ok(graceIndex < primeIndex);
  }
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

test('subtitle sidebar snapshot prefers cached YouTube parsed cues before active-source parsing', () => {
  const source = readMainSource();
  const snapshotBlock = source.match(
    /getSubtitleSidebarSnapshot:\s*async\s*\(\)\s*=>\s*\{(?<body>[\s\S]*?const resolvedSource = await resolveActiveSubtitleSidebarSourceHandler)/,
  )?.groups?.body;

  assert.ok(snapshotBlock);
  assert.match(snapshotBlock, /shouldUseCachedYoutubeParsedCues\(/);
  assert.match(snapshotBlock, /cachedMediaPath:\s*appState\.activeParsedSubtitleMediaPath/);
  assert.match(snapshotBlock, /cachedCueCount:\s*appState\.activeParsedSubtitleCues\.length/);
  assert.ok(
    snapshotBlock.indexOf('shouldUseCachedYoutubeParsedCues(') <
      snapshotBlock.indexOf('resolveActiveSubtitleSidebarSourceHandler'),
  );
});
