import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs';
import path from 'node:path';

function readMainSource(): string {
  return fs.readFileSync(path.join(process.cwd(), 'src/main.ts'), 'utf8');
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

test('media path changes clear rendered subtitle state without clearing same-youtube parsed cues', () => {
  const source = readMainSource();
  const actionBlock = source.match(
    /updateCurrentMediaPath:\s*\(path\)\s*=>\s*\{(?<body>[\s\S]*?)autoplayReadyGate\.invalidatePendingAutoplayReadyFallbacks\(\);/,
  )?.groups?.body;

  assert.ok(actionBlock);
  assert.match(actionBlock, /appState\.currentSubText = '';/);
  assert.match(actionBlock, /appState\.currentSubAssText = '';/);
  assert.match(actionBlock, /appState\.currentSubtitleData = null;/);
  assert.match(actionBlock, /isSameYoutubeMediaPath\(/);
  assert.match(actionBlock, /if \(!preserveParsedSubtitleCues\)/);
  assert.match(actionBlock, /appState\.activeParsedSubtitleCues = \[\];/);
  assert.match(actionBlock, /appState\.activeParsedSubtitleSource = null;/);
  assert.match(actionBlock, /appState\.activeParsedSubtitleMediaPath = null;/);
  assert.match(actionBlock, /lastObservedTimePos = 0;/);
  assert.match(actionBlock, /broadcastToOverlayWindows\('subtitle:set',/);
  assert.match(actionBlock, /subtitleWsService\.broadcast\(/);
  assert.match(actionBlock, /annotationSubtitleWsService\.broadcast\(/);
  assert.ok(
    actionBlock.indexOf('appState.currentSubtitleData = null;') <
      actionBlock.indexOf("broadcastToOverlayWindows('subtitle:set'"),
  );
});

test('same media path updates do not reset autoplay ready fallback state', () => {
  const source = readMainSource();
  const actionBlock = source.match(
    /updateCurrentMediaPath:\s*\(path\)\s*=>\s*\{(?<body>[\s\S]*?)\n    restoreMpvSubVisibility:/,
  )?.groups?.body;

  assert.ok(actionBlock);
  assert.match(
    actionBlock,
    /annotationSubtitleWsService\.broadcast\(resetSubtitlePayload, frequencyOptions\);\s+autoplayReadyGate\.invalidatePendingAutoplayReadyFallbacks\(\);\s+\}\s+currentMediaTokenizationGate\.updateCurrentMediaPath\(path\);/,
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
    /if \(!nextVisible\) \{\s+autoplayReadyGate\.markCurrentMediaAutoplayReady\(\);\s+cancelPendingLinuxMpvFullscreenOverlayRefreshBurst\(\);/,
  );
});

test('manual visible overlay hide clears stale overlay input state', () => {
  const source = readMainSource();
  const actionBlock = source.match(
    /function setVisibleOverlayVisible\(visible: boolean\): void \{(?<body>[\s\S]*?)\n\}/,
  )?.groups?.body;

  assert.ok(actionBlock);
  assert.match(
    actionBlock,
    /if \(!visible\) \{\s+autoplayReadyGate\.markCurrentMediaAutoplayReady\(\);\s+cancelPendingLinuxMpvFullscreenOverlayRefreshBurst\(\);\s+resetVisibleOverlayInputState\(\);/,
  );
});

test('subtitle sidebar media path tag is assigned after prefetch succeeds', () => {
  const source = readMainSource();
  const actionBlock = source.match(
    /async function refreshSubtitleSidebarFromSource\([\s\S]*?\): Promise<void> \{(?<body>[\s\S]*?)\n\}/,
  )?.groups?.body;

  assert.ok(actionBlock);
  assert.match(
    actionBlock,
    /const nextMediaPath = mediaPath\?\.trim\(\) \|\| getCurrentAutoplayMediaPath\(\);/,
  );
  assert.ok(
    actionBlock.indexOf('subtitlePrefetchInitController.initSubtitlePrefetch') <
      actionBlock.indexOf('appState.activeParsedSubtitleMediaPath = nextMediaPath;'),
  );
});

test('subtitle change re-prioritizes prefetch around live playback before tokenizing current line', () => {
  const source = readMainSource();
  const actionBlock = source.match(
    /onSubtitleChange:\s*\(text\)\s*=>\s*\{(?<body>[\s\S]*?)\n    \},\n    refreshDiscordPresence:/,
  )?.groups?.body;

  assert.ok(actionBlock);
  assert.match(actionBlock, /subtitlePrefetchService\?\.pause\(\);/);
  assert.match(actionBlock, /subtitlePrefetchService\?\.onSeek\(lastObservedTimePos\);/);
  assert.match(actionBlock, /subtitleProcessingController\.onSubtitleChange\(text\);/);
  assert.ok(
    actionBlock.indexOf('subtitlePrefetchService?.pause();') <
      actionBlock.indexOf('subtitlePrefetchService?.onSeek(lastObservedTimePos);'),
  );
  assert.ok(
    actionBlock.indexOf('subtitlePrefetchService?.onSeek(lastObservedTimePos);') <
      actionBlock.indexOf('subtitleProcessingController.onSubtitleChange(text);'),
  );
});

test('autoplay subtitle prime emits cached annotations and avoids raw fallback overlay flashes', () => {
  const source = readMainSource();
  const actionBlock = source.match(
    /function emitAutoplayPrimedSubtitle\([\s\S]*?\): boolean \{(?<body>[\s\S]*?)\n\}/,
  )?.groups?.body;

  assert.ok(actionBlock);
  assert.match(
    actionBlock,
    /const cachedPayload = subtitleProcessingController\.consumeCachedSubtitle\(text\);/,
  );
  assert.match(actionBlock, /if \(cachedPayload\) \{/);
  assert.match(actionBlock, /emitSubtitlePayload\(cachedPayload\);/);
  assert.doesNotMatch(actionBlock, /withCurrentSubtitleTiming\(\{ text, tokens: null \}\)/);
  assert.doesNotMatch(actionBlock, /broadcastToOverlayWindows\('subtitle:set', rawPayload\)/);
  assert.match(actionBlock, /subtitleProcessingController\.onSubtitleChange\(text\);/);
  assert.ok(
    actionBlock.indexOf('consumeCachedSubtitle(text)') <
      actionBlock.indexOf('subtitleProcessingController.onSubtitleChange(text);'),
  );
});

test('startup autoplay release is tied to tokenization and visible overlay measurement readiness', () => {
  const source = readMainSource();
  const gateBlock = source.match(
    /const autoplayReadyGate = createAutoplayReadyGate\(\{(?<body>[\s\S]*?)\n\}\);/,
  )?.groups?.body;
  const measurementBlock = source.match(
    /reportOverlayContentBounds:\s*\(payload: unknown\)\s*=>\s*\{(?<body>[\s\S]*?)\n      \},/,
  )?.groups?.body;

  assert.ok(gateBlock);
  assert.match(gateBlock, /isSignalTargetReady:\s*\(signal\) =>/);
  assert.match(gateBlock, /isTokenizationWarmupReady\(\)/);
  assert.match(gateBlock, /isVisibleOverlayAutoplayTargetReady\(/);
  assert.match(gateBlock, /getLatestVisibleMeasurement:/);

  assert.ok(measurementBlock);
  assert.match(measurementBlock, /overlayContentMeasurementStore\.report\(payload\)/);
  assert.match(measurementBlock, /autoplayReadyGate\.flushPendingAutoplayReadySignal\(\)/);
});

test('accepted visible overlay measurement immediately refreshes Linux pointer interaction', () => {
  const source = readMainSource();
  const measurementBlock = source.match(
    /reportOverlayContentBounds:\s*\(payload: unknown\)\s*=>\s*\{(?<body>[\s\S]*?)\n      \},/,
  )?.groups?.body;

  assert.ok(measurementBlock);
  assert.match(measurementBlock, /overlayContentMeasurementStore\.report\(payload\)/);
  assert.match(measurementBlock, /tickLinuxOverlayPointerInteractionNow\(\)/);
  assert.ok(
    measurementBlock.indexOf('overlayContentMeasurementStore.report(payload)') <
      measurementBlock.indexOf('tickLinuxOverlayPointerInteractionNow();'),
  );
});

test('subtitle sidebar open state is restored for replacement visible overlay windows', () => {
  const source = readMainSource();
  const openedBlock = source.match(
    /onOverlayModalOpened:\s*\(modal,\s*senderWindow\)\s*=>\s*\{(?<body>[\s\S]*?)\n      \},/,
  )?.groups?.body;
  const closedBlock = source.match(
    /onOverlayModalClosed:\s*\(modal,\s*senderWindow\)\s*=>\s*\{(?<body>[\s\S]*?)\n      \},/,
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

test('warm tokenization release reuses current subtitle payload instead of synthetic readiness', () => {
  const source = readMainSource();
  const warmReleaseBlock = source.match(
    /signalAutoplayReadyFromWarmTokenization = createAutoplayTokenizationWarmRelease\(\{(?<body>[\s\S]*?)\n\}\);/,
  )?.groups?.body;
  const currentPayloadBlock = source.match(
    /function getCurrentAutoplaySubtitlePayload\(\): SubtitleData \| null \{(?<body>[\s\S]*?)\n\}/,
  )?.groups?.body;

  assert.ok(warmReleaseBlock);
  assert.match(
    warmReleaseBlock,
    /signalAutoplayReady: \(\) => signalCurrentSubtitleAutoplayReady\(\)/,
  );
  assert.doesNotMatch(warmReleaseBlock, /__warm__/);

  assert.ok(currentPayloadBlock);
  assert.match(currentPayloadBlock, /appState\.currentSubtitleData/);
  assert.match(currentPayloadBlock, /payload\.text !== appState\.currentSubText/);
});

test('Linux visible overlay recreation clears stale input state before creating replacement window', () => {
  const source = readMainSource();
  const actionBlock = source.match(
    /function createLinuxVisibleOverlayWindowForCurrentMode\([\s\S]*?\): void \{(?<body>[\s\S]*?)\n\}/,
  )?.groups?.body;
  const resetBlock = source.match(
    /function resetVisibleOverlayInputState\(\): void \{(?<body>[\s\S]*?)\n\}/,
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
    /function createLinuxVisibleOverlayWindowForCurrentMode\([\s\S]*?\): void \{(?<body>[\s\S]*?)\n\}/,
  )?.groups?.body;

  assert.ok(actionBlock);
  assert.match(actionBlock, /const trackedGeometry = getCurrentTrackedOverlayGeometry\(\);/);
  assert.match(actionBlock, /if \(trackedGeometry\) \{/);
  assert.match(actionBlock, /overlayManager\.setOverlayWindowBounds\(trackedGeometry\);/);
  assert.doesNotMatch(actionBlock, /setOverlayWindowBounds\(getCurrentOverlayGeometry\(\)\)/);
});

test('known-word updates invalidate prefetched tokenizations before refreshing current subtitle', () => {
  const source = readMainSource();
  const actionBlock = source.match(
    /const refreshCurrentSubtitleAfterKnownWordUpdate = \(\): void => \{(?<body>[\s\S]*?)\n\};/,
  )?.groups?.body;

  assert.ok(actionBlock);
  assert.match(actionBlock, /subtitleProcessingController\.invalidateTokenizationCache\(\);/);
  assert.match(actionBlock, /subtitlePrefetchService\?\.onSeek\(lastObservedTimePos\);/);
  assert.match(
    actionBlock,
    /subtitleProcessingController\.refreshCurrentSubtitle\(appState\.currentSubText\);/,
  );
  assert.ok(
    actionBlock.indexOf('subtitleProcessingController.invalidateTokenizationCache();') <
      actionBlock.indexOf(
        'subtitleProcessingController.refreshCurrentSubtitle(appState.currentSubText);',
      ),
  );
});

test('manual visible overlay changes notify mpv plugin visibility state', () => {
  const source = readMainSource();
  const setBlock = source.match(
    /function setVisibleOverlayVisible\(visible: boolean\): void \{(?<body>[\s\S]*?)\n\}/,
  )?.groups?.body;
  const toggleBlock = source.match(
    /function toggleVisibleOverlay\(\): void \{(?<body>[\s\S]*?)\n\}/,
  )?.groups?.body;

  assert.ok(setBlock);
  assert.ok(toggleBlock);
  assert.match(setBlock, /notifyMpvPluginVisibleOverlayVisibility\(visible\);/);
  assert.match(toggleBlock, /const nextVisible = !overlayManager\.getVisibleOverlayVisible\(\);/);
  assert.match(toggleBlock, /notifyMpvPluginVisibleOverlayVisibility\(nextVisible\);/);
});

test('main process uses one shared mpv plugin runtime config helper', () => {
  const source = readMainSource();
  assert.match(source, /function getMpvPluginRuntimeConfig\(\)/);
  assert.equal((source.match(/socketPath: appState\.mpvSocketPath/g) ?? []).length, 1);
  assert.equal(
    (source.match(/binaryPath: getResolvedConfig\(\)\.mpv\.subminerBinaryPath/g) ?? []).length,
    0,
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
