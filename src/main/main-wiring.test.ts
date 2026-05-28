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

test('autoplay subtitle prime prefers cached annotated payload before raw fallback', () => {
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
  assert.match(actionBlock, /const rawPayload = withCurrentSubtitleTiming\(\{ text, tokens: null \}\);/);
  assert.ok(
    actionBlock.indexOf('consumeCachedSubtitle(text)') <
      actionBlock.indexOf('withCurrentSubtitleTiming({ text, tokens: null })'),
  );
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
      actionBlock.indexOf('subtitleProcessingController.refreshCurrentSubtitle(appState.currentSubText);'),
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
