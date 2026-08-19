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

test('media path changes start the YouTube media cache coordinator', () => {
  const source = readMainSource();
  const actionBlock = source.match(
    /updateCurrentMediaPath:\s*\(path\)\s*=>\s*\{(?<body>[\s\S]*?)\n    restoreMpvSubVisibility:/,
  )?.groups?.body;

  assert.ok(actionBlock);
  assert.match(actionBlock, /youtubeMediaCachePlaybackRuntime\.handleMediaPathChange\(path\);/);
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

test('mpv startup signals start overlay loading OSD before readiness work', () => {
  const source = readMainSource();
  const connectedBlock = source.match(
    /onMpvConnected:\s*\(\)\s*=>\s*\{(?<body>[\s\S]*?)\n    \},\n    maybeRunAnilistPostWatchUpdate:/,
  )?.groups?.body;
  const mediaPathBlock = source.match(
    /updateCurrentMediaPath:\s*\(path\)\s*=>\s*\{(?<body>[\s\S]*?)\n    restoreMpvSubVisibility:/,
  )?.groups?.body;
  const setVisibleBlock = source.match(
    /function setVisibleOverlayVisible\(visible: boolean\): void \{(?<body>[\s\S]*?)\n\}/,
  )?.groups?.body;

  assert.ok(connectedBlock);
  assert.ok(mediaPathBlock);
  assert.ok(setVisibleBlock);
  assert.match(connectedBlock, /maybeStartOverlayLoadingOsd\(\);/);
  assert.match(
    mediaPathBlock,
    /const normalizedPath = path\.trim\(\);\s+maybeStartOverlayLoadingOsd\(normalizedPath\);/,
  );
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

test('overlay loading dismiss notifies mpv plugin to stop early loading OSD', () => {
  const source = readSource('src/main/runtime/overlay-notifications-runtime.ts');
  const dismissBlock = source.match(
    /function dismissOverlayLoadingStatusNotification\(\): void \{(?<body>[\s\S]*?)\n  \}/,
  )?.groups?.body;

  assert.ok(dismissBlock);
  assert.match(
    dismissBlock,
    /sendMpvCommandRuntime\(deps\.getMpvClient\(\), \[\s*'script-message',\s*'subminer-overlay-loading-ready',\s*\]\);/,
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

test('remote media keeps parsed cues when the active subtitle source cannot be resolved', () => {
  const source = readMainSource();
  const actionBlock = source.match(
    /createRefreshSubtitlePrefetchFromActiveTrackHandler\(\{(?<body>[\s\S]*?)\n  \}\);/,
  )?.groups?.body;

  assert.ok(actionBlock);
  assert.match(actionBlock, /isYoutubeMediaPath\(videoPath\) \|\| isRemoteMediaPath\(videoPath\)/);
});

test('jellyfin subtitle preload seeds the tokenization prefetch directly', () => {
  const source = readMainSource();
  const actionBlock = source.match(
    /preloadJellyfinExternalSubtitlesMainDeps:\s*\{(?<body>[\s\S]*?)\n  \},/,
  )?.groups?.body;

  assert.ok(actionBlock);
  assert.match(
    actionBlock,
    /initSubtitlePrefetch: \(sourcePath\) =>\s*subtitlePrefetchRuntime\.refreshSubtitleSidebarFromSource\(sourcePath\),/,
  );
});

test('update overlay notification action triggers install flow', () => {
  const source = readMainSource();
  const runtimeSource = readSource('src/main/runtime/overlay-notifications-runtime.ts');

  assert.match(
    source,
    /handleOverlayNotificationAction:\s*\(notificationId,\s*actionId,\s*noteId\)\s*=>/,
  );
  assert.match(source, /notificationId === UPDATE_AVAILABLE_NOTIFICATION_ID/);
  assert.match(source, /actionId === INSTALL_UPDATE_ACTION_ID/);
  assert.match(source, /installWhenAvailable:\s*true/);
  assert.match(source, /actionId === OPEN_ANKI_CARD_ACTION_ID && noteId !== undefined/);
  assert.match(runtimeSource, /deps\.getAnkiIntegration\(\)\?\.openNoteInAnki\(noteId\)/);
  assert.match(
    runtimeSource,
    /deps\.getRuntimeOptionsManager\(\)\?\.getEffectiveAnkiConnectConfig/,
  );
  assert.match(
    runtimeSource,
    /new AnkiConnectClient\(\s*effectiveAnkiConfig\.url \|\| DEFAULT_CONFIG\.ankiConnect\.url/,
  );
  assert.match(runtimeSource, /fallbackClient\.openNoteInBrowser\(noteId\)/);
});

test('subtitle change pauses prefetch without restarting its run before tokenizing current line', () => {
  const source = readMainSource();
  const actionBlock = source.match(
    /onSubtitleChange:\s*\(text\)\s*=>\s*\{(?<body>[\s\S]*?)\n    \},\n    refreshDiscordPresence:/,
  )?.groups?.body;

  assert.ok(actionBlock);
  assert.match(actionBlock, /subtitlePrefetchService\?\.pause\(\);/);
  // Restarting the run per line (onSeek) discards in-flight prefetch work;
  // only real seeks restart via onTimePosUpdate.
  assert.doesNotMatch(actionBlock, /subtitlePrefetchService\?\.onSeek\(/);
  assert.match(actionBlock, /subtitleProcessingController\.onSubtitleChange\(text\)/);
  assert.ok(
    actionBlock.indexOf('subtitlePrefetchService?.pause();') <
      actionBlock.indexOf('subtitleProcessingController.onSubtitleChange(text)'),
  );
  // A repeated subtitle emits nothing, so the pause has to be released here or
  // prefetching idles until the next distinct line.
  assert.match(
    actionBlock,
    /if \(!subtitleProcessingController\.onSubtitleChange\(text\)\) \{[\s\S]*?subtitlePrefetchService\?\.resume\(\);/,
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
  assert.doesNotMatch(actionBlock, /withCurrentSubtitleTiming\(\{ text, tokens: null \}\)/);
  assert.doesNotMatch(actionBlock, /broadcastToOverlayWindows\('subtitle:set', rawPayload\)/);
  assert.match(actionBlock, /subtitleProcessingController\.onSubtitleChange\(text\);/);
  assert.ok(
    actionBlock.indexOf('consumeCachedSubtitle(text)') <
      actionBlock.indexOf('subtitleProcessingController.onSubtitleChange(text);'),
  );
});

test('autoplay subtitle prime reuses active parsed cues before synthetic warm release', () => {
  const source = readMainSource();
  const runtimeDepsBlock = source.match(
    /const autoplaySubtitlePrimingRuntime = createAutoplaySubtitlePrimingRuntime\(\{(?<body>[\s\S]*?)\n\}\);/,
  )?.groups?.body;
  const primeSource = readSource('src/main/runtime/autoplay-subtitle-priming-runtime.ts');
  const emptyTextBlock = primeSource.match(
    /if \(!text\.trim\(\) && isCurrentAutoplayMediaPath\(mediaPath\)\) \{(?<body>[\s\S]*?)\n    \}/,
  )?.groups?.body;

  assert.ok(runtimeDepsBlock);
  assert.match(
    runtimeDepsBlock,
    /getActiveParsedSubtitleCues:\s*\(\) => appState\.activeParsedSubtitleCues/,
  );

  assert.ok(emptyTextBlock);
  assert.ok(
    emptyTextBlock.indexOf('await deps.refreshSubtitlePrefetchFromActiveTrack()') <
      emptyTextBlock.indexOf(
        'await primeAutoplaySubtitleFromParsedCues(mediaPath, deps.getActiveParsedSubtitleCues())',
      ),
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

test('warm tokenization release can signal readiness before the first subtitle appears', () => {
  const source = readMainSource();
  const warmReleaseBlock = source.match(
    /signalAutoplayReadyFromWarmTokenization = createAutoplayTokenizationWarmRelease\(\{(?<body>[\s\S]*?)\n\}\);/,
  )?.groups?.body;
  const signalBlock = source.match(
    /function signalCurrentSubtitleAutoplayReady\(\): void \{(?<body>[\s\S]*?)\n\}/,
  )?.groups?.body;
  const currentPayloadBlock = source.match(
    /function getCurrentAutoplaySubtitlePayload\(\): SubtitleData \| null \{(?<body>[\s\S]*?)\n\}/,
  )?.groups?.body;

  assert.ok(warmReleaseBlock);
  assert.match(
    warmReleaseBlock,
    /signalAutoplayReady: \(\) => signalCurrentSubtitleAutoplayReady\(\)/,
  );

  assert.ok(signalBlock);
  assert.match(signalBlock, /const payload = getCurrentAutoplaySubtitlePayload\(\);/);
  assert.match(signalBlock, /if \(payload\) \{/);
  assert.match(signalBlock, /if \(!appState\.currentSubText\.trim\(\)\) \{/);
  assert.match(signalBlock, /text: '__warm__'/);

  assert.ok(currentPayloadBlock);
  assert.match(currentPayloadBlock, /appState\.currentSubtitleData/);
  assert.match(currentPayloadBlock, /payload\.text !== appState\.currentSubText/);
});

test('stats server Yomitan note creation honors configured Anki server override policy', () => {
  const source = readSource('src/main/runtime/stats-server-runtime.ts');
  const startStatsServerBlock = source.match(
    /statsServer = startStatsServer\(\{(?<body>[\s\S]*?)\n      \}\);/,
  )?.groups?.body;
  const addYomitanNoteBlock = startStatsServerBlock?.match(
    /addYomitanNote:\s*async\s*\(word: string\)\s*=>\s*\{(?<body>[\s\S]*?)\n        \},/,
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
    /function createLinuxVisibleOverlayWindowForCurrentMode\([\s\S]*?\): void \{(?<body>[\s\S]*?)\n\}/,
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
    /function createLinuxVisibleOverlayWindowForCurrentMode\([\s\S]*?\): void \{(?<body>[\s\S]*?)\n\}/,
  )?.groups?.body;

  assert.ok(actionBlock);
  assert.match(
    actionBlock,
    /const trackedGeometry = overlayGeometryRuntime\.getCurrentTrackedOverlayGeometry\(\);/,
  );
  assert.match(actionBlock, /if \(trackedGeometry\) \{/);
  assert.match(actionBlock, /overlayManager\.setOverlayWindowBounds\(trackedGeometry\);/);
  assert.doesNotMatch(actionBlock, /setOverlayWindowBounds\(getCurrentOverlayGeometry\(\)\)/);
});

test('subtitle annotation updates invalidate prefetched tokenizations before refreshing current subtitle', () => {
  const source = readMainSource();
  const actionBlock = source.match(
    /function refreshCurrentSubtitleAnnotations\(\): void \{(?<body>[\s\S]*?)\n\}/,
  )?.groups?.body;

  assert.ok(actionBlock);
  assert.match(actionBlock, /subtitleProcessingController\.invalidateTokenizationCache\(\);/);
  assert.match(actionBlock, /subtitlePrefetchService\?\.onSeek\(lastObservedTimePos\);/);
  assert.match(
    actionBlock,
    /if \(!subtitleProcessingController\.refreshCurrentSubtitle\(appState\.currentSubText\)\) \{[\s\S]*?subtitlePrefetchService\?\.resume\(\);/,
  );
  assert.ok(
    actionBlock.indexOf('subtitleProcessingController.invalidateTokenizationCache();') <
      actionBlock.indexOf(
        'subtitleProcessingController.refreshCurrentSubtitle(appState.currentSubText)',
      ),
  );
});

test('character portrait index readiness refreshes cached subtitle annotations', () => {
  const source = readMainSource();
  const lookupDeps = source.match(
    /const characterDictionaryImageLookup = createCharacterDictionaryImageLookup\(\{(?<body>[\s\S]*?)\n\}\);/,
  )?.groups?.body;

  assert.ok(lookupDeps);
  assert.match(lookupDeps, /onIndexReady: \(\) => refreshCurrentSubtitleAnnotations\(\),/);
  assert.match(
    lookupDeps,
    /onIndexReadyError: \(error\) =>[\s\S]*?logger\.warn\([\s\S]*?character portrait index became ready\.[\s\S]*?error,/,
  );
});

test('subtitle processing controller resumes prefetch on settle, not on its emits', () => {
  const source = readMainSource();
  const depsBlock = source.match(
    /createBuildSubtitleProcessingControllerMainDepsHandler\(\{(?<body>[\s\S]*?)\n  \}\);/,
  )?.groups?.body;

  assert.ok(depsBlock);
  // A controller emit can be the provisional plain payload sent before the
  // scan runs, so it must not release the prefetch pause.
  assert.match(
    depsBlock,
    /emitSubtitle: \(payload\) => emitSubtitlePayload\(payload, \{ resumePrefetch: false \}\),/,
  );
  assert.match(
    depsBlock,
    /onProcessingSettled: \(\) => \{\s+subtitlePrefetchService\?\.resume\(\);/,
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

test('manual visible overlay hide dismisses loading OSD', () => {
  const source = readMainSource();
  const setBlock = source.match(
    /function setVisibleOverlayVisible\(visible: boolean\): void \{(?<body>[\s\S]*?)\n\}/,
  )?.groups?.body;
  const toggleBlock = source.match(
    /function toggleVisibleOverlay\(\): void \{(?<body>[\s\S]*?)\n\}/,
  )?.groups?.body;
  const setOverlayBlock = source.match(
    /function setOverlayVisible\(visible: boolean\): void \{(?<body>[\s\S]*?)\n\}/,
  )?.groups?.body;

  assert.ok(setBlock);
  assert.ok(toggleBlock);
  assert.ok(setOverlayBlock);
  assert.match(setBlock, /if \(!visible\) \{[\s\S]*?dismissOverlayLoadingStatusNotification\(\);/);
  assert.match(
    toggleBlock,
    /if \(!nextVisible\) \{[\s\S]*?dismissOverlayLoadingStatusNotification\(\);/,
  );
  assert.match(
    setOverlayBlock,
    /if \(!visible\) \{[\s\S]*?dismissOverlayLoadingStatusNotification\(\);/,
  );
});

test('configured overlay notifications require visible ready overlay window', () => {
  const source = readSource('src/main/runtime/overlay-notifications-runtime.ts');
  const readinessBlock = source.match(
    /function isVisibleOverlayContentReady\(\): boolean \{(?<body>[\s\S]*?)\n  \}/,
  )?.groups?.body;
  const statusBlock = source.match(
    /function showConfiguredStatusNotification\([\s\S]*?\): void \{(?<body>[\s\S]*?)\n  \}/,
  )?.groups?.body;

  assert.ok(readinessBlock);
  assert.ok(statusBlock);
  assert.match(readinessBlock, /deps\.getVisibleOverlayVisible\(\)/);
  assert.match(readinessBlock, /isOverlayWindowReadyForNotification\(overlayWindow\)/);
  assert.doesNotMatch(readinessBlock, /isOverlayWindowContentReady\(overlayWindow\)/);
  assert.match(statusBlock, /isOverlayReady: \(\) => isVisibleOverlayContentReady\(\)/);
});

test('YouTube media cache lifecycle routes through configured status notifications', () => {
  const source = readMainSource();
  const cacheBlock = source.match(
    /const youtubeMediaCache = createYoutubeMediaCacheService\(\{(?<body>[\s\S]*?)\n\}\);\nconst waitForYoutubeMpvConnected/,
  )?.groups?.body;
  const startCacheBlock = source.match(
    /startYoutubeMediaCache:\s*\(url\)\s*=>\s*\{(?<body>[\s\S]*?)\n  \},\n  runYoutubePlaybackFlow/,
  )?.groups?.body;

  assert.ok(cacheBlock);
  assert.ok(startCacheBlock);
  assert.match(
    cacheBlock,
    /onDownloadStarted:\s*\(event\)\s*=>\s*\{[\s\S]*showConfiguredStatusNotification\(\s*'YouTube media cache is downloading\.'/,
  );
  assert.match(cacheBlock, /id:\s*'youtube-media-cache-status'/);
  assert.match(cacheBlock, /variant:\s*'progress'/);
  assert.match(cacheBlock, /persistent:\s*true/);
  assert.match(
    cacheBlock,
    /onReady:\s*\(event\)\s*=>\s*\{[\s\S]*showConfiguredStatusNotification\(\s*'YouTube media cache ready\.'/,
  );
  assert.match(cacheBlock, /variant:\s*'success'/);
  assert.match(cacheBlock, /notifyNoQueued:\s*false/);
  assert.match(
    startCacheBlock,
    /const mediaCacheConfig = configService\.getConfig\(\)\.youtube\.mediaCache;/,
  );
  assert.match(startCacheBlock, /mode:\s*mediaCacheConfig\.mode/);
  assert.match(startCacheBlock, /maxHeight:\s*mediaCacheConfig\.maxHeight/);
});

test('subtitle broadcasts share one frequency options snapshot per emitted payload', () => {
  const source = readMainSource();
  const emitBlock = source.match(
    /function emitSubtitlePayload\([\s\S]*?\): void \{(?<body>[\s\S]*?)\n\}/,
  )?.groups?.body;
  const frequencyOptionsSnapshot = emitBlock?.match(
    /const frequencyDictionary = configService\.getConfig\(\)\.subtitleStyle\.frequencyDictionary;(?<body>[\s\S]*?)\n  \};/,
  )?.[0];

  assert.ok(emitBlock);
  assert.ok(frequencyOptionsSnapshot);
  assert.match(frequencyOptionsSnapshot, /const frequencyOptions = \{/);
  assert.match(frequencyOptionsSnapshot, /enabled:\s*frequencyDictionary\.enabled/);
  assert.match(frequencyOptionsSnapshot, /topX:\s*frequencyDictionary\.topX/);
  assert.match(frequencyOptionsSnapshot, /mode:\s*frequencyDictionary\.mode/);
  assert.equal((frequencyOptionsSnapshot.match(/configService\.getConfig\(\)/g) ?? []).length, 1);
  assert.match(emitBlock, /subtitleWsService\.broadcast\(timedPayload, frequencyOptions\);/);
  assert.match(
    emitBlock,
    /annotationSubtitleWsService\.broadcast\(timedPayload, frequencyOptions\);/,
  );
});

test('annotation upgrades skip the duplicate basic websocket event', () => {
  const source = readMainSource();
  const emitBlock = source.match(
    /function emitSubtitlePayload\([\s\S]*?\): void \{(?<body>[\s\S]*?)\n\}/,
  )?.groups?.body;

  assert.ok(emitBlock);
  assert.match(
    emitBlock,
    /const isAnnotationUpgrade = isSubtitleAnnotationUpgrade\(currentSubtitleData, timedPayload\);/,
  );
  assert.match(
    emitBlock,
    /if \(!isAnnotationUpgrade\) \{\s+subtitleWsService\.broadcast\(timedPayload, frequencyOptions\);\s+\}/,
  );
  assert.equal(
    (emitBlock.match(/overlayManager\.broadcastToOverlayWindows\('subtitle:set'/g) ?? []).length,
    1,
  );
  assert.equal(
    (
      emitBlock.match(
        /annotationSubtitleWsService\.broadcast\(timedPayload, frequencyOptions\)/g,
      ) ?? []
    ).length,
    1,
  );
});

test('websocket frequency options callbacks each read one configuration snapshot', () => {
  const source = readMainSource();
  const subtitleBlock = source.match(
    /startSubtitleWebsocket:\s*\(port: number\)\s*=>\s*\{(?<body>[\s\S]*?)\n    \},\n    startAnnotationWebsocket:/,
  )?.groups?.body;
  const annotationBlock = source.match(
    /startAnnotationWebsocket:\s*\(port: number\)\s*=>\s*\{(?<body>[\s\S]*?)\n    \},\n    startTexthooker:/,
  )?.groups?.body;
  const subtitleFrequencyOptionsCallback = subtitleBlock?.match(
    /(?<body>\(\) => \{\s+const frequencyDictionary = configService\.getConfig\(\)\.subtitleStyle\.frequencyDictionary;[\s\S]*?\s+return \{[\s\S]*?\s+\};\s+\})/,
  )?.groups?.body;
  const annotationFrequencyOptionsCallback = annotationBlock?.match(
    /(?<body>\(\) => \{\s+const frequencyDictionary = configService\.getConfig\(\)\.subtitleStyle\.frequencyDictionary;[\s\S]*?\s+return \{[\s\S]*?\s+\};\s+\})/,
  )?.groups?.body;

  assert.ok(subtitleFrequencyOptionsCallback);
  assert.ok(annotationFrequencyOptionsCallback);
  for (const callback of [subtitleFrequencyOptionsCallback, annotationFrequencyOptionsCallback]) {
    assert.match(callback, /return \{[\s\S]*enabled:\s*frequencyDictionary\.enabled/);
    assert.match(callback, /topX:\s*frequencyDictionary\.topX/);
    assert.match(callback, /mode:\s*frequencyDictionary\.mode/);
    assert.equal((callback.match(/configService\.getConfig\(\)/g) ?? []).length, 1);
  }
});

test('mpv connection flushes queued configured OSD notifications', () => {
  const source = readMainSource();
  const connectedBlock = source.match(
    /onMpvConnected:\s*\(\)\s*=>\s*\{(?<body>[\s\S]*?)\n    \},\n    maybeRunAnilistPostWatchUpdate:/,
  )?.groups?.body;

  assert.ok(connectedBlock);
  assert.match(source, /flushQueuedMpvOsdNotifications/);
  assert.match(connectedBlock, /flushQueuedMpvOsdNotifications\(\);/);
});

test('manual visible overlay show primes current subtitle from mpv before relying on live events', () => {
  const source = readMainSource();
  const setBlock = source.match(
    /function setVisibleOverlayVisible\(visible: boolean\): void \{(?<body>[\s\S]*?)\n\}/,
  )?.groups?.body;
  const toggleBlock = source.match(
    /function toggleVisibleOverlay\(\): void \{(?<body>[\s\S]*?)\n\}/,
  )?.groups?.body;

  assert.ok(setBlock);
  assert.ok(toggleBlock);
  assert.match(
    setBlock,
    /if \(visible\) \{\s+maybeStartOverlayLoadingOsd\(\);\s+visibleOverlayInteractionRuntime\.resetLinuxVisibleOverlayStartupInputPrimer\(\);\s+visibleOverlayInteractionRuntime\.startLinuxVisibleOverlayStartupInputGrace\(\);\s+visibleOverlayInteractionRuntime\.restoreVisibleOverlayWindowShapeForShow\(\);\s+void ensureOverlayMpvSubtitlesHidden\(\);\s+void autoplaySubtitlePrimingRuntime\.primeCurrentSubtitleForVisibleOverlay\(\);/,
  );
  assert.match(
    toggleBlock,
    /else \{\s+maybeStartOverlayLoadingOsd\(\);\s+visibleOverlayInteractionRuntime\.resetLinuxVisibleOverlayStartupInputPrimer\(\);\s+visibleOverlayInteractionRuntime\.startLinuxVisibleOverlayStartupInputGrace\(\);\s+visibleOverlayInteractionRuntime\.restoreVisibleOverlayWindowShapeForShow\(\);\s+void ensureOverlayMpvSubtitlesHidden\(\);\s+void autoplaySubtitlePrimingRuntime\.primeCurrentSubtitleForVisibleOverlay\(\);/,
  );
});

test('Linux visible overlay show/reset does not leave an empty X11 window shape', () => {
  const source = readMainSource();
  const runtimeSource = readSource('src/main/runtime/visible-overlay-interaction-runtime.ts');
  const resetBlock = runtimeSource.match(
    /function resetVisibleOverlayInputState\(\): void \{(?<body>[\s\S]*?)\n  \}/,
  )?.groups?.body;
  const setBlock = source.match(
    /function setVisibleOverlayVisible\(visible: boolean\): void \{(?<body>[\s\S]*?)\n\}/,
  )?.groups?.body;

  assert.ok(resetBlock);
  assert.ok(setBlock);
  assert.match(resetBlock, /restoreLinuxOverlayWindowShape\(mainWindow\);/);
  assert.doesNotMatch(source, /setShape\?\.\(\[\]\)|setShape\(\[\]\)/);
  assert.doesNotMatch(runtimeSource, /setShape\?\.\(\[\]\)|setShape\(\[\]\)/);
  assert.match(
    setBlock,
    /if \(visible\) \{\s+maybeStartOverlayLoadingOsd\(\);\s+visibleOverlayInteractionRuntime\.resetLinuxVisibleOverlayStartupInputPrimer\(\);\s+visibleOverlayInteractionRuntime\.startLinuxVisibleOverlayStartupInputGrace\(\);\s+visibleOverlayInteractionRuntime\.restoreVisibleOverlayWindowShapeForShow\(\);\s+void ensureOverlayMpvSubtitlesHidden\(\);/,
  );
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
