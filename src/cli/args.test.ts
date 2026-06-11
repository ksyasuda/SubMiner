import test from 'node:test';
import assert from 'node:assert/strict';
import {
  commandNeedsOverlayStartupPrereqs,
  commandNeedsOverlayRuntime,
  hasExplicitCommand,
  isHeadlessInitialCommand,
  isStandaloneTexthookerCommand,
  parseArgs,
  shouldRunYomitanOnlyStartup,
  shouldStartApp,
} from './args';

test('parseArgs parses booleans and value flags', () => {
  const args = parseArgs([
    '--background',
    '--start',
    '--socket',
    '/tmp/mpv.sock',
    '--backend=hyprland',
    '--port',
    '6000',
    '--log-level',
    'warn',
    '--debug',
    '--jellyfin-play',
    '--jellyfin-server',
    'http://jellyfin.local:8096',
    '--jellyfin-item-id',
    'item-123',
    '--jellyfin-audio-stream-index',
    '2',
  ]);

  assert.equal(args.background, true);
  assert.equal(args.start, true);
  assert.equal(args.socketPath, '/tmp/mpv.sock');
  assert.equal(args.backend, 'hyprland');
  assert.equal(args.texthookerPort, 6000);
  assert.equal(args.logLevel, 'warn');
  assert.equal(args.debug, true);
  assert.equal(args.jellyfinPlay, true);
  assert.equal(args.jellyfinServer, 'http://jellyfin.local:8096');
  assert.equal(args.jellyfinItemId, 'item-123');
  assert.equal(args.jellyfinAudioStreamIndex, 2);
});

test('parseArgs ignores missing value after --log-level', () => {
  const args = parseArgs(['--log-level', '--start']);
  assert.equal(args.logLevel, undefined);
  assert.equal(args.start, true);
});

test('parseArgs captures update command and internal launcher paths', () => {
  const args = parseArgs([
    '--update',
    '--update-launcher-path',
    '/home/kyle/.local/bin/subminer',
    '--update-response-path',
    '/tmp/subminer-update-response.json',
  ]);

  assert.equal(args.update, true);
  assert.equal(args.updateLauncherPath, '/home/kyle/.local/bin/subminer');
  assert.equal(args.updateResponsePath, '/tmp/subminer-update-response.json');
  assert.equal(hasExplicitCommand(args), true);
  assert.equal(shouldStartApp(args), true);
  assert.equal(commandNeedsOverlayRuntime(args), false);
  assert.equal(shouldRunYomitanOnlyStartup(args), false);
});

test('parseArgs captures launch-mpv targets and keeps it out of app startup', () => {
  const args = parseArgs(['--launch-mpv', 'C:\\a.mkv', 'C:\\b.mkv']);
  assert.equal(args.launchMpv, true);
  assert.deepEqual(args.launchMpvTargets, ['C:\\a.mkv', 'C:\\b.mkv']);
  assert.equal(hasExplicitCommand(args), true);
  assert.equal(shouldStartApp(args), false);
});

test('parseArgs captures youtube startup forwarding flags', () => {
  const args = parseArgs([
    '--youtube-play',
    'https://youtube.com/watch?v=abc',
    '--youtube-mode',
    'generate',
  ]);

  assert.equal(args.youtubePlay, 'https://youtube.com/watch?v=abc');
  assert.equal(args.youtubeMode, 'generate');
  assert.equal(hasExplicitCommand(args), true);
  assert.equal(shouldStartApp(args), true);
});

test('parseArgs captures session action forwarding flags', () => {
  const args = parseArgs([
    '--toggle-stats-overlay',
    '--mark-watched',
    '--open-jimaku',
    '--open-youtube-picker',
    '--open-playlist-browser',
    '--toggle-primary-subtitle-bar',
    '--replay-current-subtitle',
    '--play-next-subtitle',
    '--shift-sub-delay-prev-line',
    '--shift-sub-delay-next-line',
    '--cycle-runtime-option',
    'anki.autoUpdateNewCards:prev',
    '--session-action',
    '{"actionId":"openCharacterDictionaryManager"}',
    '--copy-subtitle-count',
    '3',
    '--mine-sentence-count=2',
  ]);

  assert.equal(args.toggleStatsOverlay, true);
  assert.equal(args.markWatched, true);
  assert.equal(args.openJimaku, true);
  assert.equal(args.openYoutubePicker, true);
  assert.equal(args.openPlaylistBrowser, true);
  assert.equal(args.togglePrimarySubtitleBar, true);
  assert.equal(args.replayCurrentSubtitle, true);
  assert.equal(args.playNextSubtitle, true);
  assert.equal(args.shiftSubDelayPrevLine, true);
  assert.equal(args.shiftSubDelayNextLine, true);
  assert.equal(args.cycleRuntimeOptionId, 'anki.autoUpdateNewCards');
  assert.equal(args.cycleRuntimeOptionDirection, -1);
  assert.deepEqual(args.sessionAction, { actionId: 'openCharacterDictionaryManager' });
  assert.equal(args.copySubtitleCount, 3);
  assert.equal(args.mineSentenceCount, 2);
  assert.equal(hasExplicitCommand(args), true);
  assert.equal(shouldStartApp(args), true);
});

test('parseArgs captures internal playback feedback command', () => {
  const args = parseArgs(['--playback-feedback', 'You can skip by pressing TAB']);

  assert.equal(args.playbackFeedback, 'You can skip by pressing TAB');
  assert.equal(hasExplicitCommand(args), true);
  assert.equal(shouldStartApp(args), true);
  assert.equal(commandNeedsOverlayRuntime(args), true);
});

test('parseArgs ignores non-positive numeric session action counts', () => {
  const args = parseArgs(['--copy-subtitle-count=0', '--mine-sentence-count', '-1']);

  assert.equal(args.copySubtitleCount, undefined);
  assert.equal(args.mineSentenceCount, undefined);
});

test('youtube playback does not use generic overlay-runtime bootstrap classification', () => {
  const args = parseArgs(['--youtube-play', 'https://youtube.com/watch?v=abc']);

  assert.equal(commandNeedsOverlayRuntime(args), false);
  assert.equal(commandNeedsOverlayStartupPrereqs(args), true);
});

test('standalone texthooker classification excludes integrated start flow', () => {
  const standalone = parseArgs(['--texthooker']);
  const standaloneOpenBrowser = parseArgs(['--texthooker', '--open-browser']);
  const integrated = parseArgs(['--start', '--texthooker']);

  assert.equal(isStandaloneTexthookerCommand(standalone), true);
  assert.equal(standaloneOpenBrowser.texthookerOpenBrowser, true);
  assert.equal(isStandaloneTexthookerCommand(standaloneOpenBrowser), true);
  assert.equal(isStandaloneTexthookerCommand(integrated), false);
});

test('parseArgs handles jellyfin item listing controls', () => {
  const args = parseArgs([
    '--jellyfin-items',
    '--jellyfin-recursive=false',
    '--jellyfin-include-item-types',
    'Series,Movie,Folder',
  ]);

  assert.equal(args.jellyfinItems, true);
  assert.equal(args.jellyfinRecursive, false);
  assert.equal(args.jellyfinIncludeItemTypes, 'Series,Movie,Folder');
});

test('parseArgs handles space-separated jellyfin recursive control', () => {
  const args = parseArgs(['--jellyfin-items', '--jellyfin-recursive', 'false']);
  assert.equal(args.jellyfinRecursive, false);
});

test('parseArgs ignores unrecognized space-separated jellyfin recursive values', () => {
  const args = parseArgs(['--jellyfin-items', '--jellyfin-recursive', '--start']);
  assert.equal(args.jellyfinRecursive, undefined);
  assert.equal(args.start, true);
});

test('hasExplicitCommand and shouldStartApp preserve command intent', () => {
  const stopOnly = parseArgs(['--stop']);
  assert.equal(hasExplicitCommand(stopOnly), true);
  assert.equal(shouldStartApp(stopOnly), false);

  const launchMpv = parseArgs(['--launch-mpv']);
  assert.equal(launchMpv.launchMpv, true);
  assert.deepEqual(launchMpv.launchMpvTargets, []);
  assert.equal(hasExplicitCommand(launchMpv), true);
  assert.equal(shouldStartApp(launchMpv), false);

  const toggle = parseArgs(['--toggle-visible-overlay']);
  assert.equal(hasExplicitCommand(toggle), true);
  assert.equal(shouldStartApp(toggle), true);

  const noCommand = parseArgs(['--log-level', 'warn']);
  assert.equal(hasExplicitCommand(noCommand), false);
  assert.equal(shouldStartApp(noCommand), false);

  const refreshKnownWords = parseArgs(['--refresh-known-words']);
  assert.equal(refreshKnownWords.help, false);
  assert.equal(hasExplicitCommand(refreshKnownWords), true);
  assert.equal(shouldStartApp(refreshKnownWords), true);
  assert.equal(isHeadlessInitialCommand(refreshKnownWords), true);

  const update = parseArgs(['--update']);
  assert.equal(update.update, true);
  assert.equal(hasExplicitCommand(update), true);
  assert.equal(shouldStartApp(update), true);
  assert.equal(isHeadlessInitialCommand(update), true);

  const yomitan = parseArgs(['--yomitan']);
  assert.equal(yomitan.yomitan, true);
  assert.equal(hasExplicitCommand(yomitan), true);
  assert.equal(shouldStartApp(yomitan), true);
  assert.equal(shouldRunYomitanOnlyStartup(yomitan), true);

  const settings = parseArgs(['--settings']);
  assert.equal(settings.settings, true);
  assert.equal(hasExplicitCommand(settings), true);
  assert.equal(shouldStartApp(settings), true);
  assert.equal(shouldRunYomitanOnlyStartup(settings), false);
  assert.equal(commandNeedsOverlayRuntime(settings), false);
  assert.equal(commandNeedsOverlayStartupPrereqs(settings), false);

  const yomitanWithOverlay = parseArgs(['--yomitan', '--toggle-visible-overlay']);
  assert.equal(yomitanWithOverlay.yomitan, true);
  assert.equal(yomitanWithOverlay.toggleVisibleOverlay, true);
  assert.equal(shouldRunYomitanOnlyStartup(yomitanWithOverlay), false);

  const settingsDoesNotEnableYomitan = parseArgs(['--settings']);
  assert.equal(settingsDoesNotEnableYomitan.yomitan, false);

  const help = parseArgs(['--help']);
  assert.equal(help.help, true);
  assert.equal(hasExplicitCommand(help), true);
  assert.equal(shouldStartApp(help), false);
  assert.equal(shouldRunYomitanOnlyStartup(help), false);

  const appPing = parseArgs(['--app-ping']);
  assert.equal(appPing.appPing, true);
  assert.equal(hasExplicitCommand(appPing), true);
  assert.equal(shouldStartApp(appPing), false);

  const youtubePlay = parseArgs(['--youtube-play', 'https://youtube.com/watch?v=abc']);
  assert.equal(commandNeedsOverlayStartupPrereqs(youtubePlay), true);

  const anilistStatus = parseArgs(['--anilist-status']);
  assert.equal(anilistStatus.anilistStatus, true);
  assert.equal(hasExplicitCommand(anilistStatus), true);
  assert.equal(shouldStartApp(anilistStatus), false);

  const anilistRetryQueue = parseArgs(['--anilist-retry-queue']);
  assert.equal(anilistRetryQueue.anilistRetryQueue, true);
  assert.equal(hasExplicitCommand(anilistRetryQueue), true);
  assert.equal(shouldStartApp(anilistRetryQueue), false);

  const dictionaryCandidates = parseArgs([
    '--dictionary-candidates',
    '--dictionary-target',
    '/tmp/a.mkv',
  ]);
  assert.equal(dictionaryCandidates.dictionaryCandidates, true);
  assert.equal(dictionaryCandidates.dictionaryTarget, '/tmp/a.mkv');
  assert.equal(hasExplicitCommand(dictionaryCandidates), true);
  assert.equal(shouldStartApp(dictionaryCandidates), true);

  const dictionarySelect = parseArgs(['--dictionary-select', '--dictionary-anilist-id', '21355']);
  assert.equal(dictionarySelect.dictionarySelect, true);
  assert.equal(dictionarySelect.dictionaryAnilistId, 21355);
  assert.equal(hasExplicitCommand(dictionarySelect), true);
  assert.equal(shouldStartApp(dictionarySelect), true);

  const toggleStatsOverlay = parseArgs(['--toggle-stats-overlay']);
  assert.equal(toggleStatsOverlay.toggleStatsOverlay, true);
  assert.equal(hasExplicitCommand(toggleStatsOverlay), true);
  assert.equal(shouldStartApp(toggleStatsOverlay), true);

  const cycleRuntimeOption = parseArgs(['--cycle-runtime-option', 'anki.autoUpdateNewCards:next']);
  assert.equal(cycleRuntimeOption.cycleRuntimeOptionId, 'anki.autoUpdateNewCards');
  assert.equal(cycleRuntimeOption.cycleRuntimeOptionDirection, 1);
  assert.equal(hasExplicitCommand(cycleRuntimeOption), true);
  assert.equal(shouldStartApp(cycleRuntimeOption), true);
  assert.equal(commandNeedsOverlayRuntime(cycleRuntimeOption), true);

  const sessionAction = parseArgs([
    '--session-action',
    '{"actionId":"cycleRuntimeOption","payload":{"runtimeOptionId":"anki.autoUpdateNewCards","direction":-1}}',
  ]);
  assert.deepEqual(sessionAction.sessionAction, {
    actionId: 'cycleRuntimeOption',
    payload: { runtimeOptionId: 'anki.autoUpdateNewCards', direction: -1 },
  });
  assert.equal(hasExplicitCommand(sessionAction), true);
  assert.equal(shouldStartApp(sessionAction), true);
  assert.equal(commandNeedsOverlayRuntime(sessionAction), true);

  const toggleStatsOverlayRuntime = parseArgs(['--toggle-stats-overlay']);
  assert.equal(commandNeedsOverlayRuntime(toggleStatsOverlayRuntime), true);

  const markWatched = parseArgs(['--mark-watched']);
  assert.equal(markWatched.markWatched, true);
  assert.equal(hasExplicitCommand(markWatched), true);
  assert.equal(shouldStartApp(markWatched), true);
  assert.equal(commandNeedsOverlayRuntime(markWatched), true);

  const dictionary = parseArgs(['--dictionary']);
  assert.equal(dictionary.dictionary, true);
  assert.equal(hasExplicitCommand(dictionary), true);
  assert.equal(shouldStartApp(dictionary), true);
  const dictionaryTarget = parseArgs(['--dictionary', '--dictionary-target', '/tmp/example.mkv']);
  assert.equal(dictionaryTarget.dictionary, true);
  assert.equal(dictionaryTarget.dictionaryTarget, '/tmp/example.mkv');

  const stats = parseArgs([
    '--stats',
    '--stats-response-path',
    '/tmp/subminer-stats-response.json',
    '--stats-cleanup-lifetime',
  ]);
  assert.equal(stats.stats, true);
  assert.equal(stats.statsResponsePath, '/tmp/subminer-stats-response.json');
  assert.equal(stats.statsCleanup, false);
  assert.equal(stats.statsCleanupVocab, false);
  assert.equal(stats.statsCleanupLifetime, true);
  assert.equal(hasExplicitCommand(stats), true);
  assert.equal(shouldStartApp(stats), true);

  const statsBackground = parseArgs(['--stats', '--stats-background']) as typeof stats & {
    statsBackground?: boolean;
    statsStop?: boolean;
  };
  assert.equal(statsBackground.stats, true);
  assert.equal(statsBackground.statsBackground, true);
  assert.equal(statsBackground.statsStop, false);
  assert.equal(hasExplicitCommand(statsBackground), true);
  assert.equal(shouldStartApp(statsBackground), true);

  const statsStop = parseArgs(['--stats', '--stats-stop']) as typeof stats & {
    statsBackground?: boolean;
    statsStop?: boolean;
  };
  assert.equal(statsStop.stats, true);
  assert.equal(statsStop.statsStop, true);
  assert.equal(statsStop.statsBackground, false);
  assert.equal(hasExplicitCommand(statsStop), true);
  assert.equal(shouldStartApp(statsStop), true);

  const statsLifetimeRebuild = parseArgs([
    '--stats',
    '--stats-cleanup',
    '--stats-cleanup-lifetime',
  ]);
  assert.equal(statsLifetimeRebuild.stats, true);
  assert.equal(statsLifetimeRebuild.statsCleanup, true);
  assert.equal(statsLifetimeRebuild.statsCleanupLifetime, true);
  assert.equal(statsLifetimeRebuild.statsCleanupVocab, false);

  const jellyfinLibraries = parseArgs(['--jellyfin-libraries']);
  assert.equal(jellyfinLibraries.jellyfinLibraries, true);
  assert.equal(hasExplicitCommand(jellyfinLibraries), true);
  assert.equal(shouldStartApp(jellyfinLibraries), false);

  const jellyfinSetup = parseArgs(['--jellyfin']);
  assert.equal(jellyfinSetup.jellyfin, true);
  assert.equal(hasExplicitCommand(jellyfinSetup), true);
  assert.equal(shouldStartApp(jellyfinSetup), true);

  const jellyfinPlay = parseArgs(['--jellyfin-play']);
  assert.equal(jellyfinPlay.jellyfinPlay, true);
  assert.equal(hasExplicitCommand(jellyfinPlay), true);
  assert.equal(shouldStartApp(jellyfinPlay), true);

  const jellyfinSubtitles = parseArgs(['--jellyfin-subtitles', '--jellyfin-subtitle-urls']);
  assert.equal(jellyfinSubtitles.jellyfinSubtitles, true);
  assert.equal(jellyfinSubtitles.jellyfinSubtitleUrlsOnly, true);
  assert.equal(hasExplicitCommand(jellyfinSubtitles), true);
  assert.equal(shouldStartApp(jellyfinSubtitles), false);

  const jellyfinRemoteAnnounce = parseArgs(['--jellyfin-remote-announce']);
  assert.equal(jellyfinRemoteAnnounce.jellyfinRemoteAnnounce, true);
  assert.equal(hasExplicitCommand(jellyfinRemoteAnnounce), true);
  assert.equal(shouldStartApp(jellyfinRemoteAnnounce), false);

  const jellyfinPreviewAuth = parseArgs([
    '--jellyfin-preview-auth',
    '--jellyfin-response-path',
    '/tmp/subminer-jf-response.json',
  ]);
  assert.equal(jellyfinPreviewAuth.jellyfinPreviewAuth, true);
  assert.equal(jellyfinPreviewAuth.jellyfinResponsePath, '/tmp/subminer-jf-response.json');
  assert.equal(hasExplicitCommand(jellyfinPreviewAuth), true);
  assert.equal(shouldStartApp(jellyfinPreviewAuth), false);

  const background = parseArgs(['--background']);
  assert.equal(background.background, true);
  assert.equal(hasExplicitCommand(background), true);
  assert.equal(shouldStartApp(background), true);

  const managedPlayback = parseArgs(['--background', '--managed-playback']);
  assert.equal(managedPlayback.background, true);
  assert.equal(managedPlayback.managedPlayback, true);
  assert.equal(hasExplicitCommand(managedPlayback), true);
  assert.equal(shouldStartApp(managedPlayback), true);

  const setup = parseArgs(['--setup']);
  assert.equal((setup as typeof setup & { setup?: boolean }).setup, true);
  assert.equal(hasExplicitCommand(setup), true);
  assert.equal(shouldStartApp(setup), true);
});
